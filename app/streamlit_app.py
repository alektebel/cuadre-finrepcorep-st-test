"""App local de cuadre ST ↔ FINREP/COREP.

Arranque (desde la raíz del repo)::

    streamlit run app/streamlit_app.py
"""
from __future__ import annotations

import datetime as dt
import glob
import sys
from pathlib import Path

import pandas as pd
import streamlit as st

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from cuadre.catalog import rule_to_dict  # noqa: E402
from cuadre.extract import pipeline as pl  # noqa: E402
from cuadre.extract.cli import load_review, save_review  # noqa: E402
from cuadre.report import coverage_frame, facts_frame, operands_frame, results_frame, write_report  # noqa: E402
from cuadre.report.frames import conflicts_frame  # noqa: E402
from cuadre.run import run  # noqa: E402

st.set_page_config(page_title="Cuadre ST ↔ FINREP/COREP", layout="wide")

SEV_COLORS = {
    "Crítico": "#C00000", "Grave": "#FF4D4D", "Moderado": "#FFA64D", "Leve": "#FFE066",
    "REDONDEO": "#BDD7EE", "OK": "#C6EFCE", "NO_EVALUABLE": "#D9D9D9", "KO": "#FF4D4D",
}


def _style_cell(v):
    c = SEV_COLORS.get(v)
    if not c:
        return ""
    fg = "white" if v == "Crítico" else "black"
    return f"background-color: {c}; color: {fg}"


def _eur(x):
    return "" if x is None or pd.isna(x) else f"{x:,.2f}"


def _rel(p: str) -> str:
    try:
        return str(Path(p).resolve().relative_to(ROOT))
    except ValueError:
        return p


# --------------------------------------------------------------------- barra lateral
with st.sidebar:
    st.header("Ejecución")
    exercises = sorted(glob.glob(str(ROOT / "config/exercises/*.yaml")))
    sources = sorted(glob.glob(str(ROOT / "config/sources/*.yaml")))
    ex = st.selectbox("Ejercicio", exercises, format_func=_rel,
                      index=next((i for i, e in enumerate(exercises) if "synthetic" in e), 0))
    src = st.selectbox("Fuentes (entidad)", sources, format_func=_rel,
                       index=next((i for i, s in enumerate(sources) if "synthetic" in s), 0))
    if st.button("Ejecutar cuadre", type="primary", use_container_width=True):
        try:
            st.session_state["run"] = run(ex, src)
            st.session_state.pop("error", None)
        except Exception as e:  # noqa: BLE001 - se muestra al usuario
            st.session_state["error"] = f"{type(e).__name__}: {e}"
    if "error" in st.session_state:
        st.error(st.session_state["error"])

    r = st.session_state.get("run")
    if r:
        key = r.metadata["ejecutado"]
        if st.session_state.get("report_key") != key:
            path = ROOT / "out" / f"informe_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx"
            write_report(r, path)
            st.session_state.update(report_key=key, report_path=path)
        path = st.session_state["report_path"]
        st.download_button("Descargar informe Excel", path.read_bytes(), file_name=path.name,
                           use_container_width=True)
        st.caption(f"También guardado en `{_rel(str(path))}`")

st.title("Cuadre Stress Test ↔ FINREP/COREP")
r = st.session_state.get("run")

tabs = st.tabs(["Resumen", "Checks", "Visor de plantillas", "Carga de datos", "Catálogo de reglas",
                "Revisión de reglas (agentes)"])

if r is None:
    with tabs[0]:
        st.info("Elige ejercicio y fuentes en la barra lateral y pulsa **Ejecutar cuadre**. "
                "El juego `synthetic` funciona sin datos reales.")
else:
    res, ops, facts = results_frame(r), operands_frame(r), facts_frame(r)

    # ----------------------------------------------------------------- Resumen
    with tabs[0]:
        m = r.metadata
        st.caption(f"Ejercicio **{m['ejercicio']}** · referencia {m['fecha_referencia']} · marcos {m['marcos']} · "
                   f"tolerancia {m['tolerancia_eur']} EUR · ejecutado {m['ejecutado']}")
        cols = st.columns(5)
        for col, stt in zip(cols, ["KO", "NO_EVALUABLE", "REDONDEO", "OK"]):
            col.metric(stt, int((res["estado"] == stt).sum()) if len(res) else 0)
        cols[4].metric("Hechos cargados", m["n_hechos"])
        if len(res):
            ko = res[res["estado"] == "KO"]
            sev = ko["gravedad"].value_counts().reindex(list(reversed(r.severity.levels)), fill_value=0)
            c1, c2 = st.columns(2)
            c1.subheader("KO por gravedad")
            c1.dataframe(sev.rename("reglas").to_frame(), use_container_width=True)
            c2.subheader("Por origen de la regla")
            c2.dataframe(pd.crosstab(res["origen"], res["estado"]), use_container_width=True)
        if m["n_conflictos"]:
            st.warning(f"{m['n_conflictos']} datapoints con valores contradictorios entre fuentes (pestaña Carga).")

    # ------------------------------------------------------------------- Checks
    with tabs[1]:
        f1, f2, f3, f4 = st.columns([2, 2, 2, 3])
        est = f1.multiselect("Estado", ["KO", "NO_EVALUABLE", "REDONDEO", "OK"], default=["KO", "NO_EVALUABLE", "REDONDEO"])
        grav = f2.multiselect("Gravedad", list(reversed(r.severity.levels)) + ["—"])
        orig = f3.multiselect("Origen", sorted(res["origen"].unique()) if len(res) else [])
        txt = f4.text_input("Buscar (id, descripción, celda)")
        view = res.copy()
        if est:
            view = view[view["estado"].isin(est)]
        if grav:
            view = view[view["gravedad"].isin(grav)]
        if orig:
            view = view[view["origen"].isin(orig)]
        if txt:
            mask = view.apply(lambda row: txt.lower() in " ".join(map(str, row.values)).lower(), axis=1)
            view = view[mask]
        show = view[["regla", "estado", "gravedad", "descripcion", "lhs_eur", "rhs_eur", "dif_bruta_eur",
                     "ruido_redondeo_eur", "dif_no_explicada_eur", "origen"]].reset_index(drop=True)
        sel = st.dataframe(
            show.style.map(_style_cell, subset=["estado", "gravedad"]).format(
                _eur, subset=["lhs_eur", "rhs_eur", "dif_bruta_eur", "ruido_redondeo_eur", "dif_no_explicada_eur"],
                na_rep=""),
            use_container_width=True, hide_index=True, on_select="rerun", selection_mode="single-row",
        )
        rows = sel.selection.rows if sel and sel.selection else []
        if rows:
            rid = show.iloc[rows[0]]["regla"]
            rr = res[res["regla"] == rid].iloc[0]
            st.divider()
            st.subheader(f"{rid} — {rr['descripcion']}")
            a, b = st.columns([3, 2])
            with a:
                st.markdown(f"**Estado:** {rr['estado']} · **Gravedad:** {rr['gravedad']} "
                            f"(absoluta: {rr['gravedad_abs'] or '—'}, relativa: {rr['gravedad_rel'] or '—'})")
                st.code(f"LHS = {rr['lhs_expr']}\nRHS = {rr['rhs_expr']}", language="text")
                st.markdown(
                    f"LHS = **{_eur(rr['lhs_eur'])}** · RHS = **{_eur(rr['rhs_eur'])}** · "
                    f"diferencia = **{_eur(rr['dif_bruta_eur'])}** · redondeo admitido = {_eur(rr['ruido_redondeo_eur'])} · "
                    f"no explicada = **{_eur(rr['dif_no_explicada_eur'])}** EUR")
                if rr["mensaje"]:
                    st.info(rr["mensaje"])
            with b:
                st.markdown("**Fuente de la regla**")
                st.write({k: rr[k] for k in ("origen", "documento", "apartado", "pagina", "fichero_regla") if rr[k]})
                if rr["cita"]:
                    st.markdown(f"> {rr['cita']}")
            st.markdown("**Operandos y origen de cada dato**")
            od = ops[ops["regla"] == rid][["lado", "coef", "token", "referencia", "valor_fichero", "escala",
                                          "valor_eur", "contribucion_eur", "redondeo_eur", "fichero", "hoja",
                                          "celda", "detalle_origen", "nota"]].copy()
            od["fichero"] = od["fichero"].map(lambda p: _rel(p) if isinstance(p, str) else p)
            st.dataframe(od.style.format(_eur, subset=["valor_eur", "contribucion_eur", "redondeo_eur"]),
                         use_container_width=True, hide_index=True)

    # ------------------------------------------------------- Visor de plantillas
    with tabs[2]:
        if len(facts):
            c1, c2, c3 = st.columns(3)
            fw = c1.selectbox("Marco", sorted(facts["marco"].unique()))
            sub = facts[facts["marco"] == fw]
            tpl = c2.selectbox("Plantilla", sorted(sub["plantilla"].unique()))
            sub = sub[sub["plantilla"] == tpl]
            zs = sorted(sub["z"].dropna().unique())
            z = c3.selectbox("Eje z", zs) if zs else None
            if z is not None:
                sub = sub[sub["z"] == z]
            grid = sub.pivot_table(index="fila", columns="columna", values="valor_eur", aggfunc="first")
            flag = sub.pivot_table(index="fila", columns="columna", values="implicado_en", aggfunc="first")

            def _hl(df):
                return flag.reindex_like(df).map(lambda v: _style_cell(v) if isinstance(v, str) else "")

            st.caption("Valores en EUR. Rojo: celda implicada en un KO; azul: en un descuadre por redondeo.")
            st.dataframe(grid.style.apply(_hl, axis=None).format(_eur), use_container_width=True)
            st.markdown("**Origen de cada celda**")
            src_view = sub[["fila", "columna", "valor_fichero", "escala", "fichero", "hoja", "celda", "implicado_en"]].copy()
            src_view["fichero"] = src_view["fichero"].map(_rel)
            src_view["implicado_en"] = src_view["implicado_en"].fillna("")
            st.dataframe(src_view.sort_values(["fila", "columna"]), use_container_width=True, hide_index=True)

    # --------------------------------------------------------------- Carga
    with tabs[3]:
        cov = coverage_frame(r)
        if len(cov):
            cov["fichero"] = cov["fichero"].map(_rel)
        st.subheader("Cobertura por hoja / tabla")
        st.caption("Revisa aquí si el lector ha detectado bien filas y columnas antes de fiarte de los checks.")
        st.dataframe(cov, use_container_width=True, hide_index=True)
        conf = conflicts_frame(r)
        if len(conf):
            st.subheader("Conflictos entre fuentes")
            st.dataframe(conf, use_container_width=True, hide_index=True)
        st.subheader("Hashes de los ficheros de entrada")
        st.json(r.metadata["hashes"])

    # ------------------------------------------------------------- Catálogo
    with tabs[4]:
        cat = pd.DataFrame([rule_to_dict(x) for x in r.rules])
        st.dataframe(cat, use_container_width=True, hide_index=True)
        concepts = pd.DataFrame([{"concepto": c.name, "ref": str(c.ref), "etiqueta": c.label,
                                  "verificado": c.verified, "fichero": _rel(c.origin_file or "")}
                                 for c in r.concepts.concepts.values()])
        st.subheader("Mapa de conceptos del ejercicio")
        st.dataframe(concepts, use_container_width=True, hide_index=True)

# ------------------------------------------------------------ Revisión de reglas
with tabs[5]:
    work = ROOT / "work"
    review = load_review(work)
    st.caption("Reglas propuestas por el flujo de extracción (agentes + fórmulas de plantilla + propuestas por "
               "etiqueta). Genera `work/review.json` con `cuadre-extract review`.")
    if not review:
        st.info("No hay `work/review.json` todavía.")
    else:
        df = pd.DataFrame([{
            "id": it["rule"].get("id"), "origen": it["rule"].get("origen"), "auto_ok": it["auto_ok"],
            "veredicto": (it.get("verdict") or {}).get("verdict"), "decision": it["decision"],
            "descripcion": it["rule"].get("descripcion"),
        } for it in review])
        c1, c2 = st.columns(2)
        fdec = c1.multiselect("Decisión", ["pendiente", "aceptada", "rechazada"], default=["pendiente"])
        forg = c2.multiselect("Origen", sorted(df["origen"].dropna().unique()))
        v = df[df["decision"].isin(fdec)] if fdec else df
        if forg:
            v = v[v["origen"].isin(forg)]
        sel = st.dataframe(v, use_container_width=True, hide_index=True, on_select="rerun",
                           selection_mode="single-row")
        rows = sel.selection.rows if sel and sel.selection else []
        if rows:
            rid = v.iloc[rows[0]]["id"]
            idx = next(i for i, it in enumerate(review) if it["rule"].get("id") == rid)
            it = review[idx]
            st.code(pl.yaml.safe_dump(it["rule"], allow_unicode=True, sort_keys=False), language="yaml")
            st.markdown("**Checks automáticos**")
            st.dataframe(pd.DataFrame(it["checks"]), use_container_width=True, hide_index=True)
            if it.get("verdict"):
                st.markdown(f"**Verificador:** {it['verdict']['verdict']} — {it['verdict']['reason']}")
            with st.form(f"dec-{rid}"):
                dec = st.radio("Decisión", ["pendiente", "aceptada", "rechazada"],
                               index=["pendiente", "aceptada", "rechazada"].index(it["decision"]), horizontal=True)
                com = st.text_area("Comentario", it.get("comment", ""))
                if st.form_submit_button("Guardar"):
                    if dec == "aceptada" and not it["auto_ok"]:
                        st.error("No se puede aceptar una regla con checks automáticos fallidos.")
                    else:
                        it.update(decision=dec, comment=com, reviewed_at=dt.datetime.now().isoformat(timespec="seconds"))
                        save_review(work, review)
                        st.success("Guardado en work/review.json")
        if st.button("Exportar aceptadas a rules/st2025/"):
            try:
                counts = pl.export_accepted(review, ROOT / "rules/st2025", work / "export/conceptos_nuevos.yaml")
                st.success(f"Exportadas: {counts or 'ninguna'}")
            except ValueError as e:
                st.error(str(e))
