"""Juego de prueba sintético con descuadres conocidos.

Todas las plantillas llevan el prefijo ``SYN_``: **no son plantillas EBA reales** ni sus
códigos de fila/columna corresponden a los de la taxonomía. Sirven para probar el
motor, los loaders y la UI de extremo a extremo.

Construcción:
1. Se fijan importes base en EUR, múltiplos de 10.000 para que expresados en millones
   con 2 decimales sean exactos.
2. Se derivan FINREP/COREP (EUR, 2 decimales) y ST (millones, 2 decimales) de esa base,
   de modo que todo cuadra.
3. Se inyectan los descuadres listados en ``EXPECTED`` (también se escriben en
   ``expected_breaks.yaml``).
"""
from __future__ import annotations

from pathlib import Path

import yaml
from openpyxl import Workbook

M = 1_000_000

BASE = {
    "cash": 1_250_000_000,
    "loans": 38_400_000_000,
    "securities": 9_870_000_000,
    "other_assets": 1_480_000_000,
    "capital": 2_000_000_000,
    "reserves": 3_100_000_000,
    "cet1": 4_650_000_000,
    "at1": 1_000_000,
    "t2": 600_000_000,
    "rea_market": 1_200_000_000,
    "rea_oper": 3_300_000_000,
    "exp_corp": 22_000_000_000,
    "rea_corp": 15_400_000_000,
    "exp_retail": 14_500_000_000,
    "rea_retail": 5_800_000_000,
}
CRR3_RETAIL_ADJ = 120_000_000  # reexpresión CRR3 del REA minorista, explicada en ajustes

# regla -> (estado, gravedad esperada)
EXPECTED = {
    "SYN-R01": ("REDONDEO", "—"),  # FINREP total activo +3.000 EUR, dentro del redondeo ST (5.000)
    "SYN-R11": ("KO", "Moderado"),  # ...pero rompe la suma interna FINREP
    "SYN-R02": ("KO", "Crítico"),  # COREP CET1 -2,5 M
    "SYN-R03": ("KO", "Crítico"),  # ST AT1 -50.000 EUR sobre 1 M: crítico por relativo (5 %)
    "SYN-R10": ("KO", "Moderado"),  # ...y rompe la suma interna ST de fondos propios
    "SYN-R07": ("KO", "Grave"),  # ST exposición corporates +250.000 EUR
    "SYN-R14": ("NO_EVALUABLE", "—"),  # REA de mercado ST no reportado
    # Regla propuesta y económicamente mal planteada (el T2 puede hacer que fondos propios
    # > patrimonio): falla, pero su gravedad queda limitada a Leve hasta que alguien la revise.
    "SYN-R15": ("KO", "Leve"),
}


def _annex_sheet(wb: Workbook, name: str, title: str, rows: list[tuple[str, str]], cols: list[str],
                 values: dict[tuple[str, str], float]) -> None:
    ws = wb.create_sheet(name)
    ws["A1"] = title
    ws["A2"] = "SINTÉTICO — no es una plantilla EBA"
    ws.cell(4, 2, "Fila")
    for j, c in enumerate(cols):
        ws.cell(4, 4 + j, c)  # códigos de columna como texto
    for i, (rcode, label) in enumerate(rows):
        ws.cell(5 + i, 1, label)
        ws.cell(5 + i, 2, rcode)  # código de fila como texto
        for j, c in enumerate(cols):
            v = values.get((rcode, c))
            if v is not None:
                ws.cell(5 + i, 4 + j, v)


def generate(out_dir: str | Path) -> dict:
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    b = dict(BASE)
    total_assets = b["cash"] + b["loans"] + b["securities"] + b["other_assets"]
    equity = b["capital"] + b["reserves"]
    own_funds = b["cet1"] + b["at1"] + b["t2"]
    rea_credit = b["rea_corp"] + b["rea_retail"]
    rea_total = rea_credit + b["rea_market"] + b["rea_oper"]

    # --- FINREP (EUR) -------------------------------------------------------
    wb = Workbook()
    wb.remove(wb.active)
    _annex_sheet(wb, "SYN_F01.01", "Balance (activo)", [
        ("0010", "Efectivo"), ("0020", "Préstamos"), ("0030", "Valores"),
        ("0040", "Otros activos"), ("0100", "Total activo"),
    ], ["0010"], {
        ("0010", "0010"): b["cash"], ("0020", "0010"): b["loans"], ("0030", "0010"): b["securities"],
        ("0040", "0010"): b["other_assets"],
        ("0100", "0010"): total_assets + 3_000,  # descuadre inyectado (R01 redondeo, R11 KO)
    })
    wb.save(out / "finrep_annex.xlsx")

    # Export "propio" con layout libre: se lee con excel_declared + layout YAML
    wb = Workbook()
    ws = wb.active
    ws.title = "Patrimonio"
    ws["B2"], ws["B3"] = "Informe interno de patrimonio (SINTÉTICO)", "Importes en EUR"
    ws["B5"], ws["E5"] = "Capital", b["capital"]
    ws["B6"], ws["E6"] = "Reservas", b["reserves"]
    ws["B8"], ws["E8"] = "Total patrimonio", equity
    wb.save(out / "finrep_export_propio.xlsx")
    (out / "finrep_export_propio.layout.yaml").write_text(yaml.safe_dump({
        "hojas": [{"hoja": "Patrimonio", "plantilla": "SYN_F01.03",
                   "celdas": {"E5": "r0010c0010", "E6": "r0020c0010", "E8": "r0100c0010"}}],
    }, allow_unicode=True, sort_keys=False), encoding="utf-8")

    # --- COREP (EUR) --------------------------------------------------------
    wb = Workbook()
    wb.remove(wb.active)
    _annex_sheet(wb, "SYN_C01.00", "Fondos propios", [
        ("0010", "CET1"), ("0020", "AT1"), ("0030", "T2"), ("0100", "Total fondos propios"),
    ], ["0010"], {
        ("0010", "0010"): b["cet1"] - 2_500_000,  # R02
        ("0020", "0010"): b["at1"], ("0030", "0010"): b["t2"], ("0100", "0010"): own_funds,
    })
    _annex_sheet(wb, "SYN_C02.00", "APR", [
        ("0010", "Riesgo de crédito"), ("0020", "Mercado"), ("0030", "Operacional"), ("0100", "Total APR"),
    ], ["0010"], {
        ("0010", "0010"): rea_credit, ("0020", "0010"): b["rea_market"],
        ("0030", "0010"): b["rea_oper"], ("0100", "0010"): rea_total,
    })
    for cls, exp, rea in (("CORP", b["exp_corp"], b["rea_corp"]), ("RETAIL", b["exp_retail"], b["rea_retail"])):
        _annex_sheet(wb, f"SYN_C07.00_{cls}", f"Riesgo de crédito SA — {cls}", [
            ("0010", "Total exposiciones"),
        ], ["0010", "0020"], {("0010", "0010"): exp, ("0010", "0020"): rea})
    wb.save(out / "corep_annex.xlsx")

    # --- ST (millones EUR, 2 decimales) -------------------------------------
    m = lambda x: round(x / M, 2)  # noqa: E731
    wb = Workbook()
    wb.remove(wb.active)
    _annex_sheet(wb, "SYN_ST_BAL", "ST — Balance", [("0010", "Total activo")], ["0010"],
                 {("0010", "0010"): m(total_assets)})
    _annex_sheet(wb, "SYN_ST_CAP", "ST — Capital", [
        ("0010", "CET1"), ("0020", "AT1"), ("0030", "T2"), ("0040", "Total fondos propios"),
        ("0050", "Total APR"), ("0060", "APR mercado"),
    ], ["0010"], {
        ("0010", "0010"): m(b["cet1"]),
        ("0020", "0010"): m(b["at1"] - 50_000),  # R03, R10
        ("0030", "0010"): m(b["t2"]),
        ("0040", "0010"): m(own_funds),
        ("0050", "0010"): m(rea_total + CRR3_RETAIL_ADJ),
        # ("0060", "0010") no se reporta: R14 no evaluable
    })
    _annex_sheet(wb, "SYN_ST_CR", "ST — Riesgo de crédito (punto de partida)", [
        ("0010", "Corporates"), ("0020", "Retail"),
    ], ["0010", "0020"], {
        ("0010", "0010"): m(b["exp_corp"] + 250_000),  # R07
        ("0010", "0020"): m(b["rea_corp"]),
        ("0020", "0010"): m(b["exp_retail"]),
        ("0020", "0020"): m(b["rea_retail"] + CRR3_RETAIL_ADJ),  # R13 conciliado con ajuste
    })
    wb.save(out / "st_annex.xlsx")

    (out / "ajustes.yaml").write_text(yaml.safe_dump({"ajustes": [
        {"regla": "SYN-R13", "importe": CRR3_RETAIL_ADJ,
         "explicacion": "Reexpresión CRR3 del APR minorista en el punto de partida ST (sintético)"},
        {"regla": "SYN-R06", "importe": CRR3_RETAIL_ADJ,
         "explicacion": "Reexpresión CRR3 del APR minorista (sintético)"},
    ]}, allow_unicode=True, sort_keys=False), encoding="utf-8")
    expected = {k: {"estado": s, "gravedad": g} for k, (s, g) in EXPECTED.items()}
    (out / "expected_breaks.yaml").write_text(yaml.safe_dump(expected, allow_unicode=True, sort_keys=True),
                                               encoding="utf-8")
    return expected


if __name__ == "__main__":
    import sys

    generate(sys.argv[1] if len(sys.argv) > 1 else "data/synthetic")
