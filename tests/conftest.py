from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]


@pytest.fixture(scope="session")
def root() -> Path:
    return ROOT


@pytest.fixture(scope="session")
def synthetic_sources(tmp_path_factory) -> Path:
    """Regenera el juego sintético en una carpeta temporal; devuelve su config de fuentes."""
    import yaml

    from cuadre.synthetic import generate

    d = tmp_path_factory.mktemp("syn")
    generate(d)
    src = yaml.safe_load((ROOT / "config/sources/synthetic.yaml").read_text(encoding="utf-8"))
    for s in src["fuentes"]:
        s["fichero"] = str(d / Path(s["fichero"]).name)
        if "layout" in (s.get("opciones") or {}):
            s["opciones"]["layout"] = str(d / Path(s["opciones"]["layout"]).name)
    src["ajustes"] = str(d / "ajustes.yaml")
    cfg = d / "sources.yaml"
    cfg.write_text(yaml.safe_dump(src, allow_unicode=True), encoding="utf-8")
    return cfg


@pytest.fixture(scope="session")
def synthetic_run(synthetic_sources):
    """Ejecuta el cuadre sobre el juego sintético."""
    import yaml

    from cuadre.run import run

    expected = yaml.safe_load((synthetic_sources.parent / "expected_breaks.yaml").read_text(encoding="utf-8"))
    return run(ROOT / "config/exercises/synthetic.yaml", synthetic_sources), expected
