from . import excel_annex, excel_declared, st_eba, xbrl_csv, xbrl_xml
from .base import CoverageEntry, LoadResult, SourceSpec

LOADERS = {
    "excel_annex": excel_annex.load,
    "excel_declared": excel_declared.load,
    "st_eba": st_eba.load,
    "xbrl_csv": xbrl_csv.load,
    "xbrl_xml": xbrl_xml.load,
}


def load_source(spec: SourceSpec) -> LoadResult:
    try:
        fn = LOADERS[spec.loader]
    except KeyError:
        raise ValueError(f"Loader desconocido {spec.loader!r}; disponibles: {sorted(LOADERS)}") from None
    return fn(spec)


__all__ = ["LOADERS", "load_source", "SourceSpec", "LoadResult", "CoverageEntry"]
