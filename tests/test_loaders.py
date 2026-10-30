import json
import zipfile

from openpyxl import Workbook

from cuadre.loaders import SourceSpec, load_source
from cuadre.model import DatapointRef


def test_annex_detector_ignores_numeric_lookalikes(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "F 01.01"
    ws["C3"], ws["D3"] = "0010", "0020"
    ws["A4"], ws["B4"], ws["C4"], ws["D4"] = "Efectivo", "0010", 10, 20  # 10 es un importe, no un código
    ws["A5"], ws["B5"], ws["C5"] = "Préstamos", "0020", 30
    p = tmp_path / "f.xlsx"
    wb.save(p)
    res = load_source(SourceSpec("FINREP", "excel_annex", p, scale=1000, decimals=0))
    vals = {str(f.ref): (f.value, f.source.cell, f.rounding_eur) for f in res.facts}
    assert vals == {
        "FINREP:F 01.01!r0010c0010": (10_000, "C4", 500.0),
        "FINREP:F 01.01!r0010c0020": (20_000, "D4", 500.0),
        "FINREP:F 01.01!r0020c0010": (30_000, "C5", 500.0),
    }


def test_xbrl_csv_with_datapoint_map(tmp_path):
    pkg = tmp_path / "pkg.zip"
    with zipfile.ZipFile(pkg, "w") as zf:
        zf.writestr("rp/META-INF/reportPackage.json", json.dumps({"documentInfo": {}}))
        zf.writestr("rp/reports/parameters.csv", "name,value\nentityID,LEI\ndecimalsMonetary,-3\n")
        zf.writestr("rp/reports/c_01.00.csv", "datapoint,factValue\ndp1,1000\ndp2,abc\ndp9,5\n")
        zf.writestr("rp/reports/c_07.00.a.csv", "datapoint,factValue,EXC\ndp3,7,CORP\n")
    dmap = tmp_path / "map.csv"
    dmap.write_text("datapoint,template,row,col,z\ndp1,C 01.00,0010,0010,\ndp2,C 01.00,0020,0010,\n"
                    "dp3,C 07.00,0010,0200,\n", encoding="utf-8")
    res = load_source(SourceSpec("COREP", "xbrl_csv", pkg, options={"datapoint_map": str(dmap)}))
    facts = {f.ref: f for f in res.facts}
    f = facts[DatapointRef.parse("COREP:C 01.00!r0010c0010")]
    assert f.value == 1000 and f.rounding_eur == 500.0
    assert DatapointRef.parse("COREP:C 07.00!r0010c0200@EXC=CORP") in facts
    warns = " ".join(w for c in res.coverage for w in c.warnings)
    assert "sin mapeo" in warns and "no numéricos" in warns


def test_xbrl_xml_with_fact_map(tmp_path):
    inst = tmp_path / "i.xbrl"
    inst.write_text("""<?xml version="1.0"?>
<xbrli:xbrl xmlns:xbrli="http://www.xbrl.org/2003/instance" xmlns:xbrldi="http://xbrl.org/2006/xbrldi"
            xmlns:eba_met="http://www.eba.europa.eu/xbrl/crr/dict/met" xmlns:eba_dim="d" xmlns:eba_BA="m">
  <xbrli:context id="c1"><xbrli:entity><xbrli:identifier scheme="s">X</xbrli:identifier>
    <xbrli:segment><xbrldi:explicitMember dimension="eba_dim:BAS">eba_BA:x1</xbrldi:explicitMember></xbrli:segment>
  </xbrli:entity><xbrli:period><xbrli:instant>2024-12-31</xbrli:instant></xbrli:period></xbrli:context>
  <xbrli:unit id="EUR"><xbrli:measure>iso4217:EUR</xbrli:measure></xbrli:unit>
  <eba_met:mi53 contextRef="c1" unitRef="EUR" decimals="-3">123000</eba_met:mi53>
</xbrli:xbrl>""", encoding="utf-8")
    fmap = tmp_path / "fmap.csv"
    fmap.write_text("key,template,row,col,z\nmi53|BAS=x1,F 01.01,0380,0010,\n", encoding="utf-8")
    res = load_source(SourceSpec("FINREP", "xbrl_xml", inst, options={"fact_map": str(fmap)}))
    assert [(str(f.ref), f.value, f.rounding_eur) for f in res.facts] == [("FINREP:F 01.01!r0380c0010", 123000.0, 500.0)]
