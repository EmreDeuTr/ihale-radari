"""Tek seferlik analiz: DB/Autobahn/Straßen.NRW GaLaBau ilanları neden elendi?"""
# v2
import re, sys, datetime as dt
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import ihale_radari as m
P = r"\bdb\b|deutsche bahn|autobahn|straßen\.?nrw|landesbetrieb straßenbau|bundesanstalt für immobilien|wupperverband|ruhrverband|stadtwerke|wirtschaftsbetriebe"
cfg = m.load_config(); plz = m.load_plz_table()
pre = [k.rstrip("0") for k in cfg["cpv_onekleri"]]
now = dt.datetime.now()
for rel in m.iter_bkms_releases(sorted(m.CACHE.glob("*.zip"))):
    t = m.parse_bkms(rel)
    if not t or not re.search(P, t["makam"].lower()): continue
    if not any(c.startswith(p) for c in t["cpv"] for p in pre): continue
    km = None
    p5 = re.sub(r"\D", "", t["plz"])[:5]
    if p5 in plz: km = round(m.haversine(cfg["merkez_lat"], cfg["merkez_lon"], *plz[p5][:2]))
    open_ = t["son_tarih"] and t["son_tarih"] > now
    print(f"{t['tur']:6s} acik={'E' if open_ else 'H'} son={t['son_tarih'] and t['son_tarih']:%d.%m} plz={t['plz'] or '-':6s} nuts={t['nuts'] or '-':6s} km={km} | {t['makam'][:28]} | {t['baslik'][:60]}")
