"""Tek seferlik analiz: hangi kurumlar BKMS'de var, ne kadar GaLaBau ilanı yayınlıyor."""
import json, re, sys, zipfile, collections
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import ihale_radari as m

PAT = {
 "Deutsche Bahn / DB": r"\bdb\b|deutsche bahn|db netz|db infrago|db station|db energie",
 "Autobahn GmbH": r"autobahn gmbh|die autobahn",
 "Straßen.NRW": r"straßen\.?nrw|landesbetrieb straßenbau",
 "LEG": r"\bleg\b", "Vonovia": r"vonovia", "Vivawest": r"vivawest",
 "Stadtwerke": r"stadtwerke", "Technische Betriebe": r"technische betriebe",
 "Wirtschaftsbetriebe": r"wirtschaftsbetrieb", "Wohnungs-/Bau-GmbH/Genossenschaft": r"wohnungs|gemeinnützige|baugenossenschaft|bauverein|gwg|gewobau",
 "Wupperverband / Ruhrverband": r"wupperverband|ruhrverband|aggerverband|emschergenossenschaft|lippeverband",
 "RVR / Ruhr Grün": r"regionalverband ruhr|ruhr grün",
 "Bundeswehr/BImA": r"bundeswehr|bundesanstalt für immobilien|bima",
}
cfg = m.load_config()
files = sorted((m.CACHE).glob("*.zip"))
pre = [k.rstrip("0") for k in cfg["cpv_onekleri"]]
tot, gal = collections.Counter(), collections.Counter()
names = collections.defaultdict(collections.Counter)
n = 0
for rel in m.iter_bkms_releases(files):
    t = m.parse_bkms(rel)
    if not t: continue
    n += 1
    b = t["makam"].lower()
    is_gal = any(c.startswith(p) for c in t["cpv"] for p in pre)
    for k, p in PAT.items():
        if re.search(p, b):
            tot[k] += 1; gal[k] += is_gal
            if is_gal: names[k][t["makam"][:60]] += 1
print(f"Toplam ilan: {n}, gün: {len(files)}")
for k in PAT:
    print(f"{k:40s} tüm ilan: {tot[k]:5d}  GaLaBau: {gal[k]:4d}  {dict(names[k].most_common(4))}")
