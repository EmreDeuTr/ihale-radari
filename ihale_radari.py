# -*- coding: utf-8 -*-
"""
IHALE RADARI - Almanya kamu ihalelerini ücretsiz resmi kaynaklardan toplar,
bölgeye (km), CPV koduna ve anahtar kelimeye göre filtreler, Excel'e yazar.

Kaynaklar (hepsi ücretsiz, kayıt gerekmez):
  1) Bekanntmachungsservice / Datenservice Öffentlicher Einkauf (oeffentlichevergabe.de)
     -> Bund, Länder ve Kommunen ilanları; günlük OCDS açık veri paketi (CC0)
  2) TED (ted.europa.eu) Search API v3 -> AB eşiğini aşan ihaleler (yedek / çapraz kontrol)
  3) GeoNames posta kodu tablosu -> PLZ'den mesafe hesabı (ilk çalıştırmada bir kez indirilir)

Kullanım:
  python ihale_radari.py            -> config.json'daki ayarlarla çalışır
  python ihale_radari.py --gun 45   -> son 45 günü tara
  python ihale_radari.py --ac       -> bitince Excel'i aç
"""
from __future__ import annotations

import argparse
import datetime as dt
import io
import json
import math
import os
import re
import smtplib
import sys
import time
import traceback
import zipfile
from email.message import EmailMessage
from pathlib import Path

try:
    import requests
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
except ImportError:
    print("Eksik paket. Önce 'kurulum.bat' çalıştırın veya:  pip install requests openpyxl")
    sys.exit(1)

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:  # noqa: BLE001
        pass

BASE = Path(__file__).resolve().parent
VERI = BASE / "veri"
CACHE = VERI / "bkms"
CONFIG_PATH = BASE / "config.json"
STATE_PATH = VERI / "durum.json"
LOG_PATH = VERI / "log.txt"

BKMS_EXPORT = "https://oeffentlichevergabe.de/api/notice-exports?pubDay={day}&format=ocds.zip"
BKMS_DETAIL = "https://oeffentlichevergabe.de/ui/de/search/details?noticeId={uuid}"
TED_SEARCH = "https://api.ted.europa.eu/v3/notices/search"
TED_DETAIL = "https://ted.europa.eu/de/notice/-/detail/{pub}"
GEONAMES_DE = "https://download.geonames.org/export/zip/DE.zip"
UA = {"User-Agent": "IhaleRadari/1.0 (Tolgay GaLaBau; kendi kullanim)"}

# NRW Kreis/kreisfreie Stadt (NUTS-3) yaklaşık merkezleri -> PLZ yoksa yedek mesafe
NUTS3 = {
    "DEA11": ("Düsseldorf", 51.227, 6.774), "DEA12": ("Duisburg", 51.434, 6.762),
    "DEA13": ("Essen", 51.456, 7.012), "DEA14": ("Krefeld", 51.339, 6.586),
    "DEA15": ("Mönchengladbach", 51.185, 6.442), "DEA16": ("Mülheim a.d. Ruhr", 51.418, 6.884),
    "DEA17": ("Oberhausen", 51.496, 6.863), "DEA18": ("Remscheid", 51.179, 7.194),
    "DEA19": ("Solingen", 51.171, 7.083), "DEA1A": ("Wuppertal", 51.256, 7.150),
    "DEA1B": ("Kreis Kleve", 51.700, 6.200), "DEA1C": ("Kreis Mettmann", 51.250, 6.970),
    "DEA1D": ("Rhein-Kreis Neuss", 51.150, 6.650), "DEA1E": ("Kreis Viersen", 51.260, 6.390),
    "DEA1F": ("Kreis Wesel", 51.620, 6.600), "DEA22": ("Bonn", 50.737, 7.098),
    "DEA23": ("Köln", 50.938, 6.960), "DEA24": ("Leverkusen", 51.046, 7.019),
    "DEA26": ("Kreis Düren", 50.800, 6.480), "DEA27": ("Rhein-Erft-Kreis", 50.900, 6.750),
    "DEA28": ("Kreis Euskirchen", 50.580, 6.750), "DEA29": ("Kreis Heinsberg", 51.060, 6.100),
    "DEA2A": ("Oberbergischer Kreis", 51.030, 7.550), "DEA2B": ("Rheinisch-Bergischer Kreis", 51.000, 7.200),
    "DEA2C": ("Rhein-Sieg-Kreis", 50.780, 7.300), "DEA2D": ("Städteregion Aachen", 50.750, 6.150),
    "DEA31": ("Bottrop", 51.520, 6.930), "DEA32": ("Gelsenkirchen", 51.510, 7.090),
    "DEA33": ("Münster", 51.960, 7.630), "DEA34": ("Kreis Borken", 51.900, 6.850),
    "DEA35": ("Kreis Coesfeld", 51.900, 7.300), "DEA36": ("Kreis Recklinghausen", 51.620, 7.200),
    "DEA37": ("Kreis Steinfurt", 52.200, 7.550), "DEA38": ("Kreis Warendorf", 51.900, 8.000),
    "DEA41": ("Bielefeld", 52.020, 8.530), "DEA42": ("Kreis Gütersloh", 51.900, 8.380),
    "DEA43": ("Kreis Herford", 52.120, 8.670), "DEA44": ("Kreis Höxter", 51.700, 9.300),
    "DEA45": ("Kreis Lippe", 51.950, 8.850), "DEA46": ("Kreis Minden-Lübbecke", 52.300, 8.750),
    "DEA47": ("Kreis Paderborn", 51.700, 8.750), "DEA51": ("Bochum", 51.480, 7.220),
    "DEA52": ("Dortmund", 51.510, 7.470), "DEA53": ("Hagen", 51.360, 7.470),
    "DEA54": ("Hamm", 51.680, 7.820), "DEA55": ("Herne", 51.540, 7.220),
    "DEA56": ("Ennepe-Ruhr-Kreis", 51.330, 7.300), "DEA57": ("Hochsauerlandkreis", 51.300, 8.300),
    "DEA58": ("Märkischer Kreis", 51.250, 7.700), "DEA59": ("Kreis Olpe", 51.050, 7.950),
    "DEA5A": ("Kreis Siegen-Wittgenstein", 50.950, 8.150), "DEA5B": ("Kreis Soest", 51.570, 8.100),
    "DEA5C": ("Kreis Unna", 51.550, 7.650),
}

DEFAULT_CONFIG = {
    "_aciklama": "Bu dosyayi Not Defteri ile duzenleyebilirsiniz. Degisiklikten sonra kaydedip tekrar calistirin.",
    "merkez_plz": "42477",
    "merkez_ad": "Radevormwald",
    "merkez_lat": 51.2017,
    "merkez_lon": 7.3603,
    "yaricap_km": 80,
    "tarama_gun": 30,
    "cache_tut_gun": 60,
    "kaynak_bkms": True,
    "kaynak_ted": True,
    "son_tarihi_gecenleri_gizle": True,
    "min_puan": 30,
    "cpv_onekleri": {
        "45112700": "Landschaftsgärtnerische Arbeiten (Grünflächen, Parks, Friedhöfe, Spielplätze...)",
        "45236": "Flächenarbeiten Sport- und Freizeitanlagen",
        "45342000": "Zaunbau",
        "45233222": "Pflasterarbeiten",
        "773": "Gartenbau-Dienstleistungen (Grünpflege, Mähen, Baum-/Heckenschnitt)",
        "77211": "Baumfällung / Baumpflege",
        "90620000": "Winterdienst - Schneeräumung",
        "90630000": "Winterdienst - Eisbeseitigung",
        "03451": "Pflanzen (Lieferung)",
        "03452": "Bäume (Lieferung)",
        "375352": "Spielplatzgeräte"
    },
    "anahtar_kelimeler": [
        "grünpflege", "grünflächen", "grünanlage", "landschaftsbau", "galabau", "garten- und landschaft",
        "landschaftsgärtner", "mäharbeit", "mahd", "mähen", "rasen", "baumpflege", "baumfäll", "baumschnitt",
        "gehölz", "heckenschnitt", "pflanzarbeit", "pflanzung", "stauden", "spielplatz", "friedhof",
        "außenanlage", "aussenanlage", "freianlage", "freiflächen", "schulhof", "begleitgrün", "böschung",
        "wildkraut", "unkraut", "winterdienst", "schneeräum", "streudienst", "pflasterarbeit", "zaunbau",
        "sportplatz", "kunstrasen", "dachbegrünung", "renaturierung", "entsiegelung", "parkanlage",
        "grünunterhaltung", "unterhaltungspflege", "fertigstellungspflege", "entwicklungspflege"
    ],
    "haric_kelimeler": [
        "reinigung von gebäuden", "glasreinigung", "unterhaltsreinigung", "catering", "software"
    ],
    "email": {
        "aktif": False,
        "smtp_sunucu": "smtp.gmail.com",
        "smtp_port": 587,
        "kullanici": "ornek@gmail.com",
        "sifre": "UYGULAMA-SIFRESI",
        "kime": ["ornek@gmail.com"],
        "sadece_yeni_varsa": True
    }
}

DURUM_LISTESI = ["", "İncele", "Teklif Ver", "Teklif Verildi", "Pas", "Kazandık", "Kaybettik"]


# ----------------------------------------------------------------- yardımcılar
def log(msg: str) -> None:
    line = f"{dt.datetime.now():%Y-%m-%d %H:%M:%S}  {msg}"
    print(line)
    try:
        VERI.mkdir(parents=True, exist_ok=True)
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except OSError:
        pass


def load_config() -> dict:
    if not CONFIG_PATH.exists():
        CONFIG_PATH.write_text(json.dumps(DEFAULT_CONFIG, ensure_ascii=False, indent=2), encoding="utf-8")
        return json.loads(json.dumps(DEFAULT_CONFIG))
    cfg = json.loads(CONFIG_PATH.read_text(encoding="utf-8-sig"))
    merged = json.loads(json.dumps(DEFAULT_CONFIG))
    for k, v in cfg.items():
        if isinstance(v, dict) and isinstance(merged.get(k), dict) and k != "cpv_onekleri":
            merged[k].update(v)
        else:
            merged[k] = v
    return merged


def txt(x) -> str:
    if x is None:
        return ""
    if isinstance(x, dict):  # çok dilli alan
        for lang in ("deu", "DEU", "de", "eng", "ENG", "en"):
            if x.get(lang):
                return txt(x[lang])
        return txt(next(iter(x.values()), ""))
    if isinstance(x, list):
        return "; ".join(dict.fromkeys(t for t in (txt(i) for i in x) if t))
    return re.sub(r"\s+", " ", str(x)).strip()


def parse_date(s) -> dt.datetime | None:
    if not s:
        return None
    s = str(s).strip()
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2}))?", s)
    if not m:
        return None
    y, mo, d, h, mi = m.groups()
    try:
        return dt.datetime(int(y), int(mo), int(d), int(h or 0), int(mi or 0))
    except ValueError:
        return None


def haversine(lat1, lon1, lat2, lon2) -> float:
    r = 6371.0
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp, dl = p2 - p1, math.radians(lon2 - lon1)
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * r * math.asin(math.sqrt(a))


def http_get(url: str, timeout=60, tries=3):
    last = None
    for i in range(tries):
        try:
            r = requests.get(url, headers=UA, timeout=timeout)
            if r.status_code in (404, 204):
                return r
            if r.status_code >= 500 or r.status_code == 429:
                last = f"HTTP {r.status_code}"
                time.sleep(3 * (i + 1))
                continue
            return r
        except requests.RequestException as e:
            last = str(e)
            time.sleep(3 * (i + 1))
    raise RuntimeError(f"{url} alınamadı: {last}")


# ----------------------------------------------------------------- posta kodu -> koordinat
def load_plz_table() -> dict:
    path = VERI / "DE_plz.zip"
    if not path.exists():
        try:
            log("Posta kodu tablosu indiriliyor (tek seferlik, GeoNames)...")
            r = http_get(GEONAMES_DE, timeout=120)
            if r.status_code == 200 and r.content[:2] == b"PK":
                path.write_bytes(r.content)
        except Exception as e:  # noqa: BLE001
            log(f"UYARI: Posta kodu tablosu indirilemedi ({e}). Mesafe NUTS bölgesinden tahmin edilecek.")
            return {}
    acc: dict[str, list] = {}
    try:
        with zipfile.ZipFile(path) as z:
            name = next(n for n in z.namelist() if n.upper().startswith("DE") and n.endswith(".txt"))
            for line in io.TextIOWrapper(z.open(name), encoding="utf-8"):
                p = line.rstrip("\n").split("\t")
                if len(p) < 11:
                    continue
                try:
                    acc.setdefault(p[1], []).append((float(p[9]), float(p[10]), p[2]))
                except ValueError:
                    continue
    except Exception as e:  # noqa: BLE001
        log(f"UYARI: Posta kodu tablosu okunamadı: {e}")
        return {}
    return {k: (sum(a for a, _, _ in v) / len(v), sum(b for _, b, _ in v) / len(v), v[0][2]) for k, v in acc.items()}


# ----------------------------------------------------------------- Kaynak 1: Bekanntmachungsservice
def bkms_download_days(days: int, keep: int) -> list[Path]:
    CACHE.mkdir(parents=True, exist_ok=True)
    today = dt.date.today()
    files = []
    for i in range(1, days + 1):  # bugünün paketi ertesi gün çıkar
        day = today - dt.timedelta(days=i)
        f = CACHE / f"{day.isoformat()}.zip"
        if f.exists() and f.stat().st_size > 0:
            files.append(f)
            continue
        try:
            r = http_get(BKMS_EXPORT.format(day=day.isoformat()), timeout=180)
        except RuntimeError as e:
            log(f"  {day}: indirilemedi ({e})")
            continue
        if r.status_code == 200 and r.content[:2] == b"PK":
            f.write_bytes(r.content)
            files.append(f)
            log(f"  {day}: indirildi ({len(r.content) / 1e6:.1f} MB)")
        elif r.status_code == 200 and len(r.content) < 200:
            log(f"  {day}: o gün ilan yok")
        else:
            log(f"  {day}: veri yok (HTTP {r.status_code})")
    # eski önbelleği temizle
    limit = today - dt.timedelta(days=max(keep, days) + 1)
    for f in CACHE.glob("*.zip"):
        try:
            if dt.date.fromisoformat(f.stem) < limit:
                f.unlink()
        except ValueError:
            pass
    return sorted(files)


def iter_bkms_releases(files: list[Path]):
    for f in files:
        try:
            with zipfile.ZipFile(f) as z:
                for n in z.namelist():
                    if not n.lower().endswith(".json"):
                        continue
                    try:
                        data = json.loads(z.read(n))
                    except ValueError:
                        continue
                    if isinstance(data, list):
                        rels = data
                    else:
                        rels = data.get("releases") or ([data] if "tender" in data or "ocid" in data else [])
                    for rel in rels:
                        if isinstance(rel, dict):
                            yield rel
        except zipfile.BadZipFile:
            log(f"UYARI: bozuk paket silindi: {f.name}")
            f.unlink(missing_ok=True)


def _walk(o, key):
    """İç içe yapıda belirli anahtarın tüm değerlerini döndür."""
    if isinstance(o, dict):
        for k, v in o.items():
            if k == key:
                yield v
            yield from _walk(v, key)
    elif isinstance(o, list):
        for i in o:
            yield from _walk(i, key)


def parse_bkms(rel: dict) -> dict | None:
    tender = rel.get("tender") or {}
    rid = str(rel.get("id") or rel.get("ocid") or "")
    m = re.search(r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}", rid, re.I)
    uuid = m.group(0) if m else rid
    vm = re.search(r"-(\d{2})$", rid)
    version = int(vm.group(1)) if vm else 0
    tags = [t.lower() for t in (rel.get("tag") or [])]
    status = str(tender.get("status") or "").lower()

    if any("award" in t or "contract" in t for t in tags) or rel.get("awards"):
        kind = "sonuc"
    elif any("planning" in t for t in tags) or status == "planning":
        kind = "on"
    else:
        kind = "acik"
    if status in ("cancelled", "withdrawn") or any("cancel" in t for t in tags):
        kind = "iptal"

    parties = rel.get("parties") or []
    buyer = next((p for p in parties if "buyer" in (p.get("roles") or [])), None) or \
        next((p for p in parties if "procuringEntity" in (p.get("roles") or [])), {}) or {}
    baddr = buyer.get("address") or {}

    # yer: önce ifa yeri (deliveryAddresses), yoksa alıcı adresi
    places = []
    for it in tender.get("items") or []:
        for a in (it.get("deliveryAddresses") or []) + ([it["deliveryAddress"]] if it.get("deliveryAddress") else []):
            if isinstance(a, dict):
                places.append(a)
    for a in _walk(tender.get("lots") or [], "deliveryAddresses"):
        places.extend([x for x in (a or []) if isinstance(x, dict)])
    place = next((a for a in places if a.get("postalCode")), None) or (places[0] if places else {})

    cpv = []
    for it in tender.get("items") or []:
        for c in [it.get("classification") or {}] + (it.get("additionalClassifications") or []):
            if str(c.get("scheme", "")).upper() == "CPV" and c.get("id"):
                cpv.append((str(c["id"])[:8], txt(c.get("description"))))
    if tender.get("classification", {}).get("id"):
        cpv.insert(0, (str(tender["classification"]["id"])[:8], txt(tender["classification"].get("description"))))
    cpv = list(dict.fromkeys(cpv))

    deadlines = [parse_date(v.get("endDate")) for v in _walk(tender, "tenderPeriod") if isinstance(v, dict)]
    deadlines = [d for d in deadlines if d]
    deadline = min(deadlines) if deadlines else None

    lots = tender.get("lots") or []
    lot_titles = [txt(l.get("title")) for l in lots if txt(l.get("title")) and txt(l.get("title")) != txt(tender.get("title"))]
    descr = " ".join([txt(tender.get("description"))] + [txt(l.get("description")) for l in lots])

    value = None
    v = tender.get("value") or {}
    if v.get("amount"):
        value = float(v["amount"])
    else:
        s = sum(float((l.get("value") or {}).get("amount") or 0) for l in lots)
        value = s or None

    docs = []
    for doclist in _walk(tender, "documents"):
        for doc in doclist if isinstance(doclist, list) else []:
            if isinstance(doc, dict) and doc.get("url"):
                docs.append(doc["url"])

    winners, award_val = [], 0.0
    for aw in rel.get("awards") or []:
        if str(aw.get("status", "")).lower() in ("unsuccessful", "cancelled"):
            continue
        for s in aw.get("suppliers") or []:
            if s.get("name"):
                winners.append(txt(s["name"]))
        try:
            award_val += float((aw.get("value") or {}).get("amount") or 0)
        except (TypeError, ValueError):
            pass
    if not winners:
        winners = [txt(p.get("name")) for p in parties if "supplier" in (p.get("roles") or []) and p.get("name")]

    return {
        "id": f"BKMS:{uuid}", "version": version, "kaynak": "Bekanntmachungsservice", "tur": kind,
        "baslik": txt(tender.get("title")) or txt(lot_titles[:1]),
        "lotlar": "; ".join(lot_titles[:10]),
        "aciklama": descr[:3000],
        "makam": txt(buyer.get("name")),
        "yer": txt(place.get("locality")) or txt(baddr.get("locality")),
        "plz": txt(place.get("postalCode")) or txt(baddr.get("postalCode")),
        "nuts": (txt(place.get("region")) or txt(baddr.get("region")))[:5].upper(),
        "son_tarih": deadline,
        "yayin": parse_date(rel.get("date")),
        "cpv": [c for c, _ in cpv],
        "cpv_text": [t for _, t in cpv],
        "kategori": {"works": "Bau", "services": "Dienstleistung", "goods": "Lieferung"}.get(
            str(tender.get("mainProcurementCategory") or "").lower(), txt(tender.get("mainProcurementCategory"))),
        "prosedur": txt(tender.get("procurementMethodDetails")) or txt(tender.get("procurementMethod")),
        "deger": value,
        "link": BKMS_DETAIL.format(uuid=uuid),
        "belge": docs[0] if docs else "",
        "kazanan": "; ".join(dict.fromkeys(winners)),
        "kazanan_deger": award_val or None,
    }


def collect_bkms(cfg: dict, days: int) -> list[dict]:
    log(f"Bekanntmachungsservice: son {days} günün paketleri kontrol ediliyor...")
    files = bkms_download_days(days, cfg.get("cache_tut_gun", 60))
    latest: dict[str, dict] = {}
    n = 0
    for rel in iter_bkms_releases(files):
        n += 1
        try:
            t = parse_bkms(rel)
        except Exception:  # noqa: BLE001
            continue
        if not t:
            continue
        key = t["id"] + ("#S" if t["tur"] == "sonuc" else "")
        old = latest.get(key)
        if old is None or (t["version"], t["yayin"] or dt.datetime.min) >= (old["version"], old["yayin"] or dt.datetime.min):
            latest[key] = t
    # iptal edilen ilanlar açık listeden düşsün
    for k, t in list(latest.items()):
        if t["tur"] == "iptal":
            latest.pop(k, None)
            latest.pop(k.replace("#S", ""), None)
    log(f"Bekanntmachungsservice: {n} kayıt okundu, {len(latest)} benzersiz ilan.")
    return list(latest.values())


# ----------------------------------------------------------------- Kaynak 2: TED
TED_EK_CPV = ["77310000", "77311000", "77312000", "77313000", "77314000", "77315000", "77320000", "77340000",
              "77341000", "77342000", "77211400", "77211500", "45112710", "45112711", "45112712", "45112713",
              "45112714", "45112720", "45112723", "45112730", "45236100", "45236110", "45236119"]
TED_FIELDS_FULL = ["publication-number", "publication-date", "notice-title", "notice-type", "buyer-name",
                   "buyer-city", "place-of-performance", "classification-cpv",
                   "deadline-receipt-tender-date-lot", "description-lot", "links"]
TED_FIELDS_MIN = ["publication-number", "publication-date", "notice-title", "buyer-name",
                  "place-of-performance", "classification-cpv", "links"]


def collect_ted(cfg: dict, days: int) -> list[dict]:
    since = (dt.date.today() - dt.timedelta(days=days)).strftime("%Y%m%d")
    cpvs = []
    for pre in cfg["cpv_onekleri"]:
        cpvs.append(pre.ljust(8, "0") if len(pre) < 8 else pre)
    cpvs += [c for c in TED_EK_CPV if any(c.startswith(p.rstrip("0") or p) for p in cfg["cpv_onekleri"])]
    query = (f"buyer-country=DEU AND classification-cpv IN ({' '.join(dict.fromkeys(cpvs))}) "
             f"AND publication-date>={since}")
    out, fields = [], TED_FIELDS_FULL
    log("TED (AB): sorgulanıyor...")
    for page in range(1, 21):
        body = {"query": query, "fields": fields, "page": page, "limit": 100,
                "scope": "ALL", "paginationMode": "PAGE_NUMBER"}
        try:
            r = requests.post(TED_SEARCH, json=body, headers={**UA, "Accept": "application/json"}, timeout=60)
        except requests.RequestException as e:
            log(f"UYARI: TED erişilemedi: {e}")
            break
        if r.status_code == 400 and fields is TED_FIELDS_FULL:
            log("TED: alan listesi sadeleştirilip tekrar deneniyor...")
            fields = TED_FIELDS_MIN
            body["fields"] = fields
            r = requests.post(TED_SEARCH, json=body, headers={**UA, "Accept": "application/json"}, timeout=60)
        if r.status_code != 200:
            log(f"UYARI: TED yanıtı HTTP {r.status_code}: {r.text[:200]}")
            break
        data = r.json()
        notices = data.get("notices") or data.get("results") or []
        for no in notices:
            out.append(parse_ted(no))
        total = data.get("totalNoticeCount") or 0
        if not notices or page * 100 >= total:
            break
    log(f"TED: {len(out)} ilan.")
    return out


def parse_ted(no: dict) -> dict:
    pub = txt(no.get("publication-number"))
    ntype = txt(no.get("notice-type")).lower()
    kind = "sonuc" if ntype.startswith(("can", "veat")) or "result" in ntype else \
        "on" if ntype.startswith("pin") and "cfc" not in ntype else "acik"
    links = no.get("links") or {}
    link = ""
    for fmt in ("html", "htmlDirect"):
        l = links.get(fmt) or {}
        link = l.get("DEU") or l.get("deu") or l.get("ENG") or next(iter(l.values()), "") if isinstance(l, dict) else ""
        if link:
            break
    pop = no.get("place-of-performance") or []
    pop = pop if isinstance(pop, list) else [pop]
    nuts = next((p for p in pop if isinstance(p, str) and len(p) >= 5 and p.startswith("DE")), "")
    dl = no.get("deadline-receipt-tender-date-lot") or []
    dls = [parse_date(d) for d in (dl if isinstance(dl, list) else [dl])]
    dls = [d for d in dls if d]
    cpv = no.get("classification-cpv") or []
    return {
        "id": f"TED:{pub}", "version": 0, "kaynak": "TED (AB)", "tur": kind,
        "baslik": re.sub(r"^Deutschland\s+[–-]\s+.+?\s+[–-]\s+", "", txt(no.get("notice-title"))), "lotlar": "",
        "aciklama": txt(no.get("description-lot"))[:3000],
        "makam": txt(no.get("buyer-name")), "yer": txt(no.get("buyer-city")), "plz": "",
        "nuts": nuts[:5].upper(), "son_tarih": min(dls) if dls else None,
        "yayin": parse_date(txt(no.get("publication-date"))),
        "cpv": [str(c)[:8] for c in (cpv if isinstance(cpv, list) else [cpv])], "cpv_text": [],
        "kategori": "", "prosedur": "", "deger": None,
        "link": link or TED_DETAIL.format(pub=pub), "belge": "", "kazanan": "", "kazanan_deger": None,
    }


# ----------------------------------------------------------------- filtre + puan
def norm(s: str) -> str:
    return re.sub(r"[^a-z0-9äöüß]", "", (s or "").lower())


def score_all(items: list[dict], cfg: dict, plz_tab: dict) -> list[dict]:
    home = (cfg["merkez_lat"], cfg["merkez_lon"])
    if cfg.get("merkez_plz") in plz_tab:
        home = plz_tab[cfg["merkez_plz"]][:2]
    radius = float(cfg["yaricap_km"])
    prefixes = [k.rstrip("0") if len(k.rstrip("0")) >= 3 else k[:3] for k in cfg["cpv_onekleri"].keys()]
    kws = [k.lower() for k in cfg["anahtar_kelimeler"]]
    exc = [k.lower() for k in cfg.get("haric_kelimeler", [])]
    out = []
    for t in items:
        # mesafe
        km, km_kaynak = None, ""
        plz = re.sub(r"\D", "", t["plz"])[:5]
        if plz in plz_tab:
            km = haversine(*home, *plz_tab[plz][:2])
            km_kaynak = "PLZ"
            if not t["yer"]:
                t["yer"] = plz_tab[plz][2]
        elif t["nuts"] in NUTS3:
            _, la, lo = NUTS3[t["nuts"]]
            km = haversine(*home, la, lo)
            km_kaynak = "~Kreis"
        t["km"] = round(km) if km is not None else None
        t["km_kaynak"] = km_kaynak
        t["kreis"] = NUTS3.get(t["nuts"], ("",))[0]
        if km_kaynak.startswith("~"):
            t["kreis"] += " (km tahmini)"
            t["yer"] = t["yer"] or NUTS3[t["nuts"]][0]
        if km is None or km > radius:
            continue

        hay_title = (t["baslik"] + " " + t["lotlar"]).lower()
        hay_desc = t["aciklama"].lower()
        if any(e in hay_title for e in exc):
            continue
        reasons, score = [], 0
        cpv_hit = sorted({p for p in prefixes for c in t["cpv"] if c.startswith(p)})
        if cpv_hit:
            score += 45
            reasons.append("CPV " + ",".join(h.ljust(8, "x") for h in cpv_hit))
        kw_t = [k for k in kws if k in hay_title]
        kw_d = [k for k in kws if k in hay_desc and k not in kw_t]
        if kw_t:
            score += min(35, 20 + 5 * (len(kw_t) - 1))
            reasons.append("Başlık: " + ", ".join(kw_t[:4]))
        if kw_d:
            score += min(15, 5 * len(kw_d))
            reasons.append("Metin: " + ", ".join(kw_d[:3]))
        if not cpv_hit and not kw_t and not kw_d:
            continue
        score += round(20 * max(0.0, 1 - km / radius))
        t["puan"] = min(100, score)
        t["neden"] = " | ".join(reasons)
        if t["puan"] < cfg.get("min_puan", 0):
            continue
        out.append(t)
    return out


def dedupe(items: list[dict]) -> list[dict]:
    seen, out = {}, []
    for t in sorted(items, key=lambda x: 0 if x["kaynak"].startswith("Bek") else 1):
        key = (norm(t["baslik"])[:60], t["tur"])
        if key in seen and t["kaynak"].startswith("TED"):
            if not seen[key].get("link_ted"):
                seen[key]["link_ted"] = t["link"]
            continue
        seen[key] = t
        out.append(t)
    return out


# ----------------------------------------------------------------- Excel
def read_previous_marks(path: Path) -> dict:
    marks = {}
    if not path.exists():
        return marks
    try:
        wb = load_workbook(path, read_only=True)
        for ws in wb.worksheets:
            rows = ws.iter_rows(values_only=True)
            header = next(rows, None)
            if not header or "ID" not in header:
                continue
            i_id, i_d = header.index("ID"), header.index("Durum") if "Durum" in header else None
            i_n = header.index("Not") if "Not" in header else None
            for r in rows:
                if r and r[i_id]:
                    d = r[i_d] if i_d is not None else ""
                    n = r[i_n] if i_n is not None else ""
                    if d or n:
                        marks[r[i_id]] = (d or "", n or "")
        wb.close()
    except Exception as e:  # noqa: BLE001
        log(f"UYARI: önceki Excel okunamadı ({e}); Durum/Not taşınamadı.")
    return marks


HEAD_FILL = PatternFill("solid", fgColor="1F4E3D")
HEAD_FONT = Font(bold=True, color="FFFFFF")
NEW_FILL = PatternFill("solid", fgColor="E2F0D9")
RED_FILL = PatternFill("solid", fgColor="F8CBAD")
ORG_FILL = PatternFill("solid", fgColor="FFE699")
LINK_FONT = Font(color="0563C1", underline="single")


def write_sheet(ws, cols, rows, widths, link_cols=(), dv=None):
    ws.append(cols)
    for c in ws[1]:
        c.fill, c.font = HEAD_FILL, HEAD_FONT
        c.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 30
    for r in rows:
        ws.append(r)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for col in link_cols:
        ci = cols.index(col) + 1
        for row in range(2, ws.max_row + 1):
            c = ws.cell(row, ci)
            if c.value and str(c.value).startswith("http"):
                c.hyperlink = str(c.value)
                c.value = "Aç ↗"
                c.font = LINK_FONT
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(vertical="top", wrap_text=cols[c.column - 1] in ("Başlık", "Eşleşme", "Lotlar", "Not", "Kazanan"))
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{ws.max_row}"
    if dv and ws.max_row > 1:
        ws.add_data_validation(dv)
        dv.add(f"A2:A{max(ws.max_row, 500)}")


def fmt_d(d):
    return d.strftime("%d.%m.%Y %H:%M") if isinstance(d, dt.datetime) and (d.hour or d.minute) else \
        d.strftime("%d.%m.%Y") if isinstance(d, dt.datetime) else ""


def build_excel(path: Path, acik, on, sonuc, marks, new_ids, cfg, stats) -> Path:
    wb = Workbook()
    now = dt.datetime.now()

    # --- Açık ihaleler
    ws = wb.active
    ws.title = "Açık İhaleler"
    cols = ["Durum", "Not", "Yeni", "Puan", "Son Tarih", "Kalan Gün", "Başlık", "Lotlar", "İhale Makamı",
            "Yer", "PLZ", "Km", "Kreis", "Tür", "Prosedür", "Tahmini Değer €", "Eşleşme", "CPV",
            "Yayın", "Kaynak", "İlan", "Belgeler", "TED", "ID"]
    acik.sort(key=lambda t: (t["son_tarih"] or dt.datetime(2100, 1, 1), -t["puan"]))
    rows = []
    for t in acik:
        d, n = marks.get(t["id"], ("", ""))
        kalan = (t["son_tarih"] - now).days if t["son_tarih"] else None
        rows.append([d, n, "YENİ" if t["id"] in new_ids else "", t["puan"], fmt_d(t["son_tarih"]), kalan,
                     t["baslik"], t["lotlar"], t["makam"], t["yer"], t["plz"],
                     t["km"],
                     t["kreis"], t["kategori"], t["prosedur"], t["deger"], t["neden"], ", ".join(t["cpv"][:6]),
                     fmt_d(t["yayin"]), t["kaynak"], t["link"], t["belge"], t.get("link_ted", ""), t["id"]])
    dv = DataValidation(type="list", formula1='"' + ",".join(DURUM_LISTESI[1:]) + '"', allow_blank=True)
    write_sheet(ws, cols, rows,
                [14, 22, 7, 7, 16, 9, 60, 35, 34, 18, 8, 7, 20, 13, 22, 14, 40, 22, 11, 16, 8, 9, 7, 12],
                link_cols=("İlan", "Belgeler", "TED"), dv=dv)
    ci_k, ci_y, ci_v = cols.index("Kalan Gün") + 1, cols.index("Yeni") + 1, cols.index("Tahmini Değer €") + 1
    for row in range(2, ws.max_row + 1):
        k = ws.cell(row, ci_k).value
        if isinstance(k, int):
            if k < 7:
                ws.cell(row, ci_k).fill = RED_FILL
            elif k < 14:
                ws.cell(row, ci_k).fill = ORG_FILL
        if ws.cell(row, ci_y).value:
            for c in range(1, len(cols) + 1):
                if c != ci_k:
                    ws.cell(row, c).fill = NEW_FILL
        ws.cell(row, ci_v).number_format = "#,##0"

    # --- Ön bildirimler
    ws2 = wb.create_sheet("Ön Bildirimler")
    cols2 = ["Durum", "Not", "Puan", "Yayın", "Başlık", "İhale Makamı", "Yer", "Km", "Eşleşme", "Kaynak", "İlan", "ID"]
    on.sort(key=lambda t: t["yayin"] or dt.datetime.min, reverse=True)
    rows2 = [[*marks.get(t["id"], ("", "")), t["puan"], fmt_d(t["yayin"]), t["baslik"], t["makam"], t["yer"],
              t["km"], t["neden"], t["kaynak"], t["link"], t["id"]] for t in on]
    write_sheet(ws2, cols2, rows2, [14, 22, 7, 11, 60, 34, 18, 7, 40, 16, 8, 12], link_cols=("İlan",),
                dv=DataValidation(type="list", formula1='"' + ",".join(DURUM_LISTESI[1:]) + '"', allow_blank=True))

    # --- Sonuçlananlar (rakip / fiyat istihbaratı)
    ws3 = wb.create_sheet("Sonuçlananlar")
    cols3 = ["Tarih", "Başlık", "İhale Makamı", "Yer", "Km", "Kazanan", "Ihale Bedeli €", "Eşleşme", "Kaynak", "İlan", "ID"]
    sonuc.sort(key=lambda t: t["yayin"] or dt.datetime.min, reverse=True)
    rows3 = [[fmt_d(t["yayin"]), t["baslik"], t["makam"], t["yer"], t["km"], t["kazanan"], t["kazanan_deger"],
              t["neden"], t["kaynak"], t["link"], t["id"]] for t in sonuc]
    write_sheet(ws3, cols3, rows3, [11, 60, 34, 18, 7, 40, 14, 40, 16, 8, 12], link_cols=("İlan",))
    for row in range(2, ws3.max_row + 1):
        ws3.cell(row, 7).number_format = "#,##0"

    # --- Bilgi
    ws4 = wb.create_sheet("Bilgi")
    info = [
        ["İhale Radarı", ""],
        ["Son çalışma", now.strftime("%d.%m.%Y %H:%M")],
        ["Merkez", f"{cfg['merkez_ad']} ({cfg['merkez_plz']})"],
        ["Yarıçap", f"{cfg['yaricap_km']} km (kuş uçuşu)"],
        ["Taranan dönem", f"son {stats['gun']} gün"],
        ["Açık ihale", len(acik)], ["Yeni (bu çalışmada ilk kez)", len(new_ids)],
        ["Ön bildirim", len(on)], ["Sonuçlanan", len(sonuc)],
        ["Kaynaklar", stats["kaynaklar"]],
        ["", ""],
        ["Nasıl kullanılır", "‘Durum’ ve ‘Not’ sütunlarına yazdıklarınız bir sonraki çalışmada korunur (dosya kapalı olmalı)."],
        ["", "Yeşil satır = yeni ilan. Kalan gün: kırmızı < 7, sarı < 14."],
        ["", "Kreis sütununda ‘(km tahmini)’ yazıyorsa ilanda PLZ yok, mesafe Kreis merkezinden hesaplandı."],
        ["", "Ayarlar (merkez, km, CPV, kelimeler, e-posta): config.json"],
        ["", "Teklif, ilan sayfasında belirtilen ihale platformu üzerinden verilir (çoğu ücretsiz kayıt)."],
        ["Kapsam notu", "Kaynak, Bund/Land/Kommune ilanlarını toplayan resmi servistir; ama bazı küçük belediye "
                        "ilanları (ör. sadece kendi sitesinde yayınlanan freihändige/beschränkte Vergabe) bu servise "
                        "düşmeyebilir."],
    ]
    for r in info:
        ws4.append(r)
    ws4.column_dimensions["A"].width = 28
    ws4.column_dimensions["B"].width = 110
    ws4["A1"].font = Font(bold=True, size=14)

    try:
        wb.save(path)
        return path
    except PermissionError:
        alt = path.with_name(f"{path.stem}_{now:%Y%m%d_%H%M}.xlsx")
        log(f"UYARI: {path.name} açık olduğu için yazılamadı -> {alt.name} olarak kaydedildi.")
        wb.save(alt)
        return alt


# ----------------------------------------------------------------- web sayfası
def build_html(path: Path, acik, on, sonuc, new_ids, cfg, stats) -> None:
    tpl = (BASE / "sayfa_sablonu.html").read_text(encoding="utf-8")
    now = dt.datetime.now()

    def row(t):
        return {
            "id": t["id"], "puan": t["puan"], "baslik": t["baslik"], "lotlar": t["lotlar"][:300],
            "makam": t["makam"], "yer": t["yer"], "km": t["km"], "tahmini": t["km_kaynak"].startswith("~"),
            "tur": t["kategori"], "deger": t["deger"], "neden": t["neden"],
            "son": fmt_d(t["son_tarih"]), "kalan": (t["son_tarih"] - now).days if t["son_tarih"] else None,
            "yayin": fmt_d(t["yayin"])[:10], "link": t["link"], "belge": t["belge"],
            "kazanan": t["kazanan"], "bedel": t["kazanan_deger"], "yeni": t["id"] in new_ids,
        }
    data = {"acik": [row(t) for t in acik], "on": [row(t) for t in on], "sonuc": [row(t) for t in sonuc]}
    meta = {"merkez": cfg["merkez_ad"], "km": cfg["yaricap_km"], "guncel": now.strftime("%d.%m.%Y %H:%M"),
            "kaynaklar": stats["kaynaklar"]}
    def js(o):
        return json.dumps(o, ensure_ascii=False).replace("</", "<\\/")
    html = tpl.replace("__DATA__", js(data)).replace("__META__", js(meta))
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(html, encoding="utf-8")
    log(f"Web sayfası yazıldı: {path}")


# ----------------------------------------------------------------- e-posta
def send_mail(cfg: dict, xlsx: Path, new_items: list[dict]) -> None:
    m = cfg.get("email") or {}
    if not m.get("aktif"):
        return
    if m.get("sadece_yeni_varsa", True) and not new_items:
        log("E-posta: yeni ilan yok, gönderilmedi.")
        return
    msg = EmailMessage()
    msg["Subject"] = f"İhale Radarı: {len(new_items)} yeni ihale ({dt.date.today():%d.%m.%Y})"
    msg["From"] = m["kullanici"]
    msg["To"] = ", ".join(m["kime"])
    lines = [f"{len(new_items)} yeni ihale bulundu:\n"]
    for t in sorted(new_items, key=lambda x: -x["puan"])[:40]:
        lines.append(f"• [{t['puan']}] {t['baslik']}\n  {t['makam']} – {t['yer']} ({t['km']} km) – "
                     f"Son: {fmt_d(t['son_tarih']) or '?'}\n  {t['link']}\n")
    msg.set_content("\n".join(lines))
    msg.add_attachment(xlsx.read_bytes(), maintype="application",
                       subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=xlsx.name)
    try:
        with smtplib.SMTP(m["smtp_sunucu"], int(m["smtp_port"]), timeout=60) as s:
            s.starttls()
            s.login(m["kullanici"], m["sifre"])
            s.send_message(msg)
        log(f"E-posta gönderildi: {msg['To']}")
    except Exception as e:  # noqa: BLE001
        log(f"UYARI: e-posta gönderilemedi: {e}")


# ----------------------------------------------------------------- ana akış
def main() -> int:
    ap = argparse.ArgumentParser(description="İhale Radarı")
    ap.add_argument("--gun", type=int, help="kaç gün geriye taransın")
    ap.add_argument("--ac", action="store_true", help="bitince Excel'i aç")
    ap.add_argument("--cikti", default="Ihaleler.xlsx")
    ap.add_argument("--html", help="web sayfası çıktısı (örn. docs/index.html)")
    a = ap.parse_args()

    VERI.mkdir(parents=True, exist_ok=True)
    cfg = load_config()
    days = a.gun or int(cfg.get("tarama_gun", 30))
    log("=" * 60)
    log(f"İhale Radarı başladı – merkez {cfg['merkez_ad']}, {cfg['yaricap_km']} km, {days} gün")

    plz_tab = load_plz_table()
    items, sources = [], []
    if cfg.get("kaynak_bkms", True):
        try:
            items += collect_bkms(cfg, days)
            sources.append("Bekanntmachungsservice (oeffentlichevergabe.de)")
        except Exception as e:  # noqa: BLE001
            log(f"HATA Bekanntmachungsservice: {e}\n{traceback.format_exc()}")
    if cfg.get("kaynak_ted", True):
        try:
            items += collect_ted(cfg, days)
            sources.append("TED (ted.europa.eu)")
        except Exception as e:  # noqa: BLE001
            log(f"HATA TED: {e}")

    matched = dedupe(score_all(items, cfg, plz_tab))
    now = dt.datetime.now()
    acik = [t for t in matched if t["tur"] == "acik"]
    if cfg.get("son_tarihi_gecenleri_gizle", True):
        acik = [t for t in acik if not t["son_tarih"] or t["son_tarih"] >= now - dt.timedelta(hours=12)]
    on = [t for t in matched if t["tur"] == "on"]
    sonuc = [t for t in matched if t["tur"] == "sonuc"]

    state = {}
    if STATE_PATH.exists():
        try:
            state = json.loads(STATE_PATH.read_text(encoding="utf-8-sig"))
        except ValueError:
            state = {}
    seen = set(state.get("gorulen", []))
    first_run = not seen
    new_ids = set() if first_run else {t["id"] for t in acik + on if t["id"] not in seen}
    seen |= {t["id"] for t in acik + on}
    state["gorulen"] = sorted(seen)[-20000:]
    state["son_calisma"] = now.isoformat(timespec="seconds")

    out_path = BASE / a.cikti
    out_path.parent.mkdir(parents=True, exist_ok=True)
    marks = read_previous_marks(out_path)
    written = build_excel(out_path, acik, on, sonuc, marks, new_ids, cfg,
                          {"gun": days, "kaynaklar": ", ".join(sources) or "-"})
    if a.html:
        build_html(BASE / a.html, acik, on, sonuc, new_ids, cfg, {"kaynaklar": ", ".join(sources) or "-"})
    STATE_PATH.write_text(json.dumps(state, ensure_ascii=False), encoding="utf-8")
    log(f"Bitti: {len(acik)} açık ihale ({len(new_ids)} yeni), {len(on)} ön bildirim, "
        f"{len(sonuc)} sonuçlanan -> {written.name}")
    send_mail(cfg, written, [t for t in acik + on if t["id"] in new_ids])
    if a.ac and os.name == "nt":
        os.startfile(written)  # type: ignore[attr-defined]
    return 0


if __name__ == "__main__":
    sys.exit(main())
