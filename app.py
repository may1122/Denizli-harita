from __future__ import annotations

import html
import math
import re
import unicodedata
from pathlib import Path

import folium
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from folium import FeatureGroup
from folium.plugins import Fullscreen, MeasureControl
from branca.element import MacroElement, Template


# ============================================================
# STREAMLIT AYARLARI
# ============================================================

st.set_page_config(
    page_title="Denizli Eczane Grup Haritası",
    page_icon="💊",
    layout="wide",
)

BASE_DIR = Path(__file__).resolve().parent
ECZANE_FILE_NAME = "denizli_eczaneler.xlsx"


# ============================================================
# GRUP RENKLERİ
# ============================================================

GROUP_COLORS = {
    "A1": "#4A148C",
    "A2": "#7B1FA2",
    "A3": "#AB47BC",
    "B1": "#0D47A1",
    "B2": "#1976D2",
    "B3": "#64B5F6",
    "C1": "#00695C",
    "C2": "#00897B",
    "C3": "#4DB6AC",
    "D1": "#E65100",
    "D2": "#EF6C00",
    "D3": "#FFB74D",
    "E1": "#1B5E20",
    "E2": "#388E3C",
    "E3": "#81C784",
    "F1": "#B71C1C",
    "F2": "#D32F2F",
    "F3": "#EF5350",
    "G1": "#4A148C",
    "G2": "#8E24AA",
    "G3": "#CE93D8",
    "H1": "#01579B",
    "H2": "#039BE5",
    "H3": "#81D4FA",
    "K1": "#F57F17",
    "K2": "#F9A825",
    "K3": "#FFD54F",
}


# ============================================================
# DENİZLİ 9 ANA GRUP / 27 ALT GRUP
# ============================================================

GROUPS: dict[str, list[str]] = {
    "A1": [
        'AYNUR GÜLER',
        'AYŞEN',
        'BAKLAN',
        'BAŞDİL',
        'BAŞÇAVUŞ',
        'BÜŞRA BOYACI',
        'CEYDA POLAT',
        'DELİKTAŞ',
        'DEMİRTAŞ',
    ],

    "A2": [
        'DERYAM',
        'DOKUZKAVAKLAR',
        'ERMAN',
        'GAMZE',
        'GENCER',
        'KAYDIHAN',
        'KOÇAK',
    ],

    "A3": [
        'KUNDAKÇI',
        'MAVİ',
        'MİRA',
        'NEVA',
        'SEMT',
        'YENİLMEZ',
        'ZEYNEP',
        'ÇALLIOĞLU',
        'DİLEK',
    ],

    "B1": [
        'ADALET',
        'AKDENİZ',
        'CANDENİZ',
        'CANSU',
        'DEMİR',
        'EMİR',
        'FLORYA',
        'GÜLRİZ',
        'HACETTEPE',
    ],

    "B2": [
        'HAZAR',
        'KAPLAN',
        'KAYHAN',
        'KÖKNAR',
        'NÜKHET',
        'NİLGÜN',
        'PAPATYA',
    ],

    "B3": [
        'SEMİH',
        'TURAN',
        'UĞUR',
        'ÇAKMAK',
        'ÇİFTÇİ',
        'ÖZGEN',
        'ÖZSOY',
    ],

    "C1": [
        'AYKUT',
        'ESİN',
        'ÖZDERMAN',
        'BAYRAMYERİ',
        'SERGEN',
        'DAĞDEVİREN',
        'EGE',
        'ŞULE',
        'ŞİFA',
    ],

    "C2": [
        'BÜYÜK',
        'SAĞLIK',
        'DUYGU',
        'ERCAN',
        'ERDEM',
        'DENİZLİ',
        'ANAFARTALAR',
        'DİŞÇİOĞLU',
    ],

    "C3": [

        'MERVE',
        'MERKEZ',
        'PELİTLİBAĞ',
        'FATIMA ŞENTÜRK',
        'ÇETİNKAYA',
        'FATİH',
        'GÜRKAN',
        'ASMALI',
    ],

    "D1": [
        'ÇOMUT',
        'CANAN',
        'ARCA',
        'GÖKSU',
        'MORALIOĞLU',
        'DEMİRAY',
        'EFE',
        'GÜLAY',
        'IŞIL',
      
    ],

    "D2": [
        'AYGÖREN',
        'ÇAKMAKLIOĞLU',
        'VERESELİ DENİZLİ',
        'CADDE',
        'ADALI',
        'BURCU',
        'LİMONCU',
        'EZO',
        'GÜLEÇ',
    ],

    "D3": [
        'BAYRAMOĞLU',
        'GÖKHAN',
        'GÜNGÖR',
        'LOKMAN',
        'KIVILCIM',
        'NUR BAŞÇAVUŞ',
        'CEYHAN',
        'ÖZNUR',
    ],

    "E1": [
        'AYFER CEYLAN',
        'TURUNÇ',
        'GÜLERYÜZ',
        '29_EKİM',
        'ÜMİT',
        'ERTUĞRUL',
        'HASİBE KARTOĞLAN',
        'ÖZCEL',
        'SEVİM',
    ],

    "E2": [
        'SENA KELLECİ',
        'İLKE',
        'UMAY',
        'AKTÜRK',
        'ASLI',
        'KİRAZ',
        'GÖZDE GÜNDÜZ',
    ],

    "E3": [
        'OCAK',
        'TOLGAY',
        'AKKAYA',
        'ADA',
        'UZMAN',
        'DENİZİM',
        'DİNÇ',
        'EZGİ',
    ],

    "F1": [
        'SEDA BAŞDİL',
        'MERKEZEFENDİ',
        'TURKUAZ',
        'YEŞİLYURT',
        'MEHMET KAYA',
        'NİSAN',
        'NEFES',
        'BAHAR',
    ],

    "F2": [
        'ELİF PAMUKÇU',
        'OKYANUS',
        'KINIKLI',
        'SARAÇOĞLU',
        'MUTLU GÜNLER',
        'ÇAMLICA',
        'AYDIN',
    ],

    "F3": [
        'İSTİKLAL',
        'HÜRRİYET',
        'AYLİN',
        'ÇETİN',
        'ÖZKAN',
        'KÖSELER',
        'ERSAN',
        'ALBAYRAK',
        'SAHRA',
    ],

    "G1": [
        'CEMRE',
        'EKİZ',
        'ELVAN',
        'GÖKKUŞAĞI',
        'SİNEM',
        'IŞIMLIK',
        'KABAYUKA',
        'PAMUKKALE AKTÜRK',
        'ZEYNEP SULTAN',
    ],

    "G2": [
        'ÖZGÜR',
        'ALPLER',
        'FORUM ÇAMLIK',
        'BİLGE',
        'EVREN',
        'ANIL',
        'DOĞAL',
    ],

    "G3": [
        'ELİF',
        'NEŞE',
        'SEVİL',
        'SEÇKİN',
        'TEMMUZ',
        'İNCEOĞLU',
        
    ],

    "H1": [
        'BERGAMA',
        'KEKİK',
        'CEREN FİLİZER',
        'ALSANCAK',
        'CANDAN',
        'DEMİRCİOĞLU GÜL',
        'DEMİROĞLU',
        'DERMAN',
        'DEVECİ',
    ],

    "H2": [
        'EZGİ KIRDI',
        'GÖKÇE',
        'GÜRSOY',
        'KIZILTAŞ',
        'MERVE YAMUÇ',
        'PAMUKKALE',
        'SAYGIN',
    ],

    "H3": [
        'SOYLU',
        'SU',
        'TUBA',
        'ZEYTİNKÖY SEMA',
        'ÜNİVERSİTE',
        'İNANÖZ',
        'ASLAN',
    ],

    "K1": [
        'ALTINOVA',
        'GÖRKEM',
        'CADDE SAĞLIK',
        'CANSU ERKİLET',
        'CANSUYU',
        'CEYLAN',
        "ELİF'İN",
        'EMEK',
        'GÜNEŞ',
    ],

    "K2": [
        'IRMAK',
        'NAZAN',
        'OZAN',
        'PARK BOTANİK',
        'SERVET',
        'TUGAY',
        'TÜFEKÇİOĞLU',
    ],

    "K3": [
        'YEŞİLYUVA',
        'ÇAMLIK',
        'ÖZGÜ',
        'ÖZGÜN KIYAT',
        'İZMİRLİ',
        'ŞİRİN',
        
    ],

}


ALL_GROUPS = list(GROUPS.keys())


# ============================================================
# ECZANE İSMİ NORMALİZASYONU
# ============================================================

def normalize_name(value: object) -> str:

    if pd.isna(value):
        return ""

    text = unicodedata.normalize(
        "NFKC",
        str(value)
    ).strip().upper()

    text = text.translate(
        str.maketrans(
            {
                "Ç": "C",
                "Ğ": "G",
                "İ": "I",
                "Ö": "O",
                "Ş": "S",
                "Ü": "U",
            }
        )
    )

    # boşlukları kaldır
    text = re.sub(r"\s+", "", text)

    # özel karakterleri kaldır
    text = re.sub(r"[^0-9A-Z]", "", text)

    # "ECZANESI" son ekini kaldır
    if text.endswith("ECZANESI"):
        text = text[:-8]

    return text


# ============================================================
# GRUP HARİTASI
# ============================================================

def build_group_map() -> dict[str, str]:

    result: dict[str, str] = {}

    for group_name, names in GROUPS.items():

        for name in names:

            key = normalize_name(name)

            if key in result and result[key] != group_name:
                raise ValueError(
                    f"{name} birden fazla grupta tanımlanmış: "
                    f"{result[key]} ve {group_name}"
                )

            result[key] = group_name

    return result


# ============================================================
# KOORDİNAT OKUMA
# ============================================================

def parse_coord(value: object) -> float | None:

    if pd.isna(value):
        return None

    try:
        v = float(value)

    except Exception:

        txt = str(value).strip().replace(",", ".")

        try:
            v = float(txt)

        except Exception:
            return None

    # Örnek:
    # 3778015 -> 37.78015
    # 2909645 -> 29.09645
    if abs(v) > 180:
        v = v / 100000.0

    return v


# ============================================================
# EXCEL OKUMA
# ============================================================

@st.cache_data(show_spinner=False)
def read_pharmacies(
    path: str,
    file_version: int,
) -> pd.DataFrame:

    # Excel değiştiğinde Streamlit cache yenilensin
    del file_version

    raw = pd.read_excel(
        path,
        engine="openpyxl",
    )

    required = [
        "Eczane İsmi",
        "Enlem (Latitude)",
        "Boylam (Longitude)",
    ]

    missing = [
        column
        for column in required
        if column not in raw.columns
    ]

    if missing:
        raise ValueError(
            "Eksik sütun(lar): "
            + ", ".join(missing)
        )

    if "Eczane Adresi" in raw.columns:
        address_column = (
            raw["Eczane Adresi"]
            .fillna("")
            .astype(str)
            .str.strip()
        )
    else:
        address_column = pd.Series(
            [""] * len(raw),
            index=raw.index,
        )

    df = pd.DataFrame(
        {
            "Eczane": (
                raw["Eczane İsmi"]
                .fillna("")
                .astype(str)
                .str.strip()
            ),

            "Adres": address_column,

            "Latitude": (
                raw["Enlem (Latitude)"]
                .map(parse_coord)
            ),

            "Longitude": (
                raw["Boylam (Longitude)"]
                .map(parse_coord)
            ),
        }
    )

    df["Anahtar"] = df["Eczane"].map(normalize_name)

    # Koordinatı olmayanları çıkar
    df = df.dropna(
        subset=[
            "Latitude",
            "Longitude",
        ]
    )

    # Geçerli koordinat kontrolü
    df = df[
        df["Latitude"].between(-90, 90)
        & df["Longitude"].between(-180, 180)
    ]

    # Aynı eczanenin yazım varyantlarını tekilleştir
    df = (
        df.drop_duplicates(
            subset=["Anahtar"],
            keep="first",
        )
        .reset_index(drop=True)
    )

    return df


# ============================================================
# İKİ NOKTA ARASI MESAFE
# ============================================================

def latlon_distance_m(
    a: tuple[float, float],
    b: tuple[float, float],
) -> float:

    lat1, lon1 = map(
        math.radians,
        a,
    )

    lat2, lon2 = map(
        math.radians,
        b,
    )

    dlat = lat2 - lat1
    dlon = lon2 - lon1

    h = (
        math.sin(dlat / 2) ** 2
        + math.cos(lat1)
        * math.cos(lat2)
        * math.sin(dlon / 2) ** 2
    )

    return (
        2
        * 6371000.0
        * math.asin(math.sqrt(h))
    )


# ============================================================
# MINIMUM SPANNING TREE
# Aynı grubun eczanelerini gereksiz karmaşa oluşturmadan bağlar
# ============================================================

def minimum_spanning_edges(
    points: list[tuple[float, float]],
) -> list[tuple[int, int]]:

    if len(points) < 2:
        return []

    used = {0}

    edges: list[
        tuple[int, int]
    ] = []

    while len(used) < len(points):

        best: tuple[
            float,
            int,
            int,
        ] | None = None

        for i in used:

            for j in range(len(points)):

                if j in used:
                    continue

                distance = latlon_distance_m(
                    points[i],
                    points[j],
                )

                if (
                    best is None
                    or distance < best[0]
                ):
                    best = (
                        distance,
                        i,
                        j,
                    )

        assert best is not None

        _, i, j = best

        edges.append(
            (i, j)
        )

        used.add(j)

    return edges




# ============================================================
# GÖRSEL KOORDİNAT AYRIŞTIRMA
# Aynı / neredeyse aynı koordinattaki eczaneleri ekranda ayırır.
# ÖNEMLİ: Latitude / Longitude gerçek koordinat olarak korunur.
# Yoğunluk çemberi ve mesafe hesabı GERÇEK koordinatları kullanır.
# ============================================================

def add_display_coordinates(
    df: pd.DataFrame,
    overlap_threshold_m: float = 3.0,
    spread_radius_m: float = 8.0,
) -> pd.DataFrame:
    """
    Birbirine çok yakın (varsayılan <= 3 m) eczaneleri yalnızca görsel
    olarak küçük bir halka üzerine dağıtır. Gerçek koordinatlar değişmez.

    Böylece:
    - 1 eczane = 1 görünür nokta,
    - üst üste binen markerlar kaybolmaz,
    - yoğunluk çemberi gerçek konumdan hesap yapmaya devam eder.
    """

    work = df.copy().reset_index(drop=True)
    work["DisplayLatitude"] = work["Latitude"].astype(float)
    work["DisplayLongitude"] = work["Longitude"].astype(float)

    if len(work) < 2:
        return work

    # Yakın noktaları bağlı bileşen mantığıyla kümelendir.
    unvisited = set(range(len(work)))
    clusters: list[list[int]] = []

    while unvisited:
        seed = unvisited.pop()
        cluster = [seed]
        queue = [seed]

        while queue:
            i = queue.pop()
            a = (
                float(work.at[i, "Latitude"]),
                float(work.at[i, "Longitude"]),
            )

            nearby = []
            for j in list(unvisited):
                b = (
                    float(work.at[j, "Latitude"]),
                    float(work.at[j, "Longitude"]),
                )
                if latlon_distance_m(a, b) <= overlap_threshold_m:
                    nearby.append(j)

            for j in nearby:
                unvisited.remove(j)
                cluster.append(j)
                queue.append(j)

        clusters.append(cluster)

    for cluster in clusters:
        if len(cluster) <= 1:
            continue

        center_lat = sum(float(work.at[i, "Latitude"]) for i in cluster) / len(cluster)
        center_lon = sum(float(work.at[i, "Longitude"]) for i in cluster) / len(cluster)

        # Enlem / boylam derece dönüşümü (küçük mesafeler için yeterince hassas).
        lat_deg_per_m = 1.0 / 111_320.0
        cos_lat = max(0.2, math.cos(math.radians(center_lat)))
        lon_deg_per_m = 1.0 / (111_320.0 * cos_lat)

        # Nokta sayısı arttıkça halkayı çok az büyüt.
        radius_m = spread_radius_m + max(0, len(cluster) - 2) * 1.25

        for pos, idx in enumerate(cluster):
            angle = (2.0 * math.pi * pos / len(cluster)) - (math.pi / 2.0)
            north_m = math.cos(angle) * radius_m
            east_m = math.sin(angle) * radius_m

            work.at[idx, "DisplayLatitude"] = center_lat + north_m * lat_deg_per_m
            work.at[idx, "DisplayLongitude"] = center_lon + east_m * lon_deg_per_m

    return work


# ============================================================
# GRUP BAĞLANTI ÇİZGİLERİ
# ============================================================

def add_group_lines(
    map_obj: folium.Map,
    df: pd.DataFrame,
    selected_groups: set[str],
) -> None:
    """
    Bağlantıları HER HARİTA OLUŞUMUNDA doğrudan güncel GROUPS
    tanımından yeniden kurar.

    Böylece bir eczane bir gruptan başka bir gruba taşındığında:
    - eski grubun çizgi ağı o eczaneyi kesinlikle kullanmaz,
    - yeni grubun çizgi ağı eczaneyi dahil ederek baştan hesaplanır.

    df["Grup"] sütununa güvenmek yerine GROUPS içindeki güncel isimleri
    normalize edip doğrudan eşleştiriyoruz. Bu, eski grup bilgisinin çizgide
    kalması ihtimalini ortadan kaldırır.
    """

    # DataFrame'deki her eczanenin güncel anahtarını garanti altına al.
    work_df = df.copy()
    if "Anahtar" not in work_df.columns:
        work_df["Anahtar"] = work_df["Eczane"].map(normalize_name)

    for group_name in ALL_GROUPS:

        if group_name not in selected_groups:
            continue

        # ----------------------------------------------------
        # SADECE GÜNCEL GROUPS TANIMI GERÇEK KAYNAK
        # ----------------------------------------------------
        current_group_keys = {
            normalize_name(name)
            for name in GROUPS[group_name]
        }

        subset = (
            work_df.loc[
                work_df["Anahtar"].isin(current_group_keys),
                [
                    "Eczane",
                    "Anahtar",
                    "Latitude",
                    "Longitude",
                    "DisplayLatitude",
                    "DisplayLongitude",
                ],
            ]
            .dropna(subset=["Latitude", "Longitude"])
            .drop_duplicates(subset=["Anahtar"], keep="first")
            .copy()
            .reset_index(drop=True)
        )

        if len(subset) < 2:
            continue

        color = GROUP_COLORS[group_name]

        # MST hesabı GERÇEK koordinatlarla yapılır.
        real_points = [
            (
                float(row.Latitude),
                float(row.Longitude),
            )
            for row in subset.itertuples(index=False)
        ]

        # Çizim ise üst üste binmeleri ayıran GÖRSEL koordinatlarla yapılır.
        draw_points = [
            (
                float(row.DisplayLatitude),
                float(row.DisplayLongitude),
            )
            for row in subset.itertuples(index=False)
        ]

        # MST her seferinde sadece güncel grup üyeleriyle sıfırdan hesaplanır.
        edges = minimum_spanning_edges(real_points)

        line_layer = FeatureGroup(
            name=f"{group_name} bağlantıları",
            show=True,
        )

        for i, j in edges:

            row_i = subset.iloc[i]
            row_j = subset.iloc[j]

            key_i = str(row_i["Anahtar"])
            key_j = str(row_j["Anahtar"])

            # ------------------------------------------------
            # SON GÜVENLİK KONTROLÜ
            # Çizginin iki ucu da hâlâ bu grubun güncel üyesi
            # değilse o çizgiyi ASLA çizme.
            # ------------------------------------------------
            if (
                key_i not in current_group_keys
                or key_j not in current_group_keys
            ):
                continue

            pharmacy_1 = str(row_i["Eczane"])
            pharmacy_2 = str(row_j["Eczane"])

            folium.PolyLine(
                locations=[
                    draw_points[i],
                    draw_points[j],
                ],
                color=color,
                weight=2.4,
                opacity=0.82,
                dash_array="7 6",
                line_cap="round",
                line_join="round",
                tooltip=(
                    f"{group_name}: "
                    f"{pharmacy_1} ↔ {pharmacy_2}"
                ),
            ).add_to(line_layer)

        line_layer.add_to(map_obj)


# ============================================================
# ECZANE MARKERLARI
# ============================================================

def add_markers(
    map_obj: folium.Map,
    df: pd.DataFrame,
    selected_groups: set[str],
) -> None:

    for group_name in ALL_GROUPS:

        if group_name not in selected_groups:
            continue

        subset = df.loc[
            df["Grup"].eq(group_name)
        ]

        if subset.empty:
            continue

        marker_layer = FeatureGroup(
            name=f"{group_name} eczaneleri",
            show=True,
        )

        color = GROUP_COLORS[group_name]

        for _, row in subset.iterrows():

            pharmacy_name = html.escape(
                str(row["Eczane"])
            )

            address = html.escape(
                str(row["Adres"])
            )

            tooltip = (
                '<div style="'
                'font-size:13px;'
                'line-height:1.35;'
                '">'
                f"<b>{pharmacy_name}</b>"
                "<br>"
                f"{html.escape(group_name)}"
            )

            if address:
                tooltip += (
                    '<br>'
                    '<span style="color:#666">'
                    f"{address}"
                    "</span>"
                )

            tooltip += "</div>"

            folium.CircleMarker(
                location=[
                    float(row["DisplayLatitude"]),
                    float(row["DisplayLongitude"]),
                ],

                radius=6.0,

                # marker dış çerçevesi
                color="#FFFFFF",

                weight=1.6,

                fill=True,

                fill_color=color,

                fill_opacity=0.98,

                tooltip=folium.Tooltip(
                    tooltip,
                    sticky=True,
                    direction="top",
                ),
            ).add_to(
                marker_layer
            )

        marker_layer.add_to(
            map_obj
        )



# ============================================================
# YOĞUNLUK ÇEMBERİ
# Haritada kırmızı çemberi taşıyarak içindeki eczaneleri sayar
# ============================================================

class DensityCircleControl(MacroElement):
    """Leaflet üzerinde sürüklenebilir kırmızı çember ve canlı eczane sayacı."""

    def __init__(
        self,
        pharmacy_points: list[dict[str, object]],
        center_lat: float,
        center_lon: float,
    ):
        super().__init__()
        self._name = "DensityCircleControl"

        import json

        pharmacies_json = json.dumps(pharmacy_points, ensure_ascii=False)

        self._template = Template(
            r"""
{% macro script(this, kwargs) %}
(function() {
    const map = {{ this._parent.get_name() }};
    const pharmacies = {{ this.pharmacies_json | safe }};
    const total = pharmacies.length;

    const startLatLng = L.latLng(
        {{ this.center_lat }},
        {{ this.center_lon }}
    );

    const densityCircle = L.circle(startLatLng, {
        radius: 1000,
        color: '#C62828',
        weight: 3,
        opacity: 0.95,
        fillColor: '#EF5350',
        fillOpacity: 0.12,
        interactive: false
    }).addTo(map);

    const centerIcon = L.divIcon({
        className: '',
        html: '<div style="width:24px;height:24px;border-radius:50%;background:#C62828;border:4px solid white;box-shadow:0 1px 6px rgba(0,0,0,.45);cursor:move;"></div>',
        iconSize: [24, 24],
        iconAnchor: [12, 12]
    });

    const centerHandle = L.marker(startLatLng, {
        draggable: true,
        icon: centerIcon,
        zIndexOffset: 2000,
        title: 'Çember merkezini sürükle'
    }).addTo(map);

    const DensityControl = L.Control.extend({
        options: { position: 'topleft' },

        onAdd: function() {
            const div = L.DomUtil.create(
                'div',
                'ayca-density-panel'
            );

            div.innerHTML = `
                <div style="font-weight:700;font-size:15px;margin-bottom:3px;">
                    Yoğunluk Çemberi
                </div>

                <div style="font-size:11px;color:#666;margin-bottom:8px;">
                    Yalnızca haritada görünür alt gruplardaki eczaneler sayılıyor
                </div>

                <div style="display:flex;justify-content:space-between;gap:18px;font-size:14px;margin:5px 0;">
                    <span>Yarıçap</span>
                    <strong id="ayca-radius-value">1000 m</strong>
                </div>

                <div style="display:flex;justify-content:space-between;gap:18px;font-size:14px;margin:5px 0;">
                    <span>Çember içi</span>
                    <strong id="ayca-count-value">0 eczane</strong>
                </div>

                <div style="display:flex;justify-content:space-between;gap:18px;font-size:14px;margin:5px 0;">
                    <span>Toplam oran</span>
                    <strong id="ayca-share-value">0%</strong>
                </div>

                <div style="margin-top:9px;font-size:12px;font-weight:600;">
                    Yarıçapı değiştir
                </div>

                <input
                    id="ayca-radius-slider"
                    type="range"
                    min="100"
                    max="2500"
                    step="50"
                    value="1000"
                    style="width:100%;margin-top:5px;accent-color:#C62828;"
                >

                <div style="display:flex;justify-content:space-between;font-size:10px;color:#666;">
                    <span>100 m</span>
                    <span>2500 m</span>
                </div>

                <button
                    id="ayca-center-mode"
                    type="button"
                    style="width:100%;margin-top:9px;padding:7px 8px;border:1px solid #bbb;border-radius:7px;background:white;cursor:pointer;font-weight:600;"
                >
                    Haritadan merkez seç
                </button>

                <div
                    id="ayca-density-hint"
                    style="margin-top:7px;font-size:11px;color:#666;line-height:1.3;"
                >
                    Kırmızı noktayı sürükleyerek çemberi taşıyabilirsiniz.
                </div>

                <div
                    id="ayca-group-breakdown"
                    style="margin-top:9px;padding-top:8px;border-top:1px solid #e3e3e3;font-size:11px;line-height:1.45;"
                ></div>
            `;

            div.style.background = 'rgba(255,255,255,.97)';
            div.style.border = '1px solid #d8d8d8';
            div.style.borderRadius = '10px';
            div.style.padding = '12px 14px';
            div.style.minWidth = '225px';
            div.style.boxShadow = '0 2px 8px rgba(0,0,0,.15)';
            div.style.fontFamily = 'Arial, sans-serif';
            div.style.color = '#222';

            L.DomEvent.disableClickPropagation(div);
            L.DomEvent.disableScrollPropagation(div);

            return div;
        }
    });

    map.addControl(new DensityControl());

    function distanceMeters(lat1, lon1, lat2, lon2) {
        const R = 6371000;
        const toRad = d => d * Math.PI / 180;

        const dLat = toRad(lat2 - lat1);
        const dLon = toRad(lon2 - lon1);

        const a =
            Math.sin(dLat / 2) ** 2 +
            Math.cos(toRad(lat1)) *
            Math.cos(toRad(lat2)) *
            Math.sin(dLon / 2) ** 2;

        return 2 * R * Math.asin(Math.sqrt(a));
    }

    function updateDensity() {
        const center = densityCircle.getLatLng();
        const radius = densityCircle.getRadius();

        let count = 0;
        const groupCounts = {};

        pharmacies.forEach(p => {
            if (
                distanceMeters(
                    center.lat,
                    center.lng,
                    p.lat,
                    p.lon
                ) <= radius
            ) {
                count += 1;

                const groupName = p.group || 'Grupsuz';
                groupCounts[groupName] =
                    (groupCounts[groupName] || 0) + 1;
            }
        });

        const radiusEl =
            document.getElementById('ayca-radius-value');

        const countEl =
            document.getElementById('ayca-count-value');

        const shareEl =
            document.getElementById('ayca-share-value');

        const breakdownEl =
            document.getElementById('ayca-group-breakdown');

        if (radiusEl) {
            radiusEl.textContent =
                Math.round(radius) + ' m';
        }

        if (countEl) {
            countEl.textContent =
                count + ' eczane';
        }

        if (shareEl) {
            shareEl.textContent = total
                ? ((count / total) * 100)
                    .toFixed(1)
                    .replace('.', ',') + '%'
                : '0%';
        }

        if (breakdownEl) {
            const order = ['A1', 'A2', 'A3', 'B1', 'B2', 'B3', 'C1', 'C2', 'C3', 'D1', 'D2', 'D3', 'E1', 'E2', 'E3', 'F1', 'F2', 'F3', 'G1', 'G2', 'G3', 'H1', 'H2', 'H3', 'K1', 'K2', 'K3', 'Grupsuz'];

            const rows = order
                .filter(group => groupCounts[group])
                .map(
                    group =>
                        `<div style="display:flex;justify-content:space-between;gap:14px;">
                            <span>${group}</span>
                            <strong>${groupCounts[group]}</strong>
                        </div>`
                )
                .join('');

            breakdownEl.innerHTML = rows
                ? '<div style="font-weight:700;margin-bottom:4px;">Grup dağılımı</div>' + rows
                : '<span style="color:#777;">Çember içinde eczane yok.</span>';
        }
    }

    centerHandle.on('drag', function(e) {
        densityCircle.setLatLng(
            e.target.getLatLng()
        );

        updateDensity();
    });

    centerHandle.on('dragend', function(e) {
        densityCircle.setLatLng(
            e.target.getLatLng()
        );

        updateDensity();
    });

    setTimeout(function() {

        const slider =
            document.getElementById('ayca-radius-slider');

        const centerButton =
            document.getElementById('ayca-center-mode');

        const hint =
            document.getElementById('ayca-density-hint');

        let chooseCenter = false;

        if (slider) {

            slider.addEventListener(
                'input',
                function() {

                    densityCircle.setRadius(
                        Number(this.value)
                    );

                    updateDensity();
                }
            );
        }

        if (centerButton) {

            centerButton.addEventListener(
                'click',
                function() {

                    chooseCenter =
                        !chooseCenter;

                    if (chooseCenter) {

                        this.textContent =
                            'Haritada bir noktaya tıkla';

                        this.style.background =
                            '#FDECEC';

                        this.style.borderColor =
                            '#C62828';

                        if (hint) {
                            hint.textContent =
                                'Şimdi haritada çemberin merkezini istediğiniz yere tıklayın.';
                        }

                    } else {

                        this.textContent =
                            'Haritadan merkez seç';

                        this.style.background =
                            'white';

                        this.style.borderColor =
                            '#bbb';

                        if (hint) {
                            hint.textContent =
                                'Kırmızı noktayı sürükleyerek çemberi taşıyabilirsiniz.';
                        }
                    }
                }
            );
        }

        map.on(
            'click',
            function(e) {

                if (!chooseCenter) {
                    return;
                }

                centerHandle.setLatLng(
                    e.latlng
                );

                densityCircle.setLatLng(
                    e.latlng
                );

                updateDensity();

                chooseCenter = false;

                if (centerButton) {

                    centerButton.textContent =
                        'Haritadan merkez seç';

                    centerButton.style.background =
                        'white';

                    centerButton.style.borderColor =
                        '#bbb';
                }

                if (hint) {
                    hint.textContent =
                        'Merkez değiştirildi. Kırmızı noktayı da sürükleyebilirsiniz.';
                }
            }
        );

        updateDensity();

    }, 0);

})();
{% endmacro %}
"""
        )

        self.pharmacies_json = pharmacies_json
        self.center_lat = center_lat
        self.center_lon = center_lon


def add_density_circle(
    map_obj: folium.Map,
    df: pd.DataFrame,
    selected_groups: set[str],
) -> None:
    """Haritada görünür olan seçili gruplardaki eczaneleri sayar."""

    visible_df = df.loc[
        df["Grup"].isin(selected_groups)
    ].copy()

    if visible_df.empty:
        return

    pharmacy_points = []

    for _, row in visible_df.iterrows():
        group_name = str(row["Grup"])

        pharmacy_points.append(
            {
                "name": str(row["Eczane"]),
                "lat": float(row["Latitude"]),
                "lon": float(row["Longitude"]),
                "group": group_name,
            }
        )

    control = DensityCircleControl(
        pharmacy_points=pharmacy_points,
        center_lat=float(visible_df["Latitude"].median()),
        center_lon=float(visible_df["Longitude"].median()),
    )

    control.add_to(map_obj)


# ============================================================
# HARİTAYI OLUŞTUR
# ============================================================

def build_map(
    df: pd.DataFrame,
    selected_groups: set[str],
) -> folium.Map:

    center = [
        float(
            df["Latitude"].median()
        ),
        float(
            df["Longitude"].median()
        ),
    ]

    # --------------------------------------------------------
    # prefer_canvas=False
    #
    # Kesikli çizgilerin SVG olarak çizilmesini sağlıyor.
    # Grup değişikliklerinde çizgilerin güncellenmesi açısından
    # daha güvenilir.
    # --------------------------------------------------------

    m = folium.Map(
        location=center,
        zoom_start=13,
        tiles=None,
        control_scale=True,
        prefer_canvas=False,
    )

    # --------------------------------------------------------
    # HARİTA ALTLIĞI — SİVAS İLE AYNI
    # API key gerektirmez
    # --------------------------------------------------------

    folium.TileLayer(
        tiles="CartoDB positron",
        name="Sade harita",
        control=True,
        show=True,
    ).add_to(m)

    folium.TileLayer(
        tiles="OpenStreetMap",
        name="Detaylı harita",
        control=True,
        show=False,
    ).add_to(m)

    # --------------------------------------------------------
    # ÖNCE ÇİZGİLER
    # SONRA MARKERLAR
    #
    # Böylece renkli toplar çizginin üstünde kalır.
    # --------------------------------------------------------

    add_group_lines(
        m,
        df,
        selected_groups,
    )

    add_markers(
        m,
        df,
        selected_groups,
    )

    # --------------------------------------------------------
    # KIRMIZI YOĞUNLUK ÇEMBERİ
    # Tüm eczaneleri canlı olarak sayar.
    # --------------------------------------------------------

    add_density_circle(
        m,
        df,
        selected_groups,
    )

    # --------------------------------------------------------
    # HARİTA ARAÇLARI
    # --------------------------------------------------------

    Fullscreen(
        position="topright",
        title="Tam ekran",
        title_cancel="Tam ekrandan çık",
    ).add_to(m)

    MeasureControl(
        position="topright",
        primary_length_unit="meters",
    ).add_to(m)

    folium.LayerControl(
        collapsed=True,
        position="topright",
    ).add_to(m)

    # --------------------------------------------------------
    # SEÇİLİ GRUPLARA GÖRE HARİTA SINIRI
    # --------------------------------------------------------

    selected_df = df[
        df["Grup"].isin(
            selected_groups
        )
    ]

    if not selected_df.empty:
        bounds_df = selected_df
    else:
        bounds_df = df

    if not bounds_df.empty:

        m.fit_bounds(
            [
                [
                    float(
                        bounds_df[
                            "Latitude"
                        ].min()
                    ),
                    float(
                        bounds_df[
                            "Longitude"
                        ].min()
                    ),
                ],
                [
                    float(
                        bounds_df[
                            "Latitude"
                        ].max()
                    ),
                    float(
                        bounds_df[
                            "Longitude"
                        ].max()
                    ),
                ],
            ],
            padding=(25, 25),
        )

    return m


# ============================================================
# STREAMLIT SESSION STATE
# ============================================================

def init_state() -> None:

    for group_name in ALL_GROUPS:

        st.session_state.setdefault(
            f"filter_{group_name}",
            True,
        )


def select_all(
    value: bool,
) -> None:

    for group_name in ALL_GROUPS:

        st.session_state[
            f"filter_{group_name}"
        ] = value


# ============================================================
# ANA PROGRAM
# ============================================================

pharmacy_path = (
    BASE_DIR
    / ECZANE_FILE_NAME
)

if not pharmacy_path.exists():

    st.error(
        f"{ECZANE_FILE_NAME} bulunamadı. "
        "Dosyayı app.py ile aynı klasöre koyun."
    )

    st.stop()


try:

    # --------------------------------------------------------
    # EXCEL OKU
    # --------------------------------------------------------

    pharmacies = read_pharmacies(
        str(pharmacy_path),
        pharmacy_path.stat().st_mtime_ns,
    )

    # --------------------------------------------------------
    # GÜNCEL GRUP EŞLEMESİ
    # --------------------------------------------------------

    group_map = build_group_map()

    pharmacies["Grup"] = (
        pharmacies["Anahtar"]
        .map(group_map)
        .astype("string")
    )

    # --------------------------------------------------------
    # ÜST ÜSTE BİNEN MARKERLARI GÖRSEL OLARAK AYIR
    # Gerçek Latitude/Longitude değişmez.
    # --------------------------------------------------------

    pharmacies = add_display_coordinates(pharmacies)

    # --------------------------------------------------------
    # SESSION STATE
    # --------------------------------------------------------

    init_state()

    # ========================================================
    # SIDEBAR
    # ========================================================

    st.sidebar.header(
        "Denizli 9 Ana Grup / 27 Alt Grup"
    )

    c1, c2 = st.sidebar.columns(2)

    c1.button(
        "Tümünü Aç",
        on_click=select_all,
        args=(True,),
        use_container_width=True,
    )

    c2.button(
        "Temizle",
        on_click=select_all,
        args=(False,),
        use_container_width=True,
    )

    # --------------------------------------------------------
    # GRUP SEÇİMİ
    # --------------------------------------------------------

    selected_groups: set[str] = set()

    for group_name in ALL_GROUPS:

        count = int(
            (
                pharmacies["Grup"]
                == group_name
            ).sum()
        )

        checked = st.sidebar.checkbox(
            f"{group_name} ({count})",
            key=f"filter_{group_name}",
        )

        if checked:
            selected_groups.add(
                group_name
            )

    # --------------------------------------------------------
    # İSTATİSTİKLER
    # --------------------------------------------------------

    st.sidebar.divider()

    st.sidebar.metric(
        "Toplam tekil eczane",
        len(pharmacies),
    )

    st.sidebar.metric(
        "Grubu eşleşen",
        int(
            pharmacies[
                "Grup"
            ].notna().sum()
        ),
    )

    # --------------------------------------------------------
    # GRUPSUZ ECZANELER
    # --------------------------------------------------------

    missing = (
        pharmacies[
            pharmacies["Grup"].isna()
        ]["Eczane"]
        .astype(str)
        .tolist()
    )

    if missing:

        with st.sidebar.expander(
            f"Grupsuz ({len(missing)})"
        ):

            st.write(
                ", ".join(missing)
            )

    # ========================================================
    # ANA EKRAN
    # ========================================================

    st.title(
        "Denizli Eczane Grup Haritası"
    )

    st.caption(
        "9 ana grup · 27 alt grup · aynı alt gruptaki eczaneler "
        "kesikli çizgilerle birbirine bağlanır · "
        "kırmızı yoğunluk çemberi içindeki eczaneleri canlı sayar"
    )

    # --------------------------------------------------------
    # HARİTA
    # --------------------------------------------------------

    m = build_map(
        pharmacies,
        selected_groups,
    )

    components.html(
        m.get_root().render(),
        height=900,
        scrolling=False,
    )


except Exception as exc:

    st.exception(exc)
