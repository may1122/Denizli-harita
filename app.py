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
    "Grup 1": "#6A1B9A",  # mor
    "Grup 2": "#1565C0",  # mavi
    "Grup 3": "#00897B",  # turkuaz
    "Grup 4": "#EF6C00",  # turuncu
    "Grup 5": "#2E7D32",  # yeşil
    "Grup 6": "#C62828",  # kırmızı
    "Grup 7": "#7B1FA2",  # eflatun
    "Grup 8": "#42A5F5",  # açık mavi
    "Grup 9": "#F9A825",  # sarı / kehribar
}


# ============================================================
# DENİZLİ 9 GRUP
# ============================================================

GROUPS: dict[str, list[str]] = {

    "Grup 1": [
        "AYNUR GÜLER",
        "AYŞEN",
        "BAKLAN",
        "BAŞDİL",
        "BAŞÇAVUŞ",
        "BÜŞRA BOYACI",
        "CEYDA POLAT",
        "DELİKTAŞ",
        "DEMİRTAŞ",
        "DERYAM",
        "DOKUZKAVAKLAR",
        "ERMAN",
        "GAMZE",
        "GENCER",
        "KAYDIHAN",
        "KOÇAK",
        "KUNDAKÇI",
        "MAVİ",
        "MİRA",
        "NEVA",
        "SEMT",
        "YENİLMEZ",
        "ZEYNEP",
        "ÇALLIOĞLU",
        "DİLEK",
    ],

    "Grup 2": [
        "ANIL",
        "AYGÖREN",
        "AYKUT",
        "BAYRAMYERİ",
        "BÜYÜK",
        "DAĞDEVİREN",
        "DOĞAL",
        "DUYGU",
        "EGE",
        "ELİF",
        "ERCAN",
        "ERDEM",
        "ESİN",
        "NEŞE",
        "SAĞLIK",
        "SERGEN",
        "SEVİL",
        "SEÇKİN",
        "TEMMUZ",
        "ÖZDERMAN",
        "İNCEOĞLU",
        "ŞULE",
        "ŞİFA",
    ],

    "Grup 3": [
        "ADALET",
        "AKDENİZ",
        "CANDENİZ",
        "CANSU",
        "DEMİR",
        "EMİR",
        "FLORYA",
        "GÜLRİZ",
        "HACETTEPE",
        "HAZAR",
        "KAPLAN",
        "KAYHAN",
        "KÖKNAR",
        "NÜKHET",
        "NİLGÜN",
        "PAPATYA",
        "SEMİH",
        "TURAN",
        "UĞUR",
        "ÇAKMAK",
        "ÇİFTÇİ",
        "ÖZGEN",
        "ÖZSOY",
    ],

    "Grup 4": [
        "ADALI",
        "ARCA",
        "BAYRAMOĞLU",
        "BURCU",
        "CADDE",
        "CANAN",
        "DEMİRAY",
        "DENİZLİ",
        "EFE",
        "EZO",
        "GÖKHAN",
        "GÖKSU",
        "GÜLAY",
        "GÜLEÇ",
        "GÜNGÖR",
        "IŞIL",
        "KIVILCIM",
        "LOKMAN",
        "LİMONCU",
        "MERVE",
        "MORALIOĞLU",
        "NUR BAŞÇAVUŞ",
        "VERESELİ DENİZLİ",
        "ÇAKMAKLIOĞLU",
        "ÇOMUT",
    ],

    "Grup 5": [
        "ADA",
        "AKKAYA",
        "ANAFARTALAR",
        "ASMALI",
        "AYFER CEYLAN",
        "CEYHAN",
        "DENİZİM",
        "DİNÇ",
        "DİŞÇİOĞLU",
        "EZGİ",
        "FATIMA ŞENTÜRK",
        "FATİH",
        "GÜRKAN",
        "HASİBE KARTOĞLAN",
        "MERKEZ",
        "OCAK",
        "PELİTLİBAĞ",
        "SEVİM",
        "TOLGAY",
        "TURUNÇ",
        "UZMAN",
        "ÇETİNKAYA",
        "ÖZCEL",
        "ÖZNUR",
        "GÜLERYÜZ",
    ],

    "Grup 6": [
        "ALBAYRAK",
        "ASLAN",
        "AYDIN",
        "AYLİN",
        "ELİF PAMUKÇU",
        "ERSAN",
        "HÜRRİYET",
        "KINIKLI",
        "KÖSELER",
        "MEHMET KAYA",
        "MERKEZEFENDİ",
        "MUTLU GÜNLER",
        "NİSAN",
        "OKYANUS",
        "SAHRA",
        "SARAÇOĞLU",
        "SEDA BAŞDİL",
        "TURKUAZ",
        "YEŞİLYURT",
        "ÇAMLICA",
        "ÇETİN",
        "ÖZKAN",
        "İSTİKLAL",
    ],

    "Grup 7": [
        "29_EKİM",
        "AKTÜRK",
        "ASLI",
        "BERGAMA",
        "CEMRE",
        "CEREN FİLİZER",
        "EKİZ",
        "ELVAN",
        "ERTUĞRUL",
        "GÖKKUŞAĞI",
        "GÖZDE GÜNDÜZ",
        "IŞIMLIK",
        "KABAYUKA",
        "KEKİK",
        "KİRAZ",
        "PAMUKKALE AKTÜRK",
        "SENA KELLECİ",
        "SİNEM",
        "UMAY",
        "ZEYNEP SULTAN",
        "ÖZGÜR",
        "ÜMİT",
        "İLKE",
    ],

    "Grup 8": [
        "ALPLER",
        "ALSANCAK",
        "BİLGE",
        "CANDAN",
        "DEMİRCİOĞLU GÜL",
        "DEMİROĞLU",
        "DERMAN",
        "DEVECİ",
        "EVREN",
        "EZGİ KIRDI",
        "FORUM ÇAMLIK",
        "GÖKÇE",
        "GÜRSOY",
        "KIZILTAŞ",
        "MERVE YAMUÇ",
        "PAMUKKALE",
        "SAYGIN",
        "SOYLU",
        "SU",
        "TUBA",
        "ZEYTİNKÖY SEMA",
        "ÜNİVERSİTE",
        "İNANÖZ",
    ],

    "Grup 9": [
        "ALTINOVA",
        "BAHAR",
        "CADDE SAĞLIK",
        "CANSU ERKİLET",
        "CANSUYU",
        "CEYLAN",
        "ELİF'İN",
        "EMEK",
        "GÖRKEM",
        "GÜNEŞ",
        "IRMAK",
        "NAZAN",
        "OZAN",
        "PARK BOTANİK",
        "SERVET",
        "TUGAY",
        "TÜFEKÇİOĞLU",
        "YEŞİLYUVA",
        "ÇAMLIK",
        "ÖZGÜ",
        "ÖZGÜN KIYAT",
        "İZMİRLİ",
        "ŞİRİN",
        "NEFES",
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

def local_polygon_edges(
    points: list[tuple[float, float]],
    cluster_size: int = 4,
) -> list[tuple[int, int]]:
    """
    Noktaları birbirine tamamen bağlamak yerine coğrafi olarak yakın
    küçük kümelere ayırır. Her 4'lü kümede yalnızca kapalı çevre çizilir:

        A ---- B
        |      |
        D ---- C

    Böylece dörtgen/yamuk benzeri yerel şekiller oluşur ve farklı kümeler
    arasında gereksiz uzun bağlantılar kurulmaz. Kalan 3 nokta üçgen,
    kalan 2 nokta tek çizgi olarak gösterilir.
    """

    n = len(points)
    if n < 2:
        return []

    remaining = set(range(n))
    clusters: list[list[int]] = []

    while remaining:
        seed = min(remaining)
        remaining.remove(seed)
        cluster = [seed]

        if remaining:
            nearest = sorted(
                remaining,
                key=lambda j: latlon_distance_m(points[seed], points[j]),
            )
            take = nearest[: max(0, cluster_size - 1)]
            for j in take:
                remaining.remove(j)
                cluster.append(j)

        clusters.append(cluster)

    # Son kümede tek nokta kalırsa onu en yakın önceki kümeye ekle.
    if len(clusters) >= 2 and len(clusters[-1]) == 1:
        lone = clusters.pop()[0]
        best_idx = min(
            range(len(clusters)),
            key=lambda ci: min(
                latlon_distance_m(points[lone], points[j])
                for j in clusters[ci]
            ),
        )
        clusters[best_idx].append(lone)

    edge_set: set[tuple[int, int]] = set()

    for cluster in clusters:
        if len(cluster) == 2:
            a, b = sorted(cluster)
            edge_set.add((a, b))
            continue

        if len(cluster) < 2:
            continue

        # Küme merkezine göre açı sırasına diz; sadece dış çevreyi bağla.
        center_lat = sum(points[i][0] for i in cluster) / len(cluster)
        center_lon = sum(points[i][1] for i in cluster) / len(cluster)

        ordered = sorted(
            cluster,
            key=lambda i: math.atan2(
                points[i][0] - center_lat,
                points[i][1] - center_lon,
            ),
        )

        for k in range(len(ordered)):
            i = ordered[k]
            j = ordered[(k + 1) % len(ordered)]
            a, b = sorted((i, j))
            edge_set.add((a, b))

    return sorted(edge_set)


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
    tanımından yeniden kurar. Çizgiler tek bir ağaç gibi gitmez;
    yakın komşular arasında küçük ağlar ve kapalı şekiller oluşturur.

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

        points = [
            (
                float(row.Latitude),
                float(row.Longitude),
            )
            for row in subset.itertuples(index=False)
        ]

        # Noktaları 4'lü yerel kümelere ayırıp sadece kümenin çevresini çiziyoruz.
        # Böylece bütün grup tek bir ağ gibi birbirine bağlanmaz.
        edges = local_polygon_edges(points, cluster_size=4)

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
                    points[i],
                    points[j],
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
                    float(row["Latitude"]),
                    float(row["Longitude"]),
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
    # SADE HARİTA
    # --------------------------------------------------------

    folium.TileLayer(
        tiles=(
            "https://{s}.basemaps.cartocdn.com/"
            "light_all/{z}/{x}/{y}{r}.png"
        ),
        attr=(
            "&copy; OpenStreetMap contributors "
            "&copy; CARTO"
        ),
        name="Sade harita",
        control=True,
        show=True,
        subdomains="abcd",
        max_zoom=20,
    ).add_to(m)

    # --------------------------------------------------------
    # DETAYLI HARİTA
    # --------------------------------------------------------

    folium.TileLayer(
        tiles=(
            "https://{s}.tile.openstreetmap.org/"
            "{z}/{x}/{y}.png"
        ),
        attr="&copy; OpenStreetMap contributors",
        name="Detaylı harita",
        control=True,
        show=False,
        max_zoom=19,
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
    # SESSION STATE
    # --------------------------------------------------------

    init_state()

    # ========================================================
    # SIDEBAR
    # ========================================================

    st.sidebar.header(
        "Denizli 9 Grup"
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
        "9 grup · aynı gruptaki eczaneler "
        "kesikli çizgilerle birbirine bağlanır"
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
