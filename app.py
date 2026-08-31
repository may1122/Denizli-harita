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

st.set_page_config(
    page_title="Denizli Eczane Grup Haritası",
    page_icon="💊",
    layout="wide",
)

BASE_DIR = Path(__file__).resolve().parent
ECZANE_FILE_NAME = "denizli_eczaneler.xlsx"

# 9 grup için birbirinden net ayrılan renkler
GROUP_COLORS = {
    "Grup 1": "#6A1B9A",  # mor
    "Grup 2": "#1565C0",  # mavi
    "Grup 3": "#00897B",  # turkuaz
    "Grup 4": "#EF6C00",  # turuncu
    "Grup 5": "#2E7D32",  # yeşil
    "Grup 6": "#C62828",  # kırmızı
    "Grup 7": "#7B1FA2",  # eflatun
    "Grup 8": "#42A5F5",  # açık mavi
    "Grup 9": "#F9A825",  # sarı/kehribar
}

# Kullanıcının tercih ettiği 9'lu Denizli taslağı.
# Mükerrer yazım varyantlarında tek kanonik isim tutulur; koordinat dosyası da normalize edilerek tekilleştirilir.
GROUPS: dict[str, list[str]] = {
    "Grup 1": [
        "AYNUR GÜLER", "AYŞEN", "BAKLAN", "BAŞDİL", "BAŞÇAVUŞ", "BÜŞRA BOYACI",
        "CEYDA POLAT", "DELİKTAŞ", "DEMİRTAŞ", "DERYAM", "DOKUZKAVAKLAR", "ERMAN",
        "GAMZE", "GENCER", "KAYDIHAN", "KOÇAK", "KUNDAKÇI", "MAVİ", "MİRA", "NEVA",
        "SEMT", "YENİLMEZ", "ZEYNEP", "ÇALLIOĞLU", "DİLEK",
    ],
    "Grup 2": [
        "ANIL", "AYGÖREN", "AYKUT", "BAYRAMYERİ", "BÜYÜK", "DAĞDEVİREN", "DOĞAL", "DUYGU",
        "EGE", "ELİF", "ERCAN", "ERDEM", "ESİN", "NEŞE", "SAĞLIK", "SERGEN", "SEVİL",
        "SEÇKİN", "TEMMUZ", "ÖZDERMAN", "İNCEOĞLU", "ŞULE", "ŞİFA",
    ],
    "Grup 3": [
        "ADALET", "AKDENİZ", "CANDENİZ", "CANSU", "DEMİR", "EMİR", "FLORYA", "GÜLRİZ",
        "HACETTEPE", "HAZAR", "KAPLAN", "KAYHAN", "KÖKNAR", "NÜKHET", "NİLGÜN", "PAPATYA",
        "SEMİH", "TURAN", "UĞUR", "ÇAKMAK", "ÇİFTÇİ", "ÖZGEN", "ÖZSOY",
    ],
    "Grup 4": [
        "ADALI", "ARCA", "BAYRAMOĞLU", "BURCU", "CADDE", "CANAN", "DEMİRAY", "DENİZLİ", "EFE",
        "EZO", "GÖKHAN", "GÖKSU", "GÜLAY", "GÜLEÇ", "GÜNGÖR", "IŞIL", "KIVILCIM", "LOKMAN",
        "LİMONCU", "MERVE", "MORALIOĞLU", "NUR BAŞÇAVUŞ", "VERESELİ DENİZLİ", "ÇAKMAKLIOĞLU", "ÇOMUT",
    ],
    "Grup 5": [
        "ADA", "AKKAYA", "ANAFARTALAR", "ASMALI", "AYFER CEYLAN", "CEYHAN", "DENİZİM", "DİNÇ",
        "DİŞÇİOĞLU", "EZGİ", "FATIMA ŞENTÜRK", "FATİH", "GÜRKAN", "HASİBE KARTOĞLAN", "MERKEZ",
        "OCAK", "PELİTLİBAĞ", "SEVİM", "TOLGAY", "TURUNÇ", "UZMAN", "ÇETİNKAYA", "ÖZCEL", "ÖZNUR", "GÜLERYÜZ",
    ],
    "Grup 6": [
        "ALBAYRAK", "ASLAN", "AYDIN", "AYLİN", "ELİF PAMUKÇU", "ERSAN", "HÜRRİYET", "KINIKLI",
        "KÖSELER", "MEHMET KAYA", "MERKEZEFENDİ", "MUTLU GÜNLER", "NİSAN", "OKYANUS",
        "SAHRA", "SARAÇOĞLU", "SEDA BAŞDİL", "TURKUAZ", "YEŞİLYURT", "ÇAMLICA", "ÇETİN", "ÖZKAN", "İSTİKLAL",
    ],
    "Grup 7": [
        "29_EKİM", "AKTÜRK", "ASLI", "BERGAMA", "CEMRE", "CEREN FİLİZER", "EKİZ", "ELVAN", "ERTUĞRUL",
        "GÖKKUŞAĞI", "GÖZDE GÜNDÜZ", "IŞIMLIK", "KABAYUKA", "KEKİK", "KİRAZ",
        "PAMUKKALE AKTÜRK", "SENA KELLECİ", "SİNEM", "UMAY", "ZEYNEP SULTAN", "ÖZGÜR", "ÜMİT", "İLKE",
    ],
    "Grup 8": [
        "ALPLER", "ALSANCAK", "BİLGE", "CANDAN", "DEMİRCİOĞLU GÜL", "DEMİROĞLU", "DERMAN", "DEVECİ",
        "EVREN", "EZGİ KIRDI", "FORUM ÇAMLIK", "GÖKÇE", "GÜRSOY", "KIZILTAŞ", "MERVE YAMUÇ",
        "PAMUKKALE", "SAYGIN", "SOYLU", "SU", "TUBA", "ZEYTİNKÖY SEMA", "ÜNİVERSİTE", "İNANÖZ",
    ],
    "Grup 9": [
        "ALTINOVA", "BAHAR", "CADDE SAĞLIK", "CANSU ERKİLET", "CANSUYU", "CEYLAN", "ELİF'İN", "EMEK",
        "GÖRKEM", "GÜNEŞ", "IRMAK", "NAZAN", "OZAN", "PARK BOTANİK", "SERVET", "TUGAY", "TÜFEKÇİOĞLU",
        "YEŞİLYUVA", "ÇAMLIK", "ÖZGÜ", "ÖZGÜN KIYAT", "İZMİRLİ", "ŞİRİN", "NEFES",
    ],
}

ALL_GROUPS = list(GROUPS.keys())


def normalize_name(value: object) -> str:
    if pd.isna(value):
        return ""
    text = unicodedata.normalize("NFKC", str(value)).strip().upper()
    text = text.translate(str.maketrans({
        "Ç": "C", "Ğ": "G", "İ": "I", "Ö": "O", "Ş": "S", "Ü": "U"
    }))
    text = re.sub(r"\s+", "", text)
    text = re.sub(r"[^0-9A-Z]", "", text)
    if text.endswith("ECZANESI"):
        text = text[:-8]
    return text


def build_group_map() -> dict[str, str]:
    result: dict[str, str] = {}
    for group_name, names in GROUPS.items():
        for name in names:
            key = normalize_name(name)
            if key in result and result[key] != group_name:
                raise ValueError(f"{name} birden fazla grupta tanımlanmış.")
            result[key] = group_name
    return result


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

    # Kaynaktaki 3778015 -> 37.78015, 2909645 -> 29.09645 biçimi
    if abs(v) > 180:
        v = v / 100000.0
    return v


@st.cache_data(show_spinner=False)
def read_pharmacies(path: str, file_version: int) -> pd.DataFrame:
    del file_version
    raw = pd.read_excel(path, engine="openpyxl")

    required = ["Eczane İsmi", "Enlem (Latitude)", "Boylam (Longitude)"]
    missing = [c for c in required if c not in raw.columns]
    if missing:
        raise ValueError("Eksik sütun(lar): " + ", ".join(missing))

    df = pd.DataFrame({
        "Eczane": raw["Eczane İsmi"].astype(str).str.strip(),
        "Adres": raw.get("Eczane Adresi", "").fillna("").astype(str).str.strip(),
        "Latitude": raw["Enlem (Latitude)"].map(parse_coord),
        "Longitude": raw["Boylam (Longitude)"].map(parse_coord),
    })
    df["Anahtar"] = df["Eczane"].map(normalize_name)
    df = df.dropna(subset=["Latitude", "Longitude"])
    df = df[(df["Latitude"].between(-90, 90)) & (df["Longitude"].between(-180, 180))]

    # Yazım varyantı mükerrerlerini tek kayıt yap.
    df = df.drop_duplicates(subset=["Anahtar"], keep="first").reset_index(drop=True)
    return df


def latlon_distance_m(a: tuple[float, float], b: tuple[float, float]) -> float:
    lat1, lon1 = map(math.radians, a)
    lat2, lon2 = map(math.radians, b)
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    h = math.sin(dlat / 2) ** 2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon / 2) ** 2
    return 2 * 6371000.0 * math.asin(math.sqrt(h))


def minimum_spanning_edges(points: list[tuple[float, float]]) -> list[tuple[int, int]]:
    if len(points) < 2:
        return []
    used = {0}
    edges: list[tuple[int, int]] = []
    while len(used) < len(points):
        best: tuple[float, int, int] | None = None
        for i in used:
            for j in range(len(points)):
                if j in used:
                    continue
                d = latlon_distance_m(points[i], points[j])
                if best is None or d < best[0]:
                    best = (d, i, j)
        assert best is not None
        _, i, j = best
        edges.append((i, j))
        used.add(j)
    return edges


def add_group_lines(map_obj: folium.Map, df: pd.DataFrame, selected_groups: set[str]) -> None:
    for group_name in ALL_GROUPS:
        if group_name not in selected_groups:
            continue
        subset = df[df["Grup"] == group_name].reset_index(drop=True)
        if len(subset) < 2:
            continue

        layer = FeatureGroup(name=f"{group_name} bağlantıları", show=True)
        color = GROUP_COLORS[group_name]
        points = [(float(r.Latitude), float(r.Longitude)) for r in subset.itertuples()]

        for i, j in minimum_spanning_edges(points):
            folium.PolyLine(
                [points[i], points[j]],
                color=color,
                weight=3.2,
                opacity=0.88,
                dash_array="7,6",
                line_cap="round",
                line_join="round",
                tooltip=f"{group_name} bağlantısı",
            ).add_to(layer)
        layer.add_to(map_obj)


def add_markers(map_obj: folium.Map, df: pd.DataFrame, selected_groups: set[str]) -> None:
    for group_name in ALL_GROUPS:
        if group_name not in selected_groups:
            continue
        subset = df[df["Grup"] == group_name]
        if subset.empty:
            continue

        layer = FeatureGroup(name=f"{group_name} eczaneleri", show=True)
        color = GROUP_COLORS[group_name]
        for _, row in subset.iterrows():
            address = html.escape(str(row["Adres"]))
            tooltip = (
                '<div style="font-size:13px;line-height:1.35">'
                f'<b>{html.escape(str(row["Eczane"]))}</b><br>'
                f'{html.escape(group_name)}'
                + (f'<br><span style="color:#666">{address}</span>' if address else '')
                + '</div>'
            )
            folium.CircleMarker(
                location=[float(row["Latitude"]), float(row["Longitude"])],
                radius=6.0,
                color="#FFFFFF",
                weight=1.6,
                fill=True,
                fill_color=color,
                fill_opacity=0.98,
                tooltip=folium.Tooltip(tooltip, sticky=True, direction="top"),
            ).add_to(layer)
        layer.add_to(map_obj)


def build_map(df: pd.DataFrame, selected_groups: set[str]) -> folium.Map:
    center = [float(df["Latitude"].median()), float(df["Longitude"].median())]
    m = folium.Map(location=center, zoom_start=13, tiles=None, control_scale=True, prefer_canvas=True)

    folium.TileLayer(
        tiles="https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png",
        attr="&copy; OpenStreetMap contributors &copy; CARTO",
        name="Sade harita",
        control=True,
        show=True,
        subdomains="abcd",
        max_zoom=20,
    ).add_to(m)

    folium.TileLayer(
        tiles="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png",
        attr="&copy; OpenStreetMap contributors",
        name="Detaylı harita",
        control=True,
        show=False,
        max_zoom=19,
    ).add_to(m)

    # Çizgiler önce, markerlar sonra: markerlar çizgilerin üstünde kalır.
    add_group_lines(m, df, selected_groups)
    add_markers(m, df, selected_groups)

    Fullscreen(position="topright", title="Tam ekran", title_cancel="Tam ekrandan çık").add_to(m)
    MeasureControl(position="topright", primary_length_unit="meters").add_to(m)
    folium.LayerControl(collapsed=True, position="topright").add_to(m)

    selected_df = df[df["Grup"].isin(selected_groups)]
    bounds_df = selected_df if not selected_df.empty else df
    m.fit_bounds(
        [
            [float(bounds_df["Latitude"].min()), float(bounds_df["Longitude"].min())],
            [float(bounds_df["Latitude"].max()), float(bounds_df["Longitude"].max())],
        ],
        padding=(25, 25),
    )
    return m


def init_state() -> None:
    for group_name in ALL_GROUPS:
        st.session_state.setdefault(f"filter_{group_name}", True)


def select_all(value: bool) -> None:
    for group_name in ALL_GROUPS:
        st.session_state[f"filter_{group_name}"] = value


pharmacy_path = BASE_DIR / ECZANE_FILE_NAME
if not pharmacy_path.exists():
    st.error(f"{ECZANE_FILE_NAME} bulunamadı. Dosyayı app.py ile aynı klasöre koyun.")
    st.stop()

try:
    pharmacies = read_pharmacies(str(pharmacy_path), pharmacy_path.stat().st_mtime_ns)
    pharmacies["Grup"] = pharmacies["Anahtar"].map(build_group_map())

    init_state()

    st.sidebar.header("Denizli 9 Grup")
    c1, c2 = st.sidebar.columns(2)
    c1.button("Tümünü Aç", on_click=select_all, args=(True,), use_container_width=True)
    c2.button("Temizle", on_click=select_all, args=(False,), use_container_width=True)

    selected_groups: set[str] = set()
    for group_name in ALL_GROUPS:
        count = int((pharmacies["Grup"] == group_name).sum())
        if st.sidebar.checkbox(
            f"{group_name} ({count})",
            key=f"filter_{group_name}",
        ):
            selected_groups.add(group_name)

    st.sidebar.divider()
    st.sidebar.metric("Toplam tekil eczane", len(pharmacies))
    st.sidebar.metric("Grubu eşleşen", int(pharmacies["Grup"].notna().sum()))

    missing = pharmacies[pharmacies["Grup"].isna()]["Eczane"].astype(str).tolist()
    if missing:
        with st.sidebar.expander(f"Grupsuz ({len(missing)})"):
            st.write(", ".join(missing))

    st.title("Denizli Eczane Grup Haritası")
    st.caption("9 grup · aynı gruptaki eczaneler kesik çizgilerle birbirine bağlanır")

    m = build_map(pharmacies, selected_groups)
    components.html(m.get_root().render(), height=900, scrolling=False)

except Exception as exc:
    st.exception(exc)
