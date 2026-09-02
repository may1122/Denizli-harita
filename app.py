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
    page_title="Denizli Eczane Haritası",
    page_icon="💊",
    layout="wide",
)

BASE_DIR = Path(__file__).resolve().parent
ECZANE_FILE_NAME = "denizli_eczaneler.xlsx"


# ============================================================
# ŞİMDİLİK GRUP A + B + C AKTİF
# A1-A3, B1-B3 ve C1-C3 kendi içinde bağlanır.
# Diğer tüm eczaneler "DİĞER" olarak düz gösterilir.
# ============================================================

GROUPS: dict[str, list[str]] = {
    "A1": [
        "KEKİK",
        "ALBAYRAK",
        "CEREN FİLİZER",
        "CAN SUYU",
        "IRMAK",
        "OZAN",
        "ÖZGÜ",
    ],
    "A2": [
        "CEYDA POLAT",
        "AYNUR GÜLER",
        "YENİLMEZ",
        "KUNDAKÇI",
        "ERMAN",
        "DERYAM",
        "KAYDIHAN",
        "GENCER",
    ],
    "A3": [
        "CADDE SAĞLIK",
        "İZMİRLİ",
        "GÜNEŞ",
        "TÜFEKÇİOĞLU",
        "ÖZGÜN KIYAT",
        "ALTINOVA",
        "BAŞÇAVUŞ",
        "SAHRA",
    ],

    "B1": [
        "ÇİFTÇİ",
        "ÖZSOY",
        "DEMİR",
        "BERGAMA",
        "SEMİH",
        "UĞUR",
        "FLORYA",
        "ÇAKMAK",
    ],
    "B2": [
        "KAPLAN",
        "NÜKHET",
        "HACETTEPE",
        "ADALET",
        "ÖZGEN",
        "EMİR",
        "ADA",
    ],
    "B3": [
        "PAMUKKALE AKTÜRK",
        "EKİZ",
        "NİLGÜN",
        "PAPATYA",
        "CANSU",
        "GÜLRİZ",
        "HAZAR",
        "TURAN",
        "KÖKNAR",
    ],

    "C1": [
        "MERKEZEFENDİ",
        "SEDA BAŞGİL",
        "İLKE",
        "GÖZDE GÜNDÜZ",
        "SENA KELLECİ",
        "AKTÜRK",
        "SEDA BAŞDİL",
        "MEHMET KAYA",
        "ÜMİT",
        "ERTUĞRUL",
    ],
    "C2": [
        "DENİZİM",
        "UZMAN",
        "29_EKİM",
        "TURUNÇ",
        "GÜLERYÜZ",
        "AYFER CEYLAN",
        "KİRAZ",
        "ASLI",
        "UMAY",
    ],
    "C3": [
        "IŞIMLIK",
        "SİNEM",
        "ÖZGÜR",
        "CEMRE",
    ],
}

GROUP_COLORS = {
    "A1": "#4A148C",
    "A2": "#7B1FA2",
    "A3": "#AB47BC",
    "B1": "#0D47A1",
    "B2": "#1976D2",
    "B3": "#64B5F6",
    "C1": "#1B5E20",
    "C2": "#388E3C",
    "C3": "#81C784",
}

OTHER_COLOR = "#90A4AE"


# ============================================================
# ECZANE İSMİ NORMALİZASYONU
# ============================================================

def normalize_name(value: object) -> str:
    if pd.isna(value):
        return ""

    text = unicodedata.normalize("NFKC", str(value)).strip().upper()

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

    text = re.sub(r"\s+", "", text)
    text = re.sub(r"[^0-9A-Z]", "", text)

    if text.endswith("ECZANESI"):
        text = text[:-8]

    return text


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
            "Eksik sütun(lar): " + ", ".join(missing)
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

    df = df.dropna(
        subset=[
            "Latitude",
            "Longitude",
        ]
    )

    df = df[
        df["Latitude"].between(-90, 90)
        & df["Longitude"].between(-180, 180)
    ]

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

    lat1, lon1 = map(math.radians, a)
    lat2, lon2 = map(math.radians, b)

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
# ÜST ÜSTE BİNEN NOKTALARI GÖRSEL OLARAK AYIR
# Gerçek koordinatlar değişmez.
# ============================================================

def add_display_coordinates(
    df: pd.DataFrame,
    overlap_threshold_m: float = 3.0,
    spread_radius_m: float = 8.0,
) -> pd.DataFrame:

    work = df.copy().reset_index(drop=True)

    work["DisplayLatitude"] = work["Latitude"].astype(float)
    work["DisplayLongitude"] = work["Longitude"].astype(float)

    if len(work) < 2:
        return work

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

        center_lat = (
            sum(float(work.at[i, "Latitude"]) for i in cluster)
            / len(cluster)
        )

        center_lon = (
            sum(float(work.at[i, "Longitude"]) for i in cluster)
            / len(cluster)
        )

        lat_deg_per_m = 1.0 / 111_320.0

        cos_lat = max(
            0.2,
            math.cos(math.radians(center_lat)),
        )

        lon_deg_per_m = (
            1.0
            / (111_320.0 * cos_lat)
        )

        radius_m = (
            spread_radius_m
            + max(0, len(cluster) - 2) * 1.25
        )

        for pos, idx in enumerate(cluster):
            angle = (
                2.0
                * math.pi
                * pos
                / len(cluster)
            ) - (math.pi / 2.0)

            north_m = math.cos(angle) * radius_m
            east_m = math.sin(angle) * radius_m

            work.at[idx, "DisplayLatitude"] = (
                center_lat
                + north_m * lat_deg_per_m
            )

            work.at[idx, "DisplayLongitude"] = (
                center_lon
                + east_m * lon_deg_per_m
            )

    return work



# ============================================================
# A + B + C GRUBU EŞLEME
# ============================================================

def build_group_map() -> dict[str, str]:
    result: dict[str, str] = {}

    for group_name, names in GROUPS.items():
        for name in names:
            key = normalize_name(name)

            if key in result and result[key] != group_name:
                raise ValueError(
                    f"{name} birden fazla alt grupta tanımlanmış: "
                    f"{result[key]} ve {group_name}"
                )

            result[key] = group_name

    return result


# ============================================================
# MINIMUM SPANNING TREE
# Aynı A alt grubundaki eczaneleri en kısa ağ ile bağlar.
# ============================================================

def minimum_spanning_edges(
    points: list[tuple[float, float]],
) -> list[tuple[int, int]]:

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

                distance = latlon_distance_m(
                    points[i],
                    points[j],
                )

                if best is None or distance < best[0]:
                    best = (distance, i, j)

        assert best is not None

        _, i, j = best
        edges.append((i, j))
        used.add(j)

    return edges


# ============================================================
# A1-A3 / B1-B3 / C1-C3 BAĞLANTI ÇİZGİLERİ
# ============================================================

def add_group_lines(
    map_obj: folium.Map,
    df: pd.DataFrame,
) -> None:

    for group_name in ("A1", "A2", "A3", "B1", "B2", "B3", "C1", "C2", "C3"):

        subset = (
            df.loc[
                df["Grup"].eq(group_name),
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

        real_points = [
            (float(row.Latitude), float(row.Longitude))
            for row in subset.itertuples(index=False)
        ]

        draw_points = [
            (float(row.DisplayLatitude), float(row.DisplayLongitude))
            for row in subset.itertuples(index=False)
        ]

        edges = minimum_spanning_edges(real_points)

        line_layer = FeatureGroup(
            name=f"{group_name} bağlantıları",
            show=True,
        )

        for i, j in edges:
            pharmacy_1 = str(subset.iloc[i]["Eczane"])
            pharmacy_2 = str(subset.iloc[j]["Eczane"])

            folium.PolyLine(
                locations=[
                    draw_points[i],
                    draw_points[j],
                ],
                color=GROUP_COLORS[group_name],
                weight=2.8,
                opacity=0.90,
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
# TÜM ECZANELERİ TEK TİP MARKERLA GÖSTER
# ============================================================

def add_markers(
    map_obj: folium.Map,
    df: pd.DataFrame,
) -> None:

    # A1 / A2 / A3 ayrı renklerle,
    # diğer tüm eczaneler tek nötr renkle gösterilir.
    render_groups = ["A1", "A2", "A3", "B1", "B2", "B3", "C1", "C2", "C3", "DİĞER"]

    for group_name in render_groups:

        if group_name == "DİĞER":
            subset = df.loc[df["Grup"].eq("DİĞER")]
            color = OTHER_COLOR
            layer_name = "Diğer eczaneler"
        else:
            subset = df.loc[df["Grup"].eq(group_name)]
            color = GROUP_COLORS[group_name]
            layer_name = f"{group_name} eczaneleri"

        if subset.empty:
            continue

        marker_layer = FeatureGroup(
            name=layer_name,
            show=True,
        )

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
            )

            if group_name != "DİĞER":
                tooltip += (
                    "<br>"
                    f"<b>{html.escape(group_name)}</b>"
                )

            if address:
                tooltip += (
                    "<br>"
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
                radius=6.0 if group_name != "DİĞER" else 4.5,
                color="#FFFFFF",
                weight=1.6 if group_name != "DİĞER" else 1.0,
                fill=True,
                fill_color=color,
                fill_opacity=0.98 if group_name != "DİĞER" else 0.65,
                tooltip=folium.Tooltip(
                    tooltip,
                    sticky=True,
                    direction="top",
                ),
            ).add_to(marker_layer)

        marker_layer.add_to(map_obj)


# ============================================================
# YOĞUNLUK ÇEMBERİ
# TÜM ECZANELERİ SAYAR
# ============================================================

class DensityCircleControl(MacroElement):

    def __init__(
        self,
        pharmacy_points: list[dict[str, object]],
        center_lat: float,
        center_lon: float,
    ):
        super().__init__()

        self._name = "DensityCircleControl"

        import json

        pharmacies_json = json.dumps(
            pharmacy_points,
            ensure_ascii=False,
        )

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

    const centerHandle = L.marker(
        startLatLng,
        {
            draggable: true,
            icon: centerIcon,
            zIndexOffset: 2000,
            title: 'Çember merkezini sürükle'
        }
    ).addTo(map);


    const DensityControl = L.Control.extend({

        options: {
            position: 'topleft'
        },

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
                    Tüm eczaneler sayılıyor
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
            `;

            div.style.background =
                'rgba(255,255,255,.97)';

            div.style.border =
                '1px solid #d8d8d8';

            div.style.borderRadius =
                '10px';

            div.style.padding =
                '12px 14px';

            div.style.minWidth =
                '225px';

            div.style.boxShadow =
                '0 2px 8px rgba(0,0,0,.15)';

            div.style.fontFamily =
                'Arial, sans-serif';

            div.style.color =
                '#222';

            L.DomEvent.disableClickPropagation(div);
            L.DomEvent.disableScrollPropagation(div);

            return div;
        }
    });

    map.addControl(
        new DensityControl()
    );


    function distanceMeters(
        lat1,
        lon1,
        lat2,
        lon2
    ) {

        const R = 6371000;

        const toRad =
            d => d * Math.PI / 180;

        const dLat =
            toRad(lat2 - lat1);

        const dLon =
            toRad(lon2 - lon1);

        const a =
            Math.sin(dLat / 2) ** 2 +
            Math.cos(toRad(lat1)) *
            Math.cos(toRad(lat2)) *
            Math.sin(dLon / 2) ** 2;

        return (
            2
            * R
            * Math.asin(Math.sqrt(a))
        );
    }


    function updateDensity() {

        const center =
            densityCircle.getLatLng();

        const radius =
            densityCircle.getRadius();

        let count = 0;

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
            }

        });

        const radiusEl =
            document.getElementById(
                'ayca-radius-value'
            );

        const countEl =
            document.getElementById(
                'ayca-count-value'
            );

        const shareEl =
            document.getElementById(
                'ayca-share-value'
            );

        if (radiusEl) {
            radiusEl.textContent =
                Math.round(radius) + ' m';
        }

        if (countEl) {
            countEl.textContent =
                count + ' eczane';
        }

        if (shareEl) {
            shareEl.textContent =
                total
                ? (
                    (count / total) * 100
                )
                .toFixed(1)
                .replace('.', ',')
                + '%'
                : '0%';
        }
    }


    centerHandle.on(
        'drag',
        function(e) {

            densityCircle.setLatLng(
                e.target.getLatLng()
            );

            updateDensity();
        }
    );


    centerHandle.on(
        'dragend',
        function(e) {

            densityCircle.setLatLng(
                e.target.getLatLng()
            );

            updateDensity();
        }
    );


    setTimeout(
        function() {

            const slider =
                document.getElementById(
                    'ayca-radius-slider'
                );

            const centerButton =
                document.getElementById(
                    'ayca-center-mode'
                );

            const hint =
                document.getElementById(
                    'ayca-density-hint'
                );

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

        },
        0
    );

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
) -> None:

    if df.empty:
        return

    pharmacy_points = []

    for _, row in df.iterrows():
        pharmacy_points.append(
            {
                "name": str(row["Eczane"]),
                "lat": float(row["Latitude"]),
                "lon": float(row["Longitude"]),
            }
        )

    control = DensityCircleControl(
        pharmacy_points=pharmacy_points,
        center_lat=float(df["Latitude"].median()),
        center_lon=float(df["Longitude"].median()),
    )

    control.add_to(map_obj)


# ============================================================
# HARİTAYI OLUŞTUR
# ============================================================

def build_map(
    df: pd.DataFrame,
) -> folium.Map:

    center = [
        float(df["Latitude"].median()),
        float(df["Longitude"].median()),
    ]

    m = folium.Map(
        location=center,
        zoom_start=13,
        tiles=None,
        control_scale=True,
        prefer_canvas=False,
    )

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

    # Önce A1-A3 / B1-B3 / C1-C3 bağlantı çizgileri, sonra markerlar.
    add_group_lines(
        m,
        df,
    )

    add_markers(
        m,
        df,
    )

    add_density_circle(
        m,
        df,
    )

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

    if not df.empty:
        m.fit_bounds(
            [
                [
                    float(df["Latitude"].min()),
                    float(df["Longitude"].min()),
                ],
                [
                    float(df["Latitude"].max()),
                    float(df["Longitude"].max()),
                ],
            ],
            padding=(25, 25),
        )

    return m


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

    pharmacies = read_pharmacies(
        str(pharmacy_path),
        pharmacy_path.stat().st_mtime_ns,
    )

    pharmacies = add_display_coordinates(
        pharmacies
    )

    # --------------------------------------------------------
    # A + B + C GRUPLARINI EŞLEŞTİR
    # --------------------------------------------------------

    group_map = build_group_map()

    pharmacies["Grup"] = (
        pharmacies["Anahtar"]
        .map(group_map)
        .fillna("DİĞER")
    )

    # --------------------------------------------------------
    # SIDEBAR
    # Grup seçimi YOK.
    # Sadece toplam eczane sayısı gösterilir.
    # --------------------------------------------------------

    st.sidebar.header(
        "Denizli Eczaneleri"
    )

    st.sidebar.metric(
        "Toplam tekil eczane",
        len(pharmacies),
    )

    for group_name in ("A1", "A2", "A3", "B1", "B2", "B3", "C1", "C2", "C3"):
        st.sidebar.metric(
            group_name,
            int(pharmacies["Grup"].eq(group_name).sum()),
        )

    st.sidebar.metric(
        "Diğer",
        int(pharmacies["Grup"].eq("DİĞER").sum()),
    )

    st.sidebar.caption(
        "Şimdilik A1-A3, B1-B3 ve C1-C3 aktif gruplandırılmıştır. "
        "Diğer eczaneler düz/nötr nokta olarak gösterilir."
    )

    # --------------------------------------------------------
    # ANA EKRAN
    # --------------------------------------------------------

    st.title(
        "Denizli Eczane Haritası — Grup A + B + C"
    )

    st.caption(
        "A1-A3, B1-B3 ve C1-C3 kendi içlerinde kesikli çizgilerle bağlanır · "
        "diğer eczaneler şimdilik düz/nötr nokta olarak kalır · "
        "kırmızı yoğunluk çemberi tüm eczaneleri sayar"
    )

    m = build_map(
        pharmacies
    )

    components.html(
        m.get_root().render(),
        height=900,
        scrolling=False,
    )


except Exception as exc:

    st.exception(exc)
