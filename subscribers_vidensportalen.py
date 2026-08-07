import streamlit as st
import pandas as pd
import pydeck as pdk
import plotly.express as px
from pathlib import Path

st.set_page_config(page_title="Subscriber-overblik", page_icon="📍", layout="wide")

# --------------------------------------------------
# Konfiguration
# --------------------------------------------------
DEFAULT_FILE_CANDIDATES = [
    "Vidensportal_abonnenter_renset 1.xlsx",
    "Vidensportal_abonnenter_renset.xlsx",
]
DEFAULT_SHEET = "Renset_data"

# Koordinater er angivet som hovedlokation/by-center for den enkelte skole/institution.
# Justér gerne punkterne, hvis du vil have dem til at ramme en specifik adresse/campus.
SCHOOL_COORDS = {
    "Aarhus Tech": {"lat": 56.1629, "lon": 10.2039},
    "CELF": {"lat": 54.7691, "lon": 11.8743},  # Nykøbing Falster
    "College360": {"lat": 56.1697, "lon": 9.5451},  # Silkeborg
    "Den Jyske Håndværkerskole": {"lat": 56.3297, "lon": 10.0487},  # Hadsten
    "EUC Lillebælt": {"lat": 55.5377, "lon": 9.7120},  # Fredericia
    "EUC Nord": {"lat": 57.4407, "lon": 10.5366},     # Frederikshavn
    "EUC Nordvest": {"lat": 56.9567, "lon": 8.6947},  # Thisted
    "EUC Sjælland": {"lat": 55.4038, "lon": 11.3548}, # Næstved
    "EUC Syd": {"lat": 55.2494, "lon": 9.4870},       # Kolding/Haderslev-område (kan justeres)
    "HANSENBERG": {"lat": 55.4904, "lon": 9.4722},    # Kolding
    "HEG": {"lat": 56.8046, "lon": 9.5322},           # Aars/Hobro-området (kan justeres)
    "Herningsholm Erhvervsskole & Gymnasier": {"lat": 56.1362, "lon": 8.9766},
    "Learnmark Horsens": {"lat": 55.8607, "lon": 9.8503},
    "Mercantec": {"lat": 56.4532, "lon": 9.4020},
    "NEXT København": {"lat": 55.6561, "lon": 12.5187},
    "Nordvestsjællands Erhvervs- og Gymnasieuddannelser": {"lat": 55.6759, "lon": 11.0884},  # Kalundborg/Slagelse-området
    "Roskilde Tekniske Skole": {"lat": 55.6419, "lon": 12.0878},
    "Rybners": {"lat": 55.4765, "lon": 8.4594},
    "Skive College": {"lat": 56.5667, "lon": 9.0333},
    "Svendborg Erhvervsskole & Gymnasier": {"lat": 55.0598, "lon": 10.6068},
    "Syddansk Erhvervsskole": {"lat": 55.4038, "lon": 10.4024},  # Odense
    "TECHCOLLEGE": {"lat": 57.0488, "lon": 9.9217},
    "TEC": {"lat": 55.6811, "lon": 12.5273},
    "UCH Holstebro": {"lat": 56.3601, "lon": 8.6161},
    "U/NORD": {"lat": 55.9279, "lon": 12.3008},
    "UCRS": {"lat": 55.9189, "lon": 8.4908},          # Ringkøbing/Skjern-området
    "ZBC": {"lat": 55.2296, "lon": 11.7609},          # Næstved/Slagelse-området
}

TYPE_LABELS = {
    "Skole/institution (fra emaildomæne)": "Skole/institution",
    "Virksomhed/organisation/ukendt": "Virksomhed/organisation/ukendt",
}

# --------------------------------------------------
# Datainlæsning
# --------------------------------------------------
@st.cache_data
def load_data(uploaded_file=None):
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file, sheet_name=DEFAULT_SHEET, engine="openpyxl", dtype=str).fillna("")
        source_name = uploaded_file.name
        return df, source_name

    for candidate in DEFAULT_FILE_CANDIDATES:
        if Path(candidate).exists():
            df = pd.read_excel(candidate, sheet_name=DEFAULT_SHEET, engine="openpyxl", dtype=str).fillna("")
            return df, candidate

    raise FileNotFoundError(
        "Ingen standardfil fundet. Upload den rensede Excel-fil via sidebaren."
    )


# --------------------------------------------------
# Hjælpefunktioner
# --------------------------------------------------
def classify_org_type(row):
    if row.get("afsender_type") == "Skole/institution (fra emaildomæne)":
        return "Skole/institution"
    return "Virksomhed/organisation/ukendt"


def build_school_map_df(df):
    school_df = (
        df[df["subscriber_virksomhed_skole"].astype(str).str.strip() != ""]
        .assign(org_type=lambda x: x.apply(classify_org_type, axis=1))
        .groupby("subscriber_virksomhed_skole", as_index=False)
        .agg(antal_subscribers=("email_domain", "count"))
        .rename(columns={"subscriber_virksomhed_skole": "skole"})
    )

    coords_df = pd.DataFrame([
        {"skole": school, "lat": vals["lat"], "lon": vals["lon"]}
        for school, vals in SCHOOL_COORDS.items()
    ])

    school_df = school_df.merge(coords_df, on="skole", how="left")
    school_df["har_koordinater"] = school_df["lat"].notna()

    # Punktstørrelse med minimum/maximum så små skoler stadig er synlige
    if len(school_df) > 0:
        min_count = school_df["antal_subscribers"].min()
        max_count = school_df["antal_subscribers"].max()
        if max_count == min_count:
            school_df["radius"] = 12000
        else:
            school_df["radius"] = 7000 + (school_df["antal_subscribers"] - min_count) / (max_count - min_count) * 28000
    else:
        school_df["radius"] = []

    return school_df


def donut_chart(df):
    summary = (
        df.assign(type_simple=df.apply(classify_org_type, axis=1))
          .groupby("type_simple", as_index=False)
          .agg(antal=("email_domain", "count"))
    )

    fig = px.pie(
        summary,
        names="type_simple",
        values="antal",
        hole=0.55,
        color="type_simple",
        color_discrete_map={
            "Skole/institution": "#1f77b4",
            "Virksomhed/organisation/ukendt": "#ff7f0e",
        },
    )
    fig.update_traces(textposition="inside", textinfo="percent+label", hovertemplate="%{label}<br>Antal: %{value}<extra></extra>")
    fig.update_layout(margin=dict(l=10, r=10, t=10, b=10), showlegend=True)
    return fig


def school_bar_chart(school_df, top_n=15):
    plot_df = school_df.sort_values("antal_subscribers", ascending=False).head(top_n)
    fig = px.bar(
        plot_df,
        x="antal_subscribers",
        y="skole",
        orientation="h",
        text="antal_subscribers",
        color="antal_subscribers",
        color_continuous_scale="Blues",
    )
    fig.update_layout(yaxis_title="", xaxis_title="Antal subscribers", coloraxis_showscale=False, margin=dict(l=10, r=10, t=10, b=10))
    fig.update_traces(hovertemplate="%{y}<br>Antal: %{x}<extra></extra>")
    fig.update_yaxes(categoryorder="total ascending")
    return fig


def school_map(school_df):
    map_df = school_df[school_df["har_koordinater"]].copy()
    if map_df.empty:
        st.warning("Ingen skoler med koordinater. Tilføj dem i SCHOOL_COORDS i appen.")
        return

    tooltip = {
        "html": "<b>{skole}</b><br/>Antal subscribers: {antal_subscribers}",
        "style": {
            "backgroundColor": "rgba(30, 30, 30, 0.85)",
            "color": "white",
            "fontSize": "13px",
        },
    }

    view_state = pdk.ViewState(
        latitude=map_df["lat"].mean(),
        longitude=map_df["lon"].mean(),
        zoom=5.8,
        pitch=0,
    )

    layer = pdk.Layer(
        "ScatterplotLayer",
        data=map_df,
        get_position="[lon, lat]",
        get_fill_color="[31, 119, 180, 180]",
        get_line_color="[255, 255, 255, 200]",
        line_width_min_pixels=1,
        stroked=True,
        filled=True,
        pickable=True,
        get_radius="radius",
    )

    deck = pdk.Deck(
        layers=[layer],
        initial_view_state=view_state,
        tooltip=tooltip,
        map_style="light",
    )
    st.pydeck_chart(deck, use_container_width=True)


# --------------------------------------------------
# UI
# --------------------------------------------------
st.title("📍 Subscriber-overblik")
st.caption("Fordeling mellem skole/institution og virksomhed samt kort over repræsenterede skoler.")

with st.sidebar:
    st.header("Data")
    uploaded_file = st.file_uploader("Upload renset Excel-fil", type=["xlsx"])
    st.markdown("Hvis du ikke uploader en fil, forsøger appen at læse en standardfil i samme mappe.")

try:
    df, source_name = load_data(uploaded_file)
except Exception as e:
    st.error(str(e))
    st.stop()

# Grundlæggende validering
required_cols = ["email_domain", "subscriber_virksomhed_skole", "subscriber_virksomhed", "afsender_type"]
missing_cols = [c for c in required_cols if c not in df.columns]
if missing_cols:
    st.error(f"Følgende kolonner mangler i datasættet: {missing_cols}")
    st.stop()

# Filtre
with st.sidebar:
    st.header("Filtre")
    only_with_school = st.checkbox("Kun rækker med skole/institution udfyldt", value=False)
    search_term = st.text_input("Søg i skole eller virksomhed")

filtered_df = df.copy()
if only_with_school:
    filtered_df = filtered_df[filtered_df["subscriber_virksomhed_skole"].astype(str).str.strip() != ""]
if search_term:
    mask = (
        filtered_df["subscriber_virksomhed_skole"].astype(str).str.contains(search_term, case=False, na=False)
        | filtered_df["subscriber_virksomhed"].astype(str).str.contains(search_term, case=False, na=False)
    )
    filtered_df = filtered_df[mask]

school_df = build_school_map_df(filtered_df)

# KPI'er
k1, k2, k3, k4 = st.columns(4)
with k1:
    st.metric("Rækker i visningen", len(filtered_df))
with k2:
    st.metric("Skole/institution", int((filtered_df["afsender_type"] == "Skole/institution (fra emaildomæne)").sum()))
with k3:
    st.metric("Virksomhed/ukendt", int((filtered_df["afsender_type"] == "Virksomhed/organisation/ukendt").sum()))
with k4:
    st.metric("Unikke skoler på kort", int(school_df["har_koordinater"].sum()))

left, right = st.columns([1, 1])
with left:
    st.subheader("Fordeling: skole/institution vs. virksomhed")
    st.plotly_chart(donut_chart(filtered_df), use_container_width=True)

with right:
    st.subheader("Topskoler efter antal subscribers")
    st.plotly_chart(school_bar_chart(school_df), use_container_width=True)

st.subheader("Kort over repræsenterede skoler")
st.markdown("**Hold musen over en prik** for at se skolens navn og antal subscribers. Prikkerne bliver større, jo flere subscribers skolen har.")
school_map(school_df)

missing_coords = school_df[~school_df["har_koordinater"]].sort_values("antal_subscribers", ascending=False)
if not missing_coords.empty:
    with st.expander("Skoler uden koordinater (tilføj dem i SCHOOL_COORDS)"):
        st.dataframe(missing_coords[["skole", "antal_subscribers"]], use_container_width=True)