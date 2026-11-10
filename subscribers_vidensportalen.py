import streamlit as st
import pandas as pd
import pydeck as pdk
import plotly.express as px

from pathlib import Path
from io import BytesIO


# --------------------------------------------------
# Sideopsætning
# --------------------------------------------------
st.set_page_config(
    page_title="Subscriber-overblik",
    page_icon="📍",
    layout="wide"
)


# --------------------------------------------------
# Konfiguration
# --------------------------------------------------
SOURCE_SHEET = "who_is_subscribed"

# Default-filen skal ligge i samme mappe som Python-scriptet
DEFAULT_FILE = "subscribers_111026.xlsx"


# --------------------------------------------------
# Skolernes koordinater
# --------------------------------------------------
SCHOOL_COORDS = {
    "Aarhus Tech": {
        "lat": 56.1629,
        "lon": 10.2039
    },
    "Campus Bornholm": {
        "lat": 55.1009,
        "lon": 14.7066
    },
    "CELF": {
        "lat": 54.7691,
        "lon": 11.8743
    },
    "College360": {
        "lat": 56.1697,
        "lon": 9.5451
    },
    "Den Jyske Håndværkerskole": {
        "lat": 56.3297,
        "lon": 10.0487
    },
    "EUC Lillebælt": {
        "lat": 55.5377,
        "lon": 9.7120
    },
    "EUC Nord": {
        "lat": 57.4407,
        "lon": 10.5366
    },
    "EUC Nordvest": {
        "lat": 56.9567,
        "lon": 8.6947
    },
    "EUC Sjælland": {
        "lat": 55.4038,
        "lon": 11.3548
    },
    "EUC Syd": {
        "lat": 55.2494,
        "lon": 9.4870
    },
    "HANSENBERG": {
        "lat": 55.4904,
        "lon": 9.4722
    },
    "HEG": {
        "lat": 56.8046,
        "lon": 9.5322
    },
    "Herningsholm Erhvervsskole & Gymnasier": {
        "lat": 56.1362,
        "lon": 8.9766
    },
    "Learnmark Horsens": {
        "lat": 55.8607,
        "lon": 9.8503
    },
    "Mercantec": {
        "lat": 56.4532,
        "lon": 9.4020
    },
    "NEXT København": {
        "lat": 55.6561,
        "lon": 12.5187
    },
    "Nordvestsjællands Erhvervs- og Gymnasieuddannelser": {
        "lat": 55.6759,
        "lon": 11.0884
    },
    "Roskilde Tekniske Skole": {
        "lat": 55.6419,
        "lon": 12.0878
    },
    "Rybners": {
        "lat": 55.4765,
        "lon": 8.4594
    },
    "Skive College": {
        "lat": 56.5667,
        "lon": 9.0333
    },
    "Svendborg Erhvervsskole & Gymnasier": {
        "lat": 55.0598,
        "lon": 10.6068
    },
    "Syddansk Erhvervsskole": {
        "lat": 55.4038,
        "lon": 10.4024
    },
    "TECHCOLLEGE": {
        "lat": 57.0488,
        "lon": 9.9217
    },
    "TEC": {
        "lat": 55.6811,
        "lon": 12.5273
    },
    "UCH Holstebro": {
        "lat": 56.3601,
        "lon": 8.6161
    },
    "U/NORD": {
        "lat": 55.9279,
        "lon": 12.3008
    },
    "UCRS": {
        "lat": 55.9189,
        "lon": 8.4908
    },
    "ZBC": {
        "lat": 55.2296,
        "lon": 11.7609
    }
}


# --------------------------------------------------
# Kendte skoledomæner
# --------------------------------------------------
SCHOOL_DOMAINS = {
    "aarhustech.dk": "Aarhus Tech",
    "cabh.dk": "Campus Bornholm",
    "celf.dk": "CELF",
    "college360.dk": "College360",
    "djhhadsten.dk": "Den Jyske Håndværkerskole",
    "eucl.dk": "EUC Lillebælt",
    "eucnord.dk": "EUC Nord",
    "eucnordvest.dk": "EUC Nordvest",
    "eucsj.dk": "EUC Sjælland",
    "eucsyd.dk": "EUC Syd",
    "hansenberg.dk": "HANSENBERG",
    "heguddannelser.dk": "HEG",
    "herningsholm.dk": (
        "Herningsholm Erhvervsskole & Gymnasier"
    ),
    "learnmark.dk": "Learnmark Horsens",
    "mercantec.dk": "Mercantec",
    "neg.dk": (
        "Nordvestsjællands Erhvervs- og "
        "Gymnasieuddannelser"
    ),
    "nextkbh.dk": "NEXT København",
    "rts.dk": "Roskilde Tekniske Skole",
    "rybners.dk": "Rybners",
    "sde.dk": "Syddansk Erhvervsskole",
    "sesg.dk": (
        "Svendborg Erhvervsskole & Gymnasier"
    ),
    "skivecollege.dk": "Skive College",
    "tec.dk": "TEC",
    "techcollege.dk": "TECHCOLLEGE",
    "ucholstebro.dk": "UCH Holstebro",
    "ucrs.dk": "UCRS",
    "unord.dk": "U/NORD",
    "zbc.dk": "ZBC"
}


# --------------------------------------------------
# Variationer af skolenavne
# --------------------------------------------------
SCHOOL_NAME_MAP = {
    # Aarhus Tech
    "aarhus tech": "Aarhus Tech",
    "aarhustech": "Aarhus Tech",
    "århus tech": "Aarhus Tech",

    # Campus Bornholm
    "campus bornholm": "Campus Bornholm",
    "cabh": "Campus Bornholm",

    # CELF
    "celf": "CELF",
    "celf-center for erhv.rettede udd. lolland-falster": (
        "CELF"
    ),
    "celf center for erhv.rettede udd. lolland-falster": (
        "CELF"
    ),

    # College360
    "college360": "College360",
    "college 360": "College360",

    # Den Jyske Håndværkerskole
    "den jyske håndværkerskole": (
        "Den Jyske Håndværkerskole"
    ),
    "den jyske haandvaerkerskole": (
        "Den Jyske Håndværkerskole"
    ),
    "djh": "Den Jyske Håndværkerskole",
    "djh hadsten": "Den Jyske Håndværkerskole",

    # EUC Lillebælt
    "euc lillebælt": "EUC Lillebælt",
    "euc lillebaelt": "EUC Lillebælt",

    # EUC Nord
    "euc nord": "EUC Nord",
    "eucnord": "EUC Nord",

    # EUC Nordvest
    "euc nordvest": "EUC Nordvest",
    "eucnordvest": "EUC Nordvest",

    # EUC Sjælland
    "euc sj": "EUC Sjælland",
    "euc sjælland": "EUC Sjælland",
    "euc sjaelland": "EUC Sjælland",

    # EUC Syd
    "euc syd": "EUC Syd",
    "eucsyd": "EUC Syd",
    "euc syd, haderslev": "EUC Syd",

    # HANSENBERG
    "hansenberg": "HANSENBERG",

    # HEG
    "heg": "HEG",

    # Herningsholm
    "herningsholm": (
        "Herningsholm Erhvervsskole & Gymnasier"
    ),
    "herningsholm erhvervsskole": (
        "Herningsholm Erhvervsskole & Gymnasier"
    ),
    "herningsholm erhvervsskole og gymnasier": (
        "Herningsholm Erhvervsskole & Gymnasier"
    ),
    "herningsholm erhvervsskole & gymnasier": (
        "Herningsholm Erhvervsskole & Gymnasier"
    ),

    # Learnmark
    "learnmark": "Learnmark Horsens",
    "learnmark tech": "Learnmark Horsens",
    "learnmark horsens": "Learnmark Horsens",

    # Mercantec
    "mercantec": "Mercantec",
    "mercantec viborg": "Mercantec",

    # NEXT
    "next": "NEXT København",
    "next kbh": "NEXT København",
    "next københavn": "NEXT København",
    "next kobenhavn": "NEXT København",

    # NEG
    "neg": (
        "Nordvestsjællands Erhvervs- og "
        "Gymnasieuddannelser"
    ),
    "nordvestsjællands erhvervs- og "
    "gymnasieuddannelser": (
        "Nordvestsjællands Erhvervs- og "
        "Gymnasieuddannelser"
    ),

    # Roskilde Tekniske Skole
    "rts": "Roskilde Tekniske Skole",
    "roskilde tekniske skole": (
        "Roskilde Tekniske Skole"
    ),

    # Rybners
    "rybners": "Rybners",

    # Skive College
    "skive college": "Skive College",

    # Svendborg
    "svendborg erhvervsskole": (
        "Svendborg Erhvervsskole & Gymnasier"
    ),
    "svendborg erhvervsskole og gymnasier": (
        "Svendborg Erhvervsskole & Gymnasier"
    ),
    "svendborg erhvervsskole & gymnasier": (
        "Svendborg Erhvervsskole & Gymnasier"
    ),

    # Syddansk Erhvervsskole
    "sde": "Syddansk Erhvervsskole",
    "sde odense": "Syddansk Erhvervsskole",
    "sde vejle": "Syddansk Erhvervsskole",
    "syddansk erhvervsskole": (
        "Syddansk Erhvervsskole"
    ),
    "syddansk erhversskole": (
        "Syddansk Erhvervsskole"
    ),
    "syddansk erhvervsskole odense-vejle": (
        "Syddansk Erhvervsskole"
    ),
    "syddansk erhvervsskole i vejle (sde)": (
        "Syddansk Erhvervsskole"
    ),

    # TECHCOLLEGE
    "techcollege": "TECHCOLLEGE",
    "tech college": "TECHCOLLEGE",

    # TEC
    "tec": "TEC",
    "tec frederiksberg": "TEC",
    "tec fredeiksberg": "TEC",
    "tec gladsaxe": "TEC",
    "tec erhvervscenter": "TEC",
    "tec, technical education copenhagen": "TEC",

    # UCH
    "uch": "UCH Holstebro",
    "uch holstebro": "UCH Holstebro",
    "uc holstebro": "UCH Holstebro",

    # U/NORD
    "unord": "U/NORD",
    "u/nord": "U/NORD",
    "unord hillerød": "U/NORD",
    "unord. hillerød": "U/NORD",

    # UCRS
    "ucrs": "UCRS",

    # ZBC
    "zbc": "ZBC",
    "el & dataskolen, zbc": "ZBC"
}


# --------------------------------------------------
# Tekstrensning
# --------------------------------------------------
def clean_text(value):
    """
    Renser tomme værdier, dobbelte mellemrum
    og almindelige tegnkodningsfejl.
    """

    if pd.isna(value):
        return ""

    text = str(value).strip()

    replacements = {
        "\u00a0": " ",
        "Ã¸": "ø",
        "Ã˜": "Ø",
        "Ã¦": "æ",
        "Ã†": "Æ",
        "Ã¥": "å",
        "Ã…": "Å",
        "â€™": "’",
        "â€“": "–",
    }

    for old_value, new_value in replacements.items():
        text = text.replace(old_value, new_value)

    # Fjerner dobbelte og overflødige mellemrum
    text = " ".join(text.split())

    return text


def get_email_domain(email):
    """
    Finder domænet efter @ i mailadressen.
    """

    email = clean_text(email).lower()

    if "@" not in email:
        return ""

    return email.rsplit("@", 1)[-1]


def normalize_school_name(value):
    """
    Ensretter kendte variationer af skolenavne.
    """

    value = clean_text(value)

    if value == "":
        return ""

    lookup_value = value.lower().strip(" '\"")

    return SCHOOL_NAME_MAP.get(
        lookup_value,
        value
    )


# --------------------------------------------------
# Rens én subscriber
# --------------------------------------------------
def clean_subscriber_row(row):
    original_school = clean_text(
        row.get(
            "subscriber_virksomhed_skole",
            ""
        )
    )

    original_company = clean_text(
        row.get(
            "subscriber_virksomhed",
            ""
        )
    )

    organisation = clean_text(
        row.get(
            "subscriber_organisation",
            ""
        )
    )

    email = clean_text(
        row.get(
            "email",
            ""
        )
    )

    domain = get_email_domain(email)

    normalized_school = normalize_school_name(
        original_school
    )

    cleaned_school = ""
    cleaned_company = original_company

    changes = []

    # ------------------------------------------------
    # Skolen findes via maildomænet
    # ------------------------------------------------
    if domain in SCHOOL_DOMAINS:
        cleaned_school = SCHOOL_DOMAINS[domain]

        sender_type = (
            "Skole/institution (fra emaildomæne)"
        )

        if original_school != cleaned_school:
            changes.append(
                f"Skole ændret fra '{original_school}' "
                f"til '{cleaned_school}' ud fra "
                f"domænet {domain}"
            )

    # ------------------------------------------------
    # Skolen findes via det indtastede skolefelt
    # ------------------------------------------------
    elif normalized_school in SCHOOL_COORDS:
        cleaned_school = normalized_school

        sender_type = (
            "Skole/institution (fra indtastet skole)"
        )

        if original_school != cleaned_school:
            changes.append(
                f"Skolenavn normaliseret fra "
                f"'{original_school}' til "
                f"'{cleaned_school}'"
            )

    # ------------------------------------------------
    # Virksomhed, organisation eller ukendt
    # ------------------------------------------------
    else:
        sender_type = (
            "Virksomhed/organisation/ukendt"
        )

        # Brug organisationen, hvis virksomhed er tom
        if (
            cleaned_company == ""
            and organisation != ""
        ):
            cleaned_company = organisation

            changes.append(
                "Virksomhed udfyldt fra "
                "subscriber_organisation"
            )

        # Hvis skolefeltet ikke indeholder en kendt skole,
        # kan værdien være en virksomhed
        if (
            cleaned_company == ""
            and original_school != ""
        ):
            cleaned_company = original_school

            changes.append(
                "Værdi flyttet fra skolefelt "
                "til virksomhed"
            )

    # Undgå at samme navn står som både skole og virksomhed
    if (
        cleaned_school != ""
        and cleaned_company != ""
        and cleaned_school.lower()
        == cleaned_company.lower()
    ):
        cleaned_company = ""

        changes.append(
            "Dublet fjernet fra virksomhedsfelt"
        )

    return pd.Series(
        {
            "email_domain": domain,
            "subscriber_virksomhed_skole": (
                cleaned_school
            ),
            "subscriber_virksomhed": (
                cleaned_company
            ),
            "afsender_type": sender_type,
            "ændringslog": " | ".join(changes)
        }
    )


# --------------------------------------------------
# Indlæs og rens Excel
# --------------------------------------------------
@st.cache_data(show_spinner=False)
def load_and_clean_data(file_bytes):
    excel_file = BytesIO(file_bytes)

    workbook = pd.ExcelFile(
        excel_file,
        engine="openpyxl"
    )

    # Brug det forventede ark
    if SOURCE_SHEET in workbook.sheet_names:
        sheet_name = SOURCE_SHEET

    # Hvis filen kun har ét ark, bruges det automatisk
    elif len(workbook.sheet_names) == 1:
        sheet_name = workbook.sheet_names[0]

    else:
        raise ValueError(
            f"Arket '{SOURCE_SHEET}' blev ikke fundet. "
            f"Tilgængelige ark: "
            f"{workbook.sheet_names}"
        )

    excel_file.seek(0)

    df = pd.read_excel(
        excel_file,
        sheet_name=sheet_name,
        engine="openpyxl"
    )

    # Fjern helt tomme rækker
    df = df.dropna(
        how="all"
    ).copy()

    required_columns = [
        "created_at",
        "email",
        "subscriber_virksomhed_skole",
        "subscriber_organisation",
        "subscriber_virksomhed"
    ]

    missing_columns = [
        column
        for column in required_columns
        if column not in df.columns
    ]

    if missing_columns:
        raise ValueError(
            "Følgende kolonner mangler i filen: "
            f"{missing_columns}"
        )

    # Rens alle tekstkolonner
    for column in df.select_dtypes(
        include=["object"]
    ).columns:
        df[column] = df[column].apply(
            clean_text
        )

    # Beregn de rensede felter
    cleaned_columns = df.apply(
        clean_subscriber_row,
        axis=1
    )

    # Overskriv eller tilføj rensede kolonner
    for column in cleaned_columns.columns:
        df[column] = cleaned_columns[column]

    return df


# --------------------------------------------------
# Klassificering til diagram
# --------------------------------------------------
def classify_org_type(row):
    sender_type = row.get(
        "afsender_type",
        ""
    )

    if str(sender_type).startswith(
        "Skole/institution"
    ):
        return "Skole/institution"

    return "Virksomhed/organisation/ukendt"


# --------------------------------------------------
# Skoledata til kort og søjlediagram
# --------------------------------------------------
def build_school_map_df(df):
    school_rows = df[
        df["subscriber_virksomhed_skole"]
        .astype(str)
        .str.strip()
        != ""
    ].copy()

    school_df = (
        school_rows
        .groupby(
            "subscriber_virksomhed_skole",
            as_index=False
        )
        .agg(
            antal_subscribers=(
                "email",
                "count"
            )
        )
        .rename(
            columns={
                "subscriber_virksomhed_skole": (
                    "skole"
                )
            }
        )
    )

    coords_df = pd.DataFrame(
        [
            {
                "skole": school,
                "lat": coordinates["lat"],
                "lon": coordinates["lon"]
            }
            for school, coordinates
            in SCHOOL_COORDS.items()
        ]
    )

    school_df = school_df.merge(
        coords_df,
        on="skole",
        how="left"
    )

    school_df["har_koordinater"] = (
        school_df["lat"].notna()
        & school_df["lon"].notna()
    )

    # Beregn punktstørrelse på kortet
    if not school_df.empty:
        minimum = (
            school_df["antal_subscribers"]
            .min()
        )

        maximum = (
            school_df["antal_subscribers"]
            .max()
        )

        if maximum == minimum:
            school_df["radius"] = 12000

        else:
            school_df["radius"] = (
                7000
                + (
                    (
                        school_df[
                            "antal_subscribers"
                        ]
                        - minimum
                    )
                    / (maximum - minimum)
                    * 28000
                )
            )

    else:
        school_df["radius"] = pd.Series(
            dtype=float
        )

    return school_df


# --------------------------------------------------
# Fordelingsdiagram
# --------------------------------------------------
def donut_chart(df):
    summary = (
        df.assign(
            type_simple=df.apply(
                classify_org_type,
                axis=1
            )
        )
        .groupby(
            "type_simple",
            as_index=False
        )
        .agg(
            antal=(
                "email",
                "count"
            )
        )
    )

    fig = px.pie(
        summary,
        names="type_simple",
        values="antal",
        hole=0.55,
        color="type_simple",
        color_discrete_map={
            "Skole/institution": "#1f77b4",
            "Virksomhed/organisation/ukendt": (
                "#ff7f0e"
            )
        }
    )

    fig.update_traces(
        textposition="inside",
        textinfo="percent+label",
        hovertemplate=(
            "%{label}<br>"
            "Antal: %{value}"
            "<extra></extra>"
        )
    )

    fig.update_layout(
        margin=dict(
            l=10,
            r=10,
            t=10,
            b=10
        ),
        showlegend=True
    )

    return fig


# --------------------------------------------------
# Topskoler
# --------------------------------------------------
def school_bar_chart(
    school_df,
    top_n=15
):
    plot_df = (
        school_df
        .sort_values(
            "antal_subscribers",
            ascending=False
        )
        .head(top_n)
    )

    fig = px.bar(
        plot_df,
        x="antal_subscribers",
        y="skole",
        orientation="h",
        text="antal_subscribers",
        color="antal_subscribers",
        color_continuous_scale="Blues"
    )

    fig.update_layout(
        yaxis_title="",
        xaxis_title="Antal subscribers",
        coloraxis_showscale=False,
        margin=dict(
            l=10,
            r=10,
            t=10,
            b=10
        )
    )

    fig.update_traces(
        hovertemplate=(
            "%{y}<br>"
            "Antal: %{x}"
            "<extra></extra>"
        )
    )

    fig.update_yaxes(
        categoryorder="total ascending"
    )

    return fig


# --------------------------------------------------
# Data til tidsdiagrammer
# --------------------------------------------------
def prepare_timeline_data(df):
    plot_df = df.copy()

    plot_df["created_at"] = pd.to_datetime(
        plot_df["created_at"],
        errors="coerce",
        utc=True
    )

    plot_df = plot_df.dropna(
        subset=["created_at"]
    )

    # Fjern tidszone før konvertering til måned
    plot_df["created_at"] = (
        plot_df["created_at"]
        .dt.tz_convert(None)
    )

    plot_df["måned"] = (
        plot_df["created_at"]
        .dt.to_period("M")
        .astype(str)
    )

    summary = (
        plot_df
        .groupby("måned")
        .size()
        .reset_index(
            name="nye_subscribers"
        )
        .sort_values("måned")
    )

    summary["samlet"] = (
        summary["nye_subscribers"]
        .cumsum()
    )

    return summary


# --------------------------------------------------
# Nye subscribers pr. måned
# --------------------------------------------------
def monthly_subscribers_chart(df):
    summary = prepare_timeline_data(df)

    fig = px.bar(
        summary,
        x="måned",
        y="nye_subscribers",
        text="nye_subscribers",
        color="nye_subscribers",
        color_continuous_scale="Blues"
    )

    fig.update_layout(
        xaxis_title="Måned",
        yaxis_title="Nye subscribers",
        coloraxis_showscale=False
    )

    fig.update_traces(
        hovertemplate=(
            "Måned: %{x}<br>"
            "Nye subscribers: %{y}"
            "<extra></extra>"
        )
    )

    return fig


# --------------------------------------------------
# Samlet subscriber-udvikling
# --------------------------------------------------
def subscriber_timeline(df):
    summary = prepare_timeline_data(df)

    fig = px.line(
        summary,
        x="måned",
        y="samlet",
        markers=True
    )

    fig.update_layout(
        xaxis_title="Måned",
        yaxis_title="Samlet antal subscribers"
    )

    fig.update_traces(
        hovertemplate=(
            "Måned: %{x}<br>"
            "Samlet antal: %{y}"
            "<extra></extra>"
        )
    )

    return fig


# --------------------------------------------------
# Danmarkskort
# --------------------------------------------------
def school_map(school_df):
    map_df = school_df[
        school_df["har_koordinater"]
    ].copy()

    if map_df.empty:
        st.warning(
            "Der er ingen skoler med koordinater "
            "i den aktuelle visning."
        )
        return

    tooltip = {
        "html": (
            "<b>{skole}</b><br/>"
            "Antal subscribers: "
            "{antal_subscribers}"
        ),
        "style": {
            "backgroundColor": (
                "rgba(30, 30, 30, 0.85)"
            ),
            "color": "white",
            "fontSize": "13px"
        }
    }

    view_state = pdk.ViewState(
        latitude=map_df["lat"].mean(),
        longitude=map_df["lon"].mean(),
        zoom=5.8,
        pitch=0
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
        get_radius="radius"
    )

    deck = pdk.Deck(
        layers=[layer],
        initial_view_state=view_state,
        tooltip=tooltip,
        map_style="light"
    )

    st.pydeck_chart(
        deck,
        use_container_width=True
    )


# --------------------------------------------------
# Brugerflade
# --------------------------------------------------
st.title("📍 Subscriber-overblik")

st.caption(
    "Appen bruger automatisk standardfilen, "
    "men du kan uploade en nyere subscriber-fil "
    "i sidebaren."
)


# --------------------------------------------------
# Upload eller default-fil
# --------------------------------------------------
with st.sidebar:
    st.header("Data")

    uploaded_file = st.file_uploader(
        "Upload ny subscriber-fil",
        type=["xlsx"],
        help=(
            "Hvis du ikke uploader en fil, "
            "bruges standardfilen automatisk."
        )
    )


# --------------------------------------------------
# Vælg fil
# --------------------------------------------------
if uploaded_file is not None:
    file_bytes = uploaded_file.getvalue()

    source_name = uploaded_file.name
    source_type = "upload"

else:
    # Finder mappen, hvor Python-scriptet ligger
    script_folder = Path(
        __file__
    ).resolve().parent

    default_path = (
        script_folder
        / DEFAULT_FILE
    )

    if not default_path.exists():
        st.error(
            f"Standardfilen '{DEFAULT_FILE}' "
            f"blev ikke fundet i mappen:\n\n"
            f"{script_folder}"
        )

        st.info(
            "Læg Excel-filen i samme mappe som "
            "subscribers_vidensportalen.py, eller "
            "upload en fil i sidebaren."
        )

        st.stop()

    file_bytes = default_path.read_bytes()

    source_name = DEFAULT_FILE
    source_type = "standard"


# --------------------------------------------------
# Indlæs og rens data
# --------------------------------------------------
try:
    with st.spinner(
        "Indlæser og renser subscriber-data..."
    ):
        df = load_and_clean_data(
            file_bytes
        )

except Exception as error:
    st.error(
        "Filen kunne ikke indlæses eller renses."
    )

    st.exception(error)

    st.stop()


# --------------------------------------------------
# Filtre
# --------------------------------------------------
with st.sidebar:
    st.header("Filtre")

    only_with_school = st.checkbox(
        "Kun subscribers med skole",
        value=False
    )

    search_term = st.text_input(
        "Søg i skole, virksomhed eller mail"
    )


filtered_df = df.copy()


if only_with_school:
    filtered_df = filtered_df[
        filtered_df[
            "subscriber_virksomhed_skole"
        ]
        .astype(str)
        .str.strip()
        != ""
    ]


if search_term:
    search_term = search_term.strip()

    search_mask = (
        filtered_df[
            "subscriber_virksomhed_skole"
        ]
        .astype(str)
        .str.contains(
            search_term,
            case=False,
            na=False
        )
        |
        filtered_df[
            "subscriber_virksomhed"
        ]
        .astype(str)
        .str.contains(
            search_term,
            case=False,
            na=False
        )
        |
        filtered_df["email"]
        .astype(str)
        .str.contains(
            search_term,
            case=False,
            na=False
        )
    )

    filtered_df = filtered_df[
        search_mask
    ]


# --------------------------------------------------
# Hvis filtre giver nul rækker
# --------------------------------------------------
if filtered_df.empty:
    st.warning(
        "Der er ingen subscribers, som matcher "
        "de valgte filtre."
    )

    st.stop()


school_df = build_school_map_df(
    filtered_df
)


# --------------------------------------------------
# KPI'er
# --------------------------------------------------
school_count = int(
    filtered_df["afsender_type"]
    .astype(str)
    .str.startswith(
        "Skole/institution"
    )
    .sum()
)

company_count = int(
    (
        filtered_df["afsender_type"]
        == "Virksomhed/organisation/ukendt"
    ).sum()
)

unique_schools = int(
    school_df["skole"].nunique()
)

schools_on_map = int(
    school_df["har_koordinater"].sum()
)


k1, k2, k3, k4 = st.columns(4)


with k1:
    st.metric(
        "Subscribers i visningen",
        len(filtered_df)
    )


with k2:
    st.metric(
        "Skole/institution",
        school_count
    )


with k3:
    st.metric(
        "Virksomhed/ukendt",
        company_count
    )


with k4:
    st.metric(
        "Skoler på kort",
        schools_on_map,
        help=(
            f"{unique_schools} unikke skoler "
            f"findes i den aktuelle visning."
        )
    )


# --------------------------------------------------
# Fordeling og topskoler
# --------------------------------------------------
left, right = st.columns(
    [1, 1]
)


with left:
    st.subheader(
        "Fordeling: skole/institution "
        "vs. virksomhed"
    )

    st.plotly_chart(
        donut_chart(filtered_df),
        use_container_width=True
    )


with right:
    st.subheader(
        "Topskoler efter antal subscribers"
    )

    if school_df.empty:
        st.info(
            "Der er ingen skoler i den "
            "aktuelle visning."
        )

    else:
        st.plotly_chart(
            school_bar_chart(
                school_df
            ),
            use_container_width=True
        )


# --------------------------------------------------
# Kort
# --------------------------------------------------
st.subheader(
    "Kort over repræsenterede skoler"
)

st.markdown(
    "**Hold musen over en prik** for at se "
    "skolens navn og antal subscribers. "
    "Prikkerne bliver større, jo flere "
    "subscribers skolen har."
)

school_map(
    school_df
)


# --------------------------------------------------
# Nye subscribers pr. måned
# --------------------------------------------------
st.subheader(
    "Nye subscribers pr. måned"
)

st.plotly_chart(
    monthly_subscribers_chart(
        filtered_df
    ),
    use_container_width=True
)


# --------------------------------------------------
# Samlet udvikling
# --------------------------------------------------
st.subheader(
    "Samlet udvikling i subscribers"
)

st.plotly_chart(
    subscriber_timeline(
        filtered_df
    ),
    use_container_width=True
)


# --------------------------------------------------
# Rensede data
# --------------------------------------------------
with st.expander(
    "Vis rensede subscriber-data"
):
    display_columns = [
        "created_at",
        "email",
        "full_name",
        "email_domain",
        "subscriber_virksomhed_skole",
        "subscriber_organisation",
        "subscriber_virksomhed",
        "afsender_type",
        "ændringslog"
    ]

    existing_columns = [
        column
        for column in display_columns
        if column in filtered_df.columns
    ]

    st.dataframe(
        filtered_df[
            existing_columns
        ],
        use_container_width=True,
        hide_index=True
    )


# --------------------------------------------------
# Skoler uden koordinater
# --------------------------------------------------
missing_coords = (
    school_df[
        ~school_df["har_koordinater"]
    ]
    .sort_values(
        "antal_subscribers",
        ascending=False
    )
)


if not missing_coords.empty:
    with st.expander(
        "Skoler uden koordinater"
    ):
        st.warning(
            "Disse skolenavne findes i data, "
            "men matcher ikke et navn i "
            "SCHOOL_COORDS."
        )

        st.dataframe(
            missing_coords[
                [
                    "skole",
                    "antal_subscribers"
                ]
            ],
            use_container_width=True,
            hide_index=True
        )