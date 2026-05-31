from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st


# -----------------------------------------------------------------------------
# 1. App configuration and constants
# -----------------------------------------------------------------------------

APP_DIR = Path(__file__).parent
DATA_DIR = APP_DIR / "data"
ALT_DATA_DIR = APP_DIR.parent / "data_acquisition"

TRAVEL_CSV = DATA_DIR / "travel_data.csv"
BUDGET_CSV = DATA_DIR / "co2_budgets.csv"
ALT_TRAVEL_CSV = ALT_DATA_DIR / "travel_data.csv"
ALT_BUDGET_CSV = ALT_DATA_DIR / "co2_budgets.csv"
EXCEL_FALLBACK = APP_DIR / "traveldata-export.xlsx"

RFI_COLUMNS = {
    "RFI 2.0": "CO2e RFI2 (t)",
    "RFI 2.7": "CO2e RFI2.7 (t)",
}

MONTHS = {
    "Alle Monate": None,
    "Januar": 1,
    "Februar": 2,
    "Maerz": 3,
    "April": 4,
    "Mai": 5,
    "Juni": 6,
    "Juli": 7,
    "August": 8,
    "September": 9,
    "Oktober": 10,
    "November": 11,
    "Dezember": 12,
}

# -----------------------------------------------------------------------------
# Mapping-Dictionary für deutsche Spaltennamen im UI
# -----------------------------------------------------------------------------
DISPLAY_NAMES_DE = {
    "date": "Datum",
    "year": "Jahr",
    "month": "Monat",
    "business_unit": "Geschäftsbereich",
    "subunit": "Abteilung",
    "transport_mode": "Verkehrsmittel",
    "travel_purpose": "Reisezweck",
    "haul": "Distanzklasse",
    "departure_city": "Abreiseort",
    "arrival_city": "Ankunftsort",
    "departure_country": "Abreiseland",
    "arrival_country": "Ankunftsland",
    "km": "Distanz (km)",
    "cost_CHF": "Kosten (CHF)",
    "train_alternative_available": "Bahnalternative verfügbar",
    "cost_per_t_co2e": "Kosten pro t CO₂e",
    "co2_budget_t": "CO₂-Budget (t)",
    "actual_t": "Ist CO₂e (t)",
    "budget_t": "Budget (t)",
    "remaining_t": "Rest (t)",
    "budget_used_pct": "Budgetverbrauch",
    "trips": "Reisen",
    "cost_chf": "Kosten (CHF)",
    "status": "Status",
    "Einheit": "Einheit",
    "route": "Route",
    "emissions_t": "CO₂e (t)",
    "avg_km": "Ø km",
    "annual_emissions_t": "Jahres-Emissionen (t)",
    "annual_trips": "Jahres-Reisen",
    "annual_cost_chf": "Jahreskosten (CHF)",
    "report_period_emissions_t": "Perioden-Emissionen (t)",
    "report_period_trips": "Perioden-Reisen",
    "CO2e RFI2 (t)": "CO₂e RFI 2.0 (t)",
    "CO2e RFI2.7 (t)": "CO₂e RFI 2.7 (t)"
}

# Bekannte Übersetzungen für Datenwerte (Englisch → Deutsch)
# Unbekannte Werte werden automatisch bereinigt: "site_visit" → "Site Visit"
VALUE_TRANSLATIONS_DE = {
    "flight": "Flug",
    "train": "Bahn",
    "rental_car": "Mietwagen",
    "bus": "Bus",
    "car": "Auto",
    "ferry": "Fähre",
    "taxi": "Taxi",
    "long haul": "Langstrecke",
    "mid haul": "Mittelstrecke",
    "short haul": "Kurzstrecke",
    "ground": "Boden",
    "client_meeting": "Kundengespräch",
    "internal_meeting": "Internes Meeting",
    "site_visit": "Standortbesuch",
    "trade_fair": "Messe",
    "training": "Schulung",
    "workshop": "Workshop",
    "conference": "Konferenz",
    "other": "Sonstiges",
}

def translate_value(v):
    """Übersetzt einen Datenwert. Fallback: Unterstriche entfernen und Title Case."""
    s = str(v)
    if s in VALUE_TRANSLATIONS_DE:
        return VALUE_TRANSLATIONS_DE[s]
    return s.replace("_", " ").title()

STATUS_GREEN = "#1f8a5b"
STATUS_RED = "#c93838"
STATUS_AMBER = "#b7791f"
BLUE = "#2f6bff"
DARK = "#20242a"
MID = "#667085"
LIGHT = "#eef2f6"
GRID = "#d9dee7"

st.set_page_config(
    page_title="Geschäftsreisen CO₂ Dashboard",
    page_icon="",
    layout="wide",
)


# -----------------------------------------------------------------------------
# 2. Data loading and preprocessing
# -----------------------------------------------------------------------------

def _file_mtimes():
    """Gibt Änderungszeiten der Datendateien zurück — Cache wird automatisch ungültig wenn Dateien sich ändern."""
    times = []
    for p in [TRAVEL_CSV, BUDGET_CSV, ALT_TRAVEL_CSV, ALT_BUDGET_CSV, EXCEL_FALLBACK]:
        times.append(p.stat().st_mtime if p.exists() else 0)
    return tuple(times)

@st.cache_data(show_spinner=False)
def load_data(_mtimes):
    if TRAVEL_CSV.exists() and BUDGET_CSV.exists():
        travel = pd.read_csv(TRAVEL_CSV)
        budgets = pd.read_csv(BUDGET_CSV)
    elif ALT_TRAVEL_CSV.exists() and ALT_BUDGET_CSV.exists():
        travel = pd.read_csv(ALT_TRAVEL_CSV)
        budgets = pd.read_csv(ALT_BUDGET_CSV)
    elif EXCEL_FALLBACK.exists():
        travel = pd.read_excel(EXCEL_FALLBACK, sheet_name="travel_data")
        budgets = pd.read_excel(EXCEL_FALLBACK, sheet_name="co2_budgets")
    else:
        st.error(
            "Keine Daten gefunden. Erwartet werden data/travel_data.csv und "
            "data/co2_budgets.csv neben dieser Streamlit-Datei."
        )
        st.stop()

    travel = preprocess_travel_data(travel)
    budgets = preprocess_budget_data(budgets, travel)
    return travel, budgets


def preprocess_travel_data(travel):
    travel = travel.copy()
    travel["date"] = pd.to_datetime(travel["date"], errors="coerce")
    travel["year"] = travel["date"].dt.year
    travel["month"] = travel["date"].dt.month

    numeric_cols = ["CO2e RFI2 (t)", "CO2e RFI2.7 (t)", "km", "cost_CHF"]
    for col in numeric_cols:
        travel[col] = pd.to_numeric(travel[col], errors="coerce").fillna(0)

    text_cols = [
        "business_unit",
        "subunit",
        "transport_mode",
        "travel_purpose",
        "haul",
        "departure_city",
        "arrival_city",
        "departure_country",
        "arrival_country",
    ]
    for col in text_cols:
        if col in travel.columns:
            travel[col] = travel[col].fillna("Unbekannt")

    travel["train_alternative_available"] = (
        travel["train_alternative_available"].astype(str).str.lower().eq("true")
    )
    return travel


def preprocess_budget_data(budgets, travel):
    budgets = budgets.copy()
    budgets["year"] = pd.to_numeric(budgets["year"], errors="coerce").astype("Int64")
    budgets["co2_budget_t"] = pd.to_numeric(budgets["co2_budget_t"], errors="coerce")

    subunit_map = (
        travel[["subunit", "business_unit"]]
        .dropna()
        .drop_duplicates()
        .sort_values(["business_unit", "subunit"])
    )
    return budgets.merge(subunit_map, on="subunit", how="left")


# -----------------------------------------------------------------------------
# 3. Formatting helpers
# -----------------------------------------------------------------------------

def format_number(value, decimals=0):
    if pd.isna(value):
        return "n/a"
    text = f"{value:,.{decimals}f}"
    return text.replace(",", "'")


def format_tonnes(value, decimals=1):
    return f"{format_number(value, decimals)} t"


def format_chf(value):
    return f"CHF {format_number(value, 0)}"


def format_pct(value):
    if pd.isna(value):
        return "n/a"
    return f"{value:.1%}"


def budget_delta(context):
    budget = context["annual_budget"]
    remaining = context["remaining"]

    if pd.isna(budget) or budget <= 0 or pd.isna(remaining):
        return {
            "text": "n/a",
            "class": "neutral",
            "arrow": "•",
        }

    pct = remaining / budget
    if remaining >= 0:
        return {
            "text": f"+{format_tonnes(remaining)} (+{format_pct(pct)})",
            "class": "positive",
            "arrow": "↑",
        }

    return {
        "text": f"{format_tonnes(remaining)} ({format_pct(pct)})",
        "class": "negative",
        "arrow": "↓",
    }


def to_csv_bytes(df):
    return df.to_csv(index=False).encode("utf-8-sig")


# -----------------------------------------------------------------------------
# 4. Filter and metric calculations
# -----------------------------------------------------------------------------

def sidebar_filters(travel):
    with st.sidebar:
        st.title("Filter")

        view = st.radio(
            "Ansicht",
            ["BU-Leitung", "CSO Report", "Finance Report"],
            index=0,
            help="BU-Leitung ist die operative Standardsicht. CSO und Finance erhalten kompakte Report-Sichten.",
        )

        business_units = sorted(travel["business_unit"].dropna().unique())
        selected_bu = st.selectbox(DISPLAY_NAMES_DE.get("business_unit", "Business Unit"), ["Alle"] + business_units, index=1)

        if selected_bu == "Alle":
            subunits = sorted(travel["subunit"].dropna().unique())
        else:
            subunits = sorted(
                travel.loc[travel["business_unit"] == selected_bu, "subunit"]
                .dropna()
                .unique()
            )
        selected_subunit = st.selectbox(DISPLAY_NAMES_DE.get("subunit", "Subunit"), ["Alle"] + subunits, index=0)

        years = sorted(int(y) for y in travel["year"].dropna().unique())
        default_year = 2025 if 2025 in years else years[-1]
        selected_year = st.selectbox(DISPLAY_NAMES_DE.get("year", "Jahr"), years, index=years.index(default_year))

        selected_month_label = st.selectbox(DISPLAY_NAMES_DE.get("month", "Monat"), list(MONTHS.keys()), index=0)
        selected_month = MONTHS[selected_month_label]

        selected_rfi = st.radio(
            "CO₂e-Szenario",
            list(RFI_COLUMNS.keys()),
            index=1,
            help=(
                "RFI steht für Radiative Forcing Index. Er berücksichtigt, dass Flugemissionen "
                "in grosser Höhe eine stärkere Klimawirkung haben als reines CO₂ am Boden. "
                "RFI 2.0 ist die konservativere Annahme, RFI 2.7 die strengere Klimawirkungsannahme."
            ),
        )
        emission_col = RFI_COLUMNS[selected_rfi]

        with st.expander("Erweiterte Filter", expanded=False):
            modes = sorted(travel["transport_mode"].dropna().unique())
            purposes = sorted(travel["travel_purpose"].dropna().unique())
            hauls = sorted(travel["haul"].dropna().unique())

            selected_modes = st.multiselect(DISPLAY_NAMES_DE.get("transport_mode", "Transportmodus"), modes, default=modes)
            selected_purposes = st.multiselect(DISPLAY_NAMES_DE.get("travel_purpose", "Reisezweck"), purposes, default=purposes)
            selected_hauls = st.multiselect(DISPLAY_NAMES_DE.get("haul", "Haul"), hauls, default=hauls)

        return {
            "view": view,
            "business_unit": selected_bu,
            "subunit": selected_subunit,
            "year": selected_year,
            "month": selected_month,
            "month_label": selected_month_label,
            "rfi_label": selected_rfi,
            "emission_col": emission_col,
            "modes": selected_modes,
            "purposes": selected_purposes,
            "hauls": selected_hauls,
        }


def apply_scope_filters(travel, filters, include_month=True, include_advanced=True):
    df = travel[travel["year"] == filters["year"]].copy()

    if include_month and filters["month"] is not None:
        df = df[df["month"] == filters["month"]]

    if filters["business_unit"] != "Alle":
        df = df[df["business_unit"] == filters["business_unit"]]

    if filters["subunit"] != "Alle":
        df = df[df["subunit"] == filters["subunit"]]

    if include_advanced:
        df = df[
            df["transport_mode"].isin(filters["modes"])
            & df["travel_purpose"].isin(filters["purposes"])
            & df["haul"].isin(filters["hauls"])
        ]

    return df


def budget_scope(budgets, filters):
    scoped = budgets[budgets["year"] == filters["year"]].copy()

    if filters["business_unit"] != "Alle":
        scoped = scoped[scoped["business_unit"] == filters["business_unit"]]

    if filters["subunit"] != "Alle":
        scoped = scoped[scoped["subunit"] == filters["subunit"]]

    return scoped


def calculate_context(travel, budgets, filters):
    emission_col = filters["emission_col"]

    period_df = apply_scope_filters(travel, filters, include_month=True, include_advanced=True)
    annual_df = apply_scope_filters(travel, filters, include_month=False, include_advanced=False)
    annual_budget = budget_scope(budgets, filters)["co2_budget_t"].sum(min_count=1)

    annual_actual = annual_df[emission_col].sum()
    period_actual = period_df[emission_col].sum()

    if pd.isna(annual_budget) or annual_budget <= 0:
        status = "Kein Budgetwert"
        status_color = STATUS_AMBER
        budget_used = pd.NA
        remaining = pd.NA
    else:
        remaining = annual_budget - annual_actual
        budget_used = annual_actual / annual_budget
        if remaining >= 0:
            status = "Im Budget"
            status_color = STATUS_GREEN
        else:
            status = "Über Budget"
            status_color = STATUS_RED

    top_driver = "n/a"
    if not period_df.empty:
        by_mode = period_df.groupby("transport_mode")[emission_col].sum()
        if not by_mode.empty:
            top_driver = by_mode.idxmax()

    rail_scope = period_df[
        (period_df["transport_mode"] == "flight")
        & (period_df["train_alternative_available"])
    ]

    return {
        "period_df": period_df,
        "annual_df": annual_df,
        "annual_budget": annual_budget,
        "annual_actual": annual_actual,
        "period_actual": period_actual,
        "budget_used": budget_used,
        "remaining": remaining,
        "status": status,
        "status_color": status_color,
        "trips": len(period_df),
        "cost": period_df["cost_CHF"].sum(),
        "top_driver": top_driver,
        "rail_trips": len(rail_scope),
        "rail_emissions": rail_scope[emission_col].sum(),
    }


def budget_summary(actual_df, budgets, filters, group_col):
    emission_col = filters["emission_col"]

    actual = (
        actual_df.groupby(group_col, dropna=False)
        .agg(actual_t=(emission_col, "sum"), trips=(group_col, "size"), cost_chf=("cost_CHF", "sum"))
        .reset_index()
    )

    budget_base = budget_scope(budgets, filters)
    budget = (
        budget_base.groupby(group_col, dropna=False)["co2_budget_t"]
        .sum(min_count=1)
        .reset_index(name="budget_t")
    )

    result = actual.merge(budget, on=group_col, how="outer").fillna({"actual_t": 0, "trips": 0, "cost_chf": 0})
    result["remaining_t"] = result["budget_t"] - result["actual_t"]
    result["budget_used_pct"] = result["actual_t"] / result["budget_t"]
    result["status"] = result.apply(
        lambda row: "Kein Budget"
        if pd.isna(row["budget_t"]) or row["budget_t"] <= 0
        else ("Über Budget" if row["remaining_t"] < 0 else "Im Budget"),
        axis=1,
    )
    return result.sort_values("actual_t", ascending=False)


# -----------------------------------------------------------------------------
# 5. Chart helpers
# -----------------------------------------------------------------------------

def base_chart():
    return alt.Chart().configure_axis(
        labelColor=MID,
        titleColor=MID,
        gridColor=GRID,
        domain=False,
        tickColor=GRID,
    ).configure_view(strokeWidth=0).configure_title(
        color=DARK,
        fontSize=15,
        anchor="start",
    )


def budget_bullet_chart(context):
    budget = context["annual_budget"]
    actual = context["annual_actual"]

    if pd.isna(budget) or budget <= 0:
        max_x = max(actual * 1.15, 1)
        budget = 0
    else:
        max_x = max(actual, budget) * 1.15

    data = pd.DataFrame(
        {
            "label": ["Jahresbudget"],
            "actual_t": [actual],
            "budget_t": [budget],
            "max_x": [max_x],
            "status": [context["status"]],
        }
    )

    background = (
        alt.Chart(data)
        .mark_bar(color=LIGHT, cornerRadius=8, size=34)
        .encode(
            x=alt.X("max_x:Q", title=None, axis=alt.Axis(format=",.0f")),
            y=alt.Y("label:N", title=None, axis=None),
        )
    )
    actual_bar = (
        alt.Chart(data)
        .mark_bar(cornerRadius=8, size=34)
        .encode(
            x=alt.X("actual_t:Q", title="t CO₂e"),
            y=alt.Y("label:N", title=None, axis=None),
            color=alt.value(context["status_color"]),
            tooltip=[
                alt.Tooltip("actual_t:Q", title="Ist", format=",.1f"),
                alt.Tooltip("budget_t:Q", title="Budget", format=",.1f"),
                alt.Tooltip("status:N", title="Status"),
            ],
        )
    )
    budget_rule_halo = (
        alt.Chart(data)
        .mark_rule(color="#f8fafc", strokeWidth=8, opacity=0.96)
        .encode(x="budget_t:Q")
    )
    budget_rule = (
        alt.Chart(data)
        .mark_rule(color="#1f2937", strokeWidth=4, opacity=0.98)
        .encode(x="budget_t:Q")
    )
    budget_label = (
        alt.Chart(data)
        .mark_text(align="left", dx=8, dy=-26, color="#f8fafc", fontSize=13, fontWeight="bold")
        .encode(x="budget_t:Q", y=alt.Y("label:N"), text=alt.value("Budget"))
    )

    return (
        (background + actual_bar + budget_rule_halo + budget_rule + budget_label)
        .properties(height=105)
        .configure_view(strokeWidth=0)
    )


def grouped_budget_chart(summary, group_col):
    if summary.empty:
        return None

    chart_data = summary.copy().sort_values("actual_t", ascending=False)
    chart_data["Einheit"] = chart_data[group_col].astype(str)

    long_data = chart_data.melt(
        id_vars=["Einheit", "status", "budget_used_pct", "remaining_t"],
        value_vars=["actual_t", "budget_t"],
        var_name="metric",
        value_name="tonnes",
    ).dropna(subset=["tonnes"])
    
    # Hier verwenden wir direkt deutsche Bezeichnungen für die Anzeige
    long_data["metric"] = long_data["metric"].map(
        {"actual_t": "Ist CO₂e", "budget_t": "Budget"}
    )

    metric_order = ["Ist CO₂e", "Budget"]
    chart = (
        alt.Chart(long_data)
        .mark_bar(cornerRadiusEnd=3, height=16)
        .encode(
            x=alt.X("tonnes:Q", title=None, axis=alt.Axis(format=",.0f")),
            y=alt.Y("Einheit:N", title=None, sort=list(chart_data["Einheit"])),
            yOffset=alt.YOffset("metric:N", sort=metric_order),
            color=alt.Color(
                "metric:N",
                title=None,
                sort=metric_order,
                scale=alt.Scale(domain=metric_order, range=[BLUE, "#9aa0a8"]),
            ),
            tooltip=[
                alt.Tooltip("Einheit:N", title="Einheit"),
                alt.Tooltip("metric:N", title="Metrik"),
                alt.Tooltip("tonnes:Q", title="t CO₂e", format=",.1f"),
                alt.Tooltip("budget_used_pct:Q", title="Budgetverbrauch", format=".1%"),
                alt.Tooltip("remaining_t:Q", title="Budgetsaldo", format=",.1f"),
                alt.Tooltip("status:N", title="Status"),
            ],
        )
        .properties(
            title="Ist vs. Budget",
            height=max(260, min(520, 72 * len(chart_data))),
            padding={"bottom": 18, "left": 4, "right": 8, "top": 12},
        )
    )
    return chart


def ranked_bar(df, category_col, value_col, title, limit=10, color=BLUE, x_title="t CO₂e"):
    if df.empty:
        return None

    data = (
        df.groupby(category_col, dropna=False)[value_col]
        .sum()
        .reset_index(name="value")
        .sort_values("value", ascending=False)
        .head(limit)
    )

    cat_title = DISPLAY_NAMES_DE.get(category_col, "Kategorie")
    # Datenwerte übersetzen, Fallback: Unterstriche entfernen + Title Case
    data[category_col] = data[category_col].map(translate_value)
    # Spaltennamen auf deutschen Anzeigenamen umbenennen
    data = data.rename(columns={category_col: cat_title})

    return (
        alt.Chart(data)
        .mark_bar(cornerRadiusEnd=5)
        .encode(
            x=alt.X("value:Q", title=x_title),
            y=alt.Y(f"{cat_title}:N", title=None, sort="-x"),
            color=alt.value(color),
            tooltip=[
                alt.Tooltip(f"{cat_title}:N", title=cat_title),
                alt.Tooltip("value:Q", title="t CO₂e", format=",.1f"),
            ],
        )
        .properties(title=title, height=max(220, min(430, 36 * len(data))))
    )


def trend_chart(df, emission_col, group_col=None):
    if df.empty:
        return None

    group_fields = ["year"]
    if group_col:
        group_fields.append(group_col)

    data = (
        df.groupby(group_fields)[emission_col]
        .sum()
        .reset_index(name="emissions_t")
        .sort_values("year")
    )

    enc = {
        "x": alt.X("year:O", title="Jahr"),
        "y": alt.Y("emissions_t:Q", title="t CO₂e"),
        "tooltip": [
            alt.Tooltip("year:O", title="Jahr"),
            alt.Tooltip("emissions_t:Q", title="t CO₂e", format=",.1f"),
        ],
    }
    if group_col:
        enc["color"] = alt.Color(f"{group_col}:N", title=DISPLAY_NAMES_DE.get(group_col, "Gruppe"))
        enc["tooltip"].append(alt.Tooltip(f"{group_col}:N", title=DISPLAY_NAMES_DE.get(group_col, "Gruppe")))

    return alt.Chart(data).mark_line(point=True, strokeWidth=3).encode(**enc).properties(height=320)


def route_opportunity_table(df, emission_col, limit=12):
    if df.empty:
        return pd.DataFrame(columns=["route", "trips", "emissions_t", "cost_chf", "avg_km"])

    routes = df.copy()
    routes["route"] = (
        routes["departure_city"].astype(str)
        + " -> "
        + routes["arrival_city"].astype(str)
    )
    return (
        routes.groupby("route")
        .agg(
            trips=("route", "size"),
            emissions_t=(emission_col, "sum"),
            cost_chf=("cost_CHF", "sum"),
            avg_km=("km", "mean"),
        )
        .reset_index()
        .sort_values("emissions_t", ascending=False)
        .head(limit)
    )


# -----------------------------------------------------------------------------
# 6. Report builders
# -----------------------------------------------------------------------------

def build_cso_report(travel, budgets, filters):
    annual_all = travel[travel["year"] == filters["year"]].copy()
    if filters["month"] is not None:
        period_all = annual_all[annual_all["month"] == filters["month"]].copy()
    else:
        period_all = annual_all.copy()

    tmp_filters = filters.copy()
    tmp_filters["business_unit"] = "Alle"
    tmp_filters["subunit"] = "Alle"

    actual = (
        annual_all.groupby("business_unit")
        .agg(
            annual_emissions_t=(filters["emission_col"], "sum"),
            annual_trips=("business_unit", "size"),
            annual_cost_chf=("cost_CHF", "sum"),
        )
        .reset_index()
    )
    period = (
        period_all.groupby("business_unit")
        .agg(
            report_period_emissions_t=(filters["emission_col"], "sum"),
            report_period_trips=("business_unit", "size"),
        )
        .reset_index()
    )
    budget = (
        budgets[budgets["year"] == filters["year"]]
        .groupby("business_unit")["co2_budget_t"]
        .sum(min_count=1)
        .reset_index(name="budget_t")
    )

    report = actual.merge(period, on="business_unit", how="left").merge(budget, on="business_unit", how="left")
    report["remaining_t"] = report["budget_t"] - report["annual_emissions_t"]
    report["budget_used_pct"] = report["annual_emissions_t"] / report["budget_t"]
    report["status"] = report.apply(
        lambda row: "Kein Budget"
        if pd.isna(row["budget_t"]) or row["budget_t"] <= 0
        else ("Über Budget" if row["remaining_t"] < 0 else "Im Budget"),
        axis=1,
    )
    return report.sort_values("budget_used_pct", ascending=False)


def build_finance_report(period_df, emission_col):
    report = period_df.copy()
    report["cost_per_t_co2e"] = report["cost_CHF"] / report[emission_col].replace(0, pd.NA)

    columns = [
        "date",
        "business_unit",
        "subunit",
        "transport_mode",
        "departure_city",
        "arrival_city",
        "travel_purpose",
        "km",
        "cost_CHF",
        emission_col,
        "cost_per_t_co2e",
        "train_alternative_available",
    ]
    return report[columns].sort_values(["cost_CHF", emission_col], ascending=False)


# -----------------------------------------------------------------------------
# 7. UI rendering helpers
# -----------------------------------------------------------------------------

def inject_css():
    st.markdown(
        """
        <style>
        .block-container {padding-top: 1.5rem;}
        div[data-testid="stMetric"] {
            background: #ffffff;
            border: 1px solid #e5e7eb;
            border-radius: 8px;
            padding: 12px 14px;
        }
        .status-card {
            border: 1px solid #e5e7eb;
            border-radius: 8px;
            padding: 18px 20px;
            background: #ffffff;
            min-height: 128px;
        }
        .status-label {
            font-size: 0.78rem;
            text-transform: uppercase;
            letter-spacing: .04em;
            color: #667085;
            margin-bottom: 6px;
        }
        .status-value {
            font-size: 2rem;
            font-weight: 760;
            line-height: 1.05;
        }
        .context-strip {
            display: grid;
            grid-template-columns: repeat(5, minmax(0, 1fr));
            gap: 14px;
            margin-top: 12px;
        }
        .context-item {
            border-top: 1px solid #e5e7eb;
            padding-top: 9px;
        }
        .context-item-label {
            color: #98a2b3;
            font-size: 0.78rem;
            text-transform: uppercase;
            letter-spacing: .035em;
            margin-bottom: 3px;
        }
        .context-item-value {
            color: var(--text-color, #f8fafc);
            font-size: 1.25rem;
            font-weight: 820;
            line-height: 1.1;
        }
        .delta-pill {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            border-radius: 999px;
            padding: 5px 10px;
            margin-top: 8px;
            font-size: 0.95rem;
            font-weight: 760;
            line-height: 1;
            width: fit-content;
        }
        .delta-positive {
            color: #86efac;
            background: rgba(34, 197, 94, 0.18);
        }
        .delta-negative {
            color: #fca5a5;
            background: rgba(239, 68, 68, 0.18);
        }
        .delta-neutral {
            color: #cbd5e1;
            background: rgba(148, 163, 184, 0.18);
        }
        .small-note {
            color: #667085;
            font-size: 0.9rem;
        }
        @media (max-width: 900px) {
            .context-strip { grid-template-columns: repeat(2, minmax(0, 1fr)); }
        }
        .report-summary {
            display: grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            gap: 18px;
            margin: 12px 0 22px 0;
        }
        .report-summary.three {
            grid-template-columns: repeat(3, minmax(0, 1fr));
        }
        .report-item {
            border-top: 2px solid #3b4250;
            padding-top: 10px;
        }
        .report-label {
            color: #98a2b3;
            font-size: 0.82rem;
            font-weight: 680;
            text-transform: uppercase;
            letter-spacing: .035em;
            margin-bottom: 6px;
        }
        .report-value {
            color: var(--text-color, #f8fafc);
            font-size: 1.65rem;
            font-weight: 820;
            line-height: 1.08;
        }
        .report-note {
            color: #98a2b3;
            font-size: 0.92rem;
            margin-top: 6px;
        }
        @media (max-width: 900px) {
            .report-summary,
            .report-summary.three { grid-template-columns: repeat(1, minmax(0, 1fr)); }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_report_summary(items, columns=4):
    css_class = "report-summary three" if columns == 3 else "report-summary"
    html_items = []
    for item in items:
        note = f'<div class="report-note">{item.get("note", "")}</div>' if item.get("note") else ""
        html_items.append(
            '<div class="report-item">'
            f'<div class="report-label">{item["label"]}</div>'
            f'<div class="report-value">{item["value"]}</div>'
            f"{note}"
            "</div>"
        )
    html = f'<div class="{css_class}">{"".join(html_items)}</div>'
    st.markdown(html, unsafe_allow_html=True)


def render_status_context(context, filters):
    st.subheader("Kontext")
    delta = budget_delta(context)

    c1, c2 = st.columns([0.86, 1.14])
    with c1:
        st.markdown(
            f"""
            <div class="status-card">
                <div class="status-label">Budgetstatus {filters["year"]}</div>
                <div class="status-value" style="color:{context["status_color"]};">{context["status"]}</div>
                <div class="small-note">Jahres-Ist: {format_tonnes(context["annual_actual"])} /
                Budget: {format_tonnes(context["annual_budget"])}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c2:
        st.altair_chart(budget_bullet_chart(context), use_container_width=True)

    st.markdown(
        f"""
        <div class="context-strip">
            <div class="context-item">
                <div class="context-item-label">Budgetverbrauch</div>
                <div class="context-item-value">{format_pct(context["budget_used"])}</div>
            </div>
            <div class="context-item">
                <div class="context-item-label">Budgetsaldo</div>
                <div class="context-item-value">{format_tonnes(context["remaining"])}</div>
                <div class="delta-pill delta-{delta["class"]}">
                    <span>{delta["arrow"]}</span>
                    <span>{delta["text"]}</span>
                </div>
            </div>
            <div class="context-item">
                <div class="context-item-label">Zeitraum CO₂e</div>
                <div class="context-item-value">{format_tonnes(context["period_actual"])}</div>
            </div>
            <div class="context-item">
                <div class="context-item-label">Reisen</div>
                <div class="context-item-value">{format_number(context["trips"], 0)}</div>
            </div>
            <div class="context-item">
                <div class="context-item-label">Rail-First Chancen</div>
                <div class="context-item-value">{format_number(context["rail_trips"], 0)}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if filters["month"] is not None:
        st.caption(
            "Hinweis: Der Status vergleicht die Jahres-Ist-Emissionen mit dem Jahresbudget. "
            "Die übrigen Kennzahlen und Charts nutzen den gewählten Monatsfilter."
        )


def render_empty(message="Keine Daten für die gewählten Filter."):
    st.info(message)


# -----------------------------------------------------------------------------
# 8. View renderers
# -----------------------------------------------------------------------------

def render_bu_view(travel, budgets, filters, context):
    period_df = context["period_df"]
    annual_df = context["annual_df"]
    emission_col = filters["emission_col"]

    render_status_context(context, filters)

    st.subheader("Insights & Massnahmen")
    tab_budget, tab_drivers, tab_rail, tab_trend, tab_details = st.tabs(
        ["Budgetstatus", "Treiber", "Rail-First", "Trend", "Details"]
    )

    with tab_budget:
        group_col = "business_unit" if filters["business_unit"] == "Alle" else "subunit"
        summary = budget_summary(annual_df, budgets, filters, group_col)
        
        # Umbenennung des Gruppen-Keys für das Dictionary Mapping
        summary_display = summary.copy()
        if group_col in summary_display.columns:
            summary_display.rename(columns={group_col: "Einheit"}, inplace=True)
            
        chart = grouped_budget_chart(summary, group_col)
        if chart is None:
            render_empty()
        else:
            st.altair_chart(chart, use_container_width=True)
            st.dataframe(
                summary_display.rename(columns=DISPLAY_NAMES_DE),
                use_container_width=True,
                hide_index=True,
            )

    with tab_drivers:
        if period_df.empty:
            render_empty()
        else:
            col1, col2 = st.columns(2)
            with col1:
                st.altair_chart(
                    ranked_bar(period_df, "transport_mode", emission_col, "Emissionen nach Verkehrsmittel", limit=8),
                    use_container_width=True,
                )
                st.altair_chart(
                    ranked_bar(period_df, "haul", emission_col, "Emissionen nach Distanzklasse", limit=8, color="#6f59d9"),
                    use_container_width=True,
                )
            with col2:
                st.altair_chart(
                    ranked_bar(period_df, "travel_purpose", emission_col, "Emissionen nach Reisezweck", limit=8, color="#008f9c"),
                    use_container_width=True,
                )

            st.caption(
                f"Groesster Treiber im aktuellen Filter: {context['top_driver']}. "
                "Die Balken sind bewusst sortiert, damit Hotspots schnell erkennbar sind."
            )

    with tab_rail:
        rail_df = period_df[
            (period_df["transport_mode"] == "flight")
            & (period_df["train_alternative_available"])
        ].copy()

        st.markdown(
            f"Im aktuellen Filter gibt es **{format_number(len(rail_df), 0)} Flugreisen mit Bahnalternative** "
            f"mit **{format_tonnes(rail_df[emission_col].sum())}** und "
            f"**{format_chf(rail_df['cost_CHF'].sum())}** Reisekosten."
        )

        routes = route_opportunity_table(rail_df, emission_col)
        if routes.empty:
            render_empty("Keine Rail-First Chancen für die gewählten Filter.")
        else:
            st.altair_chart(
                ranked_bar(routes, "route", "emissions_t", "Top-Routen mit Bahnalternative", limit=12, color="#d97706"),
                use_container_width=True,
            )
            st.dataframe(
                routes.rename(columns=DISPLAY_NAMES_DE),
                use_container_width=True,
                hide_index=True,
            )

    with tab_trend:
        trend_scope = travel.copy()
        if filters["business_unit"] != "Alle":
            trend_scope = trend_scope[trend_scope["business_unit"] == filters["business_unit"]]
        if filters["subunit"] != "Alle":
            trend_scope = trend_scope[trend_scope["subunit"] == filters["subunit"]]
        st.altair_chart(trend_chart(trend_scope, emission_col), use_container_width=True)

    with tab_details:
        if period_df.empty:
            render_empty()
        else:
            detail_cols = [
                "date",
                "business_unit",
                "subunit",
                "transport_mode",
                "departure_city",
                "arrival_city",
                "travel_purpose",
                "km",
                "cost_CHF",
                emission_col,
                "train_alternative_available",
            ]
            st.dataframe(
                period_df[detail_cols].sort_values(emission_col, ascending=False).head(300).rename(columns=DISPLAY_NAMES_DE),
                use_container_width=True,
                hide_index=True,
            )


def render_cso_report(travel, budgets, filters, context):
    company_filters = filters.copy()
    company_filters["business_unit"] = "Alle"
    company_filters["subunit"] = "Alle"
    company_context = calculate_context(travel, budgets, company_filters)
    cso_report = build_cso_report(travel, budgets, filters)

    header_left, header_right = st.columns([0.72, 0.28])
    with header_left:
        st.subheader("CSO Report")
        st.caption("Monatlicher Summary-Report für ESG-Reporting und Netto-Null-Steuerung.")
    with header_right:
        st.download_button(
            "CSO Report CSV",
            data=to_csv_bytes(cso_report),
            file_name=f"cso_report_{filters['year']}_{filters['rfi_label'].replace(' ', '_')}.csv",
            mime="text/csv",
            use_container_width=True,
            key="cso_report_download_top",
        )

    tab_overview, tab_budget, tab_trend, tab_export = st.tabs(
        ["CSO Summary", "BU-Vergleich", "Unternehmens-Trend", "Export"]
    )

    with tab_overview:
        valid_budget = not pd.isna(company_context["annual_budget"]) and company_context["annual_budget"] > 0
        worst_bu = "n/a"
        if not cso_report.empty and cso_report["remaining_t"].notna().any():
            worst_bu = cso_report.sort_values("remaining_t").iloc[0]["business_unit"]

        render_report_summary(
            [
                {
                    "label": "Unternehmens-CO₂e",
                    "value": format_tonnes(company_context["annual_actual"]),
                    "note": f"{filters['year']} / {filters['rfi_label']}",
                },
                {
                    "label": "Unternehmensbudget",
                    "value": format_tonnes(company_context["annual_budget"]),
                    "note": "2020-2025 vollständig interpretierbar",
                },
                {
                    "label": "Budgetverbrauch",
                    "value": format_pct(company_context["budget_used"]) if valid_budget else "n/a",
                    "note": company_context["status"],
                },
                {
                    "label": "Grösste Abweichung",
                    "value": str(worst_bu),
                    "note": "nach Budgetsaldo",
                },
            ]
        )
        st.markdown("#### Aggregierte Daten für den Nachhaltigkeitsbericht")
        st.dataframe(
            cso_report.rename(columns=DISPLAY_NAMES_DE),
            use_container_width=True,
            hide_index=True,
        )

    chart_filters = filters.copy()
    chart_filters["business_unit"] = "Alle"
    chart_filters["subunit"] = "Alle"
    annual_all = travel[travel["year"] == filters["year"]].copy()
    summary = budget_summary(annual_all, budgets, chart_filters, "business_unit")

    with tab_budget:
        st.markdown("#### Actual vs Budget nach Business Unit")
        st.altair_chart(grouped_budget_chart(summary, "business_unit"), use_container_width=True)

    with tab_trend:
        st.markdown("#### Unternehmensweite Entwicklung")
        st.altair_chart(trend_chart(travel, filters["emission_col"], group_col="business_unit"), use_container_width=True)

    with tab_export:
        st.download_button(
            "CSO Report CSV herunterladen",
            data=to_csv_bytes(cso_report),
            file_name=f"cso_report_{filters['year']}_{filters['rfi_label'].replace(' ', '_')}.csv",
            mime="text/csv",
            use_container_width=True,
            key="cso_report_download_export",
        )
        st.dataframe(cso_report.rename(columns=DISPLAY_NAMES_DE), use_container_width=True, hide_index=True)


def render_finance_report(filters, context):
    period_df = context["period_df"]
    emission_col = filters["emission_col"]
    finance_report = build_finance_report(period_df, emission_col)

    header_left, header_right = st.columns([0.72, 0.28])
    with header_left:
        st.subheader("Finance Report")
        st.caption("Exportierbare Datensicht für Audit, Kosten-Emissions-Abgleich und Stichproben.")
    with header_right:
        st.download_button(
            "Finance Detailreport CSV",
            data=to_csv_bytes(finance_report),
            file_name=f"finance_report_{filters['year']}_{filters['rfi_label'].replace(' ', '_')}.csv",
            mime="text/csv",
            use_container_width=True,
            key="finance_report_download_top",
        )

    tab_quick, tab_alignment, tab_sample, tab_export = st.tabs(
        ["Finance Quick View", "Kosten-CO₂e", "Stichprobe", "Export"]
    )

    if period_df.empty:
        with tab_quick:
            render_empty()
        return

    summary = (
        period_df.groupby("transport_mode")
        .agg(
            trips=("transport_mode", "size"),
            emissions_t=(emission_col, "sum"),
            cost_chf=("cost_CHF", "sum"),
            km=("km", "sum"),
        )
        .reset_index()
    )
    summary["cost_per_t_co2e"] = summary["cost_chf"] / summary["emissions_t"].replace(0, pd.NA)

    with tab_quick:
        cost_per_t = period_df["cost_CHF"].sum() / period_df[emission_col].sum() if period_df[emission_col].sum() > 0 else pd.NA
        high_cost_mode = summary.sort_values("cost_chf", ascending=False).iloc[0]["transport_mode"]
        high_emission_mode = summary.sort_values("emissions_t", ascending=False).iloc[0]["transport_mode"]

        render_report_summary(
            [
                {
                    "label": "Kosten total",
                    "value": format_chf(period_df["cost_CHF"].sum()),
                    "note": f"{filters['month_label']} {filters['year']}",
                },
                {
                    "label": "Kosten pro t CO₂e",
                    "value": format_chf(cost_per_t),
                    "note": "Kosten-Emissions-Alignment",
                },
                {
                    "label": "Ø Kosten pro Reise",
                    "value": format_chf(period_df["cost_CHF"].mean()),
                    "note": f"{format_number(len(period_df), 0)} Transaktionen",
                },
                {
                    "label": "Haupttreiber",
                    "value": str(high_cost_mode),
                    "note": f"Emissionstreiber: {high_emission_mode}",
                },
            ]
        )

        st.dataframe(
            summary.rename(columns=DISPLAY_NAMES_DE),
            use_container_width=True,
            hide_index=True,
        )

    with tab_alignment:
        st.markdown("#### Kosten-Emissions-Alignment")
        left, right = st.columns([0.45, 0.55])
        with left:
            st.altair_chart(
                ranked_bar(
                    summary,
                    "transport_mode",
                    "cost_chf",
                    "Kosten nach Verkehrsmittel",
                    limit=8,
                    color="#344054",
                    x_title="CHF",
                ),
                use_container_width=True,
            )
        with right:
            st.altair_chart(
                ranked_bar(
                    summary,
                    "transport_mode",
                    "emissions_t",
                    "Emissionen nach Verkehrsmittel",
                    limit=8,
                    color=BLUE,
                ),
                use_container_width=True,
            )

    with tab_sample:
        st.markdown("#### High-cost / high-emission Stichprobe")
        st.dataframe(finance_report.head(300).rename(columns=DISPLAY_NAMES_DE), use_container_width=True, hide_index=True)

    with tab_export:
        st.download_button(
            "Finance Detailreport CSV",
            data=to_csv_bytes(finance_report),
            file_name=f"finance_report_{filters['year']}_{filters['rfi_label'].replace(' ', '_')}.csv",
            mime="text/csv",
            use_container_width=True,
            key="finance_report_download_export",
        )
        st.dataframe(finance_report.head(300).rename(columns=DISPLAY_NAMES_DE), use_container_width=True, hide_index=True)


# -----------------------------------------------------------------------------
# 9. Main app
# -----------------------------------------------------------------------------

def main():
    inject_css()
    travel, budgets = load_data(_mtimes=_file_mtimes())
    filters = sidebar_filters(travel)
    context = calculate_context(travel, budgets, filters)

    st.title("Geschäftsreisen CO₂ Dashboard")
    st.caption(
        "Operatives Steuerungstool für BU-Leitungen mit kompakten Report-Sichten für CSO und Finance."
    )

    if filters["year"] not in range(2020, 2026):
        st.warning(
            "Für dieses Jahr sind keine vollständigen Budgetwerte vorhanden. "
            "Budgetvergleiche sind am aussagekräftigsten für 2020-2025."
        )

    if filters["view"] == "BU-Leitung":
        render_bu_view(travel, budgets, filters, context)
    elif filters["view"] == "CSO Report":
        render_cso_report(travel, budgets, filters, context)
    else:
        render_finance_report(filters, context)


if __name__ == "__main__":
    main()
