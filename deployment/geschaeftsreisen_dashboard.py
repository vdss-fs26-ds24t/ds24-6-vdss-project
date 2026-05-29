from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st


APP_DIR = Path(__file__).parent
DATA_DIR = APP_DIR / "data"
TRAVEL_CSV = DATA_DIR / "travel_data.csv"
BUDGET_CSV = DATA_DIR / "co2_budgets.csv"
EXCEL_FALLBACK = APP_DIR / "traveldata-export.xlsx"

RFI_COLUMNS = {
    "RFI 2.0": "CO2e RFI2 (t)",
    "RFI 2.7": "CO2e RFI2.7 (t)",
}


st.set_page_config(
    page_title="Geschaeftsreisen CO2 Dashboard",
    page_icon="",
    layout="wide",
)


@st.cache_data(show_spinner=False)
def load_data():
    if TRAVEL_CSV.exists() and BUDGET_CSV.exists():
        travel = pd.read_csv(TRAVEL_CSV)
        budgets = pd.read_csv(BUDGET_CSV)
    elif EXCEL_FALLBACK.exists():
        travel = pd.read_excel(EXCEL_FALLBACK, sheet_name="travel_data")
        budgets = pd.read_excel(EXCEL_FALLBACK, sheet_name="co2_budgets")
    else:
        st.error(
            "No data found. Expected data/travel_data.csv and data/co2_budgets.csv "
            "next to this Streamlit file."
        )
        st.stop()

    travel["date"] = pd.to_datetime(travel["date"], errors="coerce")
    travel["year"] = travel["date"].dt.year

    numeric_cols = [
        "CO2e RFI2 (t)",
        "CO2e RFI2.7 (t)",
        "km",
        "cost_CHF",
    ]
    for col in numeric_cols:
        travel[col] = pd.to_numeric(travel[col], errors="coerce").fillna(0)

    travel["train_alternative_available"] = (
        travel["train_alternative_available"].astype(str).str.lower().eq("true")
    )

    budgets["year"] = pd.to_numeric(budgets["year"], errors="coerce").astype("Int64")
    budgets["co2_budget_t"] = pd.to_numeric(budgets["co2_budget_t"], errors="coerce")

    subunit_map = (
        travel[["subunit", "business_unit"]]
        .dropna()
        .drop_duplicates()
        .sort_values(["business_unit", "subunit"])
    )
    budgets = budgets.merge(subunit_map, on="subunit", how="left")

    return travel, budgets


def format_tonnes(value):
    if pd.isna(value):
        return "n/a"
    return f"{value:,.1f} t".replace(",", "'")


def format_chf(value):
    if pd.isna(value):
        return "n/a"
    return f"CHF {value:,.0f}".replace(",", "'")


def kpi_delta_text(actual, budget):
    if pd.isna(budget) or budget <= 0:
        return None
    delta = budget - actual
    pct = delta / budget
    sign = "+" if delta >= 0 else ""
    return f"{sign}{delta:,.1f} t ({sign}{pct:.1%})".replace(",", "'")


def bar_chart(data, x, y, color=None, title=None, height=320, sort="-x"):
    chart = (
        alt.Chart(data)
        .mark_bar(cornerRadiusTopRight=3, cornerRadiusBottomRight=3)
        .encode(
            x=alt.X(x, title=None),
            y=alt.Y(y, title=None, sort=sort),
            tooltip=list(data.columns),
        )
        .properties(title=title, height=height)
    )
    if color:
        chart = chart.encode(color=alt.Color(color, title=None))
    return chart


def empty_chart_message(message):
    st.info(message)


def summarize_budget(actual_df, budget_df, group_col, emission_col, year):
    actual = (
        actual_df.groupby(group_col, dropna=False)[emission_col]
        .sum()
        .reset_index(name="actual_t")
    )
    budget = (
        budget_df[budget_df["year"] == year]
        .groupby(group_col, dropna=False)["co2_budget_t"]
        .sum(min_count=1)
        .reset_index(name="budget_t")
    )
    result = actual.merge(budget, on=group_col, how="outer").fillna({"actual_t": 0})
    result["remaining_t"] = result["budget_t"] - result["actual_t"]
    result["budget_used_pct"] = result["actual_t"] / result["budget_t"]
    result["status"] = result.apply(
        lambda row: "No budget"
        if pd.isna(row["budget_t"]) or row["budget_t"] <= 0
        else ("Over budget" if row["remaining_t"] < 0 else "On track"),
        axis=1,
    )
    return result.sort_values("actual_t", ascending=False)


travel, budgets = load_data()

years = sorted(int(y) for y in travel["year"].dropna().unique())
business_units = sorted(travel["business_unit"].dropna().unique())
transport_modes = sorted(travel["transport_mode"].dropna().unique())
travel_purposes = sorted(travel["travel_purpose"].dropna().unique())
haul_types = sorted(travel["haul"].dropna().unique())

default_year = 2025 if 2025 in years else years[-1]

st.title("Geschaeftsreisen CO2 Dashboard")
st.caption("MVP fuer BU-Heads: Budgetstatus, Emissionstreiber und Rail-First Chancen.")

with st.sidebar:
    st.header("Filter")
    selected_bu = st.selectbox("Business Unit", business_units + ["All"], index=0)
    selected_year = st.selectbox(
        "Jahr",
        years,
        index=years.index(default_year),
        help="Budgetdaten sind im Datensatz fuer 2020-2025 vorhanden.",
    )
    selected_rfi = st.radio("CO2e Szenario", list(RFI_COLUMNS.keys()), index=1)
    emission_col = RFI_COLUMNS[selected_rfi]

    st.divider()
    selected_modes = st.multiselect(
        "Transportmodus",
        transport_modes,
        default=transport_modes,
    )
    selected_purposes = st.multiselect(
        "Reisezweck",
        travel_purposes,
        default=travel_purposes,
    )
    selected_haul = st.multiselect("Haul", haul_types, default=haul_types)

    st.divider()
    st.caption("Primaere Zielgruppe: BU-Heads. CSO- und Finance-Sichten sind bewusst schlank gehalten.")


base_filtered = travel[
    (travel["year"] == selected_year)
    & (travel["transport_mode"].isin(selected_modes))
    & (travel["travel_purpose"].isin(selected_purposes))
    & (travel["haul"].isin(selected_haul))
].copy()

if selected_bu != "All":
    filtered = base_filtered[base_filtered["business_unit"] == selected_bu].copy()
    budget_scope = budgets[budgets["business_unit"] == selected_bu].copy()
else:
    filtered = base_filtered.copy()
    budget_scope = budgets.copy()

total_emissions = filtered[emission_col].sum()
total_trips = len(filtered)
total_cost = filtered["cost_CHF"].sum()
train_alt_share = (
    filtered["train_alternative_available"].mean() if len(filtered) else float("nan")
)
total_budget = budget_scope.loc[budget_scope["year"] == selected_year, "co2_budget_t"].sum(
    min_count=1
)

if pd.isna(total_budget) or total_budget <= 0:
    budget_status = "No budget data"
else:
    budget_status = "Over budget" if total_emissions > total_budget else "On track"

st.subheader("BU-Status")
kpi_cols = st.columns(5)
kpi_cols[0].metric("CO2e Ist", format_tonnes(total_emissions))
kpi_cols[1].metric(
    "CO2e Budget",
    format_tonnes(total_budget),
    delta=kpi_delta_text(total_emissions, total_budget),
)
kpi_cols[2].metric("Status", budget_status)
kpi_cols[3].metric("Reisen", f"{total_trips:,}".replace(",", "'"))
kpi_cols[4].metric("Kosten", format_chf(total_cost))

if selected_year not in range(2020, 2026):
    st.warning(
        "Fuer dieses Jahr sind keine vollstaendigen Budgetwerte vorhanden. "
        "Budgetvergleiche sind am aussagekraeftigsten fuer 2020-2025."
    )

tab_bu, tab_drivers, tab_rail, tab_cso, tab_finance = st.tabs(
    [
        "Budgetsteuerung",
        "Emissionstreiber",
        "Rail-First Chancen",
        "CSO Vergleich",
        "Finance Quick View",
    ]
)

with tab_bu:
    st.markdown("#### Actual vs Budget")

    group_col = "business_unit" if selected_bu == "All" else "subunit"
    budget_summary = summarize_budget(
        filtered,
        budget_scope,
        group_col=group_col,
        emission_col=emission_col,
        year=selected_year,
    )

    if budget_summary.empty:
        empty_chart_message("Keine Daten fuer die gewaehlten Filter.")
    else:
        budget_long = budget_summary.melt(
            id_vars=[group_col, "status"],
            value_vars=["actual_t", "budget_t"],
            var_name="metric",
            value_name="tonnes",
        ).dropna(subset=["tonnes"])
        budget_long["metric"] = budget_long["metric"].map(
            {"actual_t": "Actual CO2e", "budget_t": "Budget"}
        )

        chart = (
            alt.Chart(budget_long)
            .mark_bar()
            .encode(
                x=alt.X("tonnes:Q", title="t CO2e"),
                y=alt.Y(f"{group_col}:N", title=None, sort="-x"),
                color=alt.Color(
                    "metric:N",
                    title=None,
                    scale=alt.Scale(range=["#2F6BFF", "#8A8F98"]),
                ),
                yOffset="metric:N",
                tooltip=[
                    alt.Tooltip(f"{group_col}:N", title="Einheit"),
                    alt.Tooltip("metric:N", title="Metrik"),
                    alt.Tooltip("tonnes:Q", title="t CO2e", format=",.1f"),
                    alt.Tooltip("status:N", title="Status"),
                ],
            )
            .properties(height=360)
        )
        st.altair_chart(chart, use_container_width=True)

        table = budget_summary.copy()
        table["budget_used_pct"] = table["budget_used_pct"].replace([float("inf")], pd.NA)
        st.dataframe(
            table.rename(
                columns={
                    group_col: "Einheit",
                    "actual_t": "Actual CO2e t",
                    "budget_t": "Budget t",
                    "remaining_t": "Remaining t",
                    "budget_used_pct": "Budget used %",
                    "status": "Status",
                }
            ),
            use_container_width=True,
            hide_index=True,
        )

with tab_drivers:
    st.markdown("#### Was treibt die Emissionen?")
    if filtered.empty:
        empty_chart_message("Keine Daten fuer die gewaehlten Filter.")
    else:
        c1, c2, c3 = st.columns(3)

        by_mode = (
            filtered.groupby("transport_mode")[emission_col]
            .sum()
            .reset_index(name="emissions_t")
            .sort_values("emissions_t", ascending=False)
        )
        by_purpose = (
            filtered.groupby("travel_purpose")[emission_col]
            .sum()
            .reset_index(name="emissions_t")
            .sort_values("emissions_t", ascending=False)
        )
        by_haul = (
            filtered.groupby("haul")[emission_col]
            .sum()
            .reset_index(name="emissions_t")
            .sort_values("emissions_t", ascending=False)
        )

        with c1:
            st.altair_chart(
                bar_chart(
                    by_mode,
                    "emissions_t:Q",
                    "transport_mode:N",
                    title="Nach Transportmodus",
                    height=280,
                ),
                use_container_width=True,
            )
        with c2:
            st.altair_chart(
                bar_chart(
                    by_purpose,
                    "emissions_t:Q",
                    "travel_purpose:N",
                    title="Nach Reisezweck",
                    height=280,
                ),
                use_container_width=True,
            )
        with c3:
            st.altair_chart(
                bar_chart(
                    by_haul,
                    "emissions_t:Q",
                    "haul:N",
                    title="Nach Haul",
                    height=280,
                ),
                use_container_width=True,
            )

        trend_scope = travel[
            (travel["transport_mode"].isin(selected_modes))
            & (travel["travel_purpose"].isin(selected_purposes))
            & (travel["haul"].isin(selected_haul))
        ].copy()
        if selected_bu != "All":
            trend_scope = trend_scope[trend_scope["business_unit"] == selected_bu]

        trend = (
            trend_scope.groupby("year")[emission_col]
            .sum()
            .reset_index(name="emissions_t")
            .sort_values("year")
        )
        trend_chart = (
            alt.Chart(trend)
            .mark_line(point=True)
            .encode(
                x=alt.X("year:O", title="Jahr"),
                y=alt.Y("emissions_t:Q", title="t CO2e"),
                tooltip=[
                    alt.Tooltip("year:O", title="Jahr"),
                    alt.Tooltip("emissions_t:Q", title="t CO2e", format=",.1f"),
                ],
            )
            .properties(title="CO2e Entwicklung", height=280)
        )
        st.altair_chart(trend_chart, use_container_width=True)

with tab_rail:
    st.markdown("#### Rail-First Chancen")
    rail_scope = filtered[
        (filtered["transport_mode"] == "flight")
        & (filtered["train_alternative_available"])
    ].copy()

    r1, r2, r3 = st.columns(3)
    r1.metric("Fluege mit Bahnalternative", f"{len(rail_scope):,}".replace(",", "'"))
    r2.metric("CO2e in dieser Gruppe", format_tonnes(rail_scope[emission_col].sum()))
    r3.metric("Kosten in dieser Gruppe", format_chf(rail_scope["cost_CHF"].sum()))

    if rail_scope.empty:
        empty_chart_message("Keine passenden Fluege fuer die gewaehlten Filter.")
    else:
        rail_scope["route"] = (
            rail_scope["departure_city"].fillna(rail_scope["departure_iata"].astype(str))
            + " -> "
            + rail_scope["arrival_city"].fillna(rail_scope["arrival_iata"].astype(str))
        )
        route_summary = (
            rail_scope.groupby("route")
            .agg(
                trips=("route", "size"),
                emissions_t=(emission_col, "sum"),
                cost_chf=("cost_CHF", "sum"),
                km=("km", "mean"),
            )
            .reset_index()
            .sort_values("emissions_t", ascending=False)
            .head(15)
        )

        st.altair_chart(
            bar_chart(
                route_summary,
                "emissions_t:Q",
                "route:N",
                title="Top Routen mit Bahnalternative",
                height=420,
            ),
            use_container_width=True,
        )
        st.dataframe(
            route_summary.rename(
                columns={
                    "route": "Route",
                    "trips": "Trips",
                    "emissions_t": "CO2e t",
                    "cost_chf": "Cost CHF",
                    "km": "Avg km",
                }
            ),
            use_container_width=True,
            hide_index=True,
        )

with tab_cso:
    st.markdown("#### CSO Vergleich: Business Units")
    cso_budget = summarize_budget(
        base_filtered,
        budgets,
        group_col="business_unit",
        emission_col=emission_col,
        year=selected_year,
    )
    if cso_budget.empty:
        empty_chart_message("Keine Daten fuer die gewaehlten Filter.")
    else:
        cso_budget["over_budget_t"] = (cso_budget["actual_t"] - cso_budget["budget_t"]).clip(
            lower=0
        )
        st.altair_chart(
            bar_chart(
                cso_budget,
                "actual_t:Q",
                "business_unit:N",
                color="status:N",
                title="CO2e nach Business Unit",
                height=300,
            ),
            use_container_width=True,
        )
        st.dataframe(
            cso_budget.rename(
                columns={
                    "business_unit": "Business Unit",
                    "actual_t": "Actual CO2e t",
                    "budget_t": "Budget t",
                    "remaining_t": "Remaining t",
                    "budget_used_pct": "Budget used %",
                    "status": "Status",
                    "over_budget_t": "Over budget t",
                }
            ),
            use_container_width=True,
            hide_index=True,
        )

with tab_finance:
    st.markdown("#### Finance Quick View")
    if filtered.empty:
        empty_chart_message("Keine Daten fuer die gewaehlten Filter.")
    else:
        finance = filtered.copy()
        finance["cost_per_t"] = finance["cost_CHF"] / finance[emission_col].replace(0, pd.NA)
        f1, f2, f3 = st.columns(3)
        f1.metric("Kosten total", format_chf(finance["cost_CHF"].sum()))
        f2.metric(
            "Kosten pro t CO2e",
            format_chf(finance["cost_CHF"].sum() / total_emissions)
            if total_emissions > 0
            else "n/a",
        )
        f3.metric("Durchschnittskosten pro Reise", format_chf(finance["cost_CHF"].mean()))

        finance_summary = (
            finance.groupby("transport_mode")
            .agg(
                trips=("transport_mode", "size"),
                emissions_t=(emission_col, "sum"),
                cost_chf=("cost_CHF", "sum"),
            )
            .reset_index()
        )
        finance_summary["cost_per_t"] = (
            finance_summary["cost_chf"] / finance_summary["emissions_t"].replace(0, pd.NA)
        )
        st.dataframe(
            finance_summary.sort_values("cost_chf", ascending=False).rename(
                columns={
                    "transport_mode": "Transport mode",
                    "trips": "Trips",
                    "emissions_t": "CO2e t",
                    "cost_chf": "Cost CHF",
                    "cost_per_t": "Cost per t CO2e",
                }
            ),
            use_container_width=True,
            hide_index=True,
        )

        st.markdown("##### High-cost / high-emission trips")
        detail_cols = [
            "date",
            "business_unit",
            "subunit",
            "transport_mode",
            "departure_city",
            "arrival_city",
            "travel_purpose",
            "cost_CHF",
            emission_col,
            "km",
            "train_alternative_available",
        ]
        detail = (
            finance[detail_cols]
            .sort_values([emission_col, "cost_CHF"], ascending=False)
            .head(100)
        )
        st.dataframe(detail, use_container_width=True, hide_index=True)

