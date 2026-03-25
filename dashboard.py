import streamlit as st
import pandas as pd
from PIL import Image
import plotly.express as px
from io import BytesIO

# =================================================
# PAGE CONFIG
# =================================================
st.set_page_config(page_title="Employee Dashboard", layout="wide")

# =================================================
# LOGO (MAIN PAGE)
# =================================================
logo = Image.open("logo.png")
st.columns([1, 6])[0].image(logo, width=120)
st.title("Employee Dashboard")

# =================================================
# HELPER FUNCTIONS
# =================================================
def add_sr_no(df):
    df = df.reset_index(drop=True)
    df.insert(0, "Sr. No", range(1, len(df) + 1))
    return df

def prettify_columns(df):
    return df.rename(columns=lambda x: x.title().replace(" ", "_"))

def to_excel_bytes(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Data")
    output.seek(0)
    return output

def clean_filename(text):
    return (
        text.replace(" ", "-")
        .replace("/", "-")
        .replace("–", "-")
        .replace("—", "-")
    )

def show_table(df):
    df = prettify_columns(df)
    st.dataframe(
        add_sr_no(df),
        use_container_width=True,
        hide_index=True,
        column_config={
            "Sr. No": st.column_config.NumberColumn("Sr. No", width="small"),
            "Project_Count": st.column_config.NumberColumn("Project_Count", width="small"),
            "Total_Hours": st.column_config.NumberColumn("Total_Hours", width="small"),
            "Leave_Days": st.column_config.NumberColumn("Leave_Days", width="small"),
            "Total_Leaves": st.column_config.NumberColumn("Total_Leaves", width="small"),
        }
    )

# =================================================
# LOAD EXCEL
# =================================================
df_raw = pd.read_excel(
    "Project Summary & Billability details.xlsm",
    engine="openpyxl"
)

df_raw = df_raw.loc[:, ~df_raw.columns.str.contains("^Unnamed", na=False)]

# =================================================
# DETECT DATE COLUMNS
# =================================================
rename_map, date_cols = {}, []

for col in df_raw.columns:
    parsed = pd.to_datetime(col, errors="coerce")
    if not pd.isna(parsed):
        rename_map[col] = parsed
        date_cols.append(parsed)

df_raw = df_raw.rename(columns=rename_map)

if not date_cols:
    st.error("No valid date columns found.")
    st.stop()

# =================================================
# WIDE → LONG
# =================================================
records = []

for _, row in df_raw.iterrows():
    for d in date_cols:
        val = row[d]
        hours, leave = 0, None

        if isinstance(val, (int, float)):
            hours = val
        elif isinstance(val, str):
            v = val.strip().upper()
            if v in ["AL", "CL", "PH"]:
                leave = v
            elif v.isdigit():
                hours = int(v)

        records.append({
            "employee_name": row.get("Employee Name"),
            "project_name": row.get("Project name"),
            "mars_project_name": row.get("MARS Project Name"),
            "date": d,
            "hours": hours,
            "leave_type": leave
        })

df = pd.DataFrame(records)

# =================================================
# DATE / MONTH / WEEK
# =================================================
df["date"] = pd.to_datetime(df["date"])
df["month_date"] = df["date"].dt.to_period("M").dt.to_timestamp()
df["month"] = df["month_date"].dt.strftime("%b-%Y")
df["week_start"] = df["date"].dt.to_period("W").apply(lambda r: r.start_time)
df["week_label"] = (
    df["week_start"].dt.strftime("%d %b %Y") + " - " +
    (df["week_start"] + pd.Timedelta(days=6)).dt.strftime("%d %b %Y")
)

# =================================================
# SIDEBAR FILTERS
# =================================================
st.sidebar.markdown("*Global Filters*")

selected_employees = st.sidebar.multiselect(
    "Employee Name", sorted(df["employee_name"].dropna().unique())
)

selected_projects = st.sidebar.multiselect(
    "Project Name", sorted(df["project_name"].dropna().unique())
)

month_df = df[["month", "month_date"]].drop_duplicates().sort_values("month_date")

selected_months = st.sidebar.multiselect(
    "Month", month_df["month"].tolist()
)

selected_leaves = st.sidebar.multiselect(
    "Leave Type", ["AL", "CL", "PH"]
)

st.sidebar.markdown("---")
st.sidebar.image("logo.png", use_container_width=True)

# =================================================
# APPLY FILTERS
# =================================================
filtered_df = df.copy()

if selected_employees:
    filtered_df = filtered_df[filtered_df["employee_name"].isin(selected_employees)]

if selected_projects:
    filtered_df = filtered_df[filtered_df["project_name"].isin(selected_projects)]

if selected_months:
    filtered_df = filtered_df[filtered_df["month"].isin(selected_months)]

if selected_leaves:
    filtered_df = filtered_df[filtered_df["leave_type"].isin(selected_leaves)]

if filtered_df.empty:
    st.warning("No data available.")
    st.stop()

# =================================================
# PH DEDUPLICATION
# =================================================
ph_unique = (
    filtered_df[filtered_df["leave_type"] == "PH"]
    .drop_duplicates(subset=["employee_name", "date"])
)

st.markdown("---")

# =================================================
# PROJECTS PER EMPLOYEE
# =================================================
st.subheader("Projects Per Employee")

projects_per_employee = (
    filtered_df
    .groupby("employee_name")
    .agg(
        Project_Count=("project_name", "nunique"),
        Project_Names=("project_name", lambda x: ", ".join(sorted(x.dropna().unique()))),
        Mars_Project_Names=("mars_project_name", lambda x: ", ".join(sorted(x.dropna().unique())))
    )
    .reset_index()
)

show_table(projects_per_employee)

fname = clean_filename(
    f"ProjectsPerEmployee_{'_'.join(selected_employees) if selected_employees else 'All-Employees'}.xlsx"
)

st.download_button(
    "📥 Download Projects Per Employee",
    to_excel_bytes(prettify_columns(projects_per_employee)),
    file_name=fname
)

st.markdown("---")

# =================================================
# PROJECT-WISE HOURS
# =================================================
st.subheader("Project-wise Hours")

project_hours = (
    filtered_df
    .groupby(["employee_name", "project_name", "mars_project_name"], as_index=False)
    .agg(Total_Hours=("hours", "sum"))
)

show_table(project_hours)

fname = clean_filename(
    f"ProjectWiseHours_{'_'.join(selected_projects) if selected_projects else 'All-Projects'}.xlsx"
)

st.download_button(
    "📥 Download Project-wise Hours",
    to_excel_bytes(prettify_columns(project_hours)),
    file_name=fname
)

st.markdown("---")

# =================================================
# WEEK-WISE DATA PREVIEW
# =================================================
st.subheader("Week-wise Data Preview")

c1, c2, c3 = st.columns(3)

with c1:
    preview_month = st.selectbox("Select Month", ["All"] + month_df["month"].tolist())

week_df = filtered_df.copy()
if preview_month != "All":
    week_df = week_df[week_df["month"] == preview_month]

weeks = week_df[["week_label", "week_start"]].drop_duplicates().sort_values("week_start")

with c2:
    preview_week = st.selectbox("Select Week", ["All"] + weeks["week_label"].tolist())

if preview_week != "All":
    week_df = week_df[week_df["week_label"] == preview_week]

with c3:
    preview_project = st.selectbox(
        "Select Project",
        ["All"] + sorted(week_df["project_name"].dropna().unique())
    )

if preview_project != "All":
    week_df = week_df[week_df["project_name"] == preview_project]

def calculate_leave_days(group):
    # AL & CL → count normally
    non_ph = group[group["leave_type"].isin(["AL", "CL"])]

    # PH → count once per employee per date
    ph = (
        group[group["leave_type"] == "PH"]
        .drop_duplicates(subset=["employee_name", "date"])
    )

    return non_ph.shape[0] + ph.shape[0]


week_preview = (
    week_df
    .groupby(
        ["employee_name", "project_name", "mars_project_name", "week_start", "week_label"],
        as_index=False
    )
    .apply(lambda g: pd.Series({
        "Total_Hours": g["hours"].sum(),
        "Leave_Days": calculate_leave_days(g)
    }))
    .reset_index(drop=True)
    .sort_values("week_start")
)

show_table(week_preview)

fname = clean_filename(
    f"WeekWiseData_{preview_week}_{preview_project}.xlsx"
)

st.download_button(
    "📥 Download Week-wise Data",
    to_excel_bytes(prettify_columns(week_preview)),
    file_name=fname
)

st.markdown("---")

# =================================================
# LEAVE SUMMARY
# =================================================
st.subheader("Leave Summary")

leave_non_ph = filtered_df[filtered_df["leave_type"].isin(["AL", "CL"])]
leave_combined = pd.concat([leave_non_ph, ph_unique])

leave_summary = (
    leave_combined
    .groupby(["employee_name", "leave_type"])
    .size()
    .unstack(fill_value=0)
    .reset_index()
)

leave_cols = [c for c in leave_summary.columns if c != "employee_name"]
leave_summary["Total_Leaves"] = leave_summary[leave_cols].sum(axis=1)

show_table(leave_summary)

fname = clean_filename(
    f"LeaveSummary_{'_'.join(selected_months) if selected_months else 'All-Months'}.xlsx"
)

st.download_button(
    "📥 Download Leave Summary",
    to_excel_bytes(prettify_columns(leave_summary)),
    file_name=fname
)

st.markdown("---")

# =================================================
# GRAPHICAL INSIGHTS
# =================================================
st.subheader("Insights (Based on Global Filters)")

insight_df = filtered_df

total_leaves_df = pd.concat([
    insight_df[insight_df["leave_type"].isin(["AL", "CL"])],
    insight_df[insight_df["leave_type"] == "PH"].drop_duplicates(subset=["employee_name", "date"])
])

k1, k2, k3 = st.columns(3)
k1.metric("Total Hours", int(insight_df["hours"].sum()))
k2.metric("Total Leaves", total_leaves_df.shape[0])
k3.metric("Employees", insight_df["employee_name"].nunique())

monthly = (
    insight_df
    .groupby(["month_date", "month"], as_index=False)
    .agg(Total_Hours=("hours", "sum"))
    .sort_values("month_date")
)

fig_line = px.line(
    monthly, x="month", y="Total_Hours",
    markers=True, color_discrete_sequence=["#00E5FF"],
    title="Monthly Workload Trend"
)

fig_line.update_traces(line=dict(width=4))
fig_line.update_layout(plot_bgcolor="#0E1117", paper_bgcolor="#0E1117", font_color="white")
st.plotly_chart(fig_line, use_container_width=True)

proj = insight_df.groupby("project_name", as_index=False).agg(Total_Hours=("hours", "sum"))

fig_bar = px.bar(
    proj, x="project_name", y="Total_Hours",
    color="Total_Hours", color_continuous_scale="Turbo",
    title="Project Contribution", text_auto=True
)

fig_bar.update_layout(plot_bgcolor="#0E1117", paper_bgcolor="#0E1117", font_color="white")
st.plotly_chart(fig_bar, use_container_width=True)

leave_dist = (
    total_leaves_df
    .groupby("leave_type")
    .size()
    .reset_index(name="Count")
)

if not leave_dist.empty:
    fig_leave_bar = px.bar(
        leave_dist,
        x="Count",
        y="leave_type",
        orientation="h",
        text="Count",
        title="Leave Distribution",
        color="leave_type",
        color_discrete_sequence=px.colors.qualitative.Vivid
    )

    fig_leave_bar.update_layout(
        plot_bgcolor="#0E1117",
        paper_bgcolor="#0E1117",
        font_color="white",
        showlegend=False
    )

    st.plotly_chart(fig_leave_bar, use_container_width=True)

st.markdown("---")