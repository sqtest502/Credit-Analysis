import io
from datetime import datetime

import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="SmartQ Credit Analysis",
    page_icon="📊",
    layout="wide",
)

st.title("📊 SmartQ Credit Analysis")
st.write("Upload the weekly extracted Excel report to generate daily and user-level credit analysis.")


REQUIRED_COLUMNS = [
    "User",
    "User Type",
    "Date",
    "Credits",
]

WEEKDAYS = [
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
    "Sunday",
]


def clean_columns(df):
    df = df.copy()
    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.replace(r"\s+", " ", regex=True)
    )
    return df


def prepare_data(df):
    df = clean_columns(df)

    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        raise ValueError(
            "Required columns are missing: "
            + ", ".join(missing)
        )

    # Clean User
    df["User"] = df["User"].fillna("").astype(str).str.strip()

    # Clean User Type
    df["User Type"] = (
        df["User Type"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # Convert Date
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

    invalid_dates = int(df["Date"].isna().sum())
    if invalid_dates:
        st.warning(
            f"{invalid_dates:,} rows have an invalid/missing Date and "
            "will be excluded from date-based metrics."
        )

    # Convert Credits to number
    # Handles values such as 100, "100", "₹100", "100.00"
    df["Credits"] = (
        df["Credits"]
        .astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("₹", "", regex=False)
        .str.strip()
    )
    df["Credits"] = pd.to_numeric(df["Credits"], errors="coerce").fillna(0)

    # Calculate weekday from Date rather than trusting the downloaded Week Day.
    df["Week Day"] = df["Date"].dt.day_name()

    return df


def build_user_daily(df):
    valid = df.dropna(subset=["Date"]).copy()

    user_daily = (
        valid.groupby(
            ["Date", "Week Day", "User", "User Type"],
            dropna=False,
            as_index=False,
        )["Credits"]
        .sum()
        .rename(columns={"Credits": "Daily User Credits"})
    )

    user_daily["Over 200"] = user_daily["Daily User Credits"] > 200

    user_daily["Vendor No Credit"] = (
        user_daily["User Type"]
        .str.strip()
        .str.lower()
        .eq("vendornocredit")
    )

    user_daily["Vendor Used Credit"] = (
        user_daily["Vendor No Credit"]
        & (user_daily["Daily User Credits"] > 0)
    )

    user_daily["Audit Status"] = "OK"
    user_daily.loc[
        user_daily["Over 200"] | user_daily["Vendor Used Credit"],
        "Audit Status",
    ] = "REVIEW"

    return user_daily


def build_daily_metrics(df, user_daily):
    valid = df.dropna(subset=["Date"]).copy()

    if user_daily.empty:
        return pd.DataFrame(
            columns=[
                "Date",
                "Week Day",
                "Total Users",
                "Credit Users",
                "Total Credits",
                "Avg Credit/User",
                "Users >200",
                "Vendor Credit Users",
                "Total Orders",
            ]
        )

    metrics = []

    for (date, weekday), group in user_daily.groupby(
        ["Date", "Week Day"], sort=True
    ):
        total_users = group["User"].nunique()

        credit_users = group.loc[
            group["Daily User Credits"] > 0, "User"
        ].nunique()

        total_credits = group["Daily User Credits"].sum()

        avg_credit_user = (
            total_credits / credit_users if credit_users else 0
        )

        users_over_200 = group.loc[
            group["Daily User Credits"] > 200, "User"
        ].nunique()

        vendor_credit_users = group.loc[
            group["Vendor Used Credit"], "User"
        ].nunique()

        total_orders = valid.loc[
            valid["Date"].eq(date)
        ].shape[0]

        metrics.append(
            {
                "Date": date,
                "Week Day": weekday,
                "Total Users": total_users,
                "Credit Users": credit_users,
                "Total Credits": total_credits,
                "Avg Credit/User": round(avg_credit_user, 2),
                "Users >200": users_over_200,
                "Vendor Credit Users": vendor_credit_users,
                "Total Orders": total_orders,
            }
        )

    result = pd.DataFrame(metrics)
    result = result.sort_values("Date").reset_index(drop=True)
    return result


def build_user_metrics(user_daily):
    if user_daily.empty:
        return pd.DataFrame()

    user_metrics = (
        user_daily.groupby(
            ["User", "User Type"],
            dropna=False,
            as_index=False,
        )
        .agg(
            Total_Days_Used=("Date", "nunique"),
            Total_Credits=("Daily User Credits", "sum"),
            Max_Credits_Per_Day=("Daily User Credits", "max"),
            Users_Over_200_Days=("Over 200", "sum"),
            Vendor_Credit_Days=("Vendor Used Credit", "sum"),
        )
    )

    user_metrics["Avg_Credits_Per_Day"] = (
        user_metrics["Total_Credits"]
        / user_metrics["Total_Days_Used"].replace(0, pd.NA)
    ).round(2)

    user_metrics = user_metrics.rename(
        columns={
            "Total_Days_Used": "Days Used",
            "Total_Credits": "Total Credits",
            "Max_Credits_Per_Day": "Max Credits/Day",
            "Users_Over_200_Days": ">200 Days",
            "Vendor_Credit_Days": "Vendor Credit Days",
            "Avg_Credits_Per_Day": "Avg Credits/Day",
        }
    )

    return user_metrics.sort_values(
        ["Total Credits", "User"],
        ascending=[False, True],
    ).reset_index(drop=True)


def build_over_200(user_daily):
    return (
        user_daily.loc[user_daily["Over 200"]]
        .sort_values(["Date", "Daily User Credits"], ascending=[True, False])
        .reset_index(drop=True)
    )


def build_vendor_credit_users(user_daily):
    return (
        user_daily.loc[user_daily["Vendor Used Credit"]]
        .sort_values(["Date", "Daily User Credits"], ascending=[True, False])
        .reset_index(drop=True)
    )


def build_weekday_summary(user_daily):
    if user_daily.empty:
        return pd.DataFrame()

    summary = (
        user_daily.groupby("Week Day", as_index=False)
        .agg(
            Unique_Users=("User", "nunique"),
            Credit_Users=(
                "Daily User Credits",
                lambda x: int((x > 0).sum()),
            ),
            Total_Credits=("Daily User Credits", "sum"),
            Users_Over_200=(
                "Over 200",
                "sum",
            ),
            Vendor_Credit_Users=(
                "Vendor Used Credit",
                "sum",
            ),
        )
    )

    # Use number of unique users for credit users, not number of user-day rows.
    credit_counts = (
        user_daily.loc[user_daily["Daily User Credits"] > 0]
        .groupby("Week Day")["User"]
        .nunique()
        .rename("Credit_Users")
    )

    vendor_counts = (
        user_daily.loc[user_daily["Vendor Used Credit"]]
        .groupby("Week Day")["User"]
        .nunique()
        .rename("Vendor_Credit_Users")
    )

    over200_counts = (
        user_daily.loc[user_daily["Over 200"]]
        .groupby("Week Day")["User"]
        .nunique()
        .rename("Users_Over_200")
    )

    summary = summary.drop(
        columns=["Credit_Users", "Users_Over_200", "Vendor_Credit_Users"],
        errors="ignore",
    )

    summary = summary.merge(
        credit_counts,
        on="Week Day",
        how="left",
    )
    summary = summary.merge(
        over200_counts,
        on="Week Day",
        how="left",
    )
    summary = summary.merge(
        vendor_counts,
        on="Week Day",
        how="left",
    )

    summary = summary.fillna(0)

    for col in [
        "Unique_Users",
        "Credit_Users",
        "Users_Over_200",
        "Vendor_Credit_Users",
    ]:
        summary[col] = summary[col].astype(int)

    summary["Avg_Credits_Per_Credit_User"] = (
        summary["Total_Credits"]
        / summary["Credit_Users"].replace(0, pd.NA)
    ).round(2)

    order = {day: i for i, day in enumerate(WEEKDAYS)}
    summary["_sort"] = summary["Week Day"].map(order)
    summary = summary.sort_values("_sort").drop(columns="_sort")

    summary.columns = [
        "Week Day",
        "Unique Users",
        "Total Credits",
        "Credit Users",
        "Users >200",
        "Vendor Credit Users",
        "Avg Credits/Credit User",
    ]

    # Reorder for readability
    summary = summary[
        [
            "Week Day",
            "Unique Users",
            "Credit Users",
            "Total Credits",
            "Avg Credits/Credit User",
            "Users >200",
            "Vendor Credit Users",
        ]
    ]

    return summary


def make_excel(
    raw_df,
    daily_metrics,
    user_metrics,
    over_200,
    vendor_credit_users,
    weekday_summary,
    user_daily,
):
    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl",
        datetime_format="dd-mmm-yyyy",
    ) as writer:

        daily_metrics.to_excel(
            writer,
            sheet_name="Daily_Metrics",
            index=False,
        )

        user_metrics.to_excel(
            writer,
            sheet_name="User_Metrics",
            index=False,
        )

        over_200.to_excel(
            writer,
            sheet_name="Users_Over_200",
            index=False,
        )

        vendor_credit_users.to_excel(
            writer,
            sheet_name="Vendor_Credit_Users",
            index=False,
        )

        weekday_summary.to_excel(
            writer,
            sheet_name="Weekday_Summary",
            index=False,
        )

        # User + date level calculation used for audit.
        user_daily.to_excel(
            writer,
            sheet_name="User_Daily_Audit",
            index=False,
        )

        # Separate the original extracted data by weekday.
        for day in WEEKDAYS:
            day_data = raw_df.loc[
                raw_df["Week Day"].eq(day)
            ].copy()

            if not day_data.empty:
                # Excel sheet names have a 31-character limit.
                day_data.to_excel(
                    writer,
                    sheet_name=day,
                    index=False,
                )

        # Keep a cleaned complete dataset too.
        raw_df.to_excel(
            writer,
            sheet_name="All_Data",
            index=False,
        )

        # Basic column widths for every sheet.
        for worksheet in writer.book.worksheets:
            for column_cells in worksheet.columns:
                max_length = 0
                column_letter = column_cells[0].column_letter

                for cell in column_cells[:200]:
                    try:
                        value_length = len(str(cell.value)) if cell.value is not None else 0
                        max_length = max(max_length, value_length)
                    except Exception:
                        pass

                worksheet.column_dimensions[column_letter].width = min(
                    max(max_length + 2, 10),
                    35,
                )

            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions

    output.seek(0)
    return output.getvalue()


uploaded_file = st.file_uploader(
    "Upload the extracted weekly Excel file",
    type=["xlsx", "xls"],
)

if uploaded_file is not None:
    try:
        # Read the first sheet by default.
        # If your report has a specific sheet, change sheet_name=0 below.
        raw_df = pd.read_excel(uploaded_file, sheet_name=0)

        st.success(
            f"File uploaded successfully: {uploaded_file.name}"
        )

        with st.expander("Preview uploaded data"):
            st.write(f"Rows: {len(raw_df):,}")
            st.write(f"Columns: {len(raw_df.columns):,}")
            st.dataframe(raw_df.head(10), use_container_width=True)

        with st.spinner("Processing Excel data..."):
            df = prepare_data(raw_df)

            user_daily = build_user_daily(df)
            daily_metrics = build_daily_metrics(df, user_daily)
            user_metrics = build_user_metrics(user_daily)
            over_200 = build_over_200(user_daily)
            vendor_credit_users = build_vendor_credit_users(user_daily)
            weekday_summary = build_weekday_summary(user_daily)

            excel_bytes = make_excel(
                df,
                daily_metrics,
                user_metrics,
                over_200,
                vendor_credit_users,
                weekday_summary,
                user_daily,
            )

        # ========================================================
        # KPI SECTION
        # ========================================================

        total_users = df["User"].nunique()

        credit_users = df.loc[
            df["Credits"] > 0, "User"
        ].nunique()

        total_credits = df["Credits"].sum()

        users_over_200 = over_200["User"].nunique()

        vendor_users = vendor_credit_users["User"].nunique()

        col1, col2, col3, col4, col5 = st.columns(5)

        col1.metric(
            "Total Users",
            f"{total_users:,}",
        )

        col2.metric(
            "Credit Users",
            f"{credit_users:,}",
        )

        col3.metric(
            "Total Credits",
            f"{total_credits:,.2f}",
        )

        col4.metric(
            "Users >200",
            f"{users_over_200:,}",
        )

        col5.metric(
            "Vendor Credit Users",
            f"{vendor_users:,}",
        )

        st.divider()

        # ========================================================
        # DAILY METRICS
        # ========================================================

        st.subheader("📅 Daily Metrics")

        if daily_metrics.empty:
            st.info("No valid date data was found.")
        else:
            display_daily = daily_metrics.copy()
            display_daily["Date"] = display_daily["Date"].dt.strftime(
                "%d-%b-%Y"
            )

            st.dataframe(
                display_daily,
                use_container_width=True,
                hide_index=True,
            )

        # ========================================================
        # WEEKDAY SUMMARY
        # ========================================================

        st.subheader("📆 Weekday Summary")

        if not weekday_summary.empty:
            st.dataframe(
                weekday_summary,
                use_container_width=True,
                hide_index=True,
            )

        # ========================================================
        # AUDIT RESULTS
        # ========================================================

        st.subheader("🚨 Audit Results")

        audit_col1, audit_col2 = st.columns(2)

        with audit_col1:
            st.write("**Users exceeding 200 credits in a day**")

            if over_200.empty:
                st.success("No users exceeded 200 credits.")
            else:
                st.dataframe(
                    over_200,
                    use_container_width=True,
                    hide_index=True,
                )

        with audit_col2:
            st.write("**Vendor No Credit users who used credits**")

            if vendor_credit_users.empty:
                st.success(
                    "No Vendor No Credit users used credits."
                )
            else:
                st.dataframe(
                    vendor_credit_users,
                    use_container_width=True,
                    hide_index=True,
                )

        # ========================================================
        # DOWNLOAD
        # ========================================================

        st.divider()

        output_name = (
            f"Weekly_Credit_Analysis_"
            f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )

        st.download_button(
            label="⬇️ Download Weekly Credit Analysis",
            data=excel_bytes,
            file_name=output_name,
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True,
        )

        st.caption(
            "The generated workbook contains Daily_Metrics, User_Metrics, "
            "Users_Over_200, Vendor_Credit_Users, Weekday_Summary, "
            "User_Daily_Audit, All_Data, and separate weekday sheets."
        )

    except Exception as e:
        st.error(f"Unable to process the file: {e}")
        st.info(
            "Please verify that the Excel file contains these columns: "
            + ", ".join(REQUIRED_COLUMNS)
        )
else:
    st.info(
        "Upload the weekly extracted Excel report above to start the analysis."
    )
