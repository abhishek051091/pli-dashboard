import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title="PLI Policy Dashboard", layout="wide")

st.title("📮 High-Speed PLI Policy Dashboard")
st.write(
    "Data File Upload Karein aur Single ya Bulk Policy Number Search Karein."
)


# Fast Data Loader with Caching
@st.cache_data(show_spinner="Data fast load ho raha hai...")
def load_uploaded_data(file):
  if file.name.endswith(".parquet"):
    df = pd.read_parquet(file)
  elif file.name.endswith(".csv"):
    try:
      df = pd.read_csv(file, on_bad_lines="skip")
    except Exception:
      file.seek(0)
      df = pd.read_csv(file, on_bad_lines="skip", encoding="latin1")
  else:
    df = pd.read_excel(file)

  # Column names clean karna
  df.columns = df.columns.str.strip()

  # Digits column create karna fast search ke liye
  if "policy_digits" not in df.columns and "policyNo" in df.columns:
    df["policy_digits"] = (
        df["policyNo"].astype(str).str.replace(r"\D", "", regex=True)
    )

  return df


# Function to process single policy data (Updated with Separate Current & Maturity Unpaid Lists)
def process_single_policy(policy_df):
  actual_policy_no = (
      policy_df["policyNo"].iloc[0]
      if "policyNo" in policy_df.columns
      else "N/A"
  )
  name = policy_df["name"].iloc[0] if "name" in policy_df.columns else "N/A"

  # Maturity Date extract karna
  maturity_date_str = "N/A"
  maturity_period = None
  for mat_col in ["maturityDate", "maturity_date", "MaturityDate"]:
    if mat_col in policy_df.columns:
      valid_mats = pd.to_datetime(policy_df[mat_col], errors="coerce").dropna()
      if not valid_mats.empty:
        maturity_period = valid_mats.iloc[0].to_period("M")
        maturity_date_str = valid_mats.iloc[0].strftime("%b-%Y")
        break

  premium_rate = 0
  if "premiumRate" in policy_df.columns:
    valid_rates = policy_df["premiumRate"].dropna()
    if not valid_rates.empty:
      premium_rate = valid_rates.iloc[0]

  # Dates processing
  if "scheduleMonth" in policy_df.columns:
    policy_df["dt_month"] = pd.to_datetime(
        policy_df["scheduleMonth"], format="%b-%y", errors="coerce"
    )
    if policy_df["dt_month"].isna().all():
      policy_df["dt_month"] = pd.to_datetime(
          policy_df["scheduleMonth"], errors="coerce"
      )
  elif "fromDate" in policy_df.columns:
    policy_df["dt_month"] = pd.to_datetime(
        policy_df["fromDate"], errors="coerce"
    )

  valid_dates_df = policy_df.dropna(subset=["dt_month"]).copy()

  paid_entries = []
  paid_period_set = set()

  if not valid_dates_df.empty:
    valid_dates_df = valid_dates_df.sort_values("dt_month")
    for _, row in valid_dates_df.iterrows():
      p = pd.Period(row["dt_month"], freq="M")
      paid_period_set.add(p)

      rate = (
          row["premiumRate"]
          if "premiumRate" in policy_df.columns
          and pd.notna(row["premiumRate"])
          else premium_rate
      )
      paid_entries.append({
          "Month": p.strftime("%b-%Y"),
          "Premium Amount": rate,
          "dt": row["dt_month"],
      })

    paid_df = pd.DataFrame(paid_entries)
    if not paid_df.empty:
      paid_df = (
          paid_df.drop_duplicates(subset=["Month"])
          .sort_values("dt")
          .reset_index(drop=True)
      )

    start_date = valid_dates_df["dt_month"].min().to_period("M")
    current_period = pd.Period(pd.Timestamp.now(), freq="M")

    # 1. Unpaid till Current Date
    range_current = pd.period_range(
        start=start_date, end=current_period, freq="M"
    )
    unpaid_periods_current = [
        p for p in range_current if p not in paid_period_set
    ]

    # 2. Unpaid from next month after Current Date up to Maturity Date
    unpaid_periods_maturity = []
    if maturity_period and maturity_period > current_period:
      range_maturity = pd.period_range(
          start=current_period + 1, end=maturity_period, freq="M"
      )
      unpaid_periods_maturity = [
          p for p in range_maturity if p not in paid_period_set
      ]

    paid_count = len(paid_period_set)
    unpaid_count_current = len(unpaid_periods_current)
    unpaid_count_maturity = len(unpaid_periods_maturity)

    unpaid_current_str = ", ".join(
        [p.strftime("%b-%Y") for p in unpaid_periods_current]
    )
    unpaid_maturity_str = ", ".join(
        [p.strftime("%b-%Y") for p in unpaid_periods_maturity]
    )
  else:
    paid_df = pd.DataFrame(columns=["Month", "Premium Amount"])
    paid_count = 0
    unpaid_count_current = 0
    unpaid_count_maturity = 0
    unpaid_periods_current = []
    unpaid_periods_maturity = []
    unpaid_current_str = "N/A"
    unpaid_maturity_str = "N/A"

  return {
      "actual_policy_no": actual_policy_no,
      "name": name,
      "maturity_date": maturity_date_str,
      "premium_rate": premium_rate,
      "paid_count": paid_count,
      "unpaid_count_current": unpaid_count_current,
      "unpaid_count_maturity": unpaid_count_maturity,
      "paid_df": paid_df,
      "unpaid_periods_current": unpaid_periods_current,
      "unpaid_periods_maturity": unpaid_periods_maturity,
      "unpaid_current_str": unpaid_current_str,
      "unpaid_maturity_str": unpaid_maturity_str,
  }


# Excel helper function
def to_excel(df):
  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name="Policy_Summary")
  return output.getvalue()


# File Uploader
uploaded_file = st.file_uploader(
    "Data File Upload Karein (.parquet recommended, .csv, .xlsx)",
    type=["parquet", "csv", "xlsx"],
)

if uploaded_file is not None:
  df = load_uploaded_data(uploaded_file)

  # Search Mode Selection
  search_mode = st.radio(
      "Search Mode Select Karein:",
      ["Single Policy Search", "Bulk Policy Search"],
      horizontal=True,
  )

  # ==================== SINGLE POLICY SEARCH ====================
  if search_mode == "Single Policy Search":
    policy_input = st.text_input(
        "Policy Number Enter Karein (Example: 1363132 ya APS-1363132-L)"
    ).strip()

    if policy_input:
      input_digits = re.sub(r"\D", "", policy_input)

      if input_digits and "policy_digits" in df.columns:
        policy_df = df[df["policy_digits"] == input_digits].copy()
      else:
        policy_df = pd.DataFrame()

      if not policy_df.empty:
        res = process_single_policy(policy_df)

        st.success(
            f"**Policy No:** {res['actual_policy_no']} | **Policy Holder Name:**"
            f" {res['name']} | **Maturity Date:** {res['maturity_date']}"
        )

        st.markdown("---")
        col_rate, col1, col2 = st.columns(3)

        if isinstance(res["premium_rate"], (int, float)):
          col_rate.metric(
              "Default Premium Rate", f"₹ {res['premium_rate']:,.2f}"
          )
        else:
          col_rate.metric(
              "Default Premium Rate", f"₹ {res['premium_rate']}"
          )

        col1.metric("Paid Months Count", f"{res['paid_count']} Months")
        col2.metric(
            "Unpaid (Till Current)", f"{res['unpaid_count_current']} Months"
        )

        st.markdown("---")

        # Three columns for Paid Months, Unpaid Till Current, and Unpaid Till Maturity
        col_paid, col_cur, col_mat = st.columns(3)

        with col_paid:
          st.subheader("✅ Paid Months")
          if not res["paid_df"].empty:
            display_paid_df = res["paid_df"][["Month", "Premium Amount"]]
            st.dataframe(display_paid_df, use_container_width=True, height=300)

            paid_csv = display_paid_df.to_csv(index=False).encode("utf-8")
            st.download_button(
                label="📥 Download Paid (CSV)",
                data=paid_csv,
                file_name=f"Paid_Months_{res['actual_policy_no']}.csv",
                mime="text/csv",
            )
          else:
            st.info("Koi paid records nahi mile.")

        with col_cur:
          st.subheader("📅 Unpaid (Till Current)")
          if res["unpaid_count_current"] > 0:
            cur_list = []
            for p in res["unpaid_periods_current"]:
              formatted_month = p.strftime("%b-%Y")
              st.write(f"• **{formatted_month}**")
              cur_list.append({
                  "Policy No": res["actual_policy_no"],
                  "Name": res["name"],
                  "Unpaid Month (Current)": formatted_month,
              })

            cur_df = pd.DataFrame(cur_list)
            cur_csv = cur_df.to_csv(index=False).encode("utf-8")
            st.download_button(
                label="📥 Download Current Unpaid (CSV)",
                data=cur_csv,
                file_name=f"Unpaid_Current_{res['actual_policy_no']}.csv",
                mime="text/csv",
            )
          else:
            st.success("🎉 Aaj tak ke sabhi months paid hain!")

        with col_mat:
          st.subheader("🎯 Unpaid (Till Maturity)")
          if res["unpaid_count_maturity"] > 0:
            mat_list = []
            for p in res["unpaid_periods_maturity"]:
              formatted_month = p.strftime("%b-%Y")
              st.write(f"• **{formatted_month}**")
              mat_list.append({
                  "Policy No": res["actual_policy_no"],
                  "Name": res["name"],
                  "Unpaid Month (Maturity)": formatted_month,
              })

            mat_df = pd.DataFrame(mat_list)
            mat_csv = mat_df.to_csv(index=False).encode("utf-8")
            st.download_button(
                label="📥 Download Maturity Unpaid (CSV)",
                data=mat_csv,
                file_name=f"Unpaid_Maturity_{res['actual_policy_no']}.csv",
                mime="text/csv",
            )
          else:
            st.info(
                "Maturity tak ya toh koi pending nahi hai ya Maturity date"
                " available nahi hai."
            )

      else:
        st.error(
            "Yeh Policy Number database mein nahi mila. Kripya digit check"
            " karein."
        )

  # ==================== BULK POLICY SEARCH ====================
  else:
    bulk_input = st.text_area(
        "Multiple Policy Numbers Enter Karein (Har line me ek ya comma se"
        " separate karke):",
        height=150,
        placeholder="1363132\n1363133\nAPS-1363134-L",
    )

    if st.button("🔍 Run Bulk Search", type="primary"):
      if bulk_input.strip():
        raw_inputs = re.split(r"[\n,]+", bulk_input)
        policy_list = [p.strip() for p in raw_inputs if p.strip()]

        summary_results = []

        for p_in in policy_list:
          p_digits = re.sub(r"\D", "", p_in)

          if p_digits and "policy_digits" in df.columns:
            p_df = df[df["policy_digits"] == p_digits].copy()
          else:
            p_df = pd.DataFrame()

          if not p_df.empty:
            res = process_single_policy(p_df)
            summary_results.append({
                "Input Query": p_in,
                "Policy No": res["actual_policy_no"],
                "Holder Name": res["name"],
                "Maturity Date": res["maturity_date"],
                "Default Premium Rate": res["premium_rate"],
                "Paid Months": res["paid_count"],
                "Unpaid (Current)": res["unpaid_count_current"],
                "Unpaid (Maturity)": res["unpaid_count_maturity"],
                "Current Unpaid List": res["unpaid_current_str"],
                "Maturity Unpaid List": res["unpaid_maturity_str"],
                "Status": (
                    "Complete"
                    if (
                        res["unpaid_count_current"] == 0
                        and res["unpaid_count_maturity"] == 0
                    )
                    else (
                        f"{res['unpaid_count_current']} Pending (Current),"
                        f" {res['unpaid_count_maturity']} Pending (Maturity)"
                    )
                ),
            })
          else:
            summary_results.append({
                "Input Query": p_in,
                "Policy No": "NOT FOUND",
                "Holder Name": "-",
                "Maturity Date": "-",
                "Default Premium Rate": "-",
                "Paid Months": 0,
                "Unpaid (Current)": 0,
                "Unpaid (Maturity)": 0,
                "Current Unpaid List": "-",
                "Maturity Unpaid List": "-",
                "Status": "Policy Not Found",
            })

        summary_df = pd.DataFrame(summary_results)

        st.markdown("---")
        st.subheader("📊 Bulk Search Result Summary")
        st.dataframe(summary_df, use_container_width=True)

        col_dl1, col_dl2 = st.columns(2)

        csv_data = summary_df.to_csv(index=False).encode("utf-8")
        col_dl1.download_button(
            label="📥 Download Summary Result (CSV)",
            data=csv_data,
            file_name="Bulk_PLI_Search_Summary.csv",
            mime="text/csv",
            use_container_width=True,
        )

        excel_data = to_excel(summary_df)
        col_dl2.download_button(
            label="📊 Download Summary Result (Excel)",
            data=excel_data,
            file_name="Bulk_PLI_Search_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

      else:
        st.warning(
            "Kripya search karne ke liye kam se kam ek Policy Number enter"
            " karein."
        )