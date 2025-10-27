import streamlit as st
import json
from google_sheet_utils import get_sheet
from template_generator import generate_return_template, generate_resale_template

# Load policies from Google Sheets
sheet = get_sheet()
records = sheet.get_all_records()

# Convert to dictionary
policies = {
    row['insured_name'].lower().replace(" ", "_"): {
        **row,
        "monthly_premiums": json.loads(row["premiums_json"])
    }
    for row in records
}

if not policies:
    st.error("❌ No saved policies found. Please onboard policies first.")
    st.stop()

st.set_page_config(page_title="Generate Purchase Template", layout="centered")
st.title("Life Settlement Template Generator")

# List of insured names
policy_keys = list(policies.keys())

# UI to select policy
selection = st.selectbox("Choose a policy:", options=policy_keys, format_func=lambda k: policies[k]["insured_name"])

if selection:
    policy = policies[selection]

    st.write("**Carrier:**", policy["carrier"])
    st.write("**DOB:**", policy["dob"])
    st.write("**LE Report Date:**", policy["le_report_date"])
    st.write("**LE (months):**", policy["le_months"])
    st.write("**Death Benefit:**", f"${policy['death_benefit']:,.2f}")

    # Internal Cost (optional field, shown only to user)
    internal_cost = policy.get("internal_cost", 0.0)
    try:
        internal_cost = float(internal_cost)
    except (ValueError, TypeError):
        internal_cost = 0.0
    st.write("**Internal Cost:**", f"${internal_cost:,.2f}")

    investment = st.number_input("Enter Client Cost", min_value=0.0, step=1000.0, value=internal_cost)

    col1, col2 = st.columns(2)

    with col1:
        if st.button("Generate Purchase Template"):
            monthly_premiums = {int(k): v for k, v in policy["monthly_premiums"].items()}

            output_filename = f"purchase_template_{selection}.xlsx"
            output_path = generate_return_template(
                insured_name=policy["insured_name"],
                dob=policy["dob"],
                carrier=policy["carrier"],
                le_months=policy["le_months"],
                le_report_date=policy["le_report_date"],
                death_benefit=policy["death_benefit"],
                investment=investment,
                monthly_premiums=monthly_premiums,
                output_filename=output_filename
            )

            with open(output_path, "rb") as f:
                st.success("✅ Template generated successfully!")
                st.download_button("📥 Download Excel", f, file_name=output_filename)

    with col2:
        if st.button("Generate Resale Template"):
            monthly_premiums = {int(k): v for k, v in policy["monthly_premiums"].items()}

            output_filename = f"resale_template_{selection}.xlsx"
            output_path = generate_resale_template(
                insured_name=policy["insured_name"],
                dob=policy["dob"],
                carrier=policy["carrier"],
                le_months=policy["le_months"],
                le_report_date=policy["le_report_date"],
                death_benefit=policy["death_benefit"],
                investment=investment,
                monthly_premiums=monthly_premiums,
                output_filename=output_filename
            )

            with open(output_path, "rb") as f:
                st.success("✅ Template generated successfully!")
                st.download_button("📥 Download Excel", f, file_name=output_filename)


