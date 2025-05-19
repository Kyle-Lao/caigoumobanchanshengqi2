
import streamlit as st
import json
from google_sheet_utils import get_sheet
from template_generator import generate_return_template

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

st.set_page_config(page_title="Generate Return Template", layout="centered")
st.title("📄 Select Policy & Generate Return Template")

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

    investment = st.number_input("Enter Investment Amount", min_value=0.0, step=1000.0)

    if st.button("Generate Return Template"):
        monthly_premiums = {int(k): v for k, v in policy["monthly_premiums"].items()}

        output_filename = f"return_template_{selection}.xlsx"
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
            st.success("✅ Template generated!")
            st.download_button("📥 Download Return Template", f, file_name=output_filename)
