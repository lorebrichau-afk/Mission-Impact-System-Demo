import streamlit as st
import pandas as pd
from openai import OpenAI

client = OpenAI()  # will use OPENAI_API_KEY from Streamlit later

st.title("Treeplan Impact Message Demo")

st.write("Upload the Excel_Employee_Data.xlsx file to generate Lore's impact email for 2025-10.")

uploaded_file = st.file_uploader("Upload Excel_Employee_Data.xlsx", type=["xlsx"])

if uploaded_file is not None:
    # Read the Excel sheet
    df = pd.read_excel(uploaded_file, sheet_name="Fundraiser")

    # Pick Lore 2025-10
    lore_rows = df[(df["Name"] == "Lore") & (df["Month"] == "2025-10")]

    if lore_rows.empty:
        st.error("No row found for Lore in 2025-10.")
    else:
        lore = lore_rows.iloc[0]
        donors = int(lore["Donors_Acquired"])
        tenure = float(lore["Avg_Tenure_Months"])
        monthly = float(lore["Monthly_Donation_EUR"])
        lifetime = donors * monthly * tenure
        trees = int(lifetime / 2)

        st.write("Lore's impact data:", {
            "Donors_Acquired": donors,
            "Avg_Tenure_Months": tenure,
            "Monthly_Donation_EUR": monthly,
            "Lifetime_Donation_EUR": round(lifetime, 2),
            "Trees_Funded": trees,
        })

        if st.button("Generate impact message for Lore"):
            prompt = f"""
You are writing on behalf of Treeplan, an environmental NGO, to a fundraiser named Lore.

Data for {month}:
- Donors acquired: {donors}
- Average donor tenure: {tenure} months
- Monthly donation per donor: €{monthly_donation}
- Estimated lifetime donation: €{lifetime_donation:.0f}
- Estimated trees funded: {trees}

Write a recognition email in 90-140 words:
- Start with "Hi Lore,"
- Clearly mention the {trees} trees and compare it to one concrete image
  (for example: "about two soccer fields of new forest").
- Tone: warm, appreciative, specific, and human.
- You may use one subtle tree emoji if it fits naturally, and at most one exclamation mark.
- End with "Warm regards," and "The Treeplan Team".
""".strip()

            response = client.chat.completions.create(
                model="gpt-4.1-mini",
                messages=[
                    {
                        "role": "system",
                        "content": "You write sincere, concrete appreciation emails for NGO fundraisers."
                    },
                    {"role": "user", "content": prompt},
                ],
                max_tokens=260,
                temperature=0.7,
            )

            email_text = response.choices[0].message.content.strip()

            st.subheader("Generated email")
            st.write(email_text)
