import streamlit as st
import pandas as pd
from openai import OpenAI

# Uses OPENAI_API_KEY from Streamlit Secrets
client = OpenAI()

st.title("Treeplan Impact Message Demo")

st.write(
    "Upload the Excel_Employee_Data.xlsx file to generate Lore's impact email "
    "for October 2025 based on her fundraiser results."
)

uploaded_file = st.file_uploader("Upload Excel_Employee_Data.xlsx", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Read the Excel sheet that contains the fundraiser data
        df = pd.read_excel(uploaded_file, sheet_name="Fundraiser")
    except Exception as e:
        st.error(f"Could not read sheet 'Fundraiser'. Error: {e}")
    else:
        # Filter to Lore's row for 2025-10 (your demo line)
        lore_rows = df[(df["Name"] == "Lore") & (df["Month"] == "2025-10")]

        if lore_rows.empty:
            st.error("No row found for Lore in 2025-10.")
        else:
            lore = lore_rows.iloc[0]

            # Extract values
            donors = int(lore["Donors_Acquired"])
            tenure = float(lore["Avg_Tenure_Months"])
            monthly = float(lore["Monthly_Donation_EUR"])
            month = lore["Month"]

            # Impact calculations
            lifetime_donation = donors * monthly * tenure
            trees = int(lifetime_donation / 2)  # €2 per tree

            # Show data so jury/manager sees the logic
            st.subheader("Lore's impact data")
            st.write(
                {
                    "Name": lore["Name"],
                    "Email": lore["Email"],
                    "Month": month,
                    "Donors_Acquired": donors,
                    "Avg_Tenure_Months": tenure,
                    "Monthly_Donation_EUR": monthly,
                    "Lifetime_Donation_EUR": round(lifetime_donation, 2),
                    "Trees_Funded": trees,
                }
            )

            # Button to generate message
            if st.button("Generate impact message for Lore"):
                prompt = f"""
You are writing on behalf of Treeplan, an environmental NGO, to a fundraiser named Lore.

Data for {month}:
- Donors acquired: {donors}
- Average donor tenure: {tenure} months
- Monthly donation per donor: €{monthly}
- Estimated lifetime donation: €{lifetime_donation:.0f}
- Estimated trees funded: {trees}

Write a recognition email in 90-140 words:
- Start with "Hi Lore,"
- Clearly mention the {trees} trees and compare it to one simple, concrete image
  (for example: "about two soccer fields of new forest").
- Tone: warm, appreciative, specific, and human.
- You may use one subtle tree emoji if it fits naturally, and at most one exclamation mark.
- End with "Warm regards," and "The Treeplan Team".
""".strip()

                try:
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
                except Exception as e:
                    st.error(f"Error generating message: {e}")
