import streamlit as st
import pandas as pd
from openai import OpenAI

# Uses OPENAI_API_KEY from Streamlit secrets
client = OpenAI()

st.title("Treeplan Impact Message Demo")

st.write(
    "Upload the Excel_Employee_Data.xlsx file to generate Lore's impact email "
    "for 2025-10 based on her fundraiser results."
)

uploaded_file = st.file_uploader("Upload Excel_Employee_Data.xlsx", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Read the Excel sheet
        df = pd.read_excel(uploaded_file, sheet_name="Fundraiser")
    except Exception as e:
        st.error(f"Could not read sheet 'Fundraiser'. Error: {e}")
    else:
        # Filter to Lore and the month 2025-10
        mask = (df["Name"] == "Lore") & (df["Month"] == "2025-10")
        if not mask.any():
            st.error("No row found for Lore in 2025-10.")
        else:
            lore_row = df[mask].iloc[0]

            # 3. Extract key data
            donors = int(lore_row["Donors_Acquired"])
            tenure = float(lore_row["Avg_Tenure_Months"])
            monthly_donation = float(lore_row["Monthly_Donation_EUR"])
            month = lore_row["Month"]

            # 4. Calculate total impact
            lifetime_donation = donors * monthly_donation * tenure
            trees = int(lifetime_donation / 2)  # €2 per tree

            # Show inputs so people see what the AI is based on
            st.subheader("Lore's impact data")
            st.write(
                {
                    "Name": lore_row["Name"],
                    "Email": lore_row["Email"],
                    "Month": month,
                    "Donors_Acquired": donors,
                    "Avg_Tenure_Months": tenure,
                    "Monthly_Donation_EUR": monthly_donation,
                    "Estimated_Lifetime_Donation_EUR": round(lifetime_donation, 2),
                    "Estimated_Trees_Funded": trees,
                }
            )

            # 5. Generate message button
            if st.button("Generate impact message for Lore"):
                prompt = f"""
You are writing on behalf of Treeplan, an environmental NGO,
to a fundraiser named Lore.

Data for {month}:
- Donors acquired: {donors}
- Average donor tenure: {tenure} months
- Monthly donation per donor: €{monthly_donation}
- Estimated lifetime donation: €{lifetime_donation:.0f}
- Estimated trees funded: {trees}

Write a recognition email in 90–140 words:
- Start with "Hi Lore,"
- Clearly mention the {trees} trees Lore planted and connect it to a relatable image
  (for example: "about two soccer fields of new forest").
- Tone: warm, appreciative, concrete, and professional, with a bit of heart.
- You may use one nature-related emoji if it fits naturally, and at most one exclamation mark.
- End with "Warm regards," and a Treeplan-style signature.
""".strip()

                try:
                    response = client.chat.completions.create(
                        model="gpt-4.1-mini",
                        messages=[
                            {
                                "role": "system",
                                "content": "You write authentic, concrete appreciation emails for NGO fundraisers."
                            },
                            {"role": "user", "content": prompt},
                        ],
                        max_tokens=250,
                        temperature=0.7,
                    )

                    email_text = response.choices[0].message.content.strip()

                    st.subheader("Generated email for Lore")
                    st.write(email_text)
                except Exception as e:
                    st.error(f"Error generating message: {e}")

