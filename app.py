import pandas as pd
from openai import OpenAI

client = OpenAI()  # uses your OPENAI_API_KEY from the terminal

# 1. Read the Excel file (sheet name = 'Fundraiser')
df = pd.read_excel("Excel_Employee_Data.xlsx", sheet_name="Fundraiser")

# 2. Filter to Lore and the month 2025-10
lore_row = df[(df["Name"] == "Lore") & (df["Month"] == "2025-10")].iloc[0]

# 3. Extract key data
donors = int(lore_row["Donors_Acquired"])
tenure = float(lore_row["Avg_Tenure_Months"])
monthly_donation = float(lore_row["Monthly_Donation_EUR"])
month = lore_row["Month"]

# 4. Calculate total impact
lifetime_donation = donors * monthly_donation * tenure
trees = int(lifetime_donation / 2)

# 5. Create the AI prompt
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
- Clearly mention the {trees} trees planted and connect it to a relatable image (like “two soccer fields of forest”).
- Tone: warm, appreciative, and professional.
- Emojis or exclamation marks.
- End with “Warm regards,” and a Treeplan-style signature.
"""

# 6. Send the request to OpenAI
response = client.chat.completions.create(
    model="gpt-4.1-mini",
    messages=[
        {"role": "system", "content": "You write authentic, concrete appreciation emails for NGO fundraisers."},
        {"role": "user", "content": prompt},
    ],
    max_tokens=250,
    temperature=0.7,
)

# 7. Get and display the result
email_text = response.choices[0].message.content

print("\n--- Generated Email for Lore ---\n")
print(email_text)

