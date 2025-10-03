import pandas as pd
import re

# Load file, skipping first 4 rows
# df = pd.read_excel("cuv_strongs.xlsx", skiprows=4)

# If CSV instead:
df = pd.read_csv(r"C:\Users\user\Downloads\Slide_Automation\cuv_clean.csv")

# Function to remove Strong's numbers
def clean_text(text):
    # Removes {H1234} or {G5678}
    return re.sub(r"\{[HG]\d+\}", "", str(text))


df["CleanText"] = df["Text"].apply(clean_text)

# Save cleaned full dataset
df.to_csv("cuv_clean.csv", index=False, encoding="utf-8-sig")
print("✅ Saved cleaned file as cuv_clean.csv")
