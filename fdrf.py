import pandas as pd
import os

# === CONFIGURATION ===
database_file = "Database.xlsx"         # Input file (from system or shared location)
output_file = "Final_Summary_Output.xlsx"   # Auto-generated output file

# === STEP 1: Load Database File ===
try:
    df = pd.read_excel(database_file)
    print("✅ Database file loaded successfully.")
except FileNotFoundError:
    print(f"❌ Error: '{database_file}' not found. Please check the file path.")
    exit()

# === STEP 2: Validate Required Columns ===
required_columns = [
    "Affiliate",
    "DIV_NAME",
    "HCP Selection Request ID",
    "Is PSA Created",
    "PSA Activity Executed"
]

missing_columns = [col for col in required_columns if col not in df.columns]
if missing_columns:
    print(f"❌ Missing columns in the Database file: {missing_columns}")
    exit()

# === STEP 3: Compute Summary by Affiliate and DIV_NAME ===
summary_rows = []

# Group by Affiliate & DIV_NAME
grouped = df.groupby(["Affiliate", "DIV_NAME"])

for (affiliate, div_name), group in grouped:
    hcp_count = group["HCP Selection Request ID"].nunique()
    psa_created_count = group.loc[group["Is PSA Created"] == 1, "HCP Selection Request ID"].nunique()
    psa_executed_count = group.loc[group["PSA Activity Executed"] == 1, "HCP Selection Request ID"].nunique()

    psa_created_pct = round((psa_created_count / hcp_count) * 100, 2) if hcp_count > 0 else 0
    psa_executed_pct = round((psa_executed_count / hcp_count) * 100, 2) if hcp_count > 0 else 0

    summary_rows.append({
        "Affiliate": affiliate,
        "DIV_NAME": div_name,
        "HCP Selection Request": hcp_count,
        "PSA Created": psa_created_count,
        "PSA Created %": psa_created_pct,
        "PSA Activity Executed": psa_executed_count,
        "PSA Executed %": psa_executed_pct
    })

# === STEP 4: Create DataFrame ===
summary_df = pd.DataFrame(summary_rows)

# === STEP 5: Add Grand Total Row ===
if not summary_df.empty:
    total_row = {
        "Affiliate": "Total",
        "DIV_NAME": "",
        "HCP Selection Request": summary_df["HCP Selection Request"].sum(),
        "PSA Created": summary_df["PSA Created"].sum(),
        "PSA Created %": round((summary_df["PSA Created"].sum() / summary_df["HCP Selection Request"].sum()) * 100, 2),
        "PSA Activity Executed": summary_df["PSA Activity Executed"].sum(),
        "PSA Executed %": round((summary_df["PSA Activity Executed"].sum() / summary_df["HCP Selection Request"].sum()) * 100, 2)
    }
    summary_df = pd.concat([summary_df, pd.DataFrame([total_row])], ignore_index=True)

# === STEP 6: Export to Excel ===
summary_df.to_excel(output_file, index=False)
print(f"✅ Affiliate & DIV_NAME-wise Final Summary generated → {os.path.abspath(output_file)}")

# === STEP 7: Display in Console ===
print("\n📊 Final Summary (Affiliate & DIV_NAME-wise):")
print(summary_df.to_string(index=False))
