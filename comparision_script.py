import os
import pandas as pd

prev_file = "input/previous/DAST Report.html"
curr_file = "input/current/DAST Report.html"

def extract_required_columns(html_file):
    tables = pd.read_html(html_file)[4:]
    collected = []

    for df in tables:
        df.columns = df.columns.str.strip()
        required_cols = {"Testcase Name", "Result", "Error Data"}

        if required_cols.issubset(df.columns):
            temp = df[["Testcase Name", "Result", "Error Data"]].copy()
            collected.append(temp)

    if collected:
        final_df = pd.concat(collected, ignore_index=True)
    else:
        final_df = pd.DataFrame(columns=["Testcase Name", "Result", "Error Data"])

    return final_df


# Extract from both reports
prev_all_df = extract_required_columns(prev_file)
curr_all_df = extract_required_columns(curr_file)

# Rename columns
prev_all_df = prev_all_df.rename(columns={
    "Result": "Prev_Result",
    "Error Data": "Prev_Error"
})

curr_all_df = curr_all_df.rename(columns={
    "Result": "Curr_Result",
    "Error Data": "Curr_Error"
})

# OUTER merge to detect new/removed
merged_df = pd.merge(
    prev_all_df,
    curr_all_df,
    on="Testcase Name",
    how="outer",
    indicator=True
)

# Normalize result values
merged_df["Prev_Result"] = merged_df["Prev_Result"].astype(str).str.upper().str.strip()
merged_df["Curr_Result"] = merged_df["Curr_Result"].astype(str).str.upper().str.strip()

# 4 cases
regressions_df = merged_df[
    (merged_df["Prev_Result"] == "PASS") &
    (merged_df["Curr_Result"] == "FAIL")
]

fixed_df = merged_df[
    (merged_df["Prev_Result"] == "FAIL") &
    (merged_df["Curr_Result"] == "PASS")
]

fail_both_df = merged_df[
    (merged_df["Prev_Result"] == "FAIL") &
    (merged_df["Curr_Result"] == "FAIL")
]

pass_both_df = merged_df[
    (merged_df["Prev_Result"] == "PASS") &
    (merged_df["Curr_Result"] == "PASS")
]

# New & removed
new_tests = merged_df[merged_df["_merge"] == "right_only"]
removed_tests = merged_df[merged_df["_merge"] == "left_only"]

# -------- WRITE TO SINGLE EXCEL FILE --------

output_dir = "output"
os.makedirs(output_dir, exist_ok=True)

output_file = os.path.join(output_dir, "comparison_report.xlsx")


with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    regressions_df.to_excel(writer, sheet_name="Regressions", index=False)
    fixed_df.to_excel(writer, sheet_name="Fixed", index=False)
    fail_both_df.to_excel(writer, sheet_name="Fail_in_Both", index=False)
    pass_both_df.to_excel(writer, sheet_name="Pass_in_Both", index=False)
    new_tests.to_excel(writer, sheet_name="New_Testcases", index=False)
    removed_tests.to_excel(writer, sheet_name="Removed_Testcases", index=False)

print(f"\n✅ Comparison report generated: {output_file}")

# -------- STATS --------

total_considered = len(merged_df)

print("\n========== COMPARISON STATS ==========")
print(f"Total testcases considered : {total_considered}")
print("-------------------------------------")
print(f"❌ Regressions (Pass → Fail) : {len(regressions_df)}")
print(f"✅ Fixed (Fail → Pass)       : {len(fixed_df)}")
print(f"⚠️  Fail in both              : {len(fail_both_df)}")
print(f"🙂 Pass in both              : {len(pass_both_df)}")
print("-------------------------------------")
print(f"🆕 New testcases in current  : {len(new_tests)}")
print(f"❌ Removed testcases         : {len(removed_tests)}")
print("=====================================")
