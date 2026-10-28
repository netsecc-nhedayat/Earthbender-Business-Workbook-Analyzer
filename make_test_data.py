"""Generates sample_earthbender.xlsx with known expected results.

Expected outcomes when loaded into the app:
  - 10 rows, 8 columns
  - Job IDs: J001..J008, one blank, one duplicate (J003)
  - Row for J004 has end_date before start_date
  - J006 has no revenue (missing financial)
  - J007 margin > 150% (unusual)
  - Materials: Timber (x5), Aluminium (x3), Steel (x2)
  - Lead sources include 'Referral' (2 jobs)
"""
import pandas as pd

rows = [
    # job_id, start, end, material, lead, revenue, cost, customer
    ("J001", "2024-01-05", "2024-01-20", "Timber",    "Google",     4000, 2500, "Acme"),
    ("J002", "2024-01-15", "2024-02-02", "Timber",    "Referral",   5500, 3000, "Beta Co"),
    ("J003", "2024-02-01", "2024-02-12", "Aluminium", "Facebook",   6000, 4000, "Gamma"),
    ("J003", "2024-02-01", "2024-02-12", "Aluminium", "Facebook",   6000, 4000, "Gamma"),  # dup
    ("",     "2024-02-10", "2024-02-25", "Steel",     "Google",     7000, 5000, "Delta"),  # missing id
    ("J004", "2024-03-05", "2024-03-01", "Timber",    "Walk-in",    4500, 2500, "Epsilon"),  # end<start
    ("J005", "2024-03-20", "2024-04-05", "Aluminium", "Referral",   8000, 5000, "Zeta"),    # referral
    ("J006", "2024-04-10", "2024-04-20", "Steel",     "Google",     None, 6000, "Eta"),     # missing revenue
    ("J007", "2024-04-15", "2024-04-25", "Timber",    "Facebook",   1000, -1000, "Theta"), # huge margin
    ("J008", "2024-05-01", "2024-05-15", "Timber",    "Google",     5000, 3000, "Iota"),
]

cols = ["Job ID", "Start Date", "End Date", "Material",
        "Lead Source", "Revenue", "Cost", "Customer"]
df = pd.DataFrame(rows, columns=cols)
df.to_excel("sample_earthbender.xlsx", index=False)
print("Wrote sample_earthbender.xlsx")
print(df)
print("\nExpected material counts:\n", df["Material"].value_counts())
print("\nExpected referral count:", (df["Lead Source"] == "Referral").sum())
