import pandas as pd
import numpy as np
import random
from datetime import datetime, timedelta
import json

# Makes random numbers consistent every time you run
np.random.seed(42)
random.seed(42)

# These are all possible values a lead can have
# Think of these as Salesforce picklist dropdown options
industries = ["Fintech", "SaaS", "E-Commerce", "HealthTech", "EdTech", "LogisticsTech"]
stages = ["Seed", "Series A", "Series B", "Series C", "Enterprise"]

sources = [
    "Outbound SDR",
    "Inbound Demo",
    "Partner Referral",
    "Paid LinkedIn",
    "Paid Google",
    "Event/Conference",
    "Content/SEO"
]

segments = [
    "Startup (<50 emp)",
    "Mid-Market (50-500)",
    "Enterprise (500+)"
]

print("Lists defined successfully!")
print(f"Industries: {industries}")
print(f"Sources: {sources}")

# The start date - all leads will be created sometime in 2024
start = datetime(2024, 1, 1)

# Each source has a different probability of converting to MQL
# This mirrors rael B2B data:
# - People who REQUEST a demo are alraedy interested -> high conversion
# - People who clicked a Google ad -> just browsing -> low conversion

conv_map = {
    "Outbound SDR": 0.12,
    "Inbound Demo": 0.32,
    "Partner Referral": 0.28,
    "Paid LinkedIn": 0.09,
    "Paid Google": 0.07,
    "Event/Conference": 0.18,
    "Content/SEO": 0.11
}

print("Conversion rates defined!")
print("Highest converting source: Inbound Demo at 32%")
print("Lowest converting source: Paid Google at 7%")

# Empty list - we will add 500 leads into this 
leads = []

# Loop 500 times - each loop creates one lead
for i in range(500):

    # Pick random values for this lead
    source = random.choices(sources, weights=[20, 25, 18, 12, 10, 8, 7])[0]
    industry = random.choice(industries)
    segment = random.choices(segments, weights=[40, 35, 25])[0]

    # Pick a random date in 2024
    created = start + timedelta(days=random.randint(0,365))

    #Get conversion probability for this lead's source
    p = conv_map[source]

    # Simulate funnel - each stage depends on the previous one
    is_mql = random.random() < p * 2.5
    is_sql = is_mql and random.random() < 0.45
    is_opp = is_sql and random.random() < 0.55
    is_won = is_opp and random.random() < 0.38

    # Calculate ARR - only won deals have revenue
    arr = 0
    if is_won:
        base ={
            "Startup (<50 emp)": 8000,
            "Mid-Market (50-500)": 35000,
            "Enterprise (500+)": 120000
        }[segment]
        arr = int(np.random.normal(base, base * 0.25))

    # Build one lead as a dictionary (one row in your spreadsheet)
    leads.append({
            "lead_id": f"L{i+1:04d}",
            "industry": industry,
            "segment": segment,
            "source": source,
            "created_date": created.strftime("%Y-%m-%d"),
            "is_mql": int(is_mql),
            "is_sql": int(is_sql),
            "is_opportunity": int(is_opp),
            "is_won": int(is_won),
            "arr_usd": arr,      
        })
print(f" Generated {len(leads)} leads!:")
print(f"First ead: {leads[0]}")
print(f"Last lead: {leads[-1]}")

# Convert list of 500 dictionaries into a pandas DataFrame
# A DataFrame is like an Excel table - rows and columns
df = pd.DataFrame(leads)

#Let's look at the data before saving
print("\n---FIRST 5 ROWS---")
print(df.head())

print("\n---SHAPE---")
print(f"Rows: {df.shape[0]}, Columns: {df.shape[1]}")

print("\n--- QUICK ANALYSIS ---")
print(f"Total leads: {len(df)}")
print(f"Total MQLs: {df['is_mql'].sum()}")
print(f"Total SQLs: {df['is_sql'].sum()}")
print(f"Total Won:      {df['is_won'].sum()}")
print(f"Total ARR:      ${df['arr_usd'].sum():,.0f}")
print(f"MQL Rate:       {df['is_mql'].mean()*100:.1f}%")
print(f"Win Rate:       {df['is_won'].mean()*100:.1f}%")

print("\n--- ARR BY SOURCE ---")
by_source = df.groupby("source").agg(
    leads = ("lead_id", "count"),
    mqls  = ("is_mql",  "sum"),
    won   = ("is_won",  "sum"),
    arr   = ("arr_usd", "sum")
).reset_index()
print(by_source)

# Save to CSV — you can open this in Excel!
df.to_csv("leads.csv", index=False)
print("\nSaved as leads.csv — open it in Excel!")

# Calculate summary statistics
total_leads = len(df)
total_mqls  = int(df["is_mql"].sum())
total_sqls  = int(df["is_sql"].sum())
total_opps  = int(df["is_opportunity"].sum())
total_won   = int(df["is_won"].sum())
total_arr   = int(df["arr_usd"].sum())

# Convert by_source to a list of dictionaries
by_source_list = by_source.to_dict("records")

# Build the summary dictionary
summary = {
    "total_leads":  total_leads,
    "total_mqls":   total_mqls,
    "total_sqls":   total_sqls,
    "total_opps":   total_opps,
    "total_won":    total_won,
    "total_arr":    total_arr,
    "mql_rate":     round(total_mqls / total_leads * 100, 1),
    "win_rate":     round(total_won  / total_leads * 100, 1),
    "by_source":    by_source_list,
}

# Save as JSON file
with open("summary.json", "w") as f:
    json.dump(summary, f, indent=2)

print("Saved summary.json!")
print("\n--- SUMMARY SAVED ---")
print(f"Total Leads: {total_leads}")
print(f"Total ARR:   ${total_arr:,.0f}")
print(f"MQL Rate:    {summary['mql_rate']}%")
print(f"Win Rate:    {summary['win_rate']}%")