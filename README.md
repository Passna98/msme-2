import pandas as pd

# Step 1: Load Tally Data
def load_tally_data(file_path):
    """
    Load Tally exported data with columns: 
    Supplier Name, Opening Balance, Purchases, Payments, Closing Balance
    """
    df = pd.read_excel(file_path)
    return df

# Step 2: Calculate Net Balance
def calculate_net_balance(row):
    return row['Purchases'] - row['Payments']

# Step 3: Determine MSME applicability
def msme_applicability(row, agreement_days, msme_category):
    """
    Check MSME applicability:
    - Traders are excluded
    - Agreement days max capped at 45 days
    """
    if msme_category.lower() == "trader":
        return "Not Applicable (Trader)"
    
    # Apply agreement days or default
    days = min(agreement_days, 45) if agreement_days else 15

    if row['Net Balance'] <= 0:
        return "Not Applicable (Negative/Zero Net)"
    else:
        return f"Applicable ({days} days)"

# Step 4: Generate MSME Compliance Report
def generate_msme_report(df, agreement_days_dict, msme_category_dict):
    df['Net Balance'] = df.apply(calculate_net_balance, axis=1)
    df['MSME Applicability'] = df.apply(
        lambda row: msme_applicability(
            row,
            agreement_days_dict.get(row['Supplier Name'], 0),
            msme_category_dict.get(row['Supplier Name'], "trader")
        ),
        axis=1
    )
    return df

# Step 5: Save output
def save_report(df, output_path):
    df.to_excel(output_path, index=False)
    print(f"✅ MSME Compliance Report saved at {output_path}")

# Example Run
if __name__ == "__main__":
    # File paths
    tally_file = "tally_data.xlsx"  # Input file
    output_file = "msme_compliance_report.xlsx"  # Output file

    # Load data
    data = load_tally_data(tally_file)

    # Agreement days mapping (Supplier Name -> Agreement Days)
    agreement_days = {
        "Supplier A": 30,  # Example: agreement with 30 days
        "Supplier B": 50,  # Example: agreement with 50 days (will take 45 days)
        "Supplier C": 0    # No agreement (default 15 days)
    }

    # MSME category mapping (Supplier Name -> Category)
    msme_category = {
        "Supplier A": "Manufacturer",
        "Supplier B": "Service Provider",
        "Supplier C": "Trader"  # Will be excluded
    }

    # Generate and Save Report
    report = generate_msme_report(data, agreement_days, msme_category)
    save_report(report, output_file)
# msme-2
