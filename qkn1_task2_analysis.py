import pandas as pd


def main():
    # 1. Import the data file into a data frame
    file_path = "Task 1 D598 Data Set.xlsx"
    df = pd.read_excel(file_path)

    print("=== First 5 rows of original data ===")
    print(df.head())
    print("\n=== DataFrame info ===")
    print(df.info())

    # 2. Identify any duplicate rows in the data set
    duplicate_mask = df.duplicated(keep=False)
    duplicate_rows = df[duplicate_mask]

    print("\n=== Duplicate rows (if any) ===")
    print("Number of duplicate rows:", duplicate_rows.shape[0])
    if duplicate_rows.shape[0] > 0:
        print(duplicate_rows)
    else:
        print("No duplicate rows found.")

    # 3. Group all IDs by state, then run descriptive statistics
    #    (mean, median, min, max) for all numeric variables by state
    state_stats = df.groupby("Business State").agg(["mean", "median", "min", "max"])

    # Flatten multi-level column names: (e.g., "Total Revenue_mean")
    state_stats.columns = [
        f"{col_name}_{stat_name}" for col_name, stat_name in state_stats.columns
    ]

    # Make Business State a regular column instead of index
    state_stats = state_stats.reset_index()

    print("\n=== Descriptive statistics by Business State ===")
    print(state_stats.head())

    # 4. Filter the data frame to identify businesses with negative debt-to-equity
    negative_debt_equity = df[df["Debt to Equity"] < 0]

    print("\n=== Businesses with negative Debt-to-Equity ===")
    print("Number of businesses with negative Debt-to-Equity:",
          negative_debt_equity.shape[0])
    print(negative_debt_equity.head())

    # 5. Create a new data frame that provides the debt-to-income ratio
    #    Debt-to-income = Total Long-term Debt / Total Revenue
    dti_df = df[[
        "Business ID",
        "Business State",
        "Total Long-term Debt",
        "Total Revenue"
    ]].copy()

    dti_df["Debt to Income"] = (
        dti_df["Total Long-term Debt"] / dti_df["Total Revenue"]
    )

    print("\n=== Debt-to-Income DataFrame (first 5 rows) ===")
    print(dti_df.head())

    # 6. Concatenate Debt-to-Income with the original data frame
    df["Debt to Income"] = dti_df["Debt to Income"]
    combined_df = df.copy()

    print("\n=== Combined DataFrame (original + Debt to Income) ===")
    print(combined_df.head())

    # 7. Save results to Excel files (optional for reporting)
    state_stats.to_excel("state_stats.xlsx", index=False)
    negative_debt_equity.to_excel("negative_debt_equity.xlsx", index=False)
    combined_df.to_excel("combined_df.xlsx", index=False)

    print("\nFiles saved:")
    print(" - state_stats.xlsx")
    print(" - negative_debt_equity.xlsx")
    print(" - combined_df.xlsx")


if __name__ == "__main__":
    main()
