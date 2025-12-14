import pandas as pd
import datetime

file_path = "../data/customers.xlsx"
df = pd.read_excel(file_path)

print("\n====== FINAL EXAM ======\n")

print(df.head())
print(df.info())
print(df.describe())

#--------------------------------------------------------------------

print("\n====== CLEAN NAME ======\n")

df["Name"] = df["Name"].str.strip()
df["Name"] = df["Name"].str.replace("_", " ", regex=False)
df["Name"] = df["Name"].str.title()


print(df["Name"])

#--------------------------------------------------------------------

print("\n====== CLEAN EMAIL ======\n")

df["Email"] = df["Email"].str.strip()
df["Email"] = df["Email"].str.replace(",", ".", regex=False)
df["Email"] = df["Email"].str.lower()
df.loc[df["Email"].str.endswith("."),"Email"] = \
    (df.loc[df["Email"].str.endswith("."), "Email"] + "com")


print(df["Email"])

#-----------------------------------------------------------------------

print("\n====== CLEAN COUNTRY ======\n")

df["Country"] = df["Country"].str.strip()
df["Country"] = df["Country"].str.lower()
df["Country"] = df["Country"].str.title()
df.loc[df["Country"] == "Usa", "Country"] = "USA"

print(df["Country"])

#------------------------------------------------------------------------

print("\n====== CLEAN ISACTIVE ======\n")

df["IsActive"] = df["IsActive"].str.strip().str.lower() == "yes"

print(df["IsActive"])

#------------------------------------------------------------------------

print("\n====== CLEAN TOTALSPENT ======\n")

df["TotalSpent"] = df["TotalSpent"].astype(str).str.strip()
df["TotalSpent"] = df["TotalSpent"].str.replace("$", "", regex=False)
df["TotalSpent"] = df["TotalSpent"].str.replace(",", "", regex=False)
df["TotalSpent"]= df["TotalSpent"].astype(float)

print(df["TotalSpent"])

#-------------------------------------------------------------------------

df["AnnualValue"] = df["TotalSpent"] * 12

TotalRevenue = df["TotalSpent"].sum()

print()

print(f"The total Revenue is: ${TotalRevenue}")

AverageSpend = df["TotalSpent"].mean()

print()

print(f"the average is: {AverageSpend:.2F}")

MaxCustomer = df["TotalSpent"].max()
TopCustomerName = df.loc[df["TotalSpent"] == MaxCustomer, "Name"].values[0]

print()

print(f"Top Customer: {TopCustomerName} - {MaxCustomer}")

#-----------------------------------------------------------------------------

print("\n====== CLEAN LASTPURCHASEDATE ======\n")
print()

df["LastPurchaseDate"] = pd.to_datetime(df["LastPurchaseDate"], dayfirst=True, errors="coerce")

today = datetime.datetime.now()

df["DaysSinceLastPurchase"] = (today - df["LastPurchaseDate"]).dt.days

print(df[["Name", "LastPurchaseDate", "DaysSinceLastPurchase"]])

#------------------------------------------------------------------------------

print("\n====== BEST CLIENTS ======\n")
print()

df["IsInactive"] = df["DaysSinceLastPurchase"] > 180

print(df[df["IsInactive"]][["Name", "DaysSinceLastPurchase"]])

#--------------------------------------------------------------------------------

df.to_excel("../output/Customers_Clean.xlsx")

total_customers = df.shape[0]
inactive_customers = df["IsInactive"].sum()

with open("../output/customer_kpis.txt", "w", encoding="utf-8") as file:
    file.write("CUSTOMER INSIGHTS REPORT\n")
    file.write("------------------------\n\n")
    file.write(f"Total Revenue: {TotalRevenue:.2f}\n")
    file.write(f"Average Spend per Customer: {AverageSpend:.2f}\n")
    file.write(f"Top Customer: {TopCustomerName} ({MaxCustomer:.2f})\n")
    file.write(f"Total Customers: {total_customers}\n")
    file.write(f"Inactive Customers (>180 days): {inactive_customers}\n")

