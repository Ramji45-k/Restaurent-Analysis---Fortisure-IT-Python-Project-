import pandas as pd

df = pd.read_excel("Dataset/QuickBite_Assessment.xlsx", sheet_name="orders")

print(df.head())

print(df.shape)

print(df.columns)

df.info()

print(df.isnull().sum())

print(df.dtypes)

print(df.duplicated().sum())

df = df.drop_duplicates()

print(df.shape[0])

df["order_datetime"] = pd.to_datetime(
    df["order_datetime"],
    format="mixed",
    dayfirst=True,
    errors="coerce"
)

print(df["order_datetime"].head(10))
print(df["order_datetime"].dtype)
print(df["order_datetime"].isna().sum())


df["order_amount"] = df["order_amount"].str.replace(r"Rs\.?|INR|₹|,", "", regex=True)
df["order_amount"] = pd.to_numeric(df["order_amount"], errors="coerce")

print(df["order_amount"].head())


df["payment_mode"] = df["payment_mode"].str.strip().str.title()

df["payment_mode"] = df["payment_mode"].replace({
    "Upi": "UPI",
    "Gpay": "Google Pay",
    "Phonepe": "PhonePe",
    "Cod": "Cash On Delivery",
    "Cash On Delivery": "Cash On Delivery",
    "Debit Card": "Card",
    "Credit Card": "Card",
    "Amazon Pay": "Amazon Pay",
    "Paytm Wallet": "Wallet"
})

print(df["payment_mode"].head(10))

df["order_status"] = df["order_status"].str.strip().str.lower()

print(df["order_status"].head(10))


df["delivery_minutes"] = pd.to_numeric(
    df["delivery_minutes"],
    errors="coerce"
)

print(df["delivery_minutes"].head(10))


df["coupon_code"] = df["coupon_code"].str.strip().str.upper()

print(df["coupon_code"].head(10))

areas_df = pd.read_excel(
    "Dataset/QuickBite_Assessment.xlsx",
    sheet_name="areas"
)

print(areas_df.head())

print(areas_df.shape)

areas_df = areas_df.drop_duplicates()

print(areas_df.shape[0])

areas_df["city"] = areas_df["city"].str.strip().str.lower()

areas_df["city"] = areas_df["city"].replace({
    "blr": "Bengaluru",
    "bengaluru": "Bengaluru",
    "hyd": "Hyderabad",
    "hyderabad": "Hyderabad",
    "madras": "Chennai",
    "chennai": "Chennai",
    "mumbai": "Mumbai"
})

print(areas_df["city"].head())

restaurants_df = pd.read_excel(
    "Dataset/QuickBite_Assessment.xlsx",
    sheet_name="restaurants"
)

restaurants_df["restaurant_name"] = restaurants_df["restaurant_name"].str.strip()

restaurants_df["cuisine"] = restaurants_df["cuisine"].str.strip().str.title()

restaurants_df["onboard_date"] = pd.to_datetime(
    restaurants_df["onboard_date"],
    format="mixed",
    dayfirst=True,
    errors="coerce"
)

delivery_partners_df = pd.read_excel(
    "Dataset/QuickBite_Assessment.xlsx",
    sheet_name="delivery_partners"
)

print(delivery_partners_df.head())
print(delivery_partners_df.shape)

delivery_partners_df["partner_name"] = (delivery_partners_df["partner_name"].str.strip())

delivery_partners_df["vehicle_type"] = (delivery_partners_df["vehicle_type"].str.strip())

delivery_partners_df["join_date"] = pd.to_datetime(
    delivery_partners_df["join_date"],
    format="mixed",
    dayfirst=True,
    errors="coerce"
)

customers_df = pd.read_excel(
    "Dataset/QuickBite_Assessment.xlsx",
    sheet_name="customers"
)

print(customers_df.head())
print(customers_df.shape)

customers_df = customers_df.drop_duplicates()

print(customers_df.shape[0])

customers_df["customer_name"] = customers_df["customer_name"].str.strip()

customers_df["customer_name"].head()

customers_df["phone"] = (
    customers_df["phone"]
    .str.strip()
    .str.replace("+91", "", regex=False)
    .str.replace("-", "", regex=False)
    .str.replace(" ", "", regex=False)
)
customers_df["phone"].str.len().value_counts()
customers_df[customers_df["phone"].str.len() != 10]["phone"]

customers_df[customers_df["email"] != customers_df["email"].str.strip()]["email"]

customers_df["signup_date"] = pd.to_datetime(
    customers_df["signup_date"],
    format="mixed",
    dayfirst=True,
    errors="coerce"
)

print(customers_df["signup_date"].dtype)
print(customers_df["signup_date"].isnull().sum())


feedback_df = pd.read_excel(
    "Dataset/QuickBite_Assessment.xlsx",
    sheet_name="feedback"
)

print(feedback_df.head())
print(feedback_df.shape)

feedback_df = feedback_df.drop_duplicates()

feedback_df["feedback_date"] = pd.to_datetime(
    feedback_df["feedback_date"],
    format="mixed",
    dayfirst=True,
    errors="coerce"
)

print(feedback_df["feedback_date"].dtype)
print(feedback_df["feedback_date"].isnull().sum())
print(feedback_df["feedback_date"].head(10))

feedback_df["rating"] = pd.to_numeric(
    feedback_df["rating"],
    errors="coerce"
)

import os

os.makedirs("cleaned_data", exist_ok=True)

with pd.ExcelWriter(
    "cleaned_data/Quick_Bite_Data_Cleaning.xlsx",
    engine="openpyxl"
) as writer:

    df.to_excel(writer, sheet_name="orders", index=False)
    customers_df.to_excel(writer, sheet_name="customers", index=False)
    areas_df.to_excel(writer, sheet_name="areas", index=False)
    restaurants_df.to_excel(writer, sheet_name="restaurants", index=False)
    delivery_partners_df.to_excel(writer, sheet_name="delivery_partners", index=False)
    feedback_df.to_excel(writer, sheet_name="feedback", index=False)

