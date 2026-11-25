import pandas as pd
import os

file_path = "cleaned_data/Quick_Bite_Data_Cleaning.xlsx"

orders_df = pd.read_excel(file_path, sheet_name="orders")
customers_df = pd.read_excel(file_path, sheet_name="customers")
areas_df = pd.read_excel(file_path, sheet_name="areas")
restaurants_df = pd.read_excel(file_path, sheet_name="restaurants")
delivery_partners_df = pd.read_excel(file_path, sheet_name="delivery_partners")
feedback_df = pd.read_excel(file_path, sheet_name="feedback")

print(orders_df.head())
print(orders_df.shape)

combined_df = orders_df.merge(
    customers_df[["customer_id", "customer_name"]],
    on="customer_id",
    how="left"
)

print(combined_df.head())
print(combined_df.shape)

combined_df = combined_df.merge(
    restaurants_df[
        ["restaurant_id", "restaurant_name", "cuisine", "area_id"]
    ],
    on="restaurant_id",
    how="left"
)

print(combined_df.head())
print(combined_df.shape)

combined_df = combined_df.rename(
    columns={"area_id": "restaurant_area_id"}
)

# Merge with areas
combined_df = combined_df.merge(
    areas_df[["area_id", "area_name", "city"]].rename(
        columns={"area_id": "restaurant_area_id"}
    ),
    on="restaurant_area_id",
    how="left"
)

print(combined_df.head())
print(combined_df.shape)

combined_df = combined_df.merge(
    delivery_partners_df[
        ["partner_id", "partner_name", "vehicle_type"]
    ],
    on="partner_id",
    how="left"
)

print(combined_df.head())
print(combined_df.shape)

combined_df = combined_df.merge(
    feedback_df[["order_id", "rating", "comment"]],
    on="order_id",
    how="left"
)


combined_df["order_date"] = pd.to_datetime(
    combined_df["order_datetime"],
    errors="coerce"
).dt.date

combined_df["order_time"] = pd.to_datetime(
    combined_df["order_datetime"],
    errors="coerce"
).dt.strftime("%H:%M:%S")

combined_df["order_day"] = pd.to_datetime(
    combined_df["order_datetime"],
    errors="coerce"
).dt.day_name()

# Create late delivery indicator
combined_df["is_late"] = (
    combined_df["delivery_minutes"]
    .gt(45)
    .where(combined_df["delivery_minutes"].notna(), pd.NA)
    .astype("boolean")
)
print(combined_df.head())
print(combined_df.shape)

combined_df = combined_df[
    [
        "order_id",
        "order_date",
        "order_time",
        "order_day",
        "customer_name",
        "restaurant_name",
        "cuisine",
        "area_name",
        "city",
        "partner_name",
        "vehicle_type",
        "order_amount",
        "payment_mode",
        "coupon_code",
        "order_status",
        "delivery_minutes",
        "is_late",
        "rating",
        "comment"
    ]
]


print("\nMerged Table:")
print(combined_df.head(10).to_string(index=False))

print("\nMerged Table Shape:")
print(combined_df.shape)

print("\nColumn Names:")
print(combined_df.columns.tolist())

os.makedirs("cleaned_data", exist_ok=True)

output_file = "cleaned_data/Quick_Bite_Combined_Data.xlsx"

combined_df.to_excel(
    output_file,
    index=False,
    engine="openpyxl"
)

print("\nMerged data saved successfully!")
print(output_file)


analysis_df = pd.read_excel(
    "cleaned_data/Quick_Bite_Combined_Data.xlsx"
)


delivered_df = analysis_df[
    analysis_df["order_status"] == "delivered"
]

print(delivered_df.head())
print(delivered_df.shape)


area_performance = delivered_df.groupby("area_name").agg(
    total_orders=("order_id", "count"),
    total_revenue=("order_amount", "sum"),
    average_delivery_minutes=("delivery_minutes", "mean"),
    late_orders=("is_late", "sum")
).reset_index()


area_performance = area_performance.sort_values(
    by="total_revenue",
    ascending=False
)

print("\nArea Performance:")
print(area_performance)


area_performance["late_delivery_percentage"] = (
    area_performance["late_orders"]
    / area_performance["total_orders"]
) * 100


area_performance["late_delivery_percentage"] = (
    area_performance["late_delivery_percentage"].round(2)
)

print("\nArea Performance with Late Delivery Percentage:")
print(area_performance)

cuisine_by_day = delivered_df.groupby(
    ["order_day", "cuisine"]
).agg(
    order_count=("order_id", "count")
).reset_index()

print(cuisine_by_day.head())


daily_totals = delivered_df.groupby("order_day")[
    "order_id"
].count().reset_index(name="total_day_orders")


cuisine_by_day = cuisine_by_day.merge(
    daily_totals,
    on="order_day",
    how="left"
)


cuisine_by_day["order_share_percentage"] = (
    cuisine_by_day["order_count"]
    / cuisine_by_day["total_day_orders"]
    * 100
).round(2)

top_cuisine_by_day = cuisine_by_day.loc[
    cuisine_by_day.groupby("order_day")["order_count"].idxmax()
]

top_cuisine_by_day = top_cuisine_by_day.sort_values(
    "order_day"
)

print("\nTop Cuisine by Day:")
print(top_cuisine_by_day)

vehicle_performance = delivered_df.groupby("vehicle_type").agg(
    total_orders=("order_id", "count"),
    average_delivery_minutes=("delivery_minutes", "mean"),
    late_orders=("is_late", "sum")
).reset_index()

print("\nVehicle Performance:")
print(vehicle_performance)

vehicle_performance["late_delivery_percentage"] = (
    vehicle_performance["late_orders"]
    / vehicle_performance["total_orders"]
    * 100
).round(2)

vehicle_performance["average_delivery_minutes"] = (
    vehicle_performance["average_delivery_minutes"].round(2)
)

print("\nVehicle Performance with Late Delivery Percentage:")
print(vehicle_performance)

vehicle_performance = vehicle_performance.sort_values(
    by="average_delivery_minutes",
    ascending=True
)

print("\nVehicle Performance Ranking:")
print(vehicle_performance)


os.makedirs("analysis", exist_ok=True)

output_file = "analysis/Quick_Bite_Analysis.xlsx"

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:

    area_performance.to_excel(
        writer,
        sheet_name="Area Performance",
        index=False
    )

    top_cuisine_by_day.to_excel(
        writer,
        sheet_name="Top Cuisine by Day",
        index=False
    )

    vehicle_performance.to_excel(
        writer,
        sheet_name="Vehicle Performance",
        index=False
    )

print("\nAnalysis Excel file saved successfully!")
print(output_file)