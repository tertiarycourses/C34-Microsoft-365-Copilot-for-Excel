"""Deterministic Ghana cedi sales data for the Datin labs."""

from collections import namedtuple
from datetime import datetime
from itertools import product

Period = namedtuple("Period", "filename first_month price_step")
FIRST = Period("Datin_Sales_Finance_H1_GHS.xlsx", 1, 0)
FRESH = Period("Datin_Sales_Finance_H2_GHS.xlsx", 7, 25)
CHECKPOINT = "Datin_Sales_Dashboard_Example_GHS.xlsx"
REGIONS = ["Greater Accra", "Ashanti", "Western"]
CHANNELS = ["Wholesale", "Retail"]
CATEGORIES = ["Food Staples", "Household", "Personal Care"]
PRODUCTS = [("Rice", "Maize Flour"), ("Laundry Detergent", "Dish Soap"), ("Bath Soap", "Body Lotion")]
HEADERS = [
    "Order_ID", "Order_Date", "Month_Start", "Region", "Channel", "Category",
    "Units", "List_Price_GHS", "Discount_GHS", "Net_Sales_GHS", "COGS_GHS",
    "Gross_Profit_GHS", "Product",
]


def records(period):
    for offset, region, channel, category, rep in product(range(6), range(3), range(2), range(3), range(4)):
        month = period.first_month + offset
        base = (offset * 5 + region * 3 + channel * 2 + category + rep * 4 + period.first_month) % 17
        seasonal = [2, 0, -3, 3, 9, 6][offset]
        units = 9 + base + (9, 4, 0)[region] + (5, 0)[channel] + seasonal
        price = [85, 125, 160][category] + period.price_step + (0, 12)[rep % 2]
        discount = units * (5 if channel == 0 else 2)
        net = units * price - discount
        cost = units * ([51, 74, 95][category] + period.price_step * 2 // 3 + (0, 7)[rep % 2])
        yield (
            f"D{month:02d}-{region + 1}{channel + 1}{category + 1}{rep + 1}",
            datetime(2026, month, 3 + rep * 6 + region), datetime(2026, month, 1),
            REGIONS[region], CHANNELS[channel], CATEGORIES[category], units, price,
            discount, net, cost, net - cost, PRODUCTS[category][rep % 2],
        )


def totals(rows):
    units, net, cost, gross = (sum(row[i] for row in rows) for i in (6, 9, 10, 11))
    return dict(orders=len(rows), units=units, net=net, cost=cost, gross=gross, margin=gross / net if net else 0)


def select(rows, *, region="All", channel="All", product="All"):
    return [
        row for row in rows
        if region in ("All", row[3]) and channel in ("All", row[4]) and product in ("All", row[12])
    ]


def headline(t):
    return f"GHS {t['net']:,} net sales, GHS {t['gross']:,} gross profit, {t['margin']:.2%} margin, {t['orders']:,} orders"
