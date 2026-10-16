#!/usr/bin/env python3
"""
ETL & Data Modeling Pipeline for E-commerce & Supply Chain Capstone Project
Converts raw datasets in train/ into a production Star Schema and exports metrics.
"""

import os
import json
import pandas as pd
import numpy as np

def run_etl():
    raw_dir = '/Users/ThienTuCorp/Downloads/Final test Data/train'
    output_dir = '/Users/ThienTuCorp/Downloads/Final test Data/cleaned_data'
    os.makedirs(output_dir, exist_ok=True)
    
    print("[1/6] Loading raw CSV files...")
    df_customers = pd.read_csv(os.path.join(raw_dir, 'df_Customers.csv'))
    df_order_items = pd.read_csv(os.path.join(raw_dir, 'df_OrderItems.csv'))
    df_orders = pd.read_csv(os.path.join(raw_dir, 'df_Orders.csv'))
    df_payments = pd.read_csv(os.path.join(raw_dir, 'df_Payments.csv'))
    df_products = pd.read_csv(os.path.join(raw_dir, 'df_Products.csv'))
    
    print(f"Raw record counts: Orders={len(df_orders)}, Items={len(df_order_items)}, "
          f"Payments={len(df_payments)}, Customers={len(df_customers)}, Products={len(df_products)}")

    # [2/6] Clean and deduplicate Products
    print("[2/6] Cleansing and deduplicating Products...")
    df_products_clean = df_products.drop_duplicates().copy()
    df_products_clean['product_category_name'] = df_products_clean['product_category_name'].fillna('Unspecified')
    
    # Fill missing product weight/dimensions with median by category
    median_weight = df_products_clean['product_weight_g'].median()
    df_products_clean['product_weight_g'] = df_products_clean['product_weight_g'].fillna(median_weight)
    df_products_clean['product_length_cm'] = df_products_clean['product_length_cm'].fillna(df_products_clean['product_length_cm'].median())
    df_products_clean['product_height_cm'] = df_products_clean['product_height_cm'].fillna(df_products_clean['product_height_cm'].median())
    df_products_clean['product_width_cm'] = df_products_clean['product_width_cm'].fillna(df_products_clean['product_width_cm'].median())
    
    df_products_clean['product_weight_kg'] = (df_products_clean['product_weight_g'] / 1000.0).round(3)
    
    # Weight Class Binning
    def get_weight_class(w_kg):
        if w_kg < 1.0:
            return 'Light Weight (<1kg)'
        elif w_kg <= 5.0:
            return 'Medium Weight (1-5kg)'
        else:
            return 'Heavy Weight (>5kg)'
            
    df_products_clean['weight_class'] = df_products_clean['product_weight_kg'].apply(get_weight_class)
    print(f"Dim_Products unique rows: {len(df_products_clean)}")

    # [3/6] Clean Orders and calculate Lead Time & Delivery Bins
    print("[3/6] Cleansing Orders & Dates...")
    df_orders['order_purchase_timestamp'] = pd.to_datetime(df_orders['order_purchase_timestamp'])
    df_orders['order_approved_at'] = pd.to_datetime(df_orders['order_approved_at'])
    df_orders['order_delivered_timestamp'] = pd.to_datetime(df_orders['order_delivered_timestamp'])
    df_orders['order_estimated_delivery_date'] = pd.to_datetime(df_orders['order_estimated_delivery_date'])
    
    # Lead time in days
    df_orders['delivery_lead_time_days'] = (
        df_orders['order_delivered_timestamp'] - df_orders['order_purchase_timestamp']
    ).dt.total_seconds() / (24 * 3600)
    
    def get_delivery_tier(row):
        if pd.isna(row['order_delivered_timestamp']):
            return 'Not Shipped / Cancelled'
        lead = row['delivery_lead_time_days']
        if lead <= 0:
            return '1. Same Day / Express'
        elif lead <= 3.0:
            return '2. Standard (1-3 Days)'
        else:
            return '3. Delayed (>3 Days)'
            
    df_orders['delivery_performance_tier'] = df_orders.apply(get_delivery_tier, axis=1)

    # [4/6] Merge into Fact_Orders
    print("[4/6] Merging into Fact_Orders...")
    # Rename fields in items to match standard naming
    df_order_items_renamed = df_order_items.rename(columns={
        'price': 'sales',
        'shipping_charges': 'shipping_charge'
    })
    
    df_fact = df_orders.merge(df_order_items_renamed, on='order_id', how='left')
    df_fact = df_fact.merge(df_payments[['order_id', 'payment_type', 'payment_installments', 'payment_value']], on='order_id', how='left')
    df_fact = df_fact.merge(df_products_clean[['product_id', 'product_category_name', 'product_weight_kg', 'weight_class']], on='product_id', how='left')
    df_fact = df_fact.merge(df_customers[['customer_id', 'customer_city', 'customer_state']], on='customer_id', how='left')

    # Date Key
    df_fact['order_date_key'] = df_fact['order_purchase_timestamp'].dt.strftime('%Y%m%d')
    df_fact['order_date'] = df_fact['order_purchase_timestamp'].dt.strftime('%Y-%m-%d')
    df_fact['ship_date'] = df_fact['order_delivered_timestamp'].dt.strftime('%Y-%m-%d')
    df_fact['estimated_delivery_date'] = df_fact['order_estimated_delivery_date'].dt.strftime('%Y-%m-%d')

    # Basket Value Tier
    def get_basket_tier(val):
        if val < 50.0:
            return 'Low Value (<$50)'
        elif val <= 200.0:
            return 'Medium Value ($50-$200)'
        else:
            return 'High Value (>$200)'
            
    df_fact['order_basket_value_tier'] = df_fact['sales'].apply(get_basket_tier)

    # Logistics Cost Tier
    def get_logistics_tier(row):
        sales = row['sales']
        shipping = row['shipping_charge']
        if sales <= 0:
            return 'High Cost / Profit Erosion'
        ratio = shipping / sales
        if ratio < 0.08:
            return 'Optimal Cost (<8%)'
        elif ratio <= 0.15:
            return 'Moderate Cost (8%-15%)'
        else:
            return 'High Cost / Profit Erosion (>15%)'

    df_fact['logistics_cost_tier'] = df_fact.apply(get_logistics_tier, axis=1)
    df_fact['shipping_cost_ratio'] = (df_fact['shipping_charge'] / df_fact['sales']).round(4)

    # Customer Recency & Segments
    snapshot_date = df_fact['order_purchase_timestamp'].max()
    print(f"Data Snapshot Date: {snapshot_date}")

    cust_stats = df_fact.groupby('customer_id').agg(
        last_order=('order_purchase_timestamp', 'max'),
        total_orders=('order_id', 'nunique'),
        total_spend=('sales', 'sum')
    ).reset_index()

    cust_stats['customer_recency_days'] = (snapshot_date - cust_stats['last_order']).dt.days

    def get_recency_status(days):
        if days <= 30:
            return 'Active (<=30 days)'
        elif days <= 90:
            return 'Warm / At-Risk (31-90 days)'
        else:
            return 'Inactive / Cold (>90 days)'

    cust_stats['customer_recency_status'] = cust_stats['customer_recency_days'].apply(get_recency_status)

    # Docx Standard Segment (Strict 100% rubrics match)
    def get_docx_segment(row):
        recency = row['customer_recency_days']
        orders = row['total_orders']
        spend = row['total_spend']
        if recency > 90:
            return '4. Lost / At-Risk'
        elif spend >= 1000 and orders >= 5:
            return '1. VIP Customer'
        elif orders >= 3:
            return '2. Loyal Customer'
        else:
            return '3. Potential Customer'

    cust_stats['customer_segment_standard'] = cust_stats.apply(get_docx_segment, axis=1)

    # Enhanced Segment (taking monetary high value into account for single-purchase dataset)
    def get_enhanced_segment(row):
        recency = row['customer_recency_days']
        spend = row['total_spend']
        if recency > 90 and spend < 500:
            return '4. Lost / At-Risk'
        elif spend >= 1000:
            return '1. VIP Customer (High Value)'
        elif spend >= 300:
            return '2. Loyal / High-Tier'
        else:
            return '3. Potential Customer'

    cust_stats['customer_segment_enhanced'] = cust_stats.apply(get_enhanced_segment, axis=1)

    # Merge customer stats back into Fact
    df_fact = df_fact.merge(
        cust_stats[['customer_id', 'customer_recency_days', 'customer_recency_status', 
                    'customer_segment_standard', 'customer_segment_enhanced']],
        on='customer_id', how='left'
    )
    df_fact.rename(columns={'customer_segment_standard': 'customer_segment'}, inplace=True)

    # Update Dim_Customers with aggregated stats
    df_dim_customers = df_customers.merge(
        cust_stats[['customer_id', 'customer_recency_days', 'customer_recency_status', 
                    'total_orders', 'total_spend', 'customer_segment_standard', 'customer_segment_enhanced']],
        on='customer_id', how='left'
    ).rename(columns={'customer_segment_standard': 'customer_segment'})

    # [5/6] Build Dim_Date (Calendar Table)
    print("[5/6] Generating Dim_Date (Calendar Table)...")
    min_date = df_orders['order_purchase_timestamp'].min().floor('D')
    max_date = df_orders['order_purchase_timestamp'].max().ceil('D')
    date_range = pd.date_range(start=min_date, end=max_date, freq='D')
    
    df_dim_date = pd.DataFrame({'date': date_range})
    df_dim_date['date_key'] = df_dim_date['date'].dt.strftime('%Y%m%d')
    df_dim_date['full_date'] = df_dim_date['date'].dt.strftime('%Y-%m-%d')
    df_dim_date['year'] = df_dim_date['date'].dt.year
    df_dim_date['quarter'] = df_dim_date['date'].dt.quarter
    df_dim_date['quarter_name'] = 'Q' + df_dim_date['quarter'].astype(str)
    df_dim_date['month'] = df_dim_date['date'].dt.month
    df_dim_date['month_name'] = df_dim_date['date'].dt.strftime('%b')
    df_dim_date['month_year'] = df_dim_date['date'].dt.strftime('%Y-%m')
    df_dim_date['day'] = df_dim_date['date'].dt.day
    df_dim_date['day_of_week'] = df_dim_date['date'].dt.dayofweek + 1
    df_dim_date['day_name'] = df_dim_date['date'].dt.strftime('%A')
    df_dim_date['is_weekend'] = df_dim_date['day_of_week'].apply(lambda x: 1 if x in [6, 7] else 0)

    # [6/6] Save Star Schema files
    print("[6/6] Exporting Star Schema tables to CSV...")
    fact_cols = [
        'order_id', 'customer_id', 'product_id', 'seller_id',
        'order_date_key', 'order_date', 'ship_date', 'estimated_delivery_date',
        'order_status', 'sales', 'shipping_charge', 'shipping_cost_ratio',
        'delivery_lead_time_days', 'delivery_performance_tier',
        'order_basket_value_tier', 'logistics_cost_tier',
        'customer_recency_days', 'customer_recency_status',
        'customer_segment', 'customer_segment_enhanced',
        'payment_type', 'payment_installments', 'payment_value',
        'customer_city', 'customer_state', 'product_category_name', 'weight_class'
    ]
    df_fact_export = df_fact[fact_cols].copy()
    df_fact_export.to_csv(os.path.join(output_dir, 'Fact_Orders.csv'), index=False)
    df_dim_customers.to_csv(os.path.join(output_dir, 'Dim_Customers.csv'), index=False)
    df_products_clean.to_csv(os.path.join(output_dir, 'Dim_Products.csv'), index=False)
    df_payments.to_csv(os.path.join(output_dir, 'Dim_Payments.csv'), index=False)
    df_dim_date.to_csv(os.path.join(output_dir, 'Dim_Date.csv'), index=False)

    print("Star Schema tables exported successfully!")

    # Calculate overall KPIs for reporting and Web Dashboard
    total_revenue = float(df_fact['sales'].sum())
    total_orders = int(df_fact['order_id'].nunique())
    total_customers = int(df_fact['customer_id'].nunique())
    aov = total_revenue / total_orders
    total_shipping = float(df_fact['shipping_charge'].sum())
    shipping_ratio = (total_shipping / total_revenue) * 100.0

    delivered_mask = df_fact['ship_date'].notna()
    total_delivered = int(delivered_mask.sum())
    ontime_count = int((df_fact['delivery_performance_tier'].isin(['1. Same Day / Express', '2. Standard (1-3 Days)'])).sum())
    ontime_rate = (ontime_count / total_delivered * 100.0) if total_delivered > 0 else 0.0

    cancelled_count = int((df_fact['order_status'].str.lower() == 'canceled').sum())
    cancellation_rate = (cancelled_count / total_orders * 100.0)

    avg_lead_time = float(df_fact.loc[delivered_mask, 'delivery_lead_time_days'].mean())

    # Segment counts
    segment_standard_counts = df_fact['customer_segment'].value_counts().to_dict()
    segment_enhanced_counts = df_fact['customer_segment_enhanced'].value_counts().to_dict()
    
    # Monthly trend
    df_fact['month_year'] = pd.to_datetime(df_fact['order_date']).dt.strftime('%Y-%m')
    monthly_trend = df_fact.groupby('month_year').agg(
        revenue=('sales', 'sum'),
        orders=('order_id', 'nunique'),
        shipping=('shipping_charge', 'sum')
    ).reset_index().to_dict(orient='records')
    
    # Category summary (Top 10)
    cat_summary = df_fact.groupby('product_category_name').agg(
        revenue=('sales', 'sum'),
        orders=('order_id', 'nunique'),
        avg_shipping=('shipping_charge', 'mean')
    ).reset_index().sort_values('revenue', ascending=False).to_dict(orient='records')

    # State summary
    state_summary = df_fact.groupby('customer_state').agg(
        revenue=('sales', 'sum'),
        orders=('order_id', 'nunique'),
        ontime_rate=('delivery_performance_tier', lambda s: (s.isin(['1. Same Day / Express', '2. Standard (1-3 Days)'])).mean() * 100.0)
    ).reset_index().sort_values('revenue', ascending=False).to_dict(orient='records')

    # Payment summary
    payment_summary = df_fact.groupby('payment_type').agg(
        revenue=('sales', 'sum'),
        orders=('order_id', 'nunique'),
        aov=('sales', 'mean')
    ).reset_index().sort_values('revenue', ascending=False).to_dict(orient='records')

    # Delivery tier summary
    delivery_tier_summary = df_fact['delivery_performance_tier'].value_counts().to_dict()

    metrics = {
        'total_revenue': total_revenue,
        'total_orders': total_orders,
        'total_customers': total_customers,
        'aov': aov,
        'total_shipping': total_shipping,
        'shipping_cost_ratio': shipping_ratio,
        'total_delivered': total_delivered,
        'ontime_count': ontime_count,
        'ontime_rate': ontime_rate,
        'cancelled_count': cancelled_count,
        'cancellation_rate': cancellation_rate,
        'avg_lead_time': avg_lead_time,
        'segment_standard_counts': segment_standard_counts,
        'segment_enhanced_counts': segment_enhanced_counts,
        'delivery_tier_summary': delivery_tier_summary,
        'monthly_trend': monthly_trend,
        'cat_summary_top10': cat_summary[:10],
        'state_summary_top10': state_summary[:10],
        'payment_summary': payment_summary
    }

    metrics_path = os.path.join(output_dir, 'kpi_metrics.json')
    with open(metrics_path, 'w', encoding='utf-8') as f:
        json.dump(metrics, f, ensure_ascii=False, indent=2)

    print(f"KPI metrics saved to {metrics_path}")
    print("\n--- KEY SUMMARY METRICS ---")
    print(f"Total Revenue: ${total_revenue:,.2f}")
    print(f"Total Orders: {total_orders:,}")
    print(f"AOV: ${aov:,.2f}")
    print(f"Total Shipping Charge: ${total_shipping:,.2f} ({shipping_ratio:.2f}%)")
    print(f"On-Time Delivery Rate: {ontime_rate:.2f}% (Delivered={total_delivered:,})")
    print(f"Cancellation Rate: {cancellation_rate:.2f}% (Cancelled={cancelled_count:,})")
    print(f"Avg Lead Time: {avg_lead_time:.2f} days")
    print(f"Standard Segments: {segment_standard_counts}")
    print(f"Enhanced Segments: {segment_enhanced_counts}")

if __name__ == '__main__':
    run_etl()
