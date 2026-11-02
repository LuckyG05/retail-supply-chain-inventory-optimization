/* ============================================================================
   PROJECT: DataCo Global Supply Chain, Inventory & OTIF Analytics
   TOOLS: MySQL 8.0, Excel, Power BI
   DESCRIPTION: End-to-End SQL Script covering Database Setup, Data Loading,
                Star Schema Normalization (Views), and Business Analytics.
============================================================================ */

-- ============================================================================
-- STEP 1: DATABASE & STAGING TABLE SETUP
-- ============================================================================
CREATE DATABASE IF NOT EXISTS dataco_supply_chain;
USE dataco_supply_chain;

CREATE TABLE IF NOT EXISTS raw_supply_chain (
    payment_type VARCHAR(50),
    days_shipping_real INT,
    days_shipment_scheduled INT,
    benefit_per_order DECIMAL(12,4),
    sales_per_customer DECIMAL(12,4),
    delivery_status VARCHAR(50),
    late_delivery_risk INT,
    category_id INT,
    category_name VARCHAR(100),
    customer_city VARCHAR(100),
    customer_country VARCHAR(100),
    customer_email VARCHAR(100),
    customer_fname VARCHAR(100),
    customer_id INT,
    customer_lname VARCHAR(100),
    customer_password VARCHAR(100),
    customer_segment VARCHAR(50),
    customer_state VARCHAR(50),
    customer_street VARCHAR(150),
    customer_zipcode VARCHAR(20),
    department_id INT,
    department_name VARCHAR(100),
    latitude DECIMAL(12,8),
    longitude DECIMAL(12,8),
    market VARCHAR(50),
    order_city VARCHAR(100),
    order_country VARCHAR(100),
    order_customer_id INT,
    order_date_raw VARCHAR(50),
    order_id INT,
    order_item_cardprod_id INT,
    order_item_discount DECIMAL(12,4),
    order_item_discount_rate DECIMAL(12,4),
    order_item_id INT,
    order_item_product_price DECIMAL(12,4),
    order_item_profit_ratio DECIMAL(12,4),
    order_item_quantity INT,
    sales DECIMAL(12,4),
    order_item_total DECIMAL(12,4),
    order_profit_per_order DECIMAL(12,4),
    order_region VARCHAR(100),
    order_state VARCHAR(100),
    order_status VARCHAR(50),
    order_zipcode VARCHAR(20),
    product_card_id INT,
    product_category_id INT,
    product_description TEXT,
    product_image VARCHAR(255),
    product_name VARCHAR(150),
    product_price DECIMAL(12,4),
    product_status INT,
    shipping_date_raw VARCHAR(50),
    shipping_mode VARCHAR(50)
);

-- Bulk Load 180,519 rows from CSV
/*
LOAD DATA INFILE 'C:/ProgramData/MySQL/MySQL Server 8.0/Uploads/DataCoSupplyChainDataset.csv'
INTO TABLE raw_supply_chain
CHARACTER SET latin1
FIELDS TERMINATED BY ',' 
ENCLOSED BY '"'
LINES TERMINATED BY '\r\n'
IGNORE 1 ROWS;
*/

-- ============================================================================
-- STEP 2: DATA NORMALIZATION INTO STAR SCHEMA (FACT & DIMENSION VIEWS)
-- ============================================================================

-- 1. Dimension View: Customers (PII Removed)
CREATE OR REPLACE VIEW vw_dim_customers AS
SELECT DISTINCT
    customer_id,
    customer_segment,
    customer_city,
    customer_state,
    customer_country
FROM raw_supply_chain;

-- 2. Dimension View: Products (118 Unique SKUs)
CREATE OR REPLACE VIEW vw_dim_products AS
SELECT DISTINCT
    product_card_id,
    product_name,
    category_id,
    category_name,
    department_id,
    department_name,
    product_price
FROM raw_supply_chain;

-- 3. Fact View: Orders & Shipping (Date Parsing & KPI Flags)
CREATE OR REPLACE VIEW vw_fact_orders AS
SELECT 
    order_item_id,
    order_id,
    order_customer_id AS customer_id,
    product_card_id,
    DATE(STR_TO_DATE(order_date_raw, '%m/%d/%Y %H:%i')) AS order_date,
    DATE(STR_TO_DATE(shipping_date_raw, '%m/%d/%Y %H:%i')) AS shipping_date,
    payment_type,
    order_status,
    shipping_mode,
    delivery_status,
    late_delivery_risk,
    days_shipment_scheduled,
    days_shipping_real,
    (days_shipping_real - days_shipment_scheduled) AS delay_days,
    CASE 
        WHEN delivery_status IN ('Advance shipping', 'Shipping on time') THEN 1 
        ELSE 0 
    END AS is_on_time,
    market,
    order_region,
    order_country,
    order_state,
    order_city,
    order_item_quantity,
    sales,
    order_item_discount,
    order_item_discount_rate,
    order_item_total AS net_sales,
    order_profit_per_order AS profit,
    order_item_profit_ratio AS profit_margin_ratio
FROM raw_supply_chain;

-- ============================================================================
-- STEP 3: BUSINESS ANALYTICAL QUERIES
-- ============================================================================

-- QUERY 1: Shipping Mode SLA Bottleneck Analysis
SELECT 
    shipping_mode,
    COUNT(order_item_id) AS total_shipments,
    ROUND(AVG(days_shipment_scheduled), 2) AS promised_sla_days,
    ROUND(AVG(days_shipping_real), 2) AS actual_avg_days,
    ROUND(AVG(delay_days), 2) AS avg_delay_variance_days,
    ROUND(SUM(is_on_time) * 100.0 / COUNT(*), 2) AS on_time_delivery_pct,
    ROUND(SUM(late_delivery_risk) * 100.0 / COUNT(*), 2) AS late_delivery_pct
FROM vw_fact_orders
GROUP BY shipping_mode
ORDER BY late_delivery_pct DESC;

-- QUERY 2: Regional Delivery Risk & Negative Profit (Revenue Leakage) Analysis
SELECT 
    market,
    COUNT(DISTINCT order_id) AS total_orders,
    ROUND(SUM(net_sales), 2) AS total_net_revenue,
    ROUND(SUM(profit), 2) AS total_profit,
    ROUND(SUM(CASE WHEN profit < 0 THEN profit ELSE 0 END), 2) AS total_loss_from_negative_orders,
    COUNT(CASE WHEN profit < 0 THEN 1 END) AS loss_making_order_lines,
    ROUND(SUM(late_delivery_risk) * 100.0 / COUNT(*), 2) AS late_delivery_pct
FROM vw_fact_orders
GROUP BY market
ORDER BY total_net_revenue DESC;

-- QUERY 3: Top 3 Revenue-Generating Products per Department (CTE + JOIN + DENSE_RANK)
WITH Department_Product_Stats AS (
    SELECT 
        p.department_name,
        p.product_card_id,
        p.product_name,
        ROUND(SUM(f.net_sales), 2) AS total_revenue,
        SUM(f.order_item_quantity) AS total_units_sold,
        ROUND(SUM(f.late_delivery_risk) * 100.0 / COUNT(*), 2) AS late_delivery_pct,
        DENSE_RANK() OVER (
            PARTITION BY p.department_name 
            ORDER BY SUM(f.net_sales) DESC
        ) AS dept_rank
    FROM vw_fact_orders f
    INNER JOIN vw_dim_products p 
        ON f.product_card_id = p.product_card_id
    GROUP BY p.department_name, p.product_card_id, p.product_name
)
SELECT * 
FROM Department_Product_Stats
WHERE dept_rank <= 3
ORDER BY department_name, dept_rank;

-- QUERY 4: Month-over-Month (MoM %) Revenue Growth & SLA Trend (CTE + LAG Window Function)
WITH Monthly_Performance AS (
    SELECT 
        DATE_FORMAT(order_date, '%Y-%m') AS order_year_month,
        COUNT(DISTINCT order_id) AS total_orders,
        ROUND(SUM(net_sales), 2) AS current_month_revenue,
        ROUND(SUM(is_on_time) * 100.0 / COUNT(*), 2) AS on_time_pct
    FROM vw_fact_orders
    GROUP BY DATE_FORMAT(order_date, '%Y-%m')
)
SELECT 
    order_year_month,
    total_orders,
    current_month_revenue,
    LAG(current_month_revenue, 1) OVER (ORDER BY order_year_month) AS previous_month_revenue,
    ROUND(
        (current_month_revenue - LAG(current_month_revenue, 1) OVER (ORDER BY order_year_month)) * 100.0 
        / LAG(current_month_revenue, 1) OVER (ORDER BY order_year_month), 
    2) AS mom_growth_pct,
    on_time_pct
FROM Monthly_Performance
ORDER BY order_year_month;

-- QUERY 5: ABC Inventory Pareto Classification (Cumulative Window SUM)
WITH Product_Revenue AS (
    SELECT 
        p.product_card_id,
        p.product_name,
        p.category_name,
        ROUND(SUM(f.net_sales), 2) AS product_revenue,
        SUM(f.order_item_quantity) AS total_qty
    FROM vw_fact_orders f
    INNER JOIN vw_dim_products p 
        ON f.product_card_id = p.product_card_id
    GROUP BY p.product_card_id, p.product_name, p.category_name
),
Pareto_Calc AS (
    SELECT 
        product_card_id,
        product_name,
        category_name,
        product_revenue,
        total_qty,
        ROUND(
            SUM(product_revenue) OVER (ORDER BY product_revenue DESC) * 100.0 
            / SUM(product_revenue) OVER (), 
        2) AS cumulative_revenue_pct
    FROM Product_Revenue
)
SELECT 
    product_card_id,
    product_name,
    category_name,
    product_revenue,
    cumulative_revenue_pct,
    CASE 
        WHEN cumulative_revenue_pct <= 80.00 THEN 'Category A'
        WHEN cumulative_revenue_pct <= 95.00 THEN 'Category B'
        ELSE 'Category C'
    END AS sql_abc_category
FROM Pareto_Calc
ORDER BY product_revenue DESC;