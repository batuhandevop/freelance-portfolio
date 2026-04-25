-- ============================================================================
-- BUSINESS INTELLIGENCE QUERIES
-- PostgreSQL
-- ============================================================================
-- A collection of analytical queries against the e-commerce schema that
-- answer common business questions. Each query is annotated with the
-- question it solves and the SQL techniques it demonstrates.
-- ============================================================================


-- ============================================================================
-- 1. DAILY REVENUE REPORT
-- Business question: What is the total revenue and order count for each day?
-- Techniques: aggregation, date truncation, filtering by status
-- ============================================================================
SELECT
    DATE(o.order_date)          AS order_day,
    COUNT(DISTINCT o.order_id)  AS total_orders,
    SUM(oi.line_total)          AS daily_revenue
FROM orders o
JOIN order_items oi ON oi.order_id = o.order_id
WHERE o.status NOT IN ('cancelled', 'refunded')
GROUP BY DATE(o.order_date)
ORDER BY order_day;


-- ============================================================================
-- 2. MONTHLY REVENUE REPORT WITH RUNNING TOTAL
-- Business question: What is the monthly revenue trend and cumulative total?
-- Techniques: window function (SUM OVER), date_trunc
-- ============================================================================
SELECT
    DATE_TRUNC('month', o.order_date)::DATE  AS month,
    SUM(oi.line_total)                        AS monthly_revenue,
    SUM(SUM(oi.line_total)) OVER (
        ORDER BY DATE_TRUNC('month', o.order_date)
    )                                         AS cumulative_revenue
FROM orders o
JOIN order_items oi ON oi.order_id = o.order_id
WHERE o.status NOT IN ('cancelled', 'refunded')
GROUP BY DATE_TRUNC('month', o.order_date)
ORDER BY month;


-- ============================================================================
-- 3. REVENUE BY CATEGORY
-- Business question: Which product categories generate the most revenue?
-- Techniques: multi-table JOIN, aggregation, percentage calculation
-- ============================================================================
WITH category_revenue AS (
    SELECT
        c.name                    AS category,
        SUM(oi.line_total)        AS revenue,
        COUNT(DISTINCT o.order_id) AS order_count
    FROM order_items oi
    JOIN orders   o ON o.order_id   = oi.order_id
    JOIN products p ON p.product_id = oi.product_id
    JOIN categories c ON c.category_id = p.category_id
    WHERE o.status NOT IN ('cancelled', 'refunded')
    GROUP BY c.name
)
SELECT
    category,
    revenue,
    order_count,
    ROUND(revenue * 100.0 / SUM(revenue) OVER (), 2) AS pct_of_total
FROM category_revenue
ORDER BY revenue DESC;


-- ============================================================================
-- 4. TOP 10 CUSTOMERS BY LIFETIME SPEND
-- Business question: Who are our most valuable customers?
-- Techniques: aggregation, JOIN, LIMIT
-- ============================================================================
SELECT
    c.customer_id,
    c.first_name || ' ' || c.last_name  AS full_name,
    c.city,
    COUNT(DISTINCT o.order_id)           AS total_orders,
    SUM(oi.line_total)                   AS lifetime_spend,
    ROUND(AVG(oi.line_total), 2)         AS avg_item_value
FROM customers c
JOIN orders o      ON o.customer_id = c.customer_id
JOIN order_items oi ON oi.order_id  = o.order_id
WHERE o.status NOT IN ('cancelled', 'refunded')
GROUP BY c.customer_id, c.first_name, c.last_name, c.city
ORDER BY lifetime_spend DESC
LIMIT 10;


-- ============================================================================
-- 5. PRODUCT PERFORMANCE ANALYSIS
-- Business question: Which products sell the most and have the best reviews?
-- Techniques: LEFT JOIN, COALESCE, conditional aggregation
-- ============================================================================
SELECT
    p.product_id,
    p.name,
    p.price,
    COALESCE(sales.units_sold, 0)        AS units_sold,
    COALESCE(sales.total_revenue, 0)     AS total_revenue,
    COALESCE(rev.avg_rating, 0)          AS avg_rating,
    COALESCE(rev.review_count, 0)        AS review_count
FROM products p
LEFT JOIN (
    SELECT
        oi.product_id,
        SUM(oi.quantity)     AS units_sold,
        SUM(oi.line_total)   AS total_revenue
    FROM order_items oi
    JOIN orders o ON o.order_id = oi.order_id
    WHERE o.status NOT IN ('cancelled', 'refunded')
    GROUP BY oi.product_id
) sales ON sales.product_id = p.product_id
LEFT JOIN (
    SELECT
        product_id,
        ROUND(AVG(rating), 2) AS avg_rating,
        COUNT(*)               AS review_count
    FROM reviews
    GROUP BY product_id
) rev ON rev.product_id = p.product_id
WHERE p.is_active = TRUE
ORDER BY total_revenue DESC;


-- ============================================================================
-- 6. INVENTORY ALERTS -- LOW STOCK ITEMS
-- Business question: Which products are below their reorder level and need
--                    restocking?
-- Techniques: JOIN, CASE expression, filtering
-- ============================================================================
SELECT
    p.sku,
    p.name,
    i.quantity           AS current_stock,
    i.reorder_level,
    i.reorder_quantity   AS suggested_reorder,
    CASE
        WHEN i.quantity = 0                THEN 'OUT OF STOCK'
        WHEN i.quantity <= i.reorder_level THEN 'REORDER NOW'
        WHEN i.quantity <= i.reorder_level * 1.5 THEN 'LOW STOCK'
        ELSE 'OK'
    END AS stock_status
FROM inventory i
JOIN products p ON p.product_id = i.product_id
WHERE i.quantity <= i.reorder_level * 1.5
ORDER BY i.quantity ASC;


-- ============================================================================
-- 7. CUSTOMER COHORT ANALYSIS
-- Business question: How do customers acquired in each month behave over time
--                    (retention and revenue)?
-- Techniques: CTE, date_trunc, window-less cohort pattern
-- ============================================================================
WITH customer_cohort AS (
    -- Determine each customer's acquisition month (first order date)
    SELECT
        customer_id,
        DATE_TRUNC('month', MIN(order_date))::DATE AS cohort_month
    FROM orders
    WHERE status NOT IN ('cancelled', 'refunded')
    GROUP BY customer_id
),
cohort_activity AS (
    SELECT
        cc.cohort_month,
        DATE_TRUNC('month', o.order_date)::DATE AS activity_month,
        COUNT(DISTINCT o.customer_id)            AS active_customers,
        SUM(oi.line_total)                       AS revenue
    FROM orders o
    JOIN order_items oi ON oi.order_id = o.order_id
    JOIN customer_cohort cc ON cc.customer_id = o.customer_id
    WHERE o.status NOT IN ('cancelled', 'refunded')
    GROUP BY cc.cohort_month, DATE_TRUNC('month', o.order_date)
)
SELECT
    cohort_month,
    activity_month,
    -- Months since acquisition
    EXTRACT(YEAR FROM AGE(activity_month, cohort_month)) * 12
        + EXTRACT(MONTH FROM AGE(activity_month, cohort_month)) AS months_since_acq,
    active_customers,
    revenue
FROM cohort_activity
ORDER BY cohort_month, activity_month;


-- ============================================================================
-- 8. 3-MONTH MOVING AVERAGE OF MONTHLY REVENUE
-- Business question: What is the smoothed revenue trend to filter out
--                    month-to-month noise?
-- Techniques: window function (AVG OVER with ROWS frame)
-- ============================================================================
WITH monthly AS (
    SELECT
        DATE_TRUNC('month', o.order_date)::DATE AS month,
        SUM(oi.line_total)                       AS revenue
    FROM orders o
    JOIN order_items oi ON oi.order_id = o.order_id
    WHERE o.status NOT IN ('cancelled', 'refunded')
    GROUP BY DATE_TRUNC('month', o.order_date)
)
SELECT
    month,
    revenue,
    ROUND(
        AVG(revenue) OVER (
            ORDER BY month
            ROWS BETWEEN 2 PRECEDING AND CURRENT ROW
        ), 2
    ) AS moving_avg_3m
FROM monthly
ORDER BY month;


-- ============================================================================
-- 9. YEAR-OVER-YEAR MONTHLY REVENUE GROWTH
-- Business question: How does this month's revenue compare to the same month
--                    last year?
-- Techniques: LAG window function, date extraction
-- ============================================================================
WITH monthly AS (
    SELECT
        DATE_TRUNC('month', o.order_date)::DATE AS month,
        SUM(oi.line_total)                       AS revenue
    FROM orders o
    JOIN order_items oi ON oi.order_id = o.order_id
    WHERE o.status NOT IN ('cancelled', 'refunded')
    GROUP BY DATE_TRUNC('month', o.order_date)
)
SELECT
    month,
    revenue                                     AS current_revenue,
    LAG(revenue, 12) OVER (ORDER BY month)      AS prior_year_revenue,
    CASE
        WHEN LAG(revenue, 12) OVER (ORDER BY month) IS NOT NULL
        THEN ROUND(
            (revenue - LAG(revenue, 12) OVER (ORDER BY month))
            * 100.0
            / LAG(revenue, 12) OVER (ORDER BY month),
            2
        )
    END                                         AS yoy_growth_pct
FROM monthly
ORDER BY month;


-- ============================================================================
-- 10. RFM SEGMENTATION
-- Business question: How can we segment customers by Recency, Frequency, and
--                    Monetary value for targeted marketing?
-- Techniques: CTE, NTILE window function, CASE expression
-- ============================================================================
WITH rfm_raw AS (
    SELECT
        c.customer_id,
        c.first_name || ' ' || c.last_name AS full_name,
        -- Recency: days since last order
        CURRENT_DATE - MAX(o.order_date)::DATE AS recency_days,
        -- Frequency: number of orders
        COUNT(DISTINCT o.order_id) AS frequency,
        -- Monetary: total spend
        SUM(oi.line_total) AS monetary
    FROM customers c
    JOIN orders o       ON o.customer_id = c.customer_id
    JOIN order_items oi ON oi.order_id   = o.order_id
    WHERE o.status NOT IN ('cancelled', 'refunded')
    GROUP BY c.customer_id, c.first_name, c.last_name
),
rfm_scores AS (
    SELECT
        *,
        -- Lower recency = better (more recent), so we reverse the NTILE
        5 - NTILE(5) OVER (ORDER BY recency_days ASC) + 1  AS r_score,
        NTILE(5) OVER (ORDER BY frequency ASC)              AS f_score,
        NTILE(5) OVER (ORDER BY monetary ASC)               AS m_score
    FROM rfm_raw
)
SELECT
    customer_id,
    full_name,
    recency_days,
    frequency,
    ROUND(monetary, 2) AS monetary,
    r_score,
    f_score,
    m_score,
    r_score + f_score + m_score AS rfm_total,
    CASE
        WHEN r_score >= 4 AND f_score >= 4 AND m_score >= 4 THEN 'Champions'
        WHEN r_score >= 4 AND f_score >= 2                   THEN 'Loyal Customers'
        WHEN r_score >= 3 AND f_score <= 2 AND m_score >= 3  THEN 'Big Spenders'
        WHEN r_score <= 2 AND f_score >= 3                   THEN 'At Risk'
        WHEN r_score <= 2 AND f_score <= 2                   THEN 'Lost'
        ELSE 'Potential'
    END AS segment
FROM rfm_scores
ORDER BY rfm_total DESC, monetary DESC;


-- ============================================================================
-- 11. AVERAGE ORDER VALUE BY MONTH
-- Business question: Is our average order value trending up or down?
-- Techniques: aggregation, date_trunc
-- ============================================================================
SELECT
    DATE_TRUNC('month', o.order_date)::DATE AS month,
    COUNT(DISTINCT o.order_id)               AS order_count,
    ROUND(SUM(oi.line_total) / COUNT(DISTINCT o.order_id), 2) AS avg_order_value
FROM orders o
JOIN order_items oi ON oi.order_id = o.order_id
WHERE o.status NOT IN ('cancelled', 'refunded')
GROUP BY DATE_TRUNC('month', o.order_date)
ORDER BY month;


-- ============================================================================
-- 12. TOP-SELLING PRODUCTS PER CATEGORY (TOP 3)
-- Business question: What are the best-selling products in each category?
-- Techniques: ROW_NUMBER window function, CTE, PARTITION BY
-- ============================================================================
WITH product_sales AS (
    SELECT
        c.name              AS category,
        p.name              AS product_name,
        SUM(oi.quantity)     AS units_sold,
        SUM(oi.line_total)   AS revenue,
        ROW_NUMBER() OVER (
            PARTITION BY c.category_id
            ORDER BY SUM(oi.line_total) DESC
        ) AS rank_in_category
    FROM order_items oi
    JOIN orders   o ON o.order_id   = oi.order_id
    JOIN products p ON p.product_id = oi.product_id
    JOIN categories c ON c.category_id = p.category_id
    WHERE o.status NOT IN ('cancelled', 'refunded')
    GROUP BY c.category_id, c.name, p.product_id, p.name
)
SELECT
    category,
    product_name,
    units_sold,
    revenue,
    rank_in_category
FROM product_sales
WHERE rank_in_category <= 3
ORDER BY category, rank_in_category;


-- ============================================================================
-- 13. CUSTOMERS WHO HAVE NOT ORDERED IN THE LAST 90 DAYS
-- Business question: Who are the at-risk customers we should re-engage?
-- Techniques: subquery, NOT EXISTS alternative shown as LEFT JOIN / IS NULL
-- ============================================================================
SELECT
    c.customer_id,
    c.first_name || ' ' || c.last_name AS full_name,
    c.email,
    MAX(o.order_date)::DATE             AS last_order_date,
    CURRENT_DATE - MAX(o.order_date)::DATE AS days_since_last_order
FROM customers c
JOIN orders o ON o.customer_id = c.customer_id
WHERE o.status NOT IN ('cancelled', 'refunded')
GROUP BY c.customer_id, c.first_name, c.last_name, c.email
HAVING CURRENT_DATE - MAX(o.order_date)::DATE > 90
ORDER BY days_since_last_order DESC;


-- ============================================================================
-- 14. ORDER FULFILLMENT TIME ANALYSIS
-- Business question: How long does it take us to ship and deliver orders?
-- Techniques: EXTRACT, PERCENTILE_CONT (median), aggregation
-- ============================================================================
SELECT
    DATE_TRUNC('month', order_date)::DATE  AS month,
    COUNT(*)                                AS delivered_orders,
    ROUND(AVG(EXTRACT(EPOCH FROM (shipped_date - order_date)) / 3600), 1)
        AS avg_hours_to_ship,
    ROUND(AVG(EXTRACT(EPOCH FROM (delivered_date - order_date)) / 86400), 1)
        AS avg_days_to_deliver,
    ROUND(
        PERCENTILE_CONT(0.5) WITHIN GROUP (
            ORDER BY EXTRACT(EPOCH FROM (delivered_date - order_date)) / 86400
        )::NUMERIC, 1
    ) AS median_days_to_deliver
FROM orders
WHERE status = 'delivered'
  AND shipped_date IS NOT NULL
  AND delivered_date IS NOT NULL
GROUP BY DATE_TRUNC('month', order_date)
ORDER BY month;


-- ============================================================================
-- 15. PRODUCT REVIEW SENTIMENT SUMMARY
-- Business question: What is the distribution of review ratings per product
--                    and how does it correlate with sales?
-- Techniques: conditional aggregation with FILTER, LEFT JOIN
-- ============================================================================
SELECT
    p.name                                                    AS product,
    COUNT(r.review_id)                                        AS total_reviews,
    ROUND(AVG(r.rating), 2)                                   AS avg_rating,
    COUNT(*) FILTER (WHERE r.rating >= 4)                     AS positive_reviews,
    COUNT(*) FILTER (WHERE r.rating <= 2)                     AS negative_reviews,
    COALESCE(s.units_sold, 0)                                 AS units_sold,
    COALESCE(s.revenue, 0)                                    AS revenue
FROM products p
LEFT JOIN reviews r ON r.product_id = p.product_id
LEFT JOIN (
    SELECT
        oi.product_id,
        SUM(oi.quantity)   AS units_sold,
        SUM(oi.line_total) AS revenue
    FROM order_items oi
    JOIN orders o ON o.order_id = oi.order_id
    WHERE o.status NOT IN ('cancelled', 'refunded')
    GROUP BY oi.product_id
) s ON s.product_id = p.product_id
GROUP BY p.product_id, p.name, s.units_sold, s.revenue
HAVING COUNT(r.review_id) > 0
ORDER BY avg_rating DESC, total_reviews DESC;


-- ============================================================================
-- 16. BASKET ANALYSIS -- FREQUENTLY CO-PURCHASED PRODUCTS
-- Business question: Which products are commonly bought together?
-- Techniques: self-join on order_items, aggregation, LEAST/GREATEST for
--             deduplication of pairs
-- ============================================================================
SELECT
    LEAST(p1.name, p2.name)     AS product_a,
    GREATEST(p1.name, p2.name)  AS product_b,
    COUNT(*)                     AS times_bought_together
FROM order_items oi1
JOIN order_items oi2 ON oi1.order_id = oi2.order_id
                    AND oi1.product_id < oi2.product_id
JOIN products p1 ON p1.product_id = oi1.product_id
JOIN products p2 ON p2.product_id = oi2.product_id
GROUP BY LEAST(p1.name, p2.name), GREATEST(p1.name, p2.name)
HAVING COUNT(*) >= 2
ORDER BY times_bought_together DESC;


-- ============================================================================
-- 17. REVENUE CONCENTRATION -- PARETO (80/20) ANALYSIS
-- Business question: Do 20% of our products account for 80% of revenue?
-- Techniques: CTE, cumulative SUM window, percentage calculation
-- ============================================================================
WITH product_revenue AS (
    SELECT
        p.name,
        SUM(oi.line_total) AS revenue
    FROM order_items oi
    JOIN orders o   ON o.order_id   = oi.order_id
    JOIN products p ON p.product_id = oi.product_id
    WHERE o.status NOT IN ('cancelled', 'refunded')
    GROUP BY p.product_id, p.name
),
ranked AS (
    SELECT
        name,
        revenue,
        SUM(revenue) OVER (ORDER BY revenue DESC) AS cumulative_revenue,
        SUM(revenue) OVER ()                       AS total_revenue,
        ROW_NUMBER() OVER (ORDER BY revenue DESC)  AS product_rank,
        COUNT(*) OVER ()                           AS total_products
    FROM product_revenue
)
SELECT
    product_rank,
    name,
    revenue,
    ROUND(cumulative_revenue * 100.0 / total_revenue, 2) AS cumulative_pct,
    ROUND(product_rank * 100.0 / total_products, 2)      AS pct_of_products,
    CASE
        WHEN cumulative_revenue <= total_revenue * 0.8 THEN 'Top 80% revenue'
        ELSE 'Long tail'
    END AS pareto_group
FROM ranked
ORDER BY product_rank;
