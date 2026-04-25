-- ============================================================================
-- E-COMMERCE DATABASE SCHEMA
-- PostgreSQL 15+
-- ============================================================================
-- A production-style schema for an e-commerce platform covering customers,
-- products, orders, reviews, and inventory management.
-- ============================================================================

BEGIN;

-- ----------------------------------------------------------------------------
-- CATEGORIES
-- Flat category table. A self-referencing parent_id column allows a simple
-- hierarchy (e.g., Electronics > Laptops) without a recursive closure table,
-- keeping the schema approachable while still supporting nested categories.
-- ----------------------------------------------------------------------------
CREATE TABLE categories (
    category_id   SERIAL       PRIMARY KEY,
    name          VARCHAR(100) NOT NULL UNIQUE,
    parent_id     INT          REFERENCES categories(category_id) ON DELETE SET NULL,
    description   TEXT,
    created_at    TIMESTAMPTZ  NOT NULL DEFAULT now(),
    updated_at    TIMESTAMPTZ  NOT NULL DEFAULT now()
);

CREATE INDEX idx_categories_parent ON categories(parent_id);

-- ----------------------------------------------------------------------------
-- CUSTOMERS
-- Stores account-level information. Email is unique and loosely validated with
-- a CHECK constraint. Passwords would be hashed at the application layer; the
-- column is omitted here intentionally to keep the schema focused on data
-- modelling rather than authentication.
-- ----------------------------------------------------------------------------
CREATE TABLE customers (
    customer_id   SERIAL        PRIMARY KEY,
    first_name    VARCHAR(50)   NOT NULL,
    last_name     VARCHAR(50)   NOT NULL,
    email         VARCHAR(255)  NOT NULL UNIQUE,
    phone         VARCHAR(20),
    address_line1 VARCHAR(255),
    address_line2 VARCHAR(255),
    city          VARCHAR(100),
    state         VARCHAR(100),
    postal_code   VARCHAR(20),
    country       VARCHAR(100)  NOT NULL DEFAULT 'Turkey',
    created_at    TIMESTAMPTZ   NOT NULL DEFAULT now(),
    updated_at    TIMESTAMPTZ   NOT NULL DEFAULT now(),

    CONSTRAINT chk_email_format CHECK (email ~* '^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$')
);

CREATE INDEX idx_customers_email ON customers(email);
CREATE INDEX idx_customers_country ON customers(country);

-- ----------------------------------------------------------------------------
-- PRODUCTS
-- Each product belongs to exactly one category. Price and weight must be
-- positive. The is_active flag supports soft-deletion so historical order
-- line items remain valid.
-- ----------------------------------------------------------------------------
CREATE TABLE products (
    product_id    SERIAL         PRIMARY KEY,
    category_id   INT            NOT NULL REFERENCES categories(category_id) ON DELETE RESTRICT,
    name          VARCHAR(200)   NOT NULL,
    description   TEXT,
    price         NUMERIC(10,2)  NOT NULL,
    weight_kg     NUMERIC(6,2),
    sku           VARCHAR(50)    NOT NULL UNIQUE,
    is_active     BOOLEAN        NOT NULL DEFAULT TRUE,
    created_at    TIMESTAMPTZ    NOT NULL DEFAULT now(),
    updated_at    TIMESTAMPTZ    NOT NULL DEFAULT now(),

    CONSTRAINT chk_price_positive  CHECK (price > 0),
    CONSTRAINT chk_weight_positive CHECK (weight_kg IS NULL OR weight_kg > 0)
);

CREATE INDEX idx_products_category ON products(category_id);
CREATE INDEX idx_products_active   ON products(is_active) WHERE is_active = TRUE;

-- ----------------------------------------------------------------------------
-- INVENTORY
-- One-to-one with products. Separated from the products table so that
-- frequent stock updates do not cause unnecessary row-level locks on the
-- wider products row during high-traffic order processing.
-- ----------------------------------------------------------------------------
CREATE TABLE inventory (
    inventory_id     SERIAL  PRIMARY KEY,
    product_id       INT     NOT NULL UNIQUE REFERENCES products(product_id) ON DELETE CASCADE,
    quantity         INT     NOT NULL DEFAULT 0,
    reorder_level    INT     NOT NULL DEFAULT 10,
    reorder_quantity INT     NOT NULL DEFAULT 50,
    updated_at       TIMESTAMPTZ NOT NULL DEFAULT now(),

    CONSTRAINT chk_quantity_nonneg       CHECK (quantity >= 0),
    CONSTRAINT chk_reorder_level_nonneg  CHECK (reorder_level >= 0),
    CONSTRAINT chk_reorder_qty_positive  CHECK (reorder_quantity > 0)
);

-- ----------------------------------------------------------------------------
-- ORDERS
-- status uses a CHECK constraint instead of a lookup table to keep the demo
-- self-contained. In a larger system a separate order_statuses table with
-- workflow rules would be preferable.
-- ----------------------------------------------------------------------------
CREATE TABLE orders (
    order_id       SERIAL         PRIMARY KEY,
    customer_id    INT            NOT NULL REFERENCES customers(customer_id) ON DELETE RESTRICT,
    status         VARCHAR(20)    NOT NULL DEFAULT 'pending',
    total_amount   NUMERIC(12,2)  NOT NULL DEFAULT 0,
    shipping_addr  TEXT,
    order_date     TIMESTAMPTZ    NOT NULL DEFAULT now(),
    shipped_date   TIMESTAMPTZ,
    delivered_date TIMESTAMPTZ,
    created_at     TIMESTAMPTZ    NOT NULL DEFAULT now(),
    updated_at     TIMESTAMPTZ    NOT NULL DEFAULT now(),

    CONSTRAINT chk_order_status CHECK (status IN ('pending','processing','shipped','delivered','cancelled','refunded')),
    CONSTRAINT chk_total_nonneg CHECK (total_amount >= 0)
);

CREATE INDEX idx_orders_customer   ON orders(customer_id);
CREATE INDEX idx_orders_status     ON orders(status);
CREATE INDEX idx_orders_order_date ON orders(order_date);

-- ----------------------------------------------------------------------------
-- ORDER ITEMS
-- Junction table between orders and products. unit_price is snapshotted at
-- the time of purchase so that future product price changes do not distort
-- historical revenue.
-- ----------------------------------------------------------------------------
CREATE TABLE order_items (
    order_item_id SERIAL         PRIMARY KEY,
    order_id      INT            NOT NULL REFERENCES orders(order_id) ON DELETE CASCADE,
    product_id    INT            NOT NULL REFERENCES products(product_id) ON DELETE RESTRICT,
    quantity      INT            NOT NULL,
    unit_price    NUMERIC(10,2)  NOT NULL,
    discount_pct  NUMERIC(5,2)   NOT NULL DEFAULT 0,
    line_total    NUMERIC(12,2)  GENERATED ALWAYS AS (quantity * unit_price * (1 - discount_pct / 100)) STORED,

    CONSTRAINT chk_oi_qty_positive       CHECK (quantity > 0),
    CONSTRAINT chk_oi_price_positive     CHECK (unit_price > 0),
    CONSTRAINT chk_oi_discount_range     CHECK (discount_pct >= 0 AND discount_pct <= 100)
);

CREATE INDEX idx_order_items_order   ON order_items(order_id);
CREATE INDEX idx_order_items_product ON order_items(product_id);

-- ----------------------------------------------------------------------------
-- REVIEWS
-- A customer may review a product only once (UNIQUE constraint). Rating is
-- restricted to 1-5 stars.
-- ----------------------------------------------------------------------------
CREATE TABLE reviews (
    review_id   SERIAL    PRIMARY KEY,
    product_id  INT       NOT NULL REFERENCES products(product_id) ON DELETE CASCADE,
    customer_id INT       NOT NULL REFERENCES customers(customer_id) ON DELETE CASCADE,
    rating      SMALLINT  NOT NULL,
    title       VARCHAR(200),
    body        TEXT,
    created_at  TIMESTAMPTZ NOT NULL DEFAULT now(),

    CONSTRAINT uq_one_review_per_customer UNIQUE (product_id, customer_id),
    CONSTRAINT chk_rating_range CHECK (rating BETWEEN 1 AND 5)
);

CREATE INDEX idx_reviews_product  ON reviews(product_id);
CREATE INDEX idx_reviews_customer ON reviews(customer_id);


-- ============================================================================
-- SAMPLE DATA
-- ============================================================================

-- ---- CATEGORIES -----------------------------------------------------------
INSERT INTO categories (name, description) VALUES
    ('Electronics',       'Electronic devices and accessories'),
    ('Clothing',          'Apparel for men, women, and children'),
    ('Home & Kitchen',    'Household items and kitchenware'),
    ('Books',             'Physical and digital books'),
    ('Sports & Outdoors', 'Sporting goods and outdoor gear');

-- Sub-categories
INSERT INTO categories (name, parent_id, description) VALUES
    ('Laptops',         1, 'Notebook computers'),
    ('Smartphones',     1, 'Mobile phones and accessories'),
    ('Headphones',      1, 'Audio headphones and earbuds'),
    ('Men''s Clothing', 2, 'Apparel for men'),
    ('Women''s Clothing', 2, 'Apparel for women'),
    ('Cookware',        3, 'Pots, pans, and cooking tools'),
    ('Furniture',       3, 'Home furniture'),
    ('Fiction',         4, 'Novels and fiction literature'),
    ('Non-Fiction',     4, 'Educational and non-fiction books'),
    ('Fitness',         5, 'Gym and fitness equipment');


-- ---- CUSTOMERS (25) -------------------------------------------------------
INSERT INTO customers (first_name, last_name, email, phone, city, state, country) VALUES
    ('Ahmet',    'Yilmaz',     'ahmet.yilmaz@email.com',       '+905321234567', 'Istanbul',   'Istanbul',    'Turkey'),
    ('Elif',     'Kaya',       'elif.kaya@email.com',           '+905339876543', 'Ankara',     'Ankara',      'Turkey'),
    ('Mehmet',   'Demir',      'mehmet.demir@email.com',        '+905441122334', 'Izmir',      'Izmir',       'Turkey'),
    ('Zeynep',   'Celik',      'zeynep.celik@email.com',        '+905557788990', 'Bursa',      'Bursa',       'Turkey'),
    ('Mustafa',  'Ozturk',     'mustafa.ozturk@email.com',      '+905061234567', 'Antalya',    'Antalya',     'Turkey'),
    ('Ayse',     'Arslan',     'ayse.arslan@email.com',         '+905322334455', 'Adana',      'Adana',       'Turkey'),
    ('Emre',     'Dogan',      'emre.dogan@email.com',          '+905443344556', 'Konya',      'Konya',       'Turkey'),
    ('Fatma',    'Kilic',      'fatma.kilic@email.com',         '+905554455667', 'Gaziantep',  'Gaziantep',   'Turkey'),
    ('Burak',    'Sen',        'burak.sen@email.com',           '+905065566778', 'Eskisehir',  'Eskisehir',   'Turkey'),
    ('Selin',    'Sahin',      'selin.sahin@email.com',         '+905326677889', 'Trabzon',    'Trabzon',     'Turkey'),
    ('Can',      'Yildiz',     'can.yildiz@email.com',          '+905337788990', 'Samsun',     'Samsun',      'Turkey'),
    ('Deniz',    'Aydin',      'deniz.aydin@email.com',         '+905448899001', 'Mersin',     'Mersin',      'Turkey'),
    ('Gokhan',   'Ozdemir',    'gokhan.ozdemir@email.com',      '+905559900112', 'Kayseri',    'Kayseri',     'Turkey'),
    ('Hande',    'Kurt',       'hande.kurt@email.com',          '+905060011223', 'Mugla',      'Mugla',       'Turkey'),
    ('Ibrahim',  'Tas',        'ibrahim.tas@email.com',         '+905321122334', 'Diyarbakir', 'Diyarbakir',  'Turkey'),
    ('Julide',   'Polat',      'julide.polat@email.com',        '+905332233445', 'Malatya',    'Malatya',     'Turkey'),
    ('Kerem',    'Erdogan',    'kerem.erdogan@email.com',       '+905443344556', 'Erzurum',    'Erzurum',     'Turkey'),
    ('Leyla',    'Acar',       'leyla.acar@email.com',          '+905554455667', 'Van',        'Van',         'Turkey'),
    ('Murat',    'Koc',        'murat.koc@email.com',           '+905065566778', 'Tekirdak',   'Tekirdag',    'Turkey'),
    ('Neslihan', 'Cinar',      'neslihan.cinar@email.com',      '+905326677889', 'Manisa',     'Manisa',      'Turkey'),
    ('Onur',     'Bal',        'onur.bal@email.com',            '+905337788990', 'Denizli',    'Denizli',     'Turkey'),
    ('Pinar',    'Gunes',      'pinar.gunes@email.com',         '+905448899001', 'Sanliurfa',  'Sanliurfa',   'Turkey'),
    ('Recep',    'Turan',      'recep.turan@email.com',         '+905559900112', 'Hatay',      'Hatay',       'Turkey'),
    ('Sevgi',    'Aktas',      'sevgi.aktas@email.com',         '+905060011223', 'Balikesir',  'Balikesir',   'Turkey'),
    ('Tolga',    'Bulut',      'tolga.bulut@email.com',         '+905321122335', 'Isparta',    'Isparta',     'Turkey');


-- ---- PRODUCTS (35) --------------------------------------------------------
INSERT INTO products (category_id, name, description, price, weight_kg, sku) VALUES
    (6,  'ProBook 15 Laptop',          '15.6" FHD, Intel i7, 16GB RAM, 512GB SSD',         18999.99, 1.80, 'ELEC-LAP-001'),
    (6,  'UltraSlim 14 Laptop',        '14" QHD, AMD Ryzen 7, 16GB RAM, 1TB SSD',          22499.00, 1.35, 'ELEC-LAP-002'),
    (6,  'Budget Notebook 15',         '15.6" HD, Intel i3, 8GB RAM, 256GB SSD',             8999.00, 2.10, 'ELEC-LAP-003'),
    (7,  'Galaxy Ultra Phone',         '6.8" AMOLED, 256GB, 5G',                            34999.00, 0.23, 'ELEC-PHN-001'),
    (7,  'iPhone Pro Max',             '6.7" Super Retina, 256GB',                          54999.00, 0.24, 'ELEC-PHN-002'),
    (7,  'Budget Smartphone A12',      '6.5" IPS, 128GB, 4G',                               5999.00, 0.19, 'ELEC-PHN-003'),
    (8,  'Noise-Cancel Over-Ear',      'Active noise cancelling, 30hr battery',              3499.00, 0.25, 'ELEC-AUD-001'),
    (8,  'Wireless Earbuds Pro',       'ANC, water-resistant, 8hr battery',                  1999.00, 0.05, 'ELEC-AUD-002'),
    (9,  'Classic Oxford Shirt',       '100% cotton, slim fit',                               599.00, 0.30, 'CLTH-MEN-001'),
    (9,  'Wool Blend Blazer',          'Two-button, notch lapel',                            2199.00, 0.85, 'CLTH-MEN-002'),
    (9,  'Slim Chino Pants',           'Stretch cotton, tapered leg',                         749.00, 0.45, 'CLTH-MEN-003'),
    (10, 'Cashmere Sweater',           'V-neck, lightweight cashmere blend',                 1899.00, 0.35, 'CLTH-WMN-001'),
    (10, 'High-Waist Denim Jeans',     'Organic cotton, straight leg',                       999.00, 0.60, 'CLTH-WMN-002'),
    (10, 'Silk Blouse',                'Button-front, relaxed fit',                          1299.00, 0.20, 'CLTH-WMN-003'),
    (11, 'Cast Iron Skillet 12"',      'Pre-seasoned, oven safe to 260C',                    899.00, 3.60, 'HOME-CK-001'),
    (11, 'Stainless Steel Pot Set',    '5-piece, induction compatible',                     2499.00, 5.50, 'HOME-CK-002'),
    (11, 'Non-Stick Wok 14"',          'Ceramic coating, cool-touch handle',                 649.00, 1.20, 'HOME-CK-003'),
    (12, 'Ergonomic Office Chair',     'Mesh back, lumbar support, adjustable arms',         4999.00, 15.0, 'HOME-FR-001'),
    (12, 'Standing Desk 140cm',        'Electric height adjustment, memory presets',         7999.00, 30.0, 'HOME-FR-002'),
    (12, 'Bookshelf 5-Tier',           'Solid oak, 180cm tall',                              3299.00, 22.0, 'HOME-FR-003'),
    (13, 'The Art of SQL',             'Hardcover, 420 pages',                                189.00, 0.65, 'BOOK-FIC-001'),
    (13, 'Istanbul: A Novel',          'Paperback, 380 pages',                                 89.00, 0.40, 'BOOK-FIC-002'),
    (13, 'Mystery at the Bazaar',      'Paperback, 290 pages',                                 69.00, 0.35, 'BOOK-FIC-003'),
    (14, 'Data-Driven Decisions',      'Hardcover, 510 pages',                                249.00, 0.80, 'BOOK-NF-001'),
    (14, 'Clean Architecture',         'Paperback, 430 pages',                                199.00, 0.55, 'BOOK-NF-002'),
    (14, 'Designing Data Systems',     'Hardcover, 620 pages',                                329.00, 0.90, 'BOOK-NF-003'),
    (15, 'Adjustable Dumbbell Set',    '2-24kg per dumbbell, quick-change',                  3999.00, 25.0, 'SPRT-FT-001'),
    (15, 'Yoga Mat Premium',           '6mm thick, non-slip, eco-friendly',                   399.00, 1.20, 'SPRT-FT-002'),
    (15, 'Resistance Band Kit',        '5 bands, door anchor, carry bag',                     299.00, 0.50, 'SPRT-FT-003'),
    (15, 'Treadmill Pro 3000',         'Foldable, 20km/h max, heart-rate monitor',          12999.00, 65.0, 'SPRT-FT-004'),
    (15, 'Foam Roller 45cm',           'High-density EVA foam',                               179.00, 0.30, 'SPRT-FT-005'),
    (9,  'Leather Belt Classic',       'Full-grain leather, brushed nickel buckle',           449.00, 0.25, 'CLTH-MEN-004'),
    (10, 'Running Leggings',           'Moisture-wicking, side pocket',                       599.00, 0.20, 'CLTH-WMN-004'),
    (7,  'Phone Case - Clear',         'Shockproof TPU, slim profile',                        149.00, 0.04, 'ELEC-PHN-004'),
    (8,  'Portable Speaker Mini',      'Bluetooth 5.0, waterproof, 12hr battery',             899.00, 0.35, 'ELEC-AUD-003');


-- ---- INVENTORY ------------------------------------------------------------
INSERT INTO inventory (product_id, quantity, reorder_level, reorder_quantity) VALUES
    (1,  45,  10, 30),
    (2,  22,  10, 20),
    (3,  80,  15, 50),
    (4,  35,  10, 25),
    (5,  18,   8, 20),
    (6,  120, 20, 80),
    (7,  65,  15, 40),
    (8,  90,  20, 60),
    (9,  200, 30, 100),
    (10, 50,  10, 30),
    (11, 150, 25, 80),
    (12, 40,  10, 25),
    (13, 110, 20, 60),
    (14, 70,  15, 40),
    (15, 55,  10, 30),
    (16, 30,   8, 20),
    (17, 85,  15, 50),
    (18, 12,   5, 15),
    (19, 8,    5, 10),
    (20, 25,   8, 20),
    (21, 300, 50, 150),
    (22, 180, 30, 100),
    (23, 220, 40, 120),
    (24, 95,  15, 50),
    (25, 140, 20, 60),
    (26, 60,  10, 30),
    (27, 20,   8, 20),
    (28, 250, 40, 120),
    (29, 310, 50, 150),
    (30, 3,    5, 10),
    (31, 400, 60, 200),
    (32, 75,  15, 40),
    (33, 160, 25, 80),
    (34, 500, 80, 250),
    (35, 45,  10, 30);


-- ---- ORDERS (55) ----------------------------------------------------------
-- Orders spread across 2025-Q1 through 2026-Q1 to support time-series queries.
INSERT INTO orders (customer_id, status, total_amount, order_date, shipped_date, delivered_date) VALUES
    -- 2025 Q1
    (1,  'delivered',  19598.99, '2025-01-05 10:30:00+03', '2025-01-07 09:00:00+03', '2025-01-10 14:00:00+03'),
    (2,  'delivered',   3498.00, '2025-01-12 14:15:00+03', '2025-01-13 11:00:00+03', '2025-01-16 10:00:00+03'),
    (3,  'delivered',  35748.00, '2025-01-20 09:45:00+03', '2025-01-22 08:00:00+03', '2025-01-25 16:00:00+03'),
    (4,  'delivered',   1648.00, '2025-02-02 16:20:00+03', '2025-02-03 10:00:00+03', '2025-02-06 12:00:00+03'),
    (5,  'delivered',  54999.00, '2025-02-14 11:00:00+03', '2025-02-16 09:00:00+03', '2025-02-19 15:00:00+03'),
    (6,  'delivered',   2997.00, '2025-02-20 13:30:00+03', '2025-02-21 10:00:00+03', '2025-02-24 11:00:00+03'),
    (7,  'delivered',   8999.00, '2025-03-01 10:00:00+03', '2025-03-03 09:00:00+03', '2025-03-06 14:00:00+03'),
    (8,  'delivered',  12998.00, '2025-03-10 15:45:00+03', '2025-03-12 08:00:00+03', '2025-03-15 16:00:00+03'),
    (9,  'delivered',    438.00, '2025-03-18 09:20:00+03', '2025-03-19 10:00:00+03', '2025-03-22 12:00:00+03'),
    (10, 'delivered',   4999.00, '2025-03-25 14:10:00+03', '2025-03-27 09:00:00+03', '2025-03-30 15:00:00+03'),

    -- 2025 Q2
    (1,  'delivered',  22499.00, '2025-04-03 11:00:00+03', '2025-04-05 08:00:00+03', '2025-04-08 13:00:00+03'),
    (11, 'delivered',   5999.00, '2025-04-10 16:30:00+03', '2025-04-12 10:00:00+03', '2025-04-15 14:00:00+03'),
    (12, 'delivered',   1998.00, '2025-04-18 10:15:00+03', '2025-04-19 09:00:00+03', '2025-04-22 11:00:00+03'),
    (13, 'delivered',   7999.00, '2025-05-02 14:00:00+03', '2025-05-04 08:00:00+03', '2025-05-07 16:00:00+03'),
    (14, 'delivered',    528.00, '2025-05-10 09:30:00+03', '2025-05-11 10:00:00+03', '2025-05-14 12:00:00+03'),
    (2,  'delivered',  34999.00, '2025-05-15 13:45:00+03', '2025-05-17 09:00:00+03', '2025-05-20 15:00:00+03'),
    (15, 'delivered',   3999.00, '2025-05-22 11:20:00+03', '2025-05-24 08:00:00+03', '2025-05-27 14:00:00+03'),
    (3,  'delivered',   1199.00, '2025-06-01 10:00:00+03', '2025-06-02 09:00:00+03', '2025-06-05 13:00:00+03'),
    (16, 'delivered',  12999.00, '2025-06-08 15:00:00+03', '2025-06-10 10:00:00+03', '2025-06-13 16:00:00+03'),
    (17, 'delivered',   2498.00, '2025-06-15 09:45:00+03', '2025-06-16 08:00:00+03', '2025-06-19 12:00:00+03'),

    -- 2025 Q3
    (4,  'delivered',  18999.99, '2025-07-01 11:30:00+03', '2025-07-03 09:00:00+03', '2025-07-06 14:00:00+03'),
    (18, 'delivered',    897.00, '2025-07-08 14:20:00+03', '2025-07-09 10:00:00+03', '2025-07-12 11:00:00+03'),
    (5,  'delivered',   3299.00, '2025-07-15 10:10:00+03', '2025-07-17 08:00:00+03', '2025-07-20 15:00:00+03'),
    (19, 'delivered',   1499.00, '2025-07-22 16:00:00+03', '2025-07-23 09:00:00+03', '2025-07-26 13:00:00+03'),
    (20, 'delivered',   5999.00, '2025-08-01 09:00:00+03', '2025-08-03 08:00:00+03', '2025-08-06 14:00:00+03'),
    (6,  'delivered',  22499.00, '2025-08-10 13:15:00+03', '2025-08-12 10:00:00+03', '2025-08-15 16:00:00+03'),
    (21, 'delivered',    758.00, '2025-08-18 10:30:00+03', '2025-08-19 09:00:00+03', '2025-08-22 12:00:00+03'),
    (22, 'delivered',  34999.00, '2025-08-25 15:45:00+03', '2025-08-27 08:00:00+03', '2025-08-30 15:00:00+03'),
    (7,  'delivered',   4999.00, '2025-09-02 11:00:00+03', '2025-09-04 09:00:00+03', '2025-09-07 14:00:00+03'),
    (23, 'delivered',   1798.00, '2025-09-10 14:30:00+03', '2025-09-11 10:00:00+03', '2025-09-14 11:00:00+03'),
    (8,  'delivered',   8999.00, '2025-09-18 09:15:00+03', '2025-09-20 08:00:00+03', '2025-09-23 16:00:00+03'),

    -- 2025 Q4
    (24, 'delivered',  54999.00, '2025-10-01 10:00:00+03', '2025-10-03 09:00:00+03', '2025-10-06 13:00:00+03'),
    (1,  'delivered',   6498.00, '2025-10-10 15:30:00+03', '2025-10-12 08:00:00+03', '2025-10-15 14:00:00+03'),
    (9,  'delivered',  18999.99, '2025-10-20 11:45:00+03', '2025-10-22 10:00:00+03', '2025-10-25 15:00:00+03'),
    (25, 'delivered',   2199.00, '2025-10-28 14:00:00+03', '2025-10-29 09:00:00+03', '2025-10-31 12:00:00+03'),
    (10, 'delivered',  12999.00, '2025-11-05 10:20:00+03', '2025-11-07 08:00:00+03', '2025-11-10 16:00:00+03'),
    (2,  'delivered',   3999.00, '2025-11-11 13:00:00+03', '2025-11-13 10:00:00+03', '2025-11-16 14:00:00+03'),
    (11, 'delivered',   8999.00, '2025-11-18 09:30:00+03', '2025-11-20 09:00:00+03', '2025-11-23 11:00:00+03'),
    (3,  'delivered',  34999.00, '2025-11-25 16:15:00+03', '2025-11-27 08:00:00+03', '2025-11-30 15:00:00+03'),
    (12, 'delivered',   1299.00, '2025-12-01 10:00:00+03', '2025-12-02 09:00:00+03', '2025-12-05 13:00:00+03'),
    (13, 'delivered',   4999.00, '2025-12-08 14:45:00+03', '2025-12-10 10:00:00+03', '2025-12-13 16:00:00+03'),
    (14, 'delivered',  22499.00, '2025-12-15 11:00:00+03', '2025-12-17 08:00:00+03', '2025-12-20 14:00:00+03'),
    (4,  'delivered',   7598.00, '2025-12-22 15:30:00+03', '2025-12-24 09:00:00+03', '2025-12-27 12:00:00+03'),

    -- 2026 Q1
    (15, 'delivered',  18999.99, '2026-01-05 10:00:00+03', '2026-01-07 09:00:00+03', '2026-01-10 14:00:00+03'),
    (5,  'delivered',   5999.00, '2026-01-12 14:30:00+03', '2026-01-14 08:00:00+03', '2026-01-17 15:00:00+03'),
    (16, 'delivered',   3499.00, '2026-01-20 09:15:00+03', '2026-01-21 10:00:00+03', '2026-01-24 11:00:00+03'),
    (6,  'delivered',  54999.00, '2026-02-01 13:00:00+03', '2026-02-03 09:00:00+03', '2026-02-06 16:00:00+03'),
    (17, 'delivered',   2498.00, '2026-02-10 10:45:00+03', '2026-02-11 08:00:00+03', '2026-02-14 13:00:00+03'),
    (1,  'shipped',   12999.00, '2026-02-20 15:00:00+03', '2026-02-22 10:00:00+03', NULL),
    (18, 'shipped',    1998.00, '2026-03-01 11:30:00+03', '2026-03-03 09:00:00+03', NULL),
    (8,  'processing', 4999.00, '2026-03-15 14:00:00+03', NULL,                      NULL),
    (19, 'processing', 8999.00, '2026-03-20 10:20:00+03', NULL,                      NULL),
    (20, 'pending',   34999.00, '2026-04-01 09:00:00+03', NULL,                      NULL),
    (3,  'pending',    1649.00, '2026-04-10 16:30:00+03', NULL,                      NULL),
    (21, 'cancelled',  5999.00, '2026-04-15 11:00:00+03', NULL,                      NULL);


-- ---- ORDER ITEMS ----------------------------------------------------------
-- Each order gets 1-3 line items. product_id and unit_price match the
-- products table. Discounts are applied on select items.
INSERT INTO order_items (order_id, product_id, quantity, unit_price, discount_pct) VALUES
    -- Order 1
    (1,  1,  1, 18999.99, 0),
    (1,  9,  1,   599.00, 0),
    -- Order 2
    (2,  7,  1,  3499.00, 0),
    -- Order 3
    (3,  4,  1, 34999.00, 0),
    (3, 11,  1,   749.00, 0),
    -- Order 4
    (4, 12,  1,  1899.00, 5),
    (4, 28,  1,   399.00, 0),
    -- Order 5
    (5,  5,  1, 54999.00, 0),
    -- Order 6
    (6,  8,  1,  1999.00, 0),
    (6, 13,  1,   999.00, 0),
    -- Order 7
    (7,  3,  1,  8999.00, 0),
    -- Order 8
    (8, 18,  1,  4999.00, 0),
    (8, 19,  1,  7999.00, 0),
    -- Order 9
    (9, 21,  1,   189.00, 0),
    (9, 24,  1,   249.00, 0),
    -- Order 10
    (10, 18, 1,  4999.00, 0),
    -- Order 11
    (11,  2, 1, 22499.00, 0),
    -- Order 12
    (12,  6, 1,  5999.00, 0),
    -- Order 13
    (13,  8, 1,  1999.00, 0),
    -- Order 14
    (14, 19, 1,  7999.00, 0),
    -- Order 15
    (15, 25, 1,   199.00, 0),
    (15, 26, 1,   329.00, 0),
    -- Order 16
    (16,  4, 1, 34999.00, 0),
    -- Order 17
    (17, 27, 1,  3999.00, 0),
    -- Order 18
    (18,  9, 1,   599.00, 0),
    (18, 14, 1,  1299.00, 0),
    -- Order 19
    (19, 30, 1, 12999.00, 0),
    -- Order 20
    (20,  8, 1,  1999.00, 0),
    (20, 28, 1,   399.00, 0),
    (20, 29, 1,   299.00, 0),
    -- Order 21
    (21,  1, 1, 18999.99, 0),
    -- Order 22
    (22, 28, 1,   399.00, 0),
    (22, 29, 1,   299.00, 0),
    (22, 31, 1,   179.00, 0),
    -- Order 23
    (23, 20, 1,  3299.00, 0),
    -- Order 24
    (24, 12, 1,  1899.00, 0),
    -- Order 25
    (25,  6, 1,  5999.00, 0),
    -- Order 26
    (26,  2, 1, 22499.00, 0),
    -- Order 27
    (27, 22, 1,    89.00, 0),
    (27, 23, 1,    69.00, 0),
    (27, 14, 1,  1299.00, 0),
    -- Order 28
    (28,  4, 1, 34999.00, 0),
    -- Order 29
    (29, 18, 1,  4999.00, 0),
    -- Order 30
    (30,  8, 1,  1999.00, 0),
    -- Order 31
    (31,  3, 1,  8999.00, 0),
    -- Order 32
    (32,  5, 1, 54999.00, 0),
    -- Order 33
    (33,  1, 1, 18999.99, 0),
    (33,  7, 1,  3499.00, 0),
    (33, 34, 1,   149.00, 0),
    -- Order 34
    (34,  1, 1, 18999.99, 0),
    -- Order 35
    (35, 10, 1,  2199.00, 0),
    -- Order 36
    (36, 30, 1, 12999.00, 0),
    -- Order 37
    (37,  2, 1, 22499.00, 5),
    (37, 37, 1,   449.00, 0),
    -- Order 38
    (38,  3, 1,  8999.00, 0),
    -- Order 39
    (39,  4, 1, 34999.00, 0),
    -- Order 40
    (40, 14, 1,  1299.00, 0),
    -- Order 41
    (41, 18, 1,  4999.00, 0),
    -- Order 42
    (42,  2, 1, 22499.00, 0),
    -- Order 43
    (43, 15, 1,   899.00, 0),
    (43, 16, 1,  2499.00, 0),
    (43, 17, 1,   649.00, 0),
    -- Order 44
    (44,  1, 1, 18999.99, 0),
    -- Order 45
    (45,  6, 1,  5999.00, 0),
    -- Order 46
    (46,  7, 1,  3499.00, 0),
    -- Order 47
    (47,  5, 1, 54999.00, 0),
    -- Order 48
    (48,  8, 1,  1999.00, 0),
    (48, 29, 1,   299.00, 0),
    -- Order 49
    (49, 30, 1, 12999.00, 0),
    -- Order 50
    (50, 18, 1,  4999.00, 0),
    -- Order 51
    (51,  3, 1,  8999.00, 0),
    -- Order 52
    (52,  4, 1, 34999.00, 0),
    -- Order 53
    (53, 11, 1,   749.00, 0),
    (53, 14, 1,  1299.00, 0),
    -- Order 54
    (54,  6, 1,  5999.00, 0),
    -- Order 55
    (55,  6, 1,  5999.00, 0);


-- ---- REVIEWS --------------------------------------------------------------
INSERT INTO reviews (product_id, customer_id, rating, title, body, created_at) VALUES
    (1,  1,  5, 'Excellent performance',      'Handles multitasking with ease. Battery life is impressive.',        '2025-01-15 10:00:00+03'),
    (1,  9,  4, 'Great value laptop',          'Very good for the price. Keyboard could be better.',                '2025-11-01 14:00:00+03'),
    (2,  6,  5, 'Super lightweight',           'Perfect for travel. Display quality is stunning.',                   '2025-08-20 11:00:00+03'),
    (3,  7,  3, 'Decent for basic tasks',      'Gets the job done but slows down with many tabs.',                  '2025-03-10 09:00:00+03'),
    (4,  3,  5, 'Best phone I have owned',     'Camera is incredible. Battery lasts all day.',                       '2025-02-01 16:00:00+03'),
    (4, 22,  4, 'Great phone, minor issues',   'Love the display but gets warm during gaming.',                     '2025-09-05 10:00:00+03'),
    (5,  5,  5, 'Worth every lira',            'Ecosystem integration is seamless.',                                 '2025-03-01 13:00:00+03'),
    (5, 24,  5, 'Premium experience',          'Best smartphone camera on the market.',                              '2025-10-15 15:00:00+03'),
    (6, 11,  4, 'Good budget option',          'Surprisingly capable for everyday use.',                             '2025-04-20 11:00:00+03'),
    (7,  2,  5, 'Silence is golden',           'ANC is top tier. Comfortable for long sessions.',                    '2025-01-20 14:00:00+03'),
    (8, 13,  4, 'Compact and powerful',        'Sound quality exceeds expectations for the size.',                   '2025-04-25 10:00:00+03'),
    (9,  1,  4, 'Nice fit',                    'Quality cotton, fits true to size.',                                  '2025-01-18 09:00:00+03'),
    (12, 4,  5, 'Luxuriously soft',            'The cashmere blend is wonderful. Runs slightly large.',              '2025-02-10 16:00:00+03'),
    (18, 10, 5, 'Back pain gone',              'Best investment for my home office. Lumbar support is perfect.',     '2025-04-05 11:00:00+03'),
    (18, 8,  4, 'Very comfortable',            'Great chair but assembly instructions could be clearer.',           '2025-03-20 14:00:00+03'),
    (19, 14, 5, 'Transformed my workspace',    'Electric adjustment is smooth. Memory presets are convenient.',      '2025-05-15 10:00:00+03'),
    (20, 23, 4, 'Solid construction',          'Beautiful oak finish. Assembly required patience.',                  '2025-03-28 15:00:00+03'),
    (21, 9,  5, 'Must read for SQL learners',  'Clear explanations and practical examples.',                         '2025-03-25 09:00:00+03'),
    (25, 15, 4, 'Excellent reference',         'Well organized, some chapters could be deeper.',                     '2025-05-30 11:00:00+03'),
    (27, 17, 5, 'Space-saving genius',         'Quick weight change. Replaced my full dumbbell rack.',              '2025-07-01 14:00:00+03'),
    (28, 22, 4, 'Good quality mat',            'Non-slip as advertised. Could be slightly thicker.',                '2025-07-15 10:00:00+03'),
    (30, 19, 3, 'Decent treadmill',            'Works well but the motor is a bit loud.',                            '2025-08-01 16:00:00+03'),
    (15, 8,  5, 'Perfect sear every time',     'Heavy and retains heat beautifully.',                                '2025-03-15 09:00:00+03'),
    (16, 10, 4, 'Great pot set',               'Heats evenly on induction. Handles stay cool.',                     '2025-04-10 14:00:00+03');

COMMIT;
