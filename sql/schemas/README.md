# E-Commerce Database Schema

## Overview

This schema models a complete e-commerce platform with support for customer management, product cataloging, order processing, reviews, and inventory tracking.

## Design Principles

- **Normalization**: Tables are normalized to 3NF to eliminate data redundancy.
- **Referential Integrity**: All relationships are enforced via foreign keys with appropriate ON DELETE/ON UPDATE actions.
- **Data Validation**: CHECK constraints enforce business rules at the database level (e.g., prices must be positive, ratings between 1-5, email format validation).
- **Performance**: Indexes are placed on columns frequently used in JOINs, WHERE clauses, and ORDER BY operations.
- **Audit Trail**: Timestamps (`created_at`, `updated_at`) on every table for traceability.

## Entity Relationships

```
categories (1) --- (N) products (1) --- (1) inventory
                        |
customers (1) --- (N) orders (1) --- (N) order_items (N) --- (1) products
    |                                                              |
    +----------------------- reviews (N) -------------------------+
```

## Files

- `ecommerce_schema.sql` -- Full DDL with constraints, indexes, and realistic sample data.
