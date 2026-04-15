# UN Financial Tracking & Reporting System

![Excel](https://img.shields.io/badge/MS_Excel-2016+-green?logo=microsoft-excel)
![VBA](https://img.shields.io/badge/VBA-Automated-blue)
![Power Query](https://img.shields.io/badge/Power_Query-Integrated-yellow)
![Compliance](https://img.shields.io/badge/Compliance-UN_EU_SECO_Ready-red)

## Overview

A **Database-First Financial Management System** built in Microsoft Excel for UN agencies, NGOs, and development programmes. This system manages the full lifecycle of funds from diverse sources (Regular Budget, Voluntary Contributions, Bilateral Donors like Sida, EU, ECHO) and ensures compliance with **UN-Nutrition and EU/SECO reporting standards**.

**Key Capability:** Tracks Revenue → Allocations → Expenditures in a relational data model with a dynamic "Reserve Engine" that provides real-time liquidity alerts.

![Dashboard Overview](screenshot-system.png)

---

## Core Functionality

| Module | Description | Compliance Feature |
| :--- | :--- | :--- |
| **Revenue/Receipt Log** | Tracks donor contributions, currencies, and UN exchange rates | Automated FX conversion to USD |
| **Allocation/Project Log** | Links funds to specific projects and thematic pillars | Earmarking enforcement |
| **Expenditure Log** | Manages commitments (legal obligations) vs. actual disbursements | IPSAS Cash/Accrual readiness |
| **Reserve Engine** | `Revenue - (Allocations + Expenditures) = Reserve` | **Traffic Light Alert** for low liquidity |
| **Audit Trail** | Hidden sheet logging all data modifications | User stamp + Timestamp |
| **Donor Reporting** | One-click template for EU/SECO narrative and financial reports | Aggregates unspent balances |

---

## System Architecture

This system adheres to the **Database-First** approach, avoiding the pitfalls of "spreadsheet sprawl."
