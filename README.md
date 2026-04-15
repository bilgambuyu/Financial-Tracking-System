# UN Financial Tracking & Reporting System

![Excel](https://img.shields.io/badge/MS_Excel-2016+-green?logo=microsoft-excel)
![VBA](https://img.shields.io/badge/VBA-Automated-blue)
![Power Query](https://img.shields.io/badge/Power_Query-Integrated-yellow)
![Compliance](https://img.shields.io/badge/Compliance-UN_EU_SECO_Ready-red)

## Overview

A **Database-First Financial Management System** built in Microsoft Excel for UN agencies, NGOs, and development programmes. This system manages the full lifecycle of funds from diverse sources (Regular Budget, Voluntary Contributions, Bilateral Donors like Sida, EU, ECHO) and ensures compliance with **UN-Nutrition and EU/SECO reporting standards**.

**Key Capability:** Tracks Revenue → Allocations → Expenditures in a relational data model with a dynamic "Reserve Engine" that provides real-time liquidity alerts.

![Dashboard Overview](screenshot-system.gif)

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

---

## Dashboard & Reporting Features

### 1. Executive Dashboard
The main interface provides at-a-glance portfolio health.
- **KPI Cards:** Real-time aggregation of Total Revenue, Allocated, Spent, and Reserve.
- **Reserve Liquidity Alert:** Green (>15%), Yellow (5-15%), Red (<5%).
- **Allocation Map:** Treemap/Pie chart showing fund distribution across **Agribusiness, Climate, Nutrition, Gender, and Governance** pillars.

### 2. Donor Profile View
Select a single donor (e.g., "Sida" or "Irish Aid") from the Slicer. The entire dashboard updates to show:
- Total Lifetime Contribution
- Active Projects funded by this donor
- **Remaining Unspent Balance** (Crucial for grant close-out)

### 3. Audit & Compliance
- **Hidden Audit Log (`_Audit_Log`):** Tracks `User`, `Timestamp`, `Old_Value`, `New_Value`.
- **Automated EU/SECO Template:** Pulls financial data into a format ready for narrative reporting, including utilization rates (`Expenditure / Budget`).

---

## Technology Stack & Formulas

| Feature | Implementation |
| :--- | :--- |
| **Data Validation** | Dependent Named Ranges for Donors, Currencies, and UN Thematic Pillars |
| **Currency Conversion** | `VLOOKUP([@Currency], FX_Rates, 2, FALSE) * [@Amount_Original]` |
| **Reserve Engine** | `=SUM(Revenue[Amount_USD]) - (SUM(Allocation[Amount]) + SUM(Expenditure[Disbursed]))` |
| **Traffic Light Logic** | Conditional Formatting: `=Reserve < Revenue*0.05` (Red Alert) |
| **VBA Generator** | Single macro to rebuild the entire `.xlsm` file from scratch for demos |

---

## Quick Start

### Option 1: Run the Pre-Built System
1. Download `Financial Tracking System.xlsm`
2. **Enable Macros** when prompted (required for Audit Log and Slicer interactivity).
3. Navigate to the **`EXECUTIVE_DASHBOARD`** sheet.
4. Use the **Slicers** (Donor, Stream, Pillar) to filter data.

### Option 2: Build from Source
1. Open a blank Excel workbook.
2. Press `Alt + F11` to open VBA.
3. Import `vba-source/source-code.bas`.
4. Run the macro `Financial Tracking System`.
5. The entire system will self-populate with realistic sample data in <10 seconds.

---

## Author

**Mbuyu Bilga**
- LinkedIn: [Mbuyu Bilga](https://www.linkedin.com/in/mbuyu-bilga)
- Portfolio: [bilgambuyu.com](https://bilgambuyu.com)
