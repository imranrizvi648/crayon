# Crayon Costing App - Checkpoint
## Date: January 30, 2026

---

## ✅ PROJECT STATUS: ALL CALCULATIONS VERIFIED

All calculations for both regions (Middle East & Africa) and both deal types (Normal & Ramped) are now **matching Excel exactly**.

---

## Features Implemented

### 1. Core Costing Sheet Functionality
- ✅ Customer & Agreement Details (all header fields)
- ✅ Line Items with all columns matching Excel
- ✅ Excel paste functionality (copy from Excel, paste into Part Number field)
- ✅ Category detection (Enterprise Online, Additional, On Premise)
- ✅ Auto-calculation of all derived fields

### 2. Regions Supported
- ✅ **Middle East (ME)** - FEWA Costing Sheet
- ✅ **Africa (AF)** - Africa Costing Sheet with GP Split

### 3. Deal Types Supported
- ✅ **Normal** - Single year data × 3 for totals
- ✅ **Ramped** - Separate Year 1, Year 2, Year 3 data with different values

### 4. Calculation Formulas (Verified)

#### MS Discounted Values
```
MS Disc Net = ROUND(Unit Net USD × (1 - MS Discount%) × Exchange Rate, 2)
MS Disc ERP = ROUND(Unit ERP USD × (1 - MS Discount%) × Exchange Rate, 2)
```

#### EUP (End User Price) - DIFFERENT FORMULAS BY DEAL TYPE
```
Normal:  EUP = ROUND(MS Disc Net × (1 + Crayon Markup%), 2)  [MULTIPLICATION]
Ramped:  EUP = ROUND(MS Disc Net / (1 - Crayon Markup%), 2)  [DIVISION]
```

#### Special Rule: Zero Default Markup Products
```
If Unit Net USD = Unit ERP USD (Default Markup = 0%):
  → EUP = MS Disc ERP (ignore Crayon Markup)
```

#### Totals
```
Total Net = ROUND(MS Disc Net × Unit Type × Quantity, 2)
Total ERP = ROUND(MS Disc ERP × Unit Type × Quantity, 2)
Total EUP = ROUND(EUP × Quantity × Unit Type, 2)
```

#### GP Calculations (Africa Only)
```
GP = Total EUP - Total Net
SWO GP = GP × SWO GP %
Partner GP = GP - SWO GP
```

#### Rebate
```
Rebate = Total Net × Rebate %
```

### 5. Summary Calculations

#### Normal Deal
- Year 1, 2, 3 values are identical (Y1 × 3 for totals)

#### Ramped Deal
- Year 1, 2, 3 have independent values
- Totals = Actual Y1 + Y2 + Y3

### 6. UI Features

#### Costing Form Tab
- Customer & Agreement Details section
- Deal Type dropdown (Normal/Ramped)
- Line Items table with all columns
- Year tabs for Ramped (Year 1, Year 2, Year 3)
- Copy buttons: "Copy to Year 2", "Copy to Year 3", "Copy to All Years"
- Profit Summary Box (individual years for Ramped)
- Crayon Discount/Funding section
- Bid Bond & Bank Charges section
- Other LSP Rebate section
- CIF Products section

#### Merged Tab
- Customer & Agreement Summary
- Key Metrics cards
- Cost Price / CPS Price
- Estimated Retail Price
- End User Price sections
- GP without Rebates (with Crayon/Partner split for Africa)
- Gross Profit with Rebates
- Year comparison table (Ramped shows actual values)

#### Final Price Table Tab
- **Normal**: Single table with Yr.1, Yr.2, Yr.3 columns
- **Ramped**: 3 separate year tables + Grand Total summary
- Category grouping (Enterprise Online, Additional, On Premise)
- Discount, VAT, Grand Total calculations

### 7. Excel Export
- Full costing sheet export
- Merged data export
- Final Price Table export
- All formulas preserved

### 8. Default Values Pre-populated
- Exchange Rate: 3.6725 (AED)
- VAT Rate: 5%
- Currency: AED
- Agreement Levels: D (System, Server, Application)
- Deal Type: Normal
- SWO GP %: 50% (Africa)

---

## Files Structure

```
crayon-costing-app/
├── index.html
├── package.json
├── vite.config.js
├── tailwind.config.js
├── postcss.config.js
├── README.md
├── CHECKPOINT_2026-01-30.md
└── src/
    ├── main.jsx
    ├── index.css
    ├── App.jsx          (Full version with Excel export)
    └── AppPreview.jsx   (Preview version for Claude artifacts)
```

---

## Key Technical Notes

1. **Rounding**: Excel rounds at EACH intermediate step (cell-by-cell). Our code mirrors this exactly.

2. **EUP Formula Difference**: 
   - Normal uses multiplication: `× (1 + markup)`
   - Ramped uses division: `/ (1 - markup)`
   - This is a critical business rule from the Excel templates

3. **Profit Summary Box**:
   - Normal: Shows single "Profit /Year" row
   - Ramped: Shows individual Year 1, Year 2, Year 3 rows

4. **State Management**:
   - `lineItems` - Year 1 data (always used)
   - `lineItemsY2` - Year 2 data (Ramped only)
   - `lineItemsY3` - Year 3 data (Ramped only)
   - `activeYear` - Current year tab (1, 2, or 3)

---

## Verified Against Excel Templates

- ✅ Middle_East_Costing_Sheet_FEWA.xlsx (Normal)
- ✅ Africa_Costing_Sheet.xlsx (Normal)
- ✅ Ramped_Middle_East_Costing_Sheet__FEWA_.xlsx (Ramped)
- ✅ Ramped_Africa_Costing_Sheet.xlsx (Ramped)

---

## Next Steps (Future Development)

1. Backend API integration (FastAPI/Django)
2. Database schema (PostgreSQL)
3. User authentication & RBAC
4. Approval workflow
5. Audit logging
6. Power BI integration
7. PDF export

---

## Transcript Reference

Previous conversation transcripts available in:
- `/mnt/transcripts/2026-01-23-10-35-28-fewa-costing-digitization-analysis.txt`
- `/mnt/transcripts/2026-01-30-10-46-19-ramped-deal-implementation-start.txt`
