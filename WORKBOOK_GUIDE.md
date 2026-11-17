# Excel Master Workbook – Complete Guide

## Overview
You now have **product_team_master.xlsx**, the master data structure for your PPT ↔ Excel sync engine. This workbook reflects the AS-IS/TO-BE organizational model you defined:

- **AS-IS**: Current organizational baseline on **product one-pagers** (manual entries: team count, FTE)
- **TO-BE**: Future team design on **team one-pagers** (actual teams, linked to products, with sizing and topology)

## Workbook Structure

### Sheet 1: Start here
Instructions, framework overview, and links to the Setup sheet.

### Sheet 2: Setup
**Configure your size bands here.** This is the master list of team-sizing categories.

| Column | Purpose |
|--------|---------|
| Size category | S, M, L, XL, XXL, etc. (you decide) |
| Min headcount | Minimum estimated people for this size |
| Max headcount | Maximum estimated people for this size |

**Example:**
- S: 3–5
- M: 6–9
- L: 10–15
- XL: 16–25

Add or change sizes as needed. All team sizing lookups reference this table.

### Sheet 3: Products
**One row per product.** Columns are organized by section (top strip, purpose, criticality, scope, complexity, boundaries, AS-IS, TO-BE).

#### AS-IS Section (Manual inputs)
- **AS-IS team count**: Indicative current number of teams assigned to this product
- **AS-IS FTE**: Indicative current FTE (enter as decimal: 1, 2.5, 3.75, etc.)

#### TO-BE Section (Calculated)
- **TO-BE team count**: Automatically counts linked team records with this Product ID
- **TO-BE headcount min/max**: Rolls up size-band estimates from linked teams

**All other columns** are inputs you fill as you define the product (purpose, criticality counts, complexity, vendors, etc.). See the Field map sheet for complete inventory.

### Sheet 4: Teams
**One row per actual team.** Each team can be linked to one product (TO-BE design).

#### Key Columns
- **Team ID**: Assigned by sync engine (leave blank)
- **Product ID**: The product this team belongs to (matches Products.Product ID)
- **Team name**: The team's name
- **Team size**: Dropdown referencing SizeBands (S, M, L, XL, etc.)
- **Sibling team?**: Yes/No flag; separate one-pager per actual team (flag is narrative)
- **Sizing reasons**: Five `[x]` / `[ ]` checkboxes (business criticality, service level, technical complexity, sourcing, investment)
- **Topology**: Five `[ ]` / `[x]` / `?` assessments (value flow, stream alignment, outcome ownership, handoffs/load, feedback)

#### Calculated Columns (light blue)
- **Headcount minimum/maximum**: Looked up from SizeBands based on team size

### Sheet 5: Field map
Complete inventory of all 50+ fields across Products and Teams. Includes:
- Field type (text, choice, integer, decimal, calculated, lookup, etc.)
- Expected values or format
- Mapping status (ready, pending, engine-managed, Excel-only)
- Notes

## Next Steps: Add Formulas

The calculated columns (marked light blue) need formulas to work. **In Excel, manually enter these five formulas:**

### Teams Sheet

**1. Headcount minimum** (column X, cell X2):
```
=IF([@[Team name]]="","",IF(COUNTIF(SizeBands[Size category],[@[Team size]])<>1,"",IFERROR(INDEX(SizeBands[Min headcount],MATCH([@[Team size]],SizeBands[Size category],0)),""))
```

**2. Headcount maximum** (column Y, cell Y2):
```
=IF([@[Team name]]="","",IF(COUNTIF(SizeBands[Size category],[@[Team size]])<>1,"",IFERROR(INDEX(SizeBands[Max headcount],MATCH([@[Team size]],SizeBands[Size category],0)),""))
```

### Products Sheet

**3. TO-BE team count** (column AB, cell AB2):
```
=IF([@[Product ID]]="","",COUNTIFS(Teams[Product ID],[@[Product ID]],Teams[Team name],"<>"))
```

**4. TO-BE headcount min** (column AC, cell AC2):
```
=IF([@[Product ID]]="","",IF([@[TO-BE team count]]=0,"",SUMIF(Teams[Product ID],[@[Product ID]],Teams["Headcount minimum"])))
```

**5. TO-BE headcount max** (column AD, cell AD2):
```
=IF([@[Product ID]]="","",IF([@[TO-BE team count]]=0,"",SUMIF(Teams[Product ID],[@[Product ID]],Teams["Headcount maximum"])))
```

## Test & Validate

1. **Configure sizes** on the Setup sheet (at least S, M, L).
2. **Create a test team**:
   - Team name: "Test Team"
   - Team size: "M" (or another configured size)
   - Confirm headcount min/max appear from the size band.
3. **Create a test product**:
   - Product name: "Test Product"
   - Assign the test team by entering the Product ID in the Teams sheet.
   - Confirm TO-BE counts roll up on the Products sheet.
4. **Delete test data** when done (or leave as reference).

## Conventions & Tips

### ID Fields
- **Product ID**, **Team ID**: Leave blank when creating. The sync engine will assign.
- In Teams sheet, **Product ID** must match a Product.Product ID to link correctly.

### AS-IS vs. TO-BE
- **AS-IS** is static input on the product record (represents current state before reorganization).
- **TO-BE** is calculated from team records (represents proposed future design).
- They are **not calculated from each other**.

### Sizing & Headcount Ranges
- **SizeBands** map category names (S, M, L, XL) to estimated headcount ranges.
- Ranges are **estimates** (3–5 people, not FTE hours).
- **FTE** is entered manually per product (AS-IS only), independent of headcount ranges.
- Do not conflate FTE with headcount ranges—they serve different purposes.

### Sizing Reasons & Topology
- **Sizing reasons**: `[x]` = applies, `[ ]` = does not apply. Leave blank for "not yet assessed."
- **Topology**: `[x]` = yes, `[ ]` = not assessed, `?` = no. Replace bracket symbol only.
- These are separate checkboxes, not a multi-select dropdown.

### Blank vs. Zero
- **Blank** = unknown or not yet entered.
- **Zero** = confirmed none (e.g., 0 FTE, 0 component count).
- Do not fill missing values with zeros—leave them blank.

### Decimal vs. Percentage
- **FTE**: Enter as decimal (1, 2.5, 3.75, etc.)—not percentage.
- **Make/Buy percent**: Enter as percent (0–100)—used for sourcing ratio.

### Flexibility
- **Sort and reorder rows/columns** freely; table structure and formulas are preserved.
- **Add columns** for additional inputs (the data contract is not locked).
- **Do not rename table names or headers** without updating field maps and formulas.

## Reporting & Next Steps

Once the workbook is validated:

1. **Read-only inspector** (PowerPoint reader):
   - Loads both PPT and Excel in read-only mode.
   - Surfaces field-by-field comparison.
   - Highlights discrepancies and missing values.

2. **PPT → Excel proposals**:
   - Detects changes in PPT (new slides, deleted content, edits).
   - Proposes Excel updates for review.

3. **Excel → PPT publisher**:
   - Writes approved changes back to PPT answer objects.
   - Regenerates one-pagers with updated data.

4. **Duplication & deletion handling**:
   - Marks duplicate teams/products.
   - Tracks deletion audit trail.

5. **Reporting dashboard**:
   - Headcount distribution by team type, product, sizing.
   - AS-IS vs. TO-BE comparison charts.
   - Coverage analysis (teams without products, etc.).

## Support & Troubleshooting

- **Formulas not calculating?** Ensure Setup sheet has at least one size band. Press Ctrl+Shift+F9 to recalculate all.
- **Dropdown not appearing?** Make sure the value is in the SizeBands table (column A).
- **Columns hidden?** Use Unhide All (right-click column header) if needed.
- **Validation errors?** Check the Field map sheet for expected format and type.

---

**Questions?** Refer to the Field map sheet for field-by-field definitions, or review the Notes column for guidance.
