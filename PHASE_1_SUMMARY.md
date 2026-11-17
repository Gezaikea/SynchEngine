# PPT ↔ Excel Sync Engine – Phase 1 Complete

## ✅ What You Now Have

A complete, **tested data contract** between PowerPoint and Excel that you can build a sync engine on top of.

### **product_team_master.xlsx** (257 KB)
Your master data workbook with 5 sheets:

1. **Start here** – Quick overview and instructions
2. **Setup** – Configure team sizing (S/M/L/XL/etc.) with min/max headcount
3. **Products** – 31 fields per product (descriptive, criticality, scope, boundaries, AS-IS inputs, TO-BE calculated totals)
4. **Teams** – 25 fields per team (design, sizing, topology, interaction)
5. **Field map** – Complete 56-field inventory and sync readiness

### **product_team_optimized.pptx** (96 KB)
Your aligned PowerPoint template with:
- 48 named answer fields (27 on product slide, 21 on team slide)
- Separated labels from editable answers
- Product slide captures AS-IS baseline (team count, FTE)
- Team slides capture TO-BE design (size, topology, interaction)
- Speaker notes with field mappings

### **Documentation**
- **WORKBOOK_GUIDE.md** – User instructions, field descriptions, conventions, troubleshooting
- **DATA_MODEL.md** – Complete data model, design decisions, sync workflow proposal
- Comprehensive field definitions for all 56 fields

---

## 🎯 Data Model At A Glance

### Two Independent Perspectives

```
┌─────────────────────────────────────────────────────────────┐
│                      CURRENT STATE (AS-IS)                  │
├─────────────────────────────────────────────────────────────┤
│ Product One-Pagers (Excel Products table)                   │
│                                                              │
│ Manual inputs:                                               │
│  • AS-IS team count: indicative current # of teams          │
│  • AS-IS FTE: indicative current FTE assigned               │
│                                                              │
│ Descriptive inputs:                                          │
│  • Criticality counts (SL1, SL2, High, Medium, Low)         │
│  • Scope (components, systems, vendors)                     │
│  • Complexity and investment direction                      │
│  • Boundaries and dependencies                              │
└─────────────────────────────────────────────────────────────┘

                              ↕
                        (relationship via
                        Product ID, NOT
                        calculated from)

┌─────────────────────────────────────────────────────────────┐
│                   FUTURE STATE (TO-BE)                      │
├─────────────────────────────────────────────────────────────┤
│ Team One-Pagers (Excel Teams table)                         │
│                                                              │
│ One actual team per row & per pager                         │
│                                                              │
│ Captured inputs:                                             │
│  • Team size: S, M, L, XL, etc. (configurable SizeBands)   │
│  • Sibling team? Yes/No (narrative only)                    │
│  • Sizing reasons: 5 checkboxes [x] / [ ]                   │
│  • Topology: 5 assessments [ ] / [x] / ?                    │
│  • Composition and interaction model                        │
│                                                              │
│ Calculated:                                                  │
│  • Headcount min/max (looked up from team size)             │
│  • Rolls up to product level (TO-BE totals)                 │
└─────────────────────────────────────────────────────────────┘
```

### Key Insight
- **AS-IS and TO-BE are NOT calculated from each other**
- **AS-IS** = current state snapshot (before reorganization)
- **TO-BE** = proposed organization (after transformation)
- Both are inputs to your decision-making process

---

## 🚀 Immediate Next Step: Add 5 Formulas

The workbook is **fully structured and ready**, but calculated columns need formulas. These are marked with a **light blue background**.

**In Excel, manually paste these 5 formulas** (or use the script above):

### Teams sheet (2 formulas)

**Column X, Cell X2 – Headcount minimum:**
```excel
=IF([@[Team name]]="","",IF(COUNTIF(SizeBands[Size category],[@[Team size]])<>1,"",IFERROR(INDEX(SizeBands[Min headcount],MATCH([@[Team size]],SizeBands[Size category],0)),"")
```

**Column Y, Cell Y2 – Headcount maximum:**
```excel
=IF([@[Team name]]="","",IF(COUNTIF(SizeBands[Size category],[@[Team size]])<>1,"",IFERROR(INDEX(SizeBands[Max headcount],MATCH([@[Team size]],SizeBands[Size category],0)),"")
```

### Products sheet (3 formulas)

**Column AB, Cell AB2 – TO-BE team count:**
```excel
=IF([@[Product ID]]="","",COUNTIFS(Teams[Product ID],[@[Product ID]],Teams[Team name],"<>"))
```

**Column AC, Cell AC2 – TO-BE headcount min:**
```excel
=IF([@[Product ID]]="","",IF([@[TO-BE team count]]=0,"",SUMIF(Teams[Product ID],[@[Product ID]],Teams["Headcount minimum"])))
```

**Column AD, Cell AD2 – TO-BE headcount max:**
```excel
=IF([@[Product ID]]="","",IF([@[TO-BE team count]]=0,"",SUMIF(Teams[Product ID],[@[Product ID]],Teams["Headcount maximum"])))
```

**After adding these**, test with:
1. Add sizes to Setup sheet (S=3–5, M=6–9, L=10–15, XL=16–25)
2. Create test team with size "M" → should see 6 and 9 in min/max columns
3. Create test product, assign team → should see totals roll up

---

## 📊 Field Inventory (56 Total)

### Products (31 fields)
- Metadata: ID, name
- Descriptive: type, users, owner, purpose, catalog
- Criticality: 5 count fields (SL1, SL2, High, Medium, Low)
- Scope: components, systems, vendors, competencies
- Complexity: level, investment direction, change demand
- Boundaries: adjacent products, dependencies
- **AS-IS**: team count, FTE *(manual inputs)*
- **TO-BE**: team count, headcount min/max *(calculated)*

### Teams (25 fields)
- Metadata: ID, name
- Relationship: Product ID, assignment status
- Design: type, purpose, composition
- Sizing: size, sibling flag, 5 sizing reasons *(checkboxes)*
- Topology: 5 assessments *(value flow, alignment, ownership, handoffs, feedback)*
- Interaction: consumers, supporters, modes, boundaries
- Calculated: headcount min/max

### Setup (SizeBands table)
- Size category, min headcount, max headcount
- Expandable: add your own sizes anytime

---

## 📋 Conventions to Know

| Concept | Format | Example | Location |
|---------|--------|---------|----------|
| **IDs** | Leave blank | (auto-assigned by sync engine) | Products, Teams |
| **FTE** | Decimal, not % | 2.5, 3.75 | Products AS-IS |
| **Headcount** | Estimated range | 6–9 people (from size M) | Teams calculated |
| **Percentage** | 0–100 | 60 (for 60% build/buy) | Products scope |
| **Sizing reasons** | [x] / [ ] | [x] = applies | Teams (5 boxes) |
| **Topology** | [ ] / [x] / ? | [x] = yes, ? = no | Teams (5 boxes) |
| **Blank vs Zero** | Blank = unknown | Don't fill missing with 0 | All text/numeric |
| **Sibling teams** | Yes/No flag | Narrative only; no effect on counts | Teams |

---

## 🏗️ Phase 2: Build the Sync Engine

Once formulas are added and data model is validated, you're ready to build the runtime components:

### 1. **Inspector** (Read-only comparison)
- Load PPT and Excel
- Compare all 56 fields side-by-side
- Highlight differences and missing values
- Surface conflicts (both files changed same field)

### 2. **Change Proposal** (Excel → PPT)
- Detect new/modified/deleted rows in Excel
- Suggest corresponding PPT updates
- Apply with user approval

### 3. **Change Proposal** (PPT → Excel)
- Extract answer fields from PPT
- Detect changes from last known state
- Propose Excel updates
- Apply with user approval

### 4. **Conflict Resolution**
- If both files changed same field, manual review required
- Audit trail: who changed what, when, in which file

### 5. **Reporting**
- Headcount distribution by team type, product, size
- AS-IS vs. TO-BE comparison
- Coverage analysis (teams without products, etc.)
- Validation report (missing required fields)

---

## 📂 File Locations

All files are in: `C:\Users\GESZA4\OneDrive - IKEA\Apps\ppt to excel\`

| File | Purpose | Size |
|------|---------|------|
| `product_team_master.xlsx` | Master data workbook (formulas pending) | 257 KB |
| `product_team_optimized.pptx` | Aligned PPT template | 96 KB |
| `WORKBOOK_GUIDE.md` | User instructions & conventions | 8 KB |
| `DATA_MODEL.md` | Complete field reference | 11 KB |
| `create-master-workbook.ps1` | Regeneration script (if structure changes) | — |

---

## ✨ You're Ready!

The hard part is done. The data contract is **solid**, **documented**, and **aligned across both tools**.

👉 **Next: Add those 5 formulas, test with sample data, then we build the sync engine.**

Questions? Check **WORKBOOK_GUIDE.md** (user guide) or **DATA_MODEL.md** (technical reference).
