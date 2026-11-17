# QUICK START CHECKLIST

## ✅ Phase 1 Complete
- [x] Data model finalized (AS-IS/TO-BE separation)
- [x] PowerPoint template optimized (48 named fields)
- [x] Excel master workbook created (56 fields, 5 sheets)
- [x] Field mappings documented (Field map sheet + DATA_MODEL.md)
- [x] Sizing configured as flexible lookup (SizeBands)
- [x] Documentation complete (3 user guides)

## 📋 To Do Before Phase 2

### Immediate (Today/This Week)
- [ ] Open `product_team_master.xlsx` in Excel
- [ ] Add 5 formulas to calculated columns (see WORKBOOK_GUIDE.md)
  - [ ] Teams X2: Headcount minimum
  - [ ] Teams Y2: Headcount maximum
  - [ ] Products AB2: TO-BE team count
  - [ ] Products AC2: TO-BE headcount min
  - [ ] Products AD2: TO-BE headcount max
- [ ] Test formulas with sample data
  - [ ] Add sizes to Setup sheet (S=3–5, M=6–9, L=10–15, XL=16–25)
  - [ ] Create test team, verify headcount lookups work
  - [ ] Create test product, verify team count and totals roll up
- [ ] Clear test data before final submission
- [ ] Save and close workbook

### Before Building Sync Engine
- [ ] Review Field map sheet – confirm 56 fields are correct
- [ ] Add/remove any fields needed for your use case
- [ ] Adjust field names or types if needed
- [ ] Verify PPT and Excel field names match (use DATA_MODEL.md as reference)
- [ ] Get stakeholder sign-off on data model

### Before Phase 2 (Sync Engine)
- [ ] Populate workbook with real sample data (at least 3 products, 5 teams)
- [ ] Test formulas with real data
- [ ] Confirm calculated totals make sense
- [ ] Document any additional fields or validations needed
- [ ] Plan reporting requirements

---

## 📂 File Locations

```
C:\Users\GESZA4\OneDrive - IKEA\Apps\ppt to excel\

├── product_team_master.xlsx          ← MASTER WORKBOOK (add formulas here)
├── product_team_optimized.pptx       ← ALIGNED TEMPLATE
├── WORKBOOK_GUIDE.md                 ← USER INSTRUCTIONS (read first)
├── DATA_MODEL.md                     ← TECHNICAL REFERENCE
├── PHASE_1_SUMMARY.md                ← QUICK START GUIDE
├── QUICK_START_CHECKLIST.md          ← THIS FILE
├── create-master-workbook.ps1        ← Regeneration script (if needed)
└── add-formulas-guide.ps1            ← Formula reference (can delete)
```

---

## 🎯 Key Reminders

### Data Model Concepts
- **AS-IS** = current state (product level, manual input)
- **TO-BE** = future design (team level, can be calculated)
- **Not one-to-one**: AS-IS and TO-BE are **independent** perspectives
- **SizeBands**: Flexible, configurable, expandable

### Excel Conventions
- Leave **ID fields blank** (engine assigns)
- Use **decimal for FTE** (2.5, not 2.5%)
- **Blank = unknown**, not 0
- **Sizing reasons & Topology** are sets of independent checkboxes
- **Sibling team** is narrative only; doesn't affect counting

### PPT Conventions
- 48 **named answer fields** (use these for syncing)
- **Product slide** captures AS-IS (team count, FTE)
- **Team slides** capture TO-BE (size, topology, etc.)
- Speaker notes document field mappings

---

## 🚀 Phase 2 Roadmap

Once Phase 1 is complete and formulas are tested:

### Week 1: Build Inspector
```
Read PPT → Extract all 48 answer fields
Read Excel → Extract all data from Products & Teams tables
Compare field-by-field, highlight diffs
Surface missing values and conflicts
```

### Week 2: Build Change Proposals
```
Excel → PPT: Detect new/modified/deleted Excel rows
           Propose corresponding PPT slide updates
           
PPT → Excel: Extract PPT changes (new slides, edits)
           Propose Excel row updates
           Flag conflicts if both files changed same field
```

### Week 3: Build Publish & Merge
```
Approve proposals
Write changes to target file
Handle duplicates and deletions
Maintain audit trail
```

### Week 4: Reporting & Polish
```
Build dashboard (headcount by type, AS-IS vs TO-BE, etc.)
Validation report (missing required fields)
Edge case handling
Performance optimization
```

---

## 💬 Questions?

| Question | Answer Location |
|----------|-----------------|
| How do I use the workbook? | WORKBOOK_GUIDE.md |
| What do all 56 fields mean? | DATA_MODEL.md |
| How do I add formulas? | WORKBOOK_GUIDE.md, section "Next Steps" |
| What's the data model philosophy? | DATA_MODEL.md, section "Key Design Decisions" |
| What are the field conventions? | WORKBOOK_GUIDE.md, section "Conventions & Tips" |
| How does the sync work? | DATA_MODEL.md, section "Sync Workflow" |

---

## ✨ Next Action

👉 **Open WORKBOOK_GUIDE.md and add the 5 formulas.**

That's it! Then test with sample data and you're ready for Phase 2.

Good luck! 🎉
