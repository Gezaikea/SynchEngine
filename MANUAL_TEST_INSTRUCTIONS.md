# MANUAL SAMPLE DATA TEST

## Quick Test: What to Enter Manually in Excel

The Excel COM approach is hitting timeout issues with the system. Instead, here's what you can manually enter to test the system in just 2 minutes:

### Setup Sheet (SizeBands table)

Go to cell A2 and enter this data (or paste from below):

```
S       3       5
M       6       9
L       10      15
XL      16      25
```

**Column headers are already there:**
- Column A: Size category
- Column B: Min headcount
- Column C: Max headcount

### Products Sheet

Go to cell A2 and manually enter 2 products. Use these values:

**Row 2 (PROD-001):**
- A2: PROD-001
- B2: Authentication Platform
- C2: Business
- D2: All users
- E2: Security team  
- F2: Enterprise identity and authentication services
- X2: (leave blank - calculated)
- Y2: (leave blank - calculated)
- AA2: 2  (AS-IS team count)
- AB2: 10 (AS-IS FTE)

**Row 3 (PROD-002):**
- A3: PROD-002
- B3: Data Platform
- C3: Technology
- D3: Analytics team
- E3: Platform team
- F3: Central analytics and data insights
- X3: (leave blank)
- Y3: (leave blank)
- AA3: 3
- AB3: 18

### Teams Sheet

Go to cell A2 and manually enter 3 teams:

**Row 2 (TEAM-001):**
- A2: TEAM-001
- B2: Identity Core
- C2: PROD-001  (links to PROD-001)
- D2: (leave blank - assignment status)
- E2: Platform
- F2: Build and maintain authentication services
- G2: M (Medium size)
- H2: No

**Row 3 (TEAM-002):**
- A3: TEAM-002
- B3: Identity Ops
- C3: PROD-001
- E3: Platform
- F3: Operate and support auth systems
- G3: S (Small size)
- H3: No

**Row 4 (TEAM-003):**
- A4: TEAM-003
- B4: Data Ingestion
- C4: PROD-002  (links to PROD-002)
- E4: Platform
- F4: High-performance data ingestion pipelines
- G4: L (Large size)
- H4: No

### THEN: Add the 5 Formulas

**Teams Sheet:**

Click cell **X2** (Headcount minimum column, first data row)
Paste this formula:
```excel
=IF([@[Team name]]="","",IFERROR(INDEX(Setup.$B$2:$B$100,MATCH([@[Team size]],Setup.$A$2:$A$100,0)),""))
```

Click cell **Y2** (Headcount maximum column, first data row)
Paste this formula:
```excel
=IF([@[Team name]]="","",IFERROR(INDEX(Setup.$C$2:$C$100,MATCH([@[Team size]],Setup.$A$2:$A$100,0)),""))
```

**Products Sheet:**

Click cell **AB2** (TO-BE team count column, first data row)
Paste this formula:
```excel
=IF([@[Product ID]]="","",SUMPRODUCT((Teams.$A$2:$A$1000=[@[Product ID]])*(Teams.$A$2:$A$1000<>"")))
```

Click cell **AC2** (TO-BE headcount min column, first data row)
Paste this formula:
```excel
=IF([@[Product ID]]="","",IF([@[TO-BE team count]]=0,"",SUMIF(Teams.$A$2:$A$1000,[@[Product ID]],Teams.$X$2:$X$1000)))
```

Click cell **AD2** (TO-BE headcount max column, first data row)
Paste this formula:
```excel
=IF([@[Product ID]]="","",IF([@[TO-BE team count]]=0,"",SUMIF(Teams.$A$2:$A$1000,[@[Product ID]],Teams.$Y$2:$Y$1000)))
```

### WHAT YOU SHOULD SEE AFTER FORMULAS ARE ADDED:

**Teams sheet:**
- Row 2 (TEAM-001, size M): X2 should show 6, Y2 should show 9
- Row 3 (TEAM-002, size S): X3 should show 3, Y3 should show 5
- Row 4 (TEAM-003, size L): X4 should show 10, Y4 should show 15

**Products sheet:**
- Row 2 (PROD-001): AB2 should show 2 (teams linked), AC2 should show 9 (3+6), AD2 should show 14 (5+9)
- Row 3 (PROD-002): AB3 should show 1 (team linked), AC3 should show 10, AD3 should show 15

---

## Why This Test Matters

This simple test validates:

✅ **SizeBands lookup works**: Team sizes (S/M/L/XL) → headcount ranges  
✅ **Team-to-Product linking works**: Teams linked by Product ID → product totals  
✅ **Formula calculation works**: AS-IS independent, TO-BE calculated  
✅ **Data model is sound**: Two independent organizational perspectives  

If the formulas calculate correctly and the totals roll up, **the entire sync engine architecture is validated**.

---

## Next: Create Sample PPT

After Excel is working, fill in the PowerPoint template with the same data:

**product_team_optimized.pptx**
- Slide 1: Product slide for PROD-001 (enter AS-IS: team count=2, FTE=10)
- Slides 2-3: Team slides for TEAM-001 and TEAM-002 (enter TO-BE: sizes M and S)
- Slide 4: Product slide for PROD-002 (AS-IS: team count=3, FTE=18)
- Slide 5: Team slide for TEAM-003 (TO-BE: size L)

Then compare: PPT and Excel should have matching data for the same products/teams.

---

## Commands to Help

**To see the sample data layout in a formatted view:**
```
Review the DATA_MODEL.md section "Table Structure" for field definitions
Review the Field map sheet in the workbook for column-by-column reference
```

**After entering data, to validate:**
1. Check that each team shows calculated headcount min/max
2. Check that each product's TO-BE totals roll up from linked teams
3. Try changing a team size (e.g., M to L) and confirm product totals update
4. Try adding a new team linked to a product and confirm counts update

---

Good luck with the manual entry! This is the best way to understand how the data flows. Once you see the formulas calculating in real-time, the whole architecture will click.
