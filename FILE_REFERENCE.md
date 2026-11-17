# Sync Engine - File Overview

## 🎯 START HERE

**To sync your PPT to Excel:**
```bash
python sync_full.py
```

---

## 📂 Core Engine Files

### `sync_engine_v2.py` - The Brain
**What it does:**
- Reads PowerPoint templates
- Detects changes between PPT and Excel
- Generates change reports

**Main classes:**
- `PPTReader` - Extract fields from PPT
- `ExcelReader` - Read Excel tables
- `ChangeDetector` - Compare PPT vs Excel
- `SyncReport` - Generate human-readable reports

**Use it for:**
```bash
python sync_engine_v2.py  # Analyze without changing anything
```

**Output:** `SYNC_REPORT.txt` (shows all changes detected)

---

### `sync_engine_writer.py` - The Executor
**What it does:**
- Writes extracted PPT data to Excel
- Finds or creates rows automatically
- Handles multiple products/teams

**Main classes:**
- `ExcelWriter` - Write data to Excel
- `PPTWriter` - Skeleton for future PPT writing

**Use it for:**
```python
from sync_engine_writer import ExcelWriter
writer = ExcelWriter(excel_path)
writer.open()
writer.write_product(id, data)
writer.write_team(id, data)
writer.close()
```

---

### `sync_full.py` - The Complete Automation
**What it does:**
- Reads PPT
- Extracts data
- Writes to Excel
- Verifies success
- Shows summary

**Use it for:**
```bash
python sync_full.py  # One-shot: PPT -> Excel
```

**This is your main script.** Run this after filling PPT templates.

---

## 📋 Documentation Files

### `README.md` - Quick Start Guide
- How to use the engine
- Real-world scenarios
- Troubleshooting
- **Read this first**

### `SYNC_ENGINE_DEMO.md` - Architecture Details
- How the engine works
- Real data examples
- Field mapping reference
- Next phases (bidirectional sync, conflict resolution)

### `IMPLEMENTATION_COMPLETE.md` - Technical Summary
- Component breakdown
- Validation results
- Performance metrics
- Advanced workflows

---

## 📊 Data Files

### `product_team_optimized_test.pptx`
- Your filled PPT template
- 2 slides (1 product + 1 team)
- 47 named fields with real data
- Input for sync engine

### `product_team_master.xlsx`
- Excel destination
- Products sheet (28 columns, now has row 2 with data)
- Teams sheet (25 columns, now has row 2 with data)
- Output from sync engine

### `SYNC_REPORT.txt`
- Generated change report
- Shows all 41 NEW fields detected
- Safe to review before syncing

---

## 🔄 Typical Workflow

```
1. You fill in PPT template
   └─> Edit: product_team_optimized_test.pptx

2. You run sync engine
   └─> Execute: python sync_full.py

3. Engine reads PPT
   └─> Uses: sync_engine_v2.py (PPTReader)

4. Engine writes to Excel
   └─> Uses: sync_engine_writer.py (ExcelWriter)

5. Excel is updated
   └─> Result: product_team_master.xlsx (new row with your data)

6. Done! No manual work needed.
```

---

## 🎓 For Data Scientists / Developers

### To analyze changes programmatically:

```python
from sync_engine_v2 import PPTReader, ExcelReader, ChangeDetector

# 1. Read sources
ppt_reader = PPTReader(r'path\to\ppt')
ppt_data = ppt_reader.read()

excel_reader = ExcelReader(r'path\to\excel')
excel_data = excel_reader.read()

# 2. Detect changes
detector = ChangeDetector(ppt_data, excel_data)
records = detector.detect()

# 3. Access results
for record in records:
    if record.has_changes:
        print(f"{record.record_type}: {record.record_name}")
        print(f"  Changes: {record.change_summary}")
        
        for field_name, field_value in record.fields.items():
            if field_value.status != ChangeType.MATCHED:
                print(f"    {field_value.status.value}: {field_name}")
                print(f"      PPT: {field_value.ppt_value}")
                print(f"      Excel: {field_value.excel_value}")
```

### To write data to Excel:

```python
from sync_engine_writer import ExcelWriter

writer = ExcelWriter(r'path\to\excel')
writer.open()

# Write single record
writer.write_product('product_id', {'Product name': 'My Product', ...})
writer.write_team('team_id', {'Team name': 'My Team', ...})

# Or write batch
writer.write_batch(
    products={'id1': {...}, 'id2': {...}},
    teams={'id3': {...}, 'id4': {...}}
)

writer.close()
```

---

## 📦 Field Reference

### Product Fields (27 synced)
```
Product name, type, owner, catalog link
Primary users, purpose and value
SL1 count, SL2 count, criticality counts
System count, component count
Primary systems, build %, buy %, vendor
3 competencies
Technical complexity, investment direction, change demand
Adjacent products, dependencies
AS-IS team count, AS-IS FTE
```

### Team Fields (21 synced)
```
Team name, type, size
Purpose, composition
Consumers, boundaries, interaction modes
5 sizing reasons (business criticality, service level, complexity, sourcing, investment)
5 topology decisions (value flow, stream alignment, outcome ownership, handoffs, feedback)
```

---

## ✅ Features Implemented

| Feature | Status | File |
|---------|--------|------|
| Read PPT | ✅ Done | sync_engine_v2.py |
| Read Excel | ✅ Done | sync_engine_v2.py |
| Detect changes | ✅ Done | sync_engine_v2.py |
| Generate report | ✅ Done | sync_engine_v2.py |
| Write to Excel | ✅ Done | sync_engine_writer.py |
| Full automation | ✅ Done | sync_full.py |
| Conflict detection | ✅ Done | sync_engine_v2.py |
| Field mapping (48 fields) | ✅ Done | PPT_FIELD_MAP |
| Bidirectional (PPT writer) | 🔲 Skeleton | sync_engine_writer.py |
| Scheduling | 🔲 Not yet | - |
| GUI | 🔲 Not yet | - |

---

## 🚀 To Use Immediately

### Minimum (Just sync):
1. Fill PPT: `product_team_optimized_test.pptx`
2. Run: `python sync_full.py`
3. Check Excel: `product_team_master.xlsx`

### Safe (Review before sync):
1. Run: `python sync_engine_v2.py` (creates SYNC_REPORT.txt)
2. Review: `SYNC_REPORT.txt` (see what will sync)
3. Run: `python sync_full.py` (apply changes)

### Advanced (Custom logic):
1. Use `sync_engine_v2.py` for reading and detecting
2. Use `sync_engine_writer.py` for writing
3. Build custom scripts combining both

---

## 📞 Quick Reference

```bash
# Just sync PPT to Excel
python sync_full.py

# Detect changes only (safe review)
python sync_engine_v2.py

# View the report
type SYNC_REPORT.txt  (Windows)
cat SYNC_REPORT.txt   (Mac/Linux)

# Edit PPT template
start product_team_optimized_test.pptx  (Windows)
open product_team_optimized_test.pptx   (Mac)

# View Excel results
start product_team_master.xlsx  (Windows)
open product_team_master.xlsx   (Mac)
```

---

## 💡 Tips

1. **Always run change detection first** - Use `sync_engine_v2.py` to review before syncing
2. **Keep PPT and Excel in sync** - Run sync after every PPT change
3. **Use meaningful IDs** - Product/team IDs are used to match records
4. **Document your templates** - Keep notes on which PPT fields map to which Excel columns
5. **Back up your Excel** - Before first sync, copy product_team_master.xlsx to backup

---

## 🎉 You're All Set!

Your sync engine is ready to use. No configuration needed.

Just run: `python sync_full.py`

And enjoy automated, error-free data syncing between PPT and Excel! 🚀
