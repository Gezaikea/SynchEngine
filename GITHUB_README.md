# PPT ↔ Excel Sync Engine

Automatically sync PowerPoint one-pagers with Excel data. Eliminate manual data entry, detect changes, and keep product/team information synchronized across both platforms.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.8+](https://img.shields.io/badge/Python-3.8%2B-blue.svg)](https://www.python.org/)

## 🎯 What It Does

Fill out a PowerPoint template → Run the sync engine → Your data is automatically in Excel

```bash
python sync_full.py
```

**Result:**
- ✅ All 47 PPT fields extracted automatically
- ✅ Data written to Excel in seconds
- ✅ Changes detected and reported
- ✅ Zero manual data entry
- ✅ Zero errors

## 🚀 Quick Start

### Installation

```bash
# Clone repository
git clone https://github.com/yourusername/ppt-excel-sync.git
cd ppt-excel-sync

# Install dependencies
pip install -r requirements.txt
```

### Usage

```bash
# Sync PPT to Excel (recommended)
python sync_full.py

# Detect changes only (safe review)
python sync_engine_v2.py
```

### Output

```
[STEP 1] Reading PowerPoint...
         Extracted 47 fields from PPT

[STEP 2] Parsing extracted data...
         Product data: 27 fields
         Team data: 21 fields

[STEP 3] Writing to Excel...
         [OK] Product written to Excel
         [OK] Team written to Excel

[STEP 4] Verifying written data...
         ✓ All data verified in Excel file

SYNC COMPLETE!
```

## 📊 What Data Gets Synced?

### Product Record (27 fields)
- **Identification:** ID, name, type, owner
- **Context:** Primary users, purpose & value, catalog link
- **Systems:** SL1/SL2 counts, criticality levels, component count
- **Composition:** Primary systems, build/buy %, vendors
- **Strategy:** Competencies, complexity, investment, change demand
- **Current state:** AS-IS team count, AS-IS FTE
- **Dependencies:** Adjacent products, interaction model

### Team Record (21 fields)
- **Identification:** ID, name, type, size (S/M/L/XL)
- **Purpose:** Purpose, composition, consumers, boundaries
- **Interaction:** Interaction modes, supported by
- **Sizing:** 5 sizing reasons (criticality, service level, complexity, sourcing, investment)
- **Topology:** 5 design principles (value flow, stream alignment, outcome ownership, handoffs, feedback)

## 🏗️ Architecture

```
PowerPoint Template
        ↓
  PPTReader (extract 47 named fields)
        ↓
  ChangeDetector (compare PPT vs Excel)
        ↓
  SyncReport (generate change proposals)
        ↓
  ExcelWriter (write to Excel)
        ↓
  Excel File (updated with your data)
```

## 📁 Project Structure

```
ppt-excel-sync/
├── sync_engine_v2.py           # Core reader & change detector
├── sync_engine_writer.py       # Excel writer
├── sync_full.py               # Complete automation script
├── sync_engine_validation.py  # Data validation & cleaning (TODO)
├── requirements.txt           # Python dependencies
├── LICENSE                    # MIT License
├── README.md                  # This file
├── CONTRIBUTING.md            # How to contribute
├── docs/
│   ├── ARCHITECTURE.md        # Technical architecture
│   ├── USAGE_GUIDE.md         # Detailed usage instructions
│   └── API_REFERENCE.md       # API documentation
├── templates/
│   └── product_team_optimized_test.pptx  # Example template
└── examples/
    └── sync_full.py           # Full workflow example
```

## 🎯 Real-World Workflow

### Weekly Sync Process

```
Monday 9 AM:
  1. Leaders update PPT templates
  2. Run: python sync_full.py
  3. Excel automatically updated

Monday 10 AM:
  1. Data team has latest data
  2. Build reports and analyses
  3. No manual data entry needed
```

### Adding New Products/Teams

```
1. Create new slides in PPT template
2. Fill in all required fields
3. Run: python sync_full.py
4. Engine detects NEW records
5. Excel rows created automatically
```

## 🔄 Core Components

### PPTReader
Extracts data from PowerPoint templates
```python
from sync_engine_v2 import PPTReader

reader = PPTReader('template.pptx')
data = reader.read()  # Returns all 47 fields
```

### ChangeDetector
Identifies what changed between PPT and Excel
```python
from sync_engine_v2 import ChangeDetector

detector = ChangeDetector(ppt_data, excel_data)
records = detector.detect()  # Returns changes grouped by record
```

### ExcelWriter
Writes data to Excel automatically
```python
from sync_engine_writer import ExcelWriter

writer = ExcelWriter('workbook.xlsx')
writer.open()
writer.write_product(id, data)
writer.write_team(id, data)
writer.close()
```

## 📈 Features

### Implemented ✅
- [x] Read PPT templates
- [x] Extract all 47 fields
- [x] Detect changes (NEW/MODIFIED/DELETED/CONFLICT)
- [x] Generate change reports
- [x] Write to Excel automatically
- [x] Automatic row creation
- [x] Conflict detection
- [x] Field mapping
- [x] Data verification

### In Progress 🔄
- [ ] Data validation layer
- [ ] Percentage extraction & normalization
- [ ] Placeholder text detection
- [ ] Text cleaning (special characters, typos)
- [ ] Excel formula generation (calculated fields)

### Planned 🗓️
- [ ] Bidirectional sync (Excel → PPT)
- [ ] Conflict resolution UI
- [ ] Automatic scheduling
- [ ] GUI application
- [ ] Comprehensive logging
- [ ] Unit test suite

## 📋 Requirements

- Python 3.8+
- PowerPoint file with named answer fields
- Excel file with Products and Teams tables

### Dependencies
```
python-pptx>=0.6.21
openpyxl>=3.0.0
```

## 🛠️ Configuration

### Field Mapping

All PPT named fields are mapped to Excel columns in `sync_engine_v2.py`:

```python
PPT_FIELD_MAP = {
    'sync_product_name': ('Product name', 'Products', 'text'),
    'sync_team_name': ('Team name', 'Teams', 'text'),
    # ... 46 more fields
}
```

### Customization

To add more fields:

1. **Add to PPT template:** Create a new shape with name starting with `sync_`
2. **Update field map:** Add entry to `PPT_FIELD_MAP` in `sync_engine_v2.py`
3. **Add Excel column:** Ensure Excel has the destination column
4. **Test:** Run `python sync_full.py`

## 📖 Documentation

- **[QUICK_START.md](docs/QUICK_START.md)** - Get running in 5 minutes
- **[ARCHITECTURE.md](docs/ARCHITECTURE.md)** - Technical deep-dive
- **[USAGE_GUIDE.md](docs/USAGE_GUIDE.md)** - Detailed how-to guide
- **[API_REFERENCE.md](docs/API_REFERENCE.md)** - API documentation
- **[TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md)** - Common issues & fixes

## 🧪 Testing

```bash
# Run with example data
python sync_full.py

# Generate change report (safe, no changes)
python sync_engine_v2.py

# View report
type SYNC_REPORT.txt
```

## 🤝 Contributing

We welcome contributions! Please see [CONTRIBUTING.md](CONTRIBUTING.md) for:
- Code style guidelines
- How to submit issues
- How to submit pull requests
- Development setup

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🎓 Examples

### Basic Sync
```python
from sync_engine_v2 import PPTReader
from sync_engine_writer import ExcelWriter

# Read PPT
reader = PPTReader('template.pptx')
ppt_data = reader.read()

# Write to Excel
writer = ExcelWriter('workbook.xlsx')
writer.open()
writer.write_product('product_id', ppt_data)
writer.close()
```

### Detect Changes
```python
from sync_engine_v2 import PPTReader, ExcelReader, ChangeDetector, SyncReport

# Read both sources
ppt_reader = PPTReader('template.pptx')
excel_reader = ExcelReader('workbook.xlsx')

ppt_data = ppt_reader.read()
excel_data = excel_reader.read()

# Detect changes
detector = ChangeDetector(ppt_data, excel_data)
records = detector.detect()

# Generate report
report = SyncReport(records)
print(report.generate_summary())
```

## 📞 Support

- **Issues:** Use GitHub Issues to report bugs
- **Questions:** Start a GitHub Discussion
- **Documentation:** Check the [docs](docs/) folder

## 🗺️ Roadmap

### Phase 1: PPT → Excel (Complete ✅)
- [x] Read PPT templates
- [x] Extract all fields
- [x] Write to Excel
- [x] Change detection

### Phase 2: Excel → PPT (In Progress 🔄)
- [ ] Update PPT from Excel changes
- [ ] Handle slide creation/deletion
- [ ] Conflict resolution

### Phase 3: Advanced Features (Planned 🗓️)
- [ ] Data validation & cleaning
- [ ] Automatic scheduling
- [ ] GUI application
- [ ] Database backend option

## 💡 Tips & Best Practices

1. **Always review changes first**
   ```bash
   python sync_engine_v2.py  # Safe: no changes made
   # Review SYNC_REPORT.txt
   python sync_full.py        # Apply changes
   ```

2. **Keep files closed**
   - Close Excel and PowerPoint before syncing
   - Wait 2 seconds if they were just closed

3. **Backup before first sync**
   ```bash
   cp product_team_master.xlsx product_team_master.backup.xlsx
   ```

4. **Use meaningful IDs**
   - Product/team IDs must be unique
   - Use consistent naming (e.g., "01_Engineering_Toolchain")

## 🎉 Acknowledgments

Built to solve real organizational challenges in syncing product/team information across multiple tools.

---

## ⭐ Show Your Support

If this project helps you, please consider giving it a star! It helps others discover it.

---

**Made with ❤️ for seamless data sync**
