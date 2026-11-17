# PPT ↔ Excel Sync Engine – Data Model Summary

## Completed Checkpoint

You now have a complete, **tested data contract** between PowerPoint and Excel:

✅ **PowerPoint template**: `product_team_optimized.pptx`
   - 48 named answer fields (27 on product slide, 21 on team slide)
   - Separated labels from editable answers
   - Structured for AS-IS on product, TO-BE on team one-pagers

✅ **Excel master workbook**: `product_team_master.xlsx`
   - 5 sheets: Start here, Setup, Products, Teams, Field map
   - 56 total fields across two main tables
   - Configurable sizing (SizeBands), dropdown choices, calculated totals
   - Ready for formulas and data entry

✅ **Data model**: AS-IS/TO-BE separation confirmed
   - Products store current baseline (manual inputs: team count, FTE)
   - Teams store future design (actual teams, linked to products, with sizing)
   - Sibling flag documents relationships without duplicating counts
   - Calculated fields roll up from teams to products

---

## Data Model Overview

### Core Concepts

**AS-IS (Current State)**
- Captured on **product one-pagers**
- Represents indicative current organization before restructuring
- Manual inputs: team count, FTE
- **Not calculated** from TO-BE teams

**TO-BE (Future State)**
- Captured on **team one-pagers** (one actual team per pager)
- Represents proposed organization and team design
- Linked to products via Product ID
- **Calculated** from team records (team count, headcount ranges)

**Sizing & Headcount**
- **SizeBands**: Configurable mapping (S→3–5, M→6–9, L→10–15, XL→16–25, etc.)
- **Headcount ranges**: Estimated people per team, looked up from SizeBands
- **FTE**: Entered manually per product (AS-IS), independent of headcount ranges
- These are **not interchangeable**

---

## Table Structure

### Products Table (31 columns)

#### Metadata (System-managed)
| Field | Type | Purpose |
|-------|------|---------|
| Product ID | text (auto-assigned) | Unique identifier; leave blank |
| Product name | text | Product name or code |

#### Top Strip (Descriptive)
| Field | Type | Purpose |
|-------|------|---------|
| Product type | choice | Business, Data, Technology |
| Primary users | text | Customer or user segment |
| Organisational owner | text | Org owner or executive sponsor |
| Catalogue link | URL/text | Link to product catalog/registry |

#### Purpose & Value
| Field | Type | Purpose |
|-------|------|---------|
| Purpose and value | text | Why this product exists; value delivered |

#### Service & Business Criticality (5 separate fields)
| Field | Type | Purpose |
|-------|------|---------|
| SL1 system count | integer | Number of SL1 systems in product |
| SL2 system count | integer | Number of SL2 systems |
| High criticality systems | integer | Number of high-criticality systems |
| Medium criticality systems | integer | Medium-criticality count |
| Low criticality systems | integer | Low-criticality count |

#### Scope & Sourcing
| Field | Type | Purpose |
|-------|------|---------|
| Component count | integer | Total components |
| System count | integer | Total systems |
| Primary systems | text | Key system names or identifiers |
| Build percent | percentage | % built in-house |
| Buy percent | percentage | % bought/outsourced |
| Main vendor or partner | text | Primary external partner |
| Competency 1, 2, 3 | text | Required technical skills |

#### Complexity & Sizing
| Field | Type | Purpose |
|-------|------|---------|
| Technical complexity | choice | H, M, L |
| Investment direction | choice | Invest, Divest, Retain, Keeping lights on |
| Change demand 1-2 years | choice | H, M, L |

#### Boundaries
| Field | Type | Purpose |
|-------|------|---------|
| Adjacent products and teams | text | Products/teams this interfaces with |
| Dependencies and interaction model | text | Dependency description |

#### AS-IS Baseline (Manual Inputs)
| Field | Type | Purpose |
|-------|------|---------|
| AS-IS team count | integer | Indicative current team count |
| AS-IS FTE | decimal | Indicative current FTE |

#### TO-BE Reporting (Calculated from Teams table)
| Field | Type | Purpose | Formula |
|-------|------|---------|---------|
| TO-BE team count | calculated | Count of linked teams | `=COUNTIFS(Teams[Product ID],[@[Product ID]],Teams[Team name],"<>")` |
| TO-BE headcount min | calculated | Min people estimated | `=SUMIF(Teams[Product ID],[@[Product ID]],Teams["Headcount minimum"])` |
| TO-BE headcount max | calculated | Max people estimated | `=SUMIF(Teams[Product ID],[@[Product ID]],Teams["Headcount maximum"])` |

---

### Teams Table (25 columns)

#### Metadata (System-managed)
| Field | Type | Purpose |
|-------|------|---------|
| Team ID | text (auto-assigned) | Unique identifier; leave blank |
| Team name | text | Team name |

#### Relationship
| Field | Type | Purpose |
|-------|------|---------|
| Product ID | text | Link to Products.Product ID |
| Assignment status | choice | Assigned, Intentionally unassigned, To be decided |

#### Team Type
| Field | Type | Purpose |
|-------|------|---------|
| Team type | choice | Delivery, Platform, Enabling, Subsystem |

#### Purpose/Focus
| Field | Type | Purpose |
|-------|------|---------|
| Purpose and focus | text | Team's mission and focus area |

#### Composition
| Field | Type | Purpose |
|-------|------|---------|
| Composition | text | Skills, roles, disciplines on team |

#### Size (TO-BE)
| Field | Type | Purpose |
|-------|------|---------|
| Team size | lookup | References SizeBands (S, M, L, XL, etc.) |
| Sibling team? | choice | Yes, No (narrative only; separate pager per team) |

#### Sizing Reasons (5 checkboxes: [ ] not selected, [x] selected)
| Field | Type | Purpose |
|-------|------|---------|
| Sizing reason - Business criticality | boolean | Does criticality drive this size? |
| Sizing reason - Service level | boolean | Does service level (SL1/SL2) drive this? |
| Sizing reason - Technical complexity | boolean | Does complexity drive this? |
| Sizing reason - Sourcing | boolean | Does sourcing (build vs. buy) drive this? |
| Sizing reason - Investment direction | boolean | Does investment direction drive this? |

#### Topology Checklist (5 assessments: [ ] not assessed, [x] yes, ? no)
| Field | Type | Purpose |
|-------|------|---------|
| Topology - value flow | boolean | Is value flowing correctly? |
| Topology - stream alignment | boolean | Is team stream-aligned? |
| Topology - outcome ownership | boolean | Does team own outcomes? |
| Topology - handoffs and load | boolean | Are handoffs minimized/clear? |
| Topology - feedback | boolean | Is feedback loop clear? |

#### Interaction & Guardrails
| Field | Type | Purpose |
|-------|------|---------|
| Consumers | text | Who consumes this team's output? |
| Supported by | text | Who supports or enables this team? |
| Interaction modes | text | How does this team interact (APIs, events, etc.)? |
| Boundaries | text | What is out of scope for this team? |

#### Calculated (looked up from SizeBands)
| Field | Type | Purpose | Formula |
|-------|------|---------|---------|
| Headcount minimum | calculated | Min people for this size | `=INDEX(SizeBands[Min headcount],MATCH([@[Team size]],SizeBands[Size category],0))` |
| Headcount maximum | calculated | Max people for this size | `=INDEX(SizeBands[Max headcount],MATCH([@[Team size]],SizeBands[Size category],0))` |

---

### Setup Sheet

#### SizeBands Table (Configurable)
| Column | Type | Purpose |
|--------|------|---------|
| Size category | text | S, M, L, XL, XXL, etc. |
| Min headcount | integer | Minimum estimated people |
| Max headcount | integer | Maximum estimated people |

**Example:**
- S: 3–5 people
- M: 6–9 people
- L: 10–15 people
- XL: 16–25 people
- XXL: 26+ people

Add or modify sizes as needed; all team lookups reference this table.

---

### Field Map Sheet

Complete inventory of all 56 fields:
- Table & column name
- Field key (sync identifier)
- Type (text, integer, decimal, choice, lookup, boolean, calculated)
- Allowed values/format
- PPT section
- Mapping status (ready, pending, engine-managed, Excel-only)
- Sync readiness notes

---

## Key Design Decisions

### 1. Separate Tables (Products ≠ Teams)
- Products represent **one organizational perspective** (current baseline or reporting view)
- Teams represent **another perspective** (proposed team design and structure)
- They are **independent** data sources; neither is calculated from the other
- Both feed the sync engine, but with different meanings

### 2. Manual IDs
- **Product ID** and **Team ID** are assigned by the **sync engine**, not manually
- Leave blank when creating records; the engine populates these
- Use Product ID in Teams sheet to establish the relationship

### 3. Sizing is Configurable
- Size bands (S, M, L, XL) are **not hardcoded**
- Defined in the Setup sheet so you can add XL, XXL, or custom sizes anytime
- Headcount ranges are **estimated people**, not FTE
- Independent from FTE (which is entered manually per product AS-IS)

### 4. Sibling Flag is Narrative
- **Sibling team? Yes/No** documents whether a team is part of a sibling group
- **Does not create duplicate counting units**
- One actual team = one one-pager = one row in the Teams table
- Sibling relationship is captured for context but doesn't affect roll-up calculations

### 5. Five Separate Criticality Fields
- Instead of a single "criticality" field, the product slide has **5 separate count fields**
  - SL1 count, SL2 count, High, Medium, Low
- Allows granular capture of system distribution
- Easier to compare AS-IS (current counts) with reported/known criticality

### 6. Checklists, Not Dropdowns
- **Sizing reasons** and **Topology** are sets of **independent checkboxes**
- Not a multi-select or single-choice dropdown
- Formatting: `[x]` = selected, `[ ]` = not selected, `?` = no (for topology)
- Each checkbox is a separate column, one value per cell

---

## Sync Workflow (Proposed)

1. **Review → Approve**: Users edit in their preferred tool (PPT or Excel)
2. **Inspect**: Read-only tool compares both files, highlights changes
3. **Propose**: Tool suggests which changes to sync where
4. **Approve & Apply**: User approves, changes written to target file
5. **Conflict Resolution**: Manual intervention if both files changed the same field
6. **Publish**: Regenerate one-pagers with updated data

---

## Ready for Next Phase

✅ Data structure is locked
✅ Field mappings are documented
✅ Formulas are defined (ready to add manually in Excel)
✅ Sizing is configurable
✅ Template and workbook are aligned

**Next: Build the read-only inspector and change-proposal engine.**

---

## Files

- **product_team_optimized.pptx**: PowerPoint template (48 named fields)
- **product_team_master.xlsx**: Excel workbook (5 sheets, 56 fields)
- **WORKBOOK_GUIDE.md**: User guide with conventions and tips
- **create-master-workbook.ps1**: Script to regenerate workbook
- **add-formulas-guide.ps1**: Formula reference guide
