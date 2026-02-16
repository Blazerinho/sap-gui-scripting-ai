# CLAUDE.md - SAP GUI Scripting AI Agent

## Project Overview

AI-powered SAP GUI automation agent. Uses LangGraph + LangChain with SAP Generative AI Hub to interpret natural language commands and perform SAP GUI operations via Windows COM automation (pywin32).

The user works in SAP FI/CO and related modules, receives analysis requests via email, runs SAP
queries, and replies with results. The goal is to have Claude Code handle this conversationally
— no need to save scripts unless explicitly asked.

## Tech Stack

- **Language:** Python 3.14.0
- **AI Framework:** LangGraph 1.0.7, LangChain 1.2.7
- **LLM Provider:** SAP AI Core (Generative AI Hub) - currently using gpt-4o
- **SAP GUI Access:** pywin32 (COM automation)
- **Config:** python-dotenv (.env file)

## Project Structure

```
sap_gui_scripting_ai.py   # Main AI agent application (single file)
sap_scripting.py           # Reusable SapSession & OutlookSession classes
.env                       # SAP AI Core credentials (do NOT commit)
venv/                      # Python virtual environment
```

## Setup & Run

```bash
# Activate virtual environment
venv\Scripts\activate

# Install dependencies (if needed)
pip install pywin32 langchain langgraph generative-ai-hub-sdk[langchain] python-dotenv

# Run (SAP GUI must be open with an active session)
python sap_gui_scripting_ai.py
```

## Prerequisites

- Windows OS (COM dependency)
- SAP GUI for Windows with scripting enabled (Alt+F12 → Options → Scripting)
- Server parameter: `sapgui/user_scripting = TRUE` (set via RZ11)
- Microsoft Outlook running locally (for email integration)
- Valid SAP AI Core credentials in `.env`

## Architecture

1. `SAPAutomation` class — wraps SAP GUI COM interactions (find elements, set text, press buttons, read state)
2. `@tool`-decorated functions — LangChain tools the LLM can invoke
3. LangGraph state machine — `reasoner_node` (LLM) ↔ `tools_node` (execute), loops until LLM signals done
4. GUI state is injected into each reasoning step so the LLM has context about the current screen

## SAP Connection Pattern

Always connect via the Running Object Table — never launch SAP GUI from script:
```python
import win32com.client
SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine
connection = application.Children(0)   # first connection
session = connection.Children(0)       # first session
```

## SAP GUI Scripting API — Key Methods

Reference: SAP GUI Scripting API 6.40 (PDF in project folder)

### Hierarchy
GuiApplication → GuiConnection → GuiSession → GuiMainWindow → components

### Navigation
- `session.StartTransaction("TCODE")` — go to transaction (same as SendCommand("/nTCODE"))
- `session.EndTransaction()` — back to menu (same as SendCommand("/n"))
- `session.SendCommand("/nXYZ")` — execute any command field entry
- `session.findById("wnd[0]").sendVKey(0)` — Enter; sendVKey(8) = F8/Execute; sendVKey(3) = Back

### Finding elements
- `session.findById("wnd[0]/usr/ctxtFIELD")` — by full path (most reliable)
- `session.findByName("FIELD", "ctxt")` — by SAP data dictionary name + type prefix
- `session.findAllByName("FIELD", "txt")` — returns collection of all matches
- Type prefixes: txt=TextField, ctxt=CTextField, chk=CheckBox, rad=RadioButton, btn=Button, cmb=ComboBox, lbl=Label, tbl=TableControl, tabs=TabStrip, tabp=Tab

### Reading/writing fields
- `element.text = "value"` — set field (GuiTextField, GuiCTextField)
- `element.text` — read field value (read-only for GuiPasswordField)
- `element.selected = True` — checkbox (GuiCheckBox)
- `element.select()` — radio button (GuiRadioButton)
- `element.press()` — button (GuiButton)
- `element.key = "value"` — combo box selection (GuiComboBox)

### GuiGridView (ALV grids)
- `grid.RowCount` — number of rows
- `grid.ColumnCount` — number of columns
- `grid.ColumnOrder` — collection of column ID strings
- `grid.GetCellValue(row, "COLUMN_ID")` — read cell
- `grid.GetDisplayedColumnTitle("COLUMN_ID")` — column header text
- `grid.Click(row, "COLUMN_ID")` / `grid.DoubleClick(...)` — interact with cells
- `grid.SelectedRows = "0,1,3-5"` — select rows

### GuiTableControl (dynpro tables — different from GridView)
- Cell path pattern: `table_id/ctxtFIELDNAME[col_idx, row_idx]`
- `table.RowCount`, `table.Columns`

### Status bar
- `session.findById("wnd[0]/sbar").Text` — message text
- `session.findById("wnd[0]/sbar").MessageType` — S/W/E/A/I

### Session info
- `session.Info.SystemName`, `.Client`, `.User`, `.Language`, `.Transaction`, `.Program`, `.ScreenNumber`

### Popup handling
- Modal windows are `wnd[1]`, `wnd[2]`, etc.
- Typically dismiss with: `session.findById("wnd[1]/tbar[0]/btn[0]").press()`

### Screen exploration (for unknown screens)
Walk `container.Children` to discover element IDs, types, and names:
```python
usr = session.findById("wnd[0]/usr")
for i in range(usr.Children.Count):
    child = usr.Children(i)
    print(f"{child.Id}  type={child.Type}  name={child.Name}")
```

### Session locking
- `session.LockSessionUI()` — prevent user interaction during script
- `session.UnlockSessionUI()` — release

### Visual debugging
- `element.Visualize(True)` — draw red frame around element

## Common Grid Container Paths

Try these in order when looking for an ALV grid:
1. `wnd[0]/usr/cntlRESULT_LIST/shellcont/shell`
2. `wnd[0]/usr/cntlCONTAINER/shellcont/shell`
3. `wnd[0]/usr/cntlGRID1/shellcont/shell`
4. `wnd[0]/usr/cntlGRID/shellcont/shell`
5. `wnd[0]/usr/cntlALV_CONTAINER/shellcont/shell`

## User Preferences for SAP Analysis

### Preferred approach: SE16H with aggregation
When checking data distributions or distinct values, the user prefers **SE16H** over
SE16/SE16N because it supports server-side grouping and summing. This is much faster
than reading all rows and aggregating in Python.

SE16H pattern:
1. Enter table name in `GD-TAB`
2. Enter field names in the fields table: `tblSAPLSE16HFIELDS_TABLE/ctxtGS_FIELDS-FIELDNAME[0,row]`
3. Tick group checkbox: `tblSAPLSE16HFIELDS_TABLE/chkGS_FIELDS-AGGR[4,row]`
4. Tick sum checkbox (for amounts): `tblSAPLSE16HFIELDS_TABLE/chkGS_FIELDS-SUM[5,row]`
5. Press Enter then F8

**Important:** The field name goes in column 0, and the field is then on that row index.
The group checkbox column index is 4, sum is 5. These may vary by SAP version — if they
don't work, use `explore_screen` or record a script to confirm.

### Common SAP tables
- **BSIK** — Vendor open items (key fields: BUKRS, LIFNR, UMSKZ, DMBTR, BELNR)
- **BSID** — Customer open items (key fields: BUKRS, KUNNR, UMSKZ, DMBTR, BELNR)
- **BSIS** — G/L account open items (key fields: BUKRS, HKONT, DMBTR, BELNR)
- **BSAK** — Vendor cleared items
- **BSAD** — Customer cleared items
- **BSAS** — G/L cleared items
- **EKKO** — Purchase order header (EBELN, LIFNR, BUKRS, BSART)
- **EKPO** — Purchase order items (EBELN, EBELP, MATNR, MENGE, NETWR)
- **VBAK** — Sales order header (VBELN, KUNNR, VKORG, AUART)
- **VBAP** — Sales order items (VBELN, POSNR, MATNR, KWMENG, NETWR)
- **LFA1** — Vendor master (LIFNR, NAME1, LAND1)
- **KNA1** — Customer master (KUNNR, NAME1, LAND1)
- **UMSKZ** — Special G/L indicator field (used across BSIK, BSID, BSIS)

### Common transactions
- **SE16H** — Table browser with aggregation (preferred for data analysis)
- **FBL1N** — Vendor line items
- **FBL3N** — G/L line items
- **FBL5N** — Customer line items
- **ME2M** — Purchase orders by material
- **VA05** — Sales order list
- **XK03** / **FK03** — Display vendor master
- **XD03** / **FD03** — Display customer master

## Outlook Integration

Outlook is also accessed via COM:
```python
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox
```

- Read emails: `inbox.Items`, filter with `.Restrict("[Unread] = True")`
- Sort: `items.Sort("[ReceivedTime]", True)` for most recent first
- Email properties: `.Subject`, `.Body`, `.SenderName`, `.SenderEmailAddress`, `.ReceivedTime`
- Create draft reply: `msg.Reply()` → set `.Body` → `.Save()` (NEVER `.Send()`)
- Create new draft: `outlook.CreateItem(0)` → set `.To`, `.Subject`, `.Body` → `.Save()`
- Folder constants: Inbox=6, Drafts=16, SentMail=5

## Important Rules

1. **Never send emails automatically** — always `.Save()` as draft for user review
2. **Always check the status bar** after executing — catch errors before reading grids
3. **Handle popups** — some transactions show confirmation dialogs (wnd[1])
4. **Grid paths vary** — try multiple common paths or use screen exploration
5. **Element IDs can differ** between SAP versions/layouts — if an ID fails, explore the screen or ask the user to do a quick script recording
6. **Keep code minimal** — don't save scripts to files unless the user asks; just execute inline
7. **Use SE16H** for data analysis queries; use specific transactions (FBL1N etc.) when the user asks for formatted reports

## Key Patterns

- Tools return human-readable strings as results
- `AgentState` (TypedDict) accumulates messages for conversation context
- COM initialized with `pythoncom.CoInitialize()`; SAP GUI accessed via `win32com.client.GetObject("SAPGUI")`
- Virtual key codes used for keyboard interaction (e.g., `session.SendVKey(0)` = Enter)

## Library File

`sap_scripting.py` in this project folder contains a reusable `SapSession` class and
`OutlookSession` class with all the above patterns pre-built. Import from it when useful:
```python
from sap_scripting import SapSession, OutlookSession, run_se16h
```

## Project Goal

The user wants to organise work projects better and create documentation for onboarding
new colleagues. When creating documentation or guides, produce clear, structured documents
that help new team members get up to speed quickly.

## Posting Engine Knowledge (DMLT /SLO/PECON)

Reference documentation in `Posting Engine Knowledge/` folder.
- **PECON_User_Guide_Version_2025_10.pdf** — latest user guide (PE_S4 Version 12, Oct 2025, 270 pages)
- **PECON_Posting_Engine_User_Guide_Version_2.1.pdf** — user guide (PE_S4 Version 7, Oct 2024, 251 pages)
- **PECON_FAQ_Document_Version_1.5.pdf** — troubleshooting FAQ (34 pages, 36 common errors with solutions)
- **DMLT PE Enablement_PowerPoint V2.pptx** — enablement presentation
- **SharePoint link** — DMLT Delivery Enablement Portal (requires SAP corp network)

### What is Posting Engine?

The Posting Engine (PE) is SAP's **API-based posting tool** used in DMLT (Data Management & Landscape Transformation) projects for **migrating financial open items and balances** between SAP systems. It creates new documents in the target system via SAP standard interfaces (application layer posting), making migrations fully traceable.

**Key characteristics:**
- Migrates open items and balances at key date (NOT historical data)
- Uses SAP standard APIs (ACDOCGEN, RFBIBL, BAPI) for posting
- Supports cross-system and in-system scenarios
- Pre-delivered content templates for common FI objects
- Requires master data and customizing to be in place in target system
- Content delivered via PECON transport packages (imported into client 000)

### System Architecture

Three-system landscape:
- **Sender (Source) System** — where data is read from (e.g., ERP/ECC)
- **Control System** — where Posting Engine runs and orchestrates (can be same as sender or receiver)
- **Receiver (Target) System** — where new documents are posted (e.g., S/4HANA)

Systems connected via **RFC type 3** (ABAP connections) with automatic logon. RFC user needs authorization for Finance Application Data access and Document Postings.

### Key Transactions

| Transaction | Purpose |
|---|---|
| `/SLO/PECON` | Main entry point — PECON Start Transaction (menu for all PE functions) |
| `CNV_PE_PROJ` | Project Maintenance — create/manage PE projects, areas, generation |
| `CNV_PE_EDITOR` | PE Editor — maintain transfer methods, interface mappings |
| `CNV_PE_SUPPORT_PROJ_GEN_TRULES` | Generate all transformation rules for a project |
| `CNV_PE_SUPPORT_IF_MAP_MAINT` | Maintain interface parameter mappings |
| `CNV_PE_TPM_FULL_PROJ_TRANSPORT` | Create full project transport |
| `CNV_OT_TOOL_ACCESS` | Release tool access (mark "PE" and press "allow") |

### Core Terminology

| Term | Description |
|---|---|
| **Project** | Top-level bracket around all settings and data for a migration scenario |
| **Migration Object / Area** | Business object to migrate (e.g., Vendor OI, Customer OI, G/L OI) |
| **Data Model** | Technical description — ONE header table + multiple daughter tables |
| **Worklist / Instance List** | All selected data for the business object (operation queue) |
| **Worklist Item** | Individual record selected for migration |
| **Transfer Option / Execution Rule** | Defines target system/client for transfer |
| **Transfer Method** | Concrete transformation rules + technical interface (FM, BAPI, etc.) |
| **Configuration Proposal** | Automatic assignment of Execution Rule to Worklist Items |
| **Transfer List** | Result after all field transformations — passed to transfer function |
| **Transformation / TRule** | Set of all field transformation rules (ABAP source code rules) |
| **Skip Rule** | ABAP logic to skip individual records during selection (naming: `_PE_SKIP`) |
| **Ignore Rule** | Skip specific transformation rules during processing |
| **Event Rule** | Custom logic triggered at specific processing events |
| **WL1 Modification Rule** | Modify worklist items after selection (naming: `_PE_WL1M*`) |
| **WL2 / Transfer List Modification** | Modify transfer list data before posting |
| **Extension Table (EXTTAB)** | PE-internal structure for additional data beyond standard tables |
| **Account List** | Maps source accounts to target company codes, G/L accounts, migration types |
| **Knowledge Base** | Reusable mapping tables for field transformations |

### Available Migration Object Templates

| Template Area | Description |
|---|---|
| `PECON_AP_OI` | Vendor Open Items (header table: BSIK) |
| `PECON_AP_OI_DOCSPLIT` | Vendor Open Items with document splitting |
| `PECON_AR_OI` | Customer Open Items (header table: BSID) |
| `PECON_AR_OI_DOCSPLIT` | Customer Open Items with document splitting |
| `PECON_GL_OI` | G/L Open Items (header table: BSIS) |
| `PECON_GL_OI_DOCSPLIT` | G/L Open Items with document splitting |
| `PECON_GL_OI_LEDGERSPEZ` | G/L Ledger Specific Clearing |
| `PECON_FI_POSTING_UPLOAD` | FI Postings from Excel upload |
| `PECON_POSTING_WITH_REFERENCE` | Create Posting with Reference (FBR2) |
| `PECON_TOTALS_FI_TT` | FI Totals from /SLO/PECON_FI_TT |
| `PECON_FI_AA_S4_OVERTAKE_LEGACY` | Legacy Fixed Asset Migration to S/4 via Overtake |
| `PECON_FI_AA_S4_TRANSFER` | Asset Transfer on S/4 |

### Execution Rules (Posting Methods)

**Sender-side (clearing in source):**
- `SIMPLE_CLEAR_FI_AP/AR/GL_OI_SND` — Direct posting clearing
- `RFBIBL_CLEAR_FI_AP/AR/GL_OI_SND` — Batch input clearing (also for Net Documents)
- `SLO_PECON_FI_CLEAR_SIMPLE_M` — Bulk clearing capability for G/L

**Receiver-side (posting in target):**
- `ACDOCGEN_FI_AP/AR/GL_OI_RCV` — Standard ACDOCGEN posting for open items
- `ACDOCGEN_FI_AP/AR_EX_RCV` — Extended posting (gross down payments, net documents)
- `RFBIBL_FI_AP/AR/GL_OI_RCV` — RFBIBL-based posting (legacy, not supported in newer S/4)

**Special topics supported:** Deferred Tax, Withholding Tax, Down Payments (gross/net), Document Splitting, Asset Accounting (classic & new)

### Configuration Parameters (Config IDs)

| Config ID | Purpose |
|---|---|
| `GROUP_ID` | Grouping of worklist items for processing |
| `CURR` | Currency handling and conversion |
| `RECON` | Reconciliation account determination |
| `OFFSETTAB` | Offset table for building counter-entries |
| `CLRSCLS` | Simple clearing exit class |
| `RETIRE` | Asset retirement (AA) |
| `ANLB` | Asset depreciation areas (AA) |
| `AQUI` | Asset acquisition (AA) |
| `ENRICH` | Asset data enrichment (AA) |

### Project Implementation Workflow

1. **Install PECON content** — Import transport packages, release tool access (`CNV_OT_TOOL_ACCESS`)
2. **Create Project** (`CNV_PE_PROJ`) — Define name, RFC destinations for sender/receiver
3. **Set Active Project** — Set project as active for your user, choose Test/Productive mode
4. **Create Migration Objects** — Copy from pre-delivered templates (e.g., PECON_AP_OI)
5. **Configure Data Model** — Review header/daughter tables, add extension tables if needed
6. **Configure Worklist** — Add custom fields, create WL modification rules
7. **Configure Account List** — Map source to target company codes, G/L accounts
8. **Create Transfer Options** — Define sender/receiver execution rules
9. **Create Transfer Methods** — Configure field transformations, TRules
10. **Implement Exits** — Custom ABAP logic for clearing, simulation, transfer
11. **Configure Proposals** — Automatic execution rule assignment
12. **Generate** — Run task generation (creates DDIC objects, function modules)
13. **Transport** — Move PE project to target landscape

### Project Execution Workflow

1. **Prepare** — Create number range for SLO Transaction ID, prepare account list, set RFC connections
2. **Re-Generate** — Ensure all tasks are generated after any config changes
3. **Select Data** — Run selection programs to fill worklist (DMSEL, Selection Control, or enhanced PECON selection)
4. **Build** — Build transfer list from worklist items (applies transformations)
5. **Simulate** — Test posting without actually posting (validates data)
6. **Transfer** — Execute actual posting in target system
7. **Result Link Update** — Enrich worklist with target document numbers
8. **Undo** — Reverse posted documents if needed

### Selection Methods

| Method | Description |
|---|---|
| **Restricted by Source Header Table (DMSEL)** | Default method, selection criteria for header table only, AND-linked, single RFC call |
| **Restricted by Source Header Table by Portions** | Supports AND/OR linking, parallelizable, uses MWB migration object |
| **No Restrictions** | Selection criteria for all data model tables, complex SQL-like conditions, uses OBT |
| **Selection Control** | Enhanced method (from PECON Package 13+), designed for complex projects with multiple selections |

### Key Data Model Tables per Area

**Vendor Open Items (PECON_AP_OI):**
- Header: BSIK
- Daughter tables: BKPF, BSEG, BSEC, etc.
- Extension table: EXTTAB001 (additional document data)
- Note: BSAK (cleared items) is NOT part of data model — only open items can be selected

**Customer Open Items (PECON_AR_OI):**
- Header: BSID
- Similar structure to AP_OI

**G/L Open Items (PECON_GL_OI):**
- Header: BSIS

**FI Totals:**
- Uses pre-delivered table `/SLO/PECON_FI_TT`
- Standard fields: company code, account, profit center, business area, segment
- Extendable via SE11 Append Structure + Custom Selection Class

### Worklist Monitor Functions

| Function | Description |
|---|---|
| **Build** | Apply transformations, create transfer list |
| **Simulate** | Test posting without committing |
| **Transfer** | Execute actual posting in target |
| **Undo** | Reverse posted documents |
| **Result Link Update** | Enrich worklist with target document numbers |

### Mass Processing (Job Scheduling)

- **Plan Jobs** — Define job plans for batch processing
- **Schedule Jobs** — Fast scheduling function
- **Job Release Management** — Control job execution
- **Processing Control Cockpit** — Monitor and manage processing
- **Job Monitor** — Track job status

### Important PE Notes

- **Logon language must be ENGLISH** — descriptions and functions may not work in other languages
- **RFBIBL is deprecated** in S/4 releases >= OP1809 SPS5, OP1909 SPS3, OP2020 — use ACDOCGEN execution rules instead
- **Retention period** — Default 365 days from S/4HANA OP2022; delete if PE project should not auto-expire
- **Test Mode vs Productive Mode** — Test mode prevents actual transfers; switch per user
- **Compatibility** — Projects from lower S/4 releases may be "Read Only" in OP2023+ (SAP Note 3321187)
- **Generation is required** after any configuration change before execution
- **Tasks must be generated** before DDIC objects, function modules, or worklist structures become available

### Common FAQ / Troubleshooting

| Error | Cause | Solution |
|---|---|---|
| `CNV_OT_APPL_PE053` — T_ACCIT import failed | Inconsistent data format | Regenerate DDIC objects, rebuild transfer list |
| `CNV_PE451` — Transformation rule WAERS error | Currency settings not read | "Read Customizing" at area level with "overwrite" |
| `RW 033` — Balance in Transaction Currency | Document amounts don't balance to zero | Check offset lines, EXTTAB entries, ACCIT-BSCHL/SHKZG signs |
| `CNV_PE205` — Reference document not posted | Reference doc (REBZG) must be posted first | Post reference documents before dependent items |
| `F5 266` — One-time account name/city | Missing client in T_ACCFI | Fill T_ACCFI-MANDT in transformation |
| `CNV_PE_GEN131` — SIM generation error | Bug in basis component | Install SAP Note 3309487 |
| `CNV_PE_IF_RT 053` — POST generation missing params | Missing interface parameter mapping | Maintain worklist result links in CNV_PE_EDITOR |
| `TR 771` — Transport creation error | Missing authorization or wrong transport owner | Check SU53, change transport owner |
| Function group not generated | Missing generation step | Run `CNV_PE_GEN_RTO_AREA_FUGR` at area level |
| Syntax error `_GC_PE_WL1_ID` unknown | Incorrectly generated function group | Re-run function group generation with "Delete Object, prior to creation" |
| RFC Callback Whitelist error | Missing whitelist entries for RFC callbacks | Add `DD_GET_UCLEN` and `RFC_GET_STRUCTURE_DEFINITION` to SM59 |
| Missing number range | Number range not maintained in target | Complete number range in customizing |

### Key SAP Notes for PE

- **3321187** — Cannot edit projects from lower releases in OP2023+
- **3246235 / 3226942** — Fix for "Adding/updating global include" error (Allowance List bug)
- **3309487** — Fix for SIM function module generation error
- **3293908 / 3325830** — Fix for SIM_M generation error (S/4 basis issue)
- **3447984** — Fix for table type assignment bug in OP2022
- **3147442** — Translation/terminology reference for PE UI terms
- **2572945** — DMIS compatibility with S/4HANA

### Key PE Reports

| Report | Purpose |
|---|---|
| `/SLO/PECON_SCHEDULE_FI_OI_SEL` | Schedule FI open item selection |
| `/SLO/PECON_SELECT_FI_TT` | Selection program for FI Totals |
| `/SLO/PECON_ADD_WL_FIELDS` | Add fields to worklist structure |
| `/SLO/PECON_ADD_DM_STRUCT` | Add fields to data model structure |
| `/SLO/PECON_ADD_SCN_DEF` | Create scenario definition entries (for OP2022 table changes) |
| `CNV_PE_SUPPORT_PROJ_GEN_TRULES` | Generate all transformation rules for project |
| `CNV_PE_SUPPORT_IF_MAP_MAINT` | Maintain interface parameter mappings |

## Important Notes

- **No tests** — no test framework is configured
- **No requirements.txt** — dependencies managed in venv directly
- **Single-file app** — main agent logic lives in `sap_gui_scripting_ai.py`
- **Sensitive credentials** in `.env` — never commit this file
- SAP session must already be open before running the agent
