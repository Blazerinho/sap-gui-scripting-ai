# SAP GUI Scripting AI Agent

AI-powered SAP GUI automation agent using LangGraph + LangChain with SAP Generative AI Hub to interpret natural language commands and perform SAP GUI operations via Windows COM automation.

## Overview

This project automates SAP GUI operations for FI/CO analysis tasks. The user receives analysis requests via email, runs SAP queries, and replies with results. The agent handles this conversationally using natural language.

## Tech Stack

- **Language:** Python 3.14.0
- **AI Framework:** LangGraph 1.0.7, LangChain 1.2.7
- **LLM Provider:** SAP AI Core (Generative AI Hub) - gpt-4o
- **SAP GUI Access:** pywin32 (COM automation)
- **Config:** python-dotenv (.env file)

## Project Structure

```
sap_gui_scripting_ai.py   # Main AI agent application
sap_scripting.py           # Reusable SapSession & OutlookSession classes
create_excel.py            # Posting Engine error analysis Excel generator
.env                       # SAP AI Core credentials (NOT committed)
venv/                      # Python virtual environment
```

## Scripts

### create_excel.py - Posting Engine Error Analysis

Generates a formatted Excel report analyzing common Posting Engine errors from SAP DMLT projects.

**Features:**
- Analyzes 10 most common PE errors with root cause analysis
- Provides actionable solutions and SAP Note references
- Creates summary sheet with priority order for resolution
- Color-coded severity levels (High/Medium/Low)
- Professional formatting for sharing with project teams

**Usage:**
```bash
python create_excel.py
```

**Output:**
- `PE_Error_Analysis_ZFI_SDT.xlsx` - Detailed error analysis workbook
- Two sheets: "Error Analysis" (detailed) and "Summary" (prioritized action plan)

**Common Errors Analyzed:**
- NR 751 - Number range intervals missing
- CNV_PE 451 - Transformation rule failures
- CNV_OT_CX 000 - Generic exceptions
- F5A 002/003 - Master data deletion flags
- CNV_PE 205 - Reference document dependencies
- And more...

## Setup & Prerequisites

### Windows Requirements
- Windows OS (COM dependency)
- SAP GUI for Windows with scripting enabled (Alt+F12 → Options → Scripting)
- Server parameter: `sapgui/user_scripting = TRUE` (set via RZ11)
- Microsoft Outlook running locally (for email integration)

### Python Setup
```bash
# Create virtual environment
python -m venv venv

# Activate virtual environment
venv\Scripts\activate

# Install dependencies
pip install pywin32 langchain langgraph generative-ai-hub-sdk[langchain] python-dotenv openpyxl
```

### Configuration
Create a `.env` file with your SAP AI Core credentials:
```
# SAP AI Core credentials (never commit this file!)
SAP_AI_CORE_URL=https://your-instance.com
SAP_AI_CORE_CLIENT_ID=your-client-id
SAP_AI_CORE_CLIENT_SECRET=your-secret
```

## Running the Agent

```bash
# Ensure SAP GUI is open with an active session
venv\Scripts\activate
python sap_gui_scripting_ai.py
```

## SAP Posting Engine Knowledge

This repository includes expertise on SAP Posting Engine (DMLT /SLO/PECON) for migrating FI/CO open items and balances between SAP systems.

**Key Transactions:**
- `/SLO/PECON` - Main Posting Engine entry point
- `CNV_PE_PROJ` - Project maintenance
- `CNV_PE_EDITOR` - Transfer method configuration

**Documentation:** See `Posting Engine Knowledge/` folder for PDFs and guides.

## Common SAP Tables

- **BSIK** - Vendor open items
- **BSID** - Customer open items
- **BSIS** - G/L account open items
- **EKKO/EKPO** - Purchase orders
- **VBAK/VBAP** - Sales orders

## Contributing

This is a personal project for SAP FI/CO automation. Feel free to fork and adapt for your own needs.

## License

Internal use - SAP proprietary tools and access required.

## Author

Herbert Fromm (herbert.fromm@sap.com)
