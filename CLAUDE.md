# CLAUDE.md - AI Assistant Guide for STIG-Control-CCI

## Project Overview

This repository provides tools for mapping NIST 800-53 security controls to Control Correlation Identifiers (CCIs) and organizing them by Defense Levels (DL-1 through DL-6). It generates formatted Excel reference sheets for compliance and security assessment workflows.

**Purpose**: Create easy reference sheets showing which controls are required at each defense level, with CCI mappings and counts for STIG (Security Technical Implementation Guide) compliance.

## Repository Structure

```
STIG-Control-CCI/
├── CLAUDE.md                           # This file - AI assistant guidance
├── generate_level_sheets.py            # Main Python script for generating Excel reports
├── level_data.json                     # Configurable input: control IDs by defense level
├── r4controls.json                     # NIST 800-53 Rev 4 controls data
├── r5controls.json                     # NIST 800-53 Rev 5 controls data (primary)
├── rev4cci.json                        # CCI mappings for Rev 4 controls
├── rev5cci.json                        # CCI mappings for Rev 5 controls (primary)
└── STIG_Control_Level_Reference.xlsx   # Generated output (Excel workbook)
```

## Data Files

### Control Files (`r4controls.json`, `r5controls.json`)
JSON arrays containing NIST 800-53 controls with:
- `Control Identifier`: e.g., "AC-01", "AC-02(01)"
- `Control (or Control Enhancement) Name`: Human-readable name
- `Control Text`: Full control requirements
- `Discussion`: Implementation guidance
- `Related Controls`: Cross-references

### CCI Files (`rev4cci.json`, `rev5cci.json`)
JSON arrays mapping controls to CCIs with:
- `Index`: Sub-control reference (e.g., "AC-1 a 1 (a)")
- `Control`: Control identifier (e.g., "AC-01")
- `CCI Number`: e.g., "CCI-000002"
- `Description`: CCI requirement description

### Level Data (`level_data.json`)
Configurable JSON mapping defense levels to control IDs:
```json
{
    "DL-1 DODIN": ["AT-01", "AT-02", ...],
    "DL-2 MCEN": ["AC-04", "AC-04(01)", ...],
    ...
}
```

## Key Conventions

### Control ID Formatting
**IMPORTANT**: All control IDs use double-digit format:
- Base controls: `AC-01`, `AT-02` (NOT `AC-1`, `AT-2`)
- Enhancements: `AC-02(01)`, `PE-02(03)` (NOT `AC-2(1)`, `PE-2(3)`)

The script automatically normalizes IDs via `normalize_control_id()` function.

### Defense Levels
| Level | Name | Description |
|-------|------|-------------|
| DL-1 | DODIN | DoD Information Network |
| DL-2 | MCEN | Mission Partner Environment Network |
| DL-3 | MITSC/IPN/ISN/Data Center | Infrastructure & Data Center |
| DL-4 | - | Physical/Environmental |
| DL-5 | System HW/SW/OS | Hardware, Software, Operating System |
| DL-6 | Application | Application Layer |

### Control Families
Common families in this context:
- `AC` - Access Control
- `AT` - Awareness and Training
- `AU` - Audit and Accountability
- `CM` - Configuration Management
- `IA` - Identification and Authentication
- `PE` - Physical and Environmental Protection
- `SC` - System and Communications Protection
- `SI` - System and Information Integrity

## Development Workflow

### Running the Script
```bash
# Basic usage (uses default data embedded in script)
python generate_level_sheets.py

# With custom level data JSON
python generate_level_sheets.py --input level_data.json

# With detailed CCI breakdown sheets
python generate_level_sheets.py --input level_data.json --detailed-cci

# Custom output path
python generate_level_sheets.py --output my_report.xlsx

# Use Rev 4 data instead of Rev 5
python generate_level_sheets.py --controls r4controls.json --cci rev4cci.json
```

### Dependencies
- Python 3.7+
- `pandas` - Data manipulation
- `openpyxl` - Excel file generation with charts

Install: `pip install pandas openpyxl`

### Input Formats
The script accepts level data in two formats:

1. **JSON** (recommended):
```json
{
    "Level Name": ["CTRL-01", "CTRL-02(01)", ...]
}
```

2. **CSV** (columns are level names, rows are controls):
```csv
DL-1 DODIN,DL-2 MCEN,DL-3,...
AT-01,AC-04,AC-19(04),...
```

## Output Structure

The generated Excel workbook contains:

1. **Summary Sheet** (first tab):
   - Level overview table (controls count, CCI totals, averages)
   - Bar chart: Controls per Level
   - Family breakdown table across all levels
   - Stacked bar chart: Control Families by Level
   - CCI count by family table

2. **Level Sheets** (one per defense level):
   - Control ID, Name, Text
   - CCI numbers (comma-separated)
   - CCI count per control
   - Control family

3. **Detailed CCI Sheets** (optional, with `--detailed-cci`):
   - One row per CCI mapping
   - Control ID, Name, CCI Number, CCI Description

## Common Tasks for AI Assistants

### Adding New Controls to a Level
1. Edit `level_data.json`
2. Add control IDs in double-digit format
3. Re-run the script

### Updating Control/CCI Data
Replace the JSON files (`r5controls.json`, `rev5cci.json`) with updated versions maintaining the same schema.

### Customizing Output
Modify `generate_level_sheets.py`:
- `create_level_sheet()` - Individual level sheet format
- `create_summary_sheet()` - Summary charts and tables
- `create_cci_detail_sheet()` - Detailed CCI breakdown

### Troubleshooting

**Control not found**: Verify double-digit formatting (AC-01 not AC-1)

**No CCIs mapped**: Some controls may not have CCI mappings in the data files

**Chart errors**: Ensure openpyxl is updated (`pip install --upgrade openpyxl`)

## Code Style

- Python 3.7+ compatible
- Type hints where practical
- Functions are documented with docstrings
- Constants at module level (e.g., `DEFAULT_LEVEL_DATA`)
- Use `pathlib.Path` for file operations

## Testing

Verify output by:
1. Running script with sample data
2. Opening generated Excel file
3. Checking control counts match input
4. Verifying CCI mappings are populated

## Version Control

- Commit generated `.xlsx` files only when intentionally sharing final reports
- Primary version-controlled files: `.py`, `.json` source files
- Use meaningful commit messages describing data or logic changes
