# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Purpose

Desktop application (future web) to automate generation of qualification Excel reports for biomedical equipment (stability chambers). Reduces manual time filling a 10-sheet Excel template from MadgeTech sensor data.

## Commands

```bash
# Install dependencies
pip install -r requirements.txt

# Run the application
python -m app.main

# Run tests
python -m pytest tests/

# Run a single test
python -m pytest tests/test_parser_primarios.py::test_name -v

# Build .exe
pyinstaller --onefile --windowed --name CalificadorIA app/main.py
```

## Architecture

### Data Flow
```
[primarios.xlsx]  ──┐
                    ├─► core/parser_*.py ──► core/data_cleaner.py ──► core/data_simulator.py ──► core/excel_writer.py ──► [plantilla_llena.xlsx]
[fallas.xlsx]     ──┘
```

### Layer Responsibilities

**`app/core/`** — Business logic, no GUI dependencies. Safe to reuse in Flask.
- `parser_primarios.py`: Reads MadgeTech MultiChannel Excel. Detects sensor columns by Serial (row 3), merges :00/:02 duplicate rows into single 5-min intervals, filters to user-specified 24h window.
- `parser_fallas.py`: Reads fallas Excel. Segments by test type (PO: 30min corte + 30min recovery | PP: 5min door + 30min recovery) using user-provided start time.
- `data_cleaner.py`: Validates values against equipment range, corrects out-of-range outliers via smooth interpolation (equipment doesn't produce sudden jumps).
- `data_simulator.py`: Generates synthetic data for missing sensors, mirroring trend of present sensors, constrained within range.
- `excel_writer.py`: Fills the 10-sheet template using openpyxl. **Only writes to**: Primarios rows 7-295 (cols A-K, M-U), Fallas Tem rows 7-66, Fallas HR rows 7-66, and Análisis cells C2-C7 + A39. Never overwrites formula cells.

**`app/models/`** — Data models.
- `sensor.py`: Sensor dataclass (serial, model, position, has_temp, has_humidity, data DataFrame).
- `proyecto.py`: Project config (equipment range, setpoints, client info, test type, start datetime).

**`app/gui/`** — customtkinter UI. All user inputs collected before processing.
- `main_window.py`: Step-by-step wizard: load files → configure sensors → set ranges/dates → process → save.

### Critical Rules for Excel Template
- Sheets **T, HR, GT, GHR, Grafico Fallas Tem, Grafico Fallas HR** are formula/chart sheets — **never touch them**.
- Primarios sheet: rows 1-5 = sensor metadata, row 6 = headers (do not overwrite). Data starts row 7.
- 289 samples exactly in Primarios (rows 7-295), 60 samples in each Fallas sheet (rows 7-66).
- Column layout in Primarios: A=sample#, B=datetime, C-K=temperature (9 sensors), M-U=humidity (9 sensors). Column L is intentionally skipped.

### MadgeTech MultiChannel Excel Format
- Row 1: Device name, Row 3: Serial number (sensor identifier), Row 7: channel headers.
- Columns come in pairs per sensor: "Canal 1 Temperatura (°C)" + "Canal 2 Humedad (% RH)".
- RTD sensors (93410-5) have temperature only; RHTemp101A have both temp and humidity.
- Rows at :00 and :02 seconds within the same 5-min interval belong to different channels — merge by grouping on the 5-minute timestamp floor.

### Sensor Configuration (User Input Required)
Always ask the user before processing:
1. How many sensors and which serials/references are expected
2. Equipment temperature and humidity range (for validation and simulation)
3. Start date and time of the 24h qualification window
4. Test type for fallas file: PO or PP
5. Start time of the fallas test (not marked in the file)
6. Client info: company, brand, location, equipment code, setpoints

## Stack
- **pandas** — data manipulation and 5-min interval grouping
- **openpyxl** — reading/writing Excel while preserving formulas and charts (`keep_vba=False`, load with `data_only=False` to preserve formulas)
- **customtkinter** — GUI
- **PyInstaller** — packaging to .exe
