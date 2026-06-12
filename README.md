# H-AnodPlot

**GUI tool for Anodization and Electrochemical Data Visualization**

![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python)
![License](https://img.shields.io/badge/License-MIT-green)
![Version](https://img.shields.io/badge/Version-2.5-informational)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey?logo=windows)

H-AnodPlot is a Python desktop application for visualizing and analyzing anodization experiment data. It is part of the **H-SciTools** scientific software suite, developed to support materials engineering research.

---

## Features

**Data loading**
- Import `.dat` files exported by potentiostats (whitespace-delimited format)
- Automatic stripping of inline comments (lines and inline segments starting with `*`)
- Multiple samples loaded simultaneously

**Data processing**
- Electrode area input per sample (cm²)
- Automatic normalization of raw current (mA) to current density (mA/cm²)
- Peak and final value extraction for potential and current density

**Style and appearance**
- Color, line width, and line style configurable per sample
- Sequential automatic color assignment (matplotlib tab10 palette)
- Grid toggle (on/off)

**Graph interactivity**
- Embedded matplotlib canvas directly in the main window
- Double-click on plot titles, axis labels, or legend entries to edit them inline
- Sample name editing from both the sidebar and the legend

**Export**
- PNG (300 DPI, white background, publication-ready)
- PDF report with embedded charts and per-sample statistics (via ReportLab)
- Excel `.xlsx` with one sheet per sample and a summary tab with peak/final values

---

## Interface

The layout is split into a scrollable control panel (left) and the plot area (right):

```
┌────────────────────┬──────────────────────────────────────────┐
│  H-AnodPlot        │                                          │
│  ─────────────     │                                          │
│  1. Files          │           Plot area                      │
│  2. Output folder  │       (interactive matplotlib)           │
│  3. Electrode areas│                                          │
│  4. Process data   │                                          │
│  ─────────────     │                                          │
│  Visualization     │                                          │
│  Style controls    │                                          │
│  Label editing     │                                          │
│  Export            │                                          │
└────────────────────┴──────────────────────────────────────────┘
```

---

## Requirements

Python 3.10 or higher.

```
numpy
matplotlib
openpyxl
reportlab
```

Install dependencies:

```bash
pip install numpy matplotlib openpyxl reportlab
```

> `tkinter` is included in standard Python distributions. If missing on Linux, install via `sudo apt install python3-tk`.

---

## Usage

### Running directly

```bash
python H-AnodPlot_GUI_v2_5.py
```

### Packaging with PyInstaller

```bash
pyinstaller --onefile --windowed --icon=AnodPlot.ico H-AnodPlot_GUI_v2_5.py
```

The resulting executable in `dist/` can be distributed without a Python installation.

> The `AnodPlot.ico` icon file is optional. If not found, the application starts normally without an icon.

### Workflow

1. **Select files** — choose one or more `.dat` files from your potentiostat
2. **Set output folder** — define where exports will be saved
3. **Enter electrode areas** — input the exposed area (cm²) for each sample
4. **Process data** — applies area normalization and prepares datasets for plotting
5. **Visualize** — renders Potential vs. Time and Current Density vs. Time side by side
6. **Customize & Export** — adjust per-sample styles, edit labels inline, and export PNG / PDF / Excel

---

## Supported file format

H-AnodPlot reads `.dat` files with the following column layout:

```
0.0   ...  0.452  ...  12.34  ...
0.5   ...  0.467  ...  12.21  ...
...
```

| Column index | Physical quantity |
|---|---|
| 0 | Time (s) |
| 2 | Potential (V) |
| 4 | Current (mA) |

Inline comments after `*` are stripped automatically. Lines that cannot be parsed are silently skipped.

---

## Code structure

```
H-AnodPlot_GUI_v2_5.py
│
├── resource_path()         # PyInstaller-compatible asset resolver
├── Experimento             # Per-sample data model (time, potential, density, style)
├── AnodPlot                # Core model: loading, normalization, plotting, export
│   ├── carregar_arquivos() # Registers .dat files as Experimento instances
│   ├── aplicar_areas()     # Parses files and computes current density
│   ├── plotar()            # Standalone matplotlib window with editable labels
│   ├── gerar_imagens()     # Renders and saves PNG charts
│   ├── exportar_excel()    # Writes .xlsx workbook with summary sheet
│   └── gerar_pdf()         # Builds PDF report via ReportLab
├── GraficoEmbutidoAnod     # Embedded matplotlib canvas with interactive editing
└── App                     # Main tkinter window and pipeline UI
```

---

## Part of the H-SciTools suite

| Tool | Purpose |
|---|---|
| H-AnodPlot | Electrochemical anodization curves |
| H-TGAPlot  | Thermogravimetric Analysis (TGA) |
| H-DMAPlot  | Dynamic Mechanical Analysis (DMA) |
| H-DRXPlot  | X-Ray Diffraction (XRD) |

---

## Author

**Carlos Henrique Amaro da Silva**  
M.Sc. in Materials Technology and Industrial Processes — Universidade Feevale (2025)  
B.Sc. in Chemical Engineering (2023)

Research focus: surface treatments, anodization, and electrodeposition with biomedical applications.

GitHub: https://github.com/Leindsher

---

## License

This project is licensed under the [MIT License](LICENSE).
