# F1 Pit Strategy Calculator

**Interactive Excel tool for Formula 1 pit stop strategy optimization**

![Excel](https://img.shields.io/badge/Excel-VBA%20Enabled-green.svg)
![F1](https://img.shields.io/badge/F1-Hungary%202023-red.svg)
![License](https://img.shields.io/badge/license-MIT-blue.svg)

---

## Overview

An Excel-based pit stop strategy calculator that models tyre degradation and compares race strategies for the 2023 Hungarian Grand Prix. Features VBA automation for input validation and one-click calculations.

**Use case:** Quick "what-if" analysis for F1 strategy decisions without needing Python or coding knowledge.

---

## Features

### Interactive Calculator
- **Dropdown menus** for tyre compound selection (Soft, Medium, Hard)
- **Editable pit stop laps** for custom strategy testing
- **Automatic calculations** update in real-time
- **VBA buttons** for validation and reset

### Strategy Comparison
- Compare your strategy against:
  - **Piastri's Actual** (M-H-M, Pit laps 18, 42)
  - **Optimal 2-Stop** (M-H-M, Pit laps 25, 47)
  - **1-Stop** (M-H, Pit lap 34)

### Sensitivity Analysis
- See how pit stop duration (20s, 22s, 24s, 26s) affects strategy choice
- Visual chart showing crossover points between strategies

### Built-in Documentation
- Assumptions clearly stated
- Data sources referenced
- Limitations acknowledged

---

## Screenshots

### Calculator Sheet
```
┌─────────────────────────────────────────────────────────────┐
│  F1 PIT STRATEGY CALCULATOR                                 │
├─────────────────────────────────────────────────────────────┤
│  RACE INPUTS                      │  QUICK SUMMARY          │
│  Track: Hungary                   │  Time Saved vs Piastri: │
│  Total Laps: 70                   │  7.24s                  │
│  Base Lap Time: 82.5s             │  Optimal Pit Lap: 34    │
│  Pit Stop Loss: 22s               │                         │
├─────────────────────────────────────────────────────────────┤
│  STRATEGY INPUTS                                            │
│  Stint 1 Compound: [MEDIUM ▼]                               │
│  Stint 2 Compound: [HARD ▼]                                 │
│  Stint 3 Compound: [MEDIUM ▼]                               │
│  Pit 1 After Lap:  18                                       │
│  Pit 2 After Lap:  42                                       │
├─────────────────────────────────────────────────────────────┤
│  TOTAL RACE TIME: 5857.96s                                  │
└─────────────────────────────────────────────────────────────┘
```

---

## How to Use

### Step 1: Enable Macros
When you open the file, Excel will ask about macros. Click **Enable Content** to allow VBA functionality.

### Step 2: Select Your Strategy
1. Use the **dropdown menus** (cells B10-B12) to select tyre compounds for each stint
2. Enter **pit stop laps** (cells B13-B14)

### Step 3: Calculate
- Click the **"Calculate"** button to validate inputs and update results
- Or just modify cells — formulas update automatically

### Step 4: Compare
- Check the **Strategy Comparison** table to see how your strategy ranks
- Review the **Sensitivity Analysis** chart for pit time impact

### Step 5: Reset
- Click **"Reset"** to restore Piastri's actual strategy (default values)

---

## Workbook Structure

| Sheet | Purpose |
|-------|---------|
| **Data** | Tyre compound parameters, race constants |
| **Calculator** | Main interactive interface |
| **Documentation** | User guide, assumptions, data sources |

---

## Tyre Data (Hungary 2023)

| Compound | Pace Offset | Deg Rate (s/lap) | Min Stint | Max Stint |
|----------|-------------|------------------|-----------|-----------|
| **SOFT** | -0.80s | 0.070 | 10 laps | 22 laps |
| **MEDIUM** | 0.00s | 0.040 | 15 laps | 32 laps |
| **HARD** | +0.40s | 0.025 | 20 laps | 45 laps |

*Data extracted from FastF1 telemetry using linear regression (see Python analysis)*

---

## VBA Macros

### CalculateStrategy
- Validates pit lap inputs (within race length, minimum stint lengths)
- Displays warning messages for invalid configurations
- Triggers recalculation of all formulas

### ResetInputs
- Restores default values (Piastri's actual strategy)
- Compounds: MEDIUM → HARD → MEDIUM
- Pit laps: 18, 42

---

## Key Formulas

### Average Lap Time per Stint
```
Avg Lap Time = Base Pace + Pace Offset + (Deg Rate × Stint Length / 2)
```

### Stint Time
```
Stint Time = Avg Lap Time × Stint Length
```

### Total Race Time
```
Total = Stint 1 Time + Stint 2 Time + Stint 3 Time + (Pit Stops × Pit Loss)
```

---

## Strategy Results

| Strategy | Stint 1 | Stint 2 | Stint 3 | Pit Laps | Total Time | Gap |
|----------|---------|---------|---------|----------|------------|-----|
| **1-Stop (M-H)** | MEDIUM | HARD | - | 34 | 5850.72s | — |
| Optimal 2-Stop | MEDIUM | HARD | MEDIUM | 25, 47 | 5856.93s | +6.21s |
| Piastri Actual | MEDIUM | HARD | MEDIUM | 18, 42 | 5857.96s | +7.24s |

**Finding:** 1-stop strategy is fastest in pure race time, but real races include Safety Cars and track position effects not captured here.

---

## Assumptions & Limitations

### Assumptions
- ✓ Linear tyre degradation (constant loss per lap)
- ✓ Base lap time: 82.5s (clean-air average)
- ✓ Pit stop loss: 22s (pit lane + stationary time)
- ✓ No fuel effect modeled

### Limitations
- ✗ No Safety Car or VSC events
- ✗ No track position / dirty air effects
- ✗ No non-linear degradation (tyre cliff)
- ✗ No temperature effects
- ✗ Single race only (Hungary 2023)

---

## Related Projects

| Project | Description |
|---------|-------------|
| [F1-Strategy-Quantitative-Investigation](https://github.com/egyan175/F1-Strategy-Quantitative-Investigation) | Full Python investigation and analysis with Monte Carlo simulation |
| [F1-Strategy-ML](https://github.com/egyan175/F1-Strategy-ML) | XGBoost lap time prediction (MAE 0.263s) |

---

## Getting Started

### Requirements
- Microsoft Excel 2016+ (Windows) or Excel 365
- Macros must be enabled for VBA functionality

### Download
1. Download `F1_Pit_Strategy_Calculator.xlsm`
2. Open in Excel
3. Click **Enable Content** when prompted
4. Start experimenting

---

## About

**Author:** Emmanuel Gyan  
**Education:** Final Year Aerospace Engineering, KNUST (Ghana)  
**Contact:** egyan175@gmail.com  


This tool complements the Python-based F1 strategy analysis by providing an accessible, no-code interface for strategy engineers and enthusiasts.

---

## License

MIT License — free to use and modify.

---

*This is an independent project for educational and portfolio purposes. Not affiliated with Formula 1, McLaren Racing, or any F1 team.*


