# VBA Code Documentation

This document explains the VBA macros in `F1_Pit_Strategy_Calculator.xlsm`.

---

## Module: Strategy Calculator Macros

### Macro 1: CalculateStrategy

**Purpose:** Validates user inputs and triggers calculation

**Trigger:** "Calculate" button on Calculator sheet

**Logic:**
```vba
Sub CalculateStrategy()
    ' Get input values
    Dim pit1 As Integer
    Dim pit2 As Integer
    Dim totalLaps As Integer
    
    pit1 = Range("B13").Value
    pit2 = Range("B14").Value
    totalLaps = Range("B5").Value
    
    ' Validation checks
    If pit1 < 10 Then
        MsgBox "Pit 1 must be at least lap 10 (minimum stint length)", vbExclamation
        Exit Sub
    End If
    
    If pit2 <= pit1 Then
        MsgBox "Pit 2 must be after Pit 1", vbExclamation
        Exit Sub
    End If
    
    If pit2 > totalLaps - 5 Then
        MsgBox "Pit 2 must allow at least 5 laps in final stint", vbExclamation
        Exit Sub
    End If
    
    If pit2 - pit1 < 10 Then
        MsgBox "Stint 2 must be at least 10 laps", vbExclamation
        Exit Sub
    End If
    
    ' Force recalculation
    Application.Calculate
    
    ' Success message
    MsgBox "Strategy calculated successfully!" & vbCrLf & _
           "Total Race Time: " & Format(Range("B27").Value, "#,##0.00") & " seconds", _
           vbInformation
End Sub
```

**Validation Rules:**
1. First pit stop must be at least lap 10
2. Second pit stop must be after first pit stop
3. Final stint must have at least 5 laps
4. Each stint must be at least 10 laps

---

### Macro 2: ResetInputs

**Purpose:** Restores default values (Piastri's actual strategy)

**Trigger:** "Reset" button on Calculator sheet

**Logic:**
```vba
Sub ResetInputs()
    ' Reset compound selections to Piastri's actual strategy
    Range("B10").Value = "MEDIUM"    ' Stint 1
    Range("B11").Value = "HARD"      ' Stint 2
    Range("B12").Value = "MEDIUM"    ' Stint 3
    
    ' Reset pit stop laps
    Range("B13").Value = 18          ' Pit 1
    Range("B14").Value = 42          ' Pit 2
    
    ' Force recalculation
    Application.Calculate
    
    ' Confirmation message
    MsgBox "Inputs reset to Piastri's actual Hungary 2023 strategy:" & vbCrLf & _
           "MEDIUM → HARD → MEDIUM" & vbCrLf & _
           "Pit stops: Lap 18, Lap 42", vbInformation
End Sub
```

---

## How to Access VBA Code

1. Open the Excel file
2. Press **Alt + F11** to open the VBA Editor
3. In the Project Explorer, expand **VBAProject (F1_Pit_Strategy_Calculator.xlsm)**
4. Double-click **Module1** to view the code

---

## How to Edit VBA Code

1. Open VBA Editor (Alt + F11)
2. Modify the code in Module1
3. Press **Ctrl + S** to save
4. Close the VBA Editor
5. Save the workbook as `.xlsm` (macro-enabled)

---

## Button Assignment

| Button | Macro | Location |
|--------|-------|----------|
| Calculate | `CalculateStrategy` | Calculator sheet, row 2 |
| Reset | `ResetInputs` | Calculator sheet, row 5 |

To reassign a button:
1. Right-click the button
2. Select **Assign Macro**
3. Choose the macro from the list
4. Click OK

---

## Error Handling

The macros include basic error handling:
- Invalid pit lap inputs trigger warning messages
- User must acknowledge before continuing
- No data is changed if validation fails

---

*This VBA implementation demonstrates practical Excel automation skills used in F1 strategy engineering departments.*

