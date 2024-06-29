# Excel Formula Dragging Automation with VBA

Automate the formula dragging across columns in Excel with this VBA script.

## Description

This repository contains a VBA script (`DragFormulas`) designed to automate the dragging of formulas from column D to column E in an Excel worksheet named "Ref." This script fills down formulas from cells D3:E3 to the last populated row in column C.

## Usage

To use this script:

1. **Download the VBA Script**: Copy the `DragFormulas` subroutine code from this repository.

2. **Open Your Excel Workbook**:
   - Press `Alt + F11` to open the Visual Basic for Applications (VBA) editor.
   - Import the VBA script into your workbook by copying and pasting the code into a new module.

3. **Run the Script**:
   - Once imported, execute the `DragFormulas` subroutine to automate the formula dragging across columns.

## Example

```vba
Sub DragFormulas()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("SHEETNAME")
    
    ' Select range D3:E3 and fill down formulas
    With ws
        .Range("D3:E3").AutoFill Destination:=.Range("D3:E" & .Range("C" & .Rows.Count).End(xlUp).Row)
    End With
End Sub
