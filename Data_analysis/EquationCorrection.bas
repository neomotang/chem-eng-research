Attribute VB_Name = "EquationCorrection"

Sub CorrectionFunction()

'Compiled November 2022

'This sub adapts UpdateFormulas.bas (https://github.com/neomotang/vba-analyses/blob/main/UpdateFormulas.bas) to quickly correct equations across multiple cells in a spreadsheet, over different tabs. 
'Very useful if e.g. the form of an equation that has already been applied to multiple cells now needs to be updated, or if constant variables are stored in a different tab and they need to be included in the equations.
'i.e. no need to manually check if things are consistent!

'In the example below, experimental vapour-liquid equilibrium data for a chemical system consisting of 1-dodecanol and carbon dioxide was considered.
'Temperature, pressure, volume and frequency data were collected at various compositions. Data for individual compositions were saved in separate tabs.
'The code updates the pressure, temperature, volume and frequency calibration formulas, referencing calibration factors that are stored in a separate tab ("PTVfCalibration").
'The format of each data tab is as follows:
'  Columns A-E - contains notes, initial mass, etc. for the experiment at the current composition
'    The total mass of material is saved in cell D16, to be used in density calculations
'  Columns F-M - columns for various experimental measurements and calibration corrections
'    Row 1 has headings
'    The rest of the rows contain blocks of data, with measurements at the same temperature forming a block. Up to 5 temperatures were used at each composition.
'    In the last row of each temperature block, column F is left blank and average values for the preceding rows are entered in the other columns.
'    A blank row appears before the next data block, without repeating column headings.
    
    Workbooks("1 Dodecanol.xlsx").Activate 'for correcting 1-dodecanol data
    
    SheetID = 2 'First sheet with VLE data, where the Sheet 1 is reserved for notes
    SheetTotal = Sheets.Count

    'Retrieve the corrected temperature values for each of the 5 setpoint temperatures (35 - 75 degrees C)
    T35 = Sheets("PTVfCalibration").Range("E3").Value
    T45 = Sheets("PTVfCalibration").Range("I3").Value
    T55 = Sheets("PTVfCalibration").Range("M3").Value
    T65 = Sheets("PTVfCalibration").Range("Q3").Value
    T75 = Sheets("PTVfCalibration").Range("U3").Value
   
    Do While SheetID <= SheetTotal - 5 'because the last few sheets have summaries, graphs, comparison to literature, etc.
        Sheets(SheetID).Activate
                
        Sheets(SheetID).Range("F1").Offset(1, 0).Select 'find the beginning of the block of data, assuming columns A - E have notes about the current experiment and measured data is entered from column F, with headings in row 1
        ff = ActiveCell.Value 'intialize flag that should be raised if the end is reached, signified by an empty cell
        
        'loop to change formulas on the current tab
        Do While Not IsEmpty(ff)
            RowID = ActiveCell.Row 'gets the row ID of the beginning of the data block
            ColID = ActiveCell.Column 'gets the column ID of the beginning of the data block
            RootNum = ActiveCell.Value 'reads the value from column F, which is experimental temperature
            
            Do While Not IsEmpty(RootNum) 'loop condition to check if the cell in column F is empty or not; run only if not empty
                
                Sheets(SheetID).Cells(RowID, ColID + 1).Value = "=F" & RowID & "-PTVfCalibration!$G$28-PTVfCalibration!$H$28*F" & RowID 'Temperature correction formula is entered in column G
                
                'Pressure correction formula is entered in column I, calling on the custom function from Interpolate.bas to perform a double-interpolation
                Select Case Sheets(SheetID).Cells(RowID, ColID + 1).Value 'Use the corrected temperature in column G to determine the appropriate limits for the double-interpolation
                    Case Is < T35: Sheets(SheetID).Cells(RowID, ColID + 3).Value = "=H" & RowID & "+InterpolateP(G" & RowID & ",H" & RowID & ",PTVfCalibration!$E$3,PTVfCalibration!$I$3,PTVfCalibration!$B$5:$B$24,PTVfCalibration!$F$5:$F$24,PTVfCalibration!$C$5:$C$24,PTVfCalibration!$G$5:$G$24)+1.01325" 'P correction, for 35 - 45 dC interval, extrapolated for values lower than 35 dC
                    Case Is < T45: Sheets(SheetID).Cells(RowID, ColID + 3).Value = "=H" & RowID & "+InterpolateP(G" & RowID & ",H" & RowID & ",PTVfCalibration!$E$3,PTVfCalibration!$I$3,PTVfCalibration!$B$5:$B$24,PTVfCalibration!$F$5:$F$24,PTVfCalibration!$C$5:$C$24,PTVfCalibration!$G$5:$G$24)+1.01325" 'P correction, for 35 - 45 dC interval
                    Case Is < T55: Sheets(SheetID).Cells(RowID, ColID + 3).Value = "=H" & RowID & "+InterpolateP(G" & RowID & ",H" & RowID & ",PTVfCalibration!$I$3,PTVfCalibration!$M$3,PTVfCalibration!$F$5:$F$24,PTVfCalibration!$J$5:$J$24,PTVfCalibration!$G$5:$G$24,PTVfCalibration!$K$5:$K$24)+1.01325" 'P correction, for 45 - 55 dC interval
                    Case Is < T65: Sheets(SheetID).Cells(RowID, ColID + 3).Value = "=H" & RowID & "+InterpolateP(G" & RowID & ",H" & RowID & ",PTVfCalibration!$M$3,PTVfCalibration!$Q$3,PTVfCalibration!$J$5:$J$24,PTVfCalibration!$N$5:$N$24,PTVfCalibration!$K$5:$K$24,PTVfCalibration!$O$5:$O$24)+1.01325" 'P correction, for 55 - 65 dC interval
                    Case Is < T75: Sheets(SheetID).Cells(RowID, ColID + 3).Value = "=H" & RowID & "+InterpolateP(G" & RowID & ",H" & RowID & ",PTVfCalibration!$Q$3,PTVfCalibration!$U$3,PTVfCalibration!$N$5:$N$24,PTVfCalibration!$R$5:$R$24,PTVfCalibration!$O$5:$O$24,PTVfCalibration!$S$5:$S$24)+1.01325" 'P correction, for 65 - 75 dC interval
                    Case Is > T75: Sheets(SheetID).Cells(RowID, ColID + 3).Value = "=H" & RowID & "+InterpolateP(G" & RowID & ",H" & RowID & ",PTVfCalibration!$Q$3,PTVfCalibration!$U$3,PTVfCalibration!$N$5:$N$24,PTVfCalibration!$R$5:$R$24,PTVfCalibration!$O$5:$O$24,PTVfCalibration!$S$5:$S$24)+1.01325" 'P correction, for 65 - 75 dC interval, extrapolated for values higher than 75 dC
                    End Select
                    
                Sheets(SheetID).Cells(RowID, ColID + 5).Value = "=$D$16/(PTVfCalibration!$D$32*(J" & RowID & ") + PTVfCalibration!$C$32)*1000" 'Volume correction and density calculation formula is entered in column K
                
                RowID = RowID + 1 'will move to the next row to check if empty in the next line of code and loop condition
                RootNum = Sheets(SheetID).Cells(RowID, ColID).Value
                Loop

            'The first row where column F is empty is reserved for calculating the average corrected temperature, pressure and density for the preceding rows, as well as calculating viscosity in column M
            varcase = WorksheetFunction.MRound(Sheets(SheetID).Cells(RowID, ColID + 1).Value - 5, 10) + 5 'rounds the average tempature to the closest multiple of 5, by first rounding to a multiple of 10 and then adding 5
            Select Case varcase 'Use the rounded corrected temperature to determine the appropriate frequency correction and viscosity calibration factors, to be entered in column M
                Case 35: Sheets(SheetID).Cells(RowID, ColID + 7).Value = "=(('PTVfCalibration'!$D$41-L" & RowID & ")/('PTVfCalibration'!$D$39+'PTVfCalibration'!$D$40*I" & RowID & "*10))^2/(PI()*L" & RowID & "*K" & RowID & ")" 'viscosity calculation for 35 dC
                Case 45: Sheets(SheetID).Cells(RowID, ColID + 7).Value = "=(('PTVfCalibration'!$E$41-L" & RowID & ")/('PTVfCalibration'!$E$39+'PTVfCalibration'!$E$40*I" & RowID & "*10))^2/(PI()*L" & RowID & "*K" & RowID & ")" 'viscosity calculation for 45 dC
                Case 55: Sheets(SheetID).Cells(RowID, ColID + 7).Value = "=(('PTVfCalibration'!$F$41-L" & RowID & ")/('PTVfCalibration'!$F$39+'PTVfCalibration'!$F$40*I" & RowID & "*10))^2/(PI()*L" & RowID & "*K" & RowID & ")" 'viscosity calculation for 55 dC
                Case 65: Sheets(SheetID).Cells(RowID, ColID + 7).Value = "=(('PTVfCalibration'!$G$41-L" & RowID & ")/('PTVfCalibration'!$G$39+'PTVfCalibration'!$G$40*I" & RowID & "*10))^2/(PI()*L" & RowID & "*K" & RowID & ")" 'viscosity calculation for 65 dC
                Case 75: Sheets(SheetID).Cells(RowID, ColID + 7).Value = "=(('PTVfCalibration'!$H$41-L" & RowID & ")/('PTVfCalibration'!$H$39+'PTVfCalibration'!$H$40*I" & RowID & "*10))^2/(PI()*L" & RowID & "*K" & RowID & ")" 'viscosity calculation for 75 dC
                Case Else: MsgBox "Error in viscosity due to temperature out of range"
                End Select
            
            ActiveCell.End(xlDown).Offset(0, 0).Select 'find the end of the current block of data
            ActiveCell.End(xlDown).Offset(0, 0).Select 'find the beginning of the next block of data (assuming no headings repeated) or end of column if all number blocks have been identified
            ff = ActiveCell.Value 'should be changed here based on the value of the selected cell
            
            Loop
        Sheets(SheetID).Range("A1").Select 'go to the top of the tab since analysis on the tab is complete
                
        SheetID = SheetID + 1 'go to the next tab of data
        Loop
    
    Sheets(2).Activate
    ActiveWorkbook.Save
    MsgBox "All done!"
    
End Sub
