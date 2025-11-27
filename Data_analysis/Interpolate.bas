Attribute VB_Name = "Interpolate"

Public Function InterpolateP(ByRef Texp As Double, ByRef Pexp As Double, ByRef TCal1 As Double, ByRef TCal2 As Double, ByRef PCal1 As Range, ByRef PCal2 As Range, ByRef CorrP1 As Range, ByRef CorrP2 As Range) As Double

'Compiled July 2020

    'Registers a function within the MS Excel environment that can be used to determine the appropriate pressure correction through a double interpolation between two specified pressure calibration temperatures.

    'To use:
    '   Copy all code into a new module in Visual Basic editor
    '   "Run" from within the editor only once (F5)
    '   A message dialog box should appear, indicating that the function has been registered successfully
    '   Save the Excel macro file, e.g. as "Interpolate.xlsm"
    '   The function is now ready to be used
    '   Use the function in any other spreadsheet by keeping "Interpolate.xlsm" open, and navigating to the function "Interpolate.xlsm!InterpolateP" in the Engineering functions category

    'Description of input arguments in function:
    '  Texp - Measured experimental temperature, in degrees C
    '  Pexp - Measured experimental pressure, in bar
    '  TCal1 - Lower calibration temperature, in degrees C
    '  TCal2 - Upper calibration temperature, in degrees C
    '  PCal1 - Array of pressures measured during calibration at TCal1, in bar
    '  PCal2 - Array of pressures measured during calibration at TCal2, in bar
    '  CorrP1 - Array of pressure corrections from calibration at TCal1, in bar
    '  CorrP2 - Array of pressure corrections from calibration at TCal2, in bar
    
    Dim iRow1 As Long     'loop counter to store the position of a search term
    Dim iRow2 As Long     'loop counter to store the position of a search term
    Dim LowV1 As Double   'lower value to be used in interpolating at TCal1
    Dim UpV1 As Double    'upper value to be used in interpolating at TCal1
    Dim LowV2 As Double   'lower value to be used in interpolating at TCal2
    Dim UpV2 As Double    'upper value to be used in interpolating at TCal2
    Dim Fslope1 As Double 'fraction corresponding to slope in PCal1 vs. CorrP1
    Dim Fslope2 As Double 'fraction corresponding to slope in PCal2 vs. CorrP2
    Dim FslopeT As Double 'fraction corresponding to slope in {TCal1, TCal2} vs. {CorrPT1, CorrPT2}
    Dim CorrPT1 As Double 'pressure correction interpolated at TCal1
    Dim CorrPT2 As Double 'pressure correction interpolated at TCal2
    
    'finds the position of the calibration pressure closest to the Pexp at TCal1
    iRow1 = 1
    Do While Pexp >= PCal1(iRow1, 1).Value And (iRow1 < PCal1.Rows.Count)
        iRow1 = iRow1 + 1
    Loop
    
    'finds the position of the calibration pressure closest to the Pexp at TCal2
    iRow2 = 1
    Do While Pexp >= PCal2(iRow2, 1).Value And (iRow2 < PCal2.Rows.Count)
        iRow2 = iRow2 + 1
    Loop
        
    'if Pexp falls outside either PCal1 or PCal2, the function result is an error value
    If (Pexp < PCal1(1, 1).Value) Or (Pexp > PCal1(iRow1, 1).Value) Or (Pexp < PCal2(1, 1).Value) Or (Pexp > PCal2(iRow2, 1).Value) Then
        MsgBox "The experimental pressure selected falls outside the calibration range!", vbInformation, "Interpolation error"
        InterpolateP = (1 / 0)
        Exit Function
    End If
    
    'retrieves the 4 data points at TCal1 and TCal2 to use in interpolation calculations
    LowV1 = CorrP1.Cells(iRow1 - 1, 1).Value
    UpV1 = CorrP1.Cells(iRow1, 1).Value
    LowV2 = CorrP2.Cells(iRow2 - 1, 1).Value
    UpV2 = CorrP2.Cells(iRow2, 1).Value

    'lever rule type of calculations to get the fractions corresponding to Pexp at TCal1 and TCal2, and Texp
    Fslope1 = (Pexp - PCal1.Cells(iRow1 - 1, 1).Value) / (PCal1.Cells(iRow1, 1).Value - PCal1.Cells(iRow1 - 1, 1).Value)
    Fslope2 = (Pexp - PCal2.Cells(iRow2 - 1, 1).Value) / (PCal2.Cells(iRow2, 1).Value - PCal2.Cells(iRow2 - 1, 1).Value)
    FslopeT = (Texp - TCal1) / (TCal2 - TCal1)

    '2 intermediate results after interpolation on Pexp at TCal1 and TCal2
    CorrPT1 = LowV1 + Fslope1 * (UpV1 - LowV1)
    CorrPT2 = LowV2 + Fslope2 * (UpV2 - LowV2)
    
    'final result after interpolation at Texp
    InterpolateP = CorrPT1 + FslopeT * (CorrPT2 - CorrPT1)
    
End Function

Sub RegisterFunction()

    'Registers the description of the interpolation function and its arguments in Excel so they appear as tooltips
    
    Dim FuncName As String
    Dim FuncDesc As String
    Dim FuncCat As Variant
    
    'The function has eight arguments, so eight variables are declared
    Dim ArgDesc(1 To 8) As String
    
    FuncName = "InterpolateP"
    
    FuncDesc = "Determines the appropriate pressure correction through a double interpolation between two specified pressure calibration temperatures."
    
    FuncCat = "Engineering" 'The function is moved from the User Defined category to Engineering
    
    ArgDesc(1) = "Measured experimental temperature, in degrees C"
    ArgDesc(2) = "Measured experimental pressure, in bar"
    ArgDesc(3) = "Lower calibration temperature, in degrees C"
    ArgDesc(4) = "Upper calibration temperature, in degrees C"
    ArgDesc(5) = "Array of pressures measured during calibration at TCal1, in bar"
    ArgDesc(6) = "Array of pressures measured during calibration at TCal2, in bar"
    ArgDesc(7) = "Array of pressure corrections from calibration at TCal1, in bar"
    ArgDesc(8) = "Array of pressure corrections from calibration at TCal2, in bar"

    Application.MacroOptions _
        Macro:=FuncName, _
        Description:=FuncDesc, _
        Category:=FuncCat, _
        ArgumentDescriptions:=ArgDesc
    
    'Inform the user about the process.
    MsgBox FuncName & " was successfully added to the " & FuncCat & " category!", vbInformation, "Done"
    
End Sub
