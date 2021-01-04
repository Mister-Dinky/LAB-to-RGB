VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LABtoRGB 
   Caption         =   "UserForm1"
   ClientHeight    =   3150
   ClientLeft      =   -180
   ClientTop       =   -585
   ClientWidth     =   5685
   OleObjectBlob   =   "LABtoRGB.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LABtoRGB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub CommandButton1_Click()
    'dim ranges, check if values are ranges
    Dim InputRng As String: InputRng = InputRange.value
    Dim OutputRng As String: OutputRng = OutputRange.value
    
    'Make the range check a boolean
    Dim IsRangeInput As Boolean
    Dim IsRangeOutput As Boolean
    
    'if it is not a range, make the boolean False
    On Error Resume Next
    IsRangeInput = IsObject(Range(InputRng))
    On Error GoTo 0
    
    'Check if the input range is not empty or invalid
    If InputRange.value = "" Then MsgBox "Input range is empty": Exit Sub
    If IsRangeInput = False Then MsgBox "Input range is invalid": Exit Sub
    
    'if it is not a range, make the boolean False
    On Error Resume Next
    IsRangeOutput = IsObject(Range(OutputRng))
    On Error GoTo 0
    
    'Check if the input range is not empty or invalid
    If OutputRange.value = "" Then MsgBox "Output range is empty": Exit Sub
    If IsRangeOutput = False Then MsgBox "Output range is invalid": Exit Sub
    
    'Run macro and close it
    LABtoRGBMacro
    Unload LABtoRGB
End Sub

Public Sub UserForm_Initialize()
    
    'Selecting the right radio box + combobox item
    StandardsOB2.value = True
    IlluminantCB1.AddItem "Illuminant D65 2°"
    IlluminantCB1.AddItem "Illuminant D65 10°"
    IlluminantCB1.ListIndex = 1
    RGBCB1.value = True
End Sub

Public Sub LABtoRGBMacro()
    'dim values
    Dim Illuminant As String
    Dim CRow As Integer
    Dim Xref As Double, Yref As Double, Zref As Double, x As Double, y As Double 'White Reference values
    Dim Epsilon As Double, Kappa As Double 'CIE standard values
    Dim L As Double, a As Double, b As Double 'Measurement data
    Dim fx As Double, fy As Double, fz As Double 'needed to determine Xr, Yr and Zr
    Dim Xr As Double, Yr As Double, Zr As Double 'When multiplied with Xref, Yref and Zref respectively will result in Xval, Yval and Zval
    Dim Xval As Double, Yval As Double, Zval As Double 'LAB data converted to XYZ
    Dim Mxr As Double, Myr As Double, Mxg As Double, Myg As Double, Mxb As Double, Myb As Double 'Values needed to create [XYZ] matrix values
    Dim MXred As Double, MYred As Double, MZred As Double, MXgreen As Double, MYgreen As Double, MZgreen As Double, MXblue As Double, MYblue As Double, MZblue As Double 'Matrix Value
    Dim XYZ(1 To 3, 1 To 3) As Variant, XYZw(1 To 3, 1 To 1) As Variant, iXYZ As Variant, Srgb As Variant, M(1 To 3, 1 To 3) As Variant, iM As Variant  'Create Matrices
    Dim XYZval(1 To 3, 1 To 1) As Variant, rgbt As Variant, rt As Double, gt As Double, bt As Double 'For turning our XYZ values into a matrix to multiply with M, giving the values r, g, b
    Dim Rf As Double, Gf As Double, Bf As Double
    Dim InpRange As Range
    Dim OutRange As Range
    
    Illuminant = LABtoRGB.IlluminantCB1.value
    Set InpRange = Range(LABtoRGB.InputRange.value)
    Set OutRange = Range(LABtoRGB.OutputRange.value)
    
    If LABtoRGB.IlluminantCB1.ListIndex = 0 Then 'get White Ref data D65 10°
        Xref = 0.95047
        Yref = 1
        Zref = 1.08883
        x = 0.31271
        y = 0.32902
    ElseIf LABtoRGB.IlluminantCB1.ListIndex = 1 Then
        Xref = 0.94811
        Yref = 1
        Zref = 1.07304
        x = 0.31382
        y = 0.331
    End If
    
    'Select correct Epsilon and Kappa
    If LABtoRGB.StandardsOB1 = True Then Epsilon = 0.008856: Kappa = 903.3
    If LABtoRGB.StandardsOB1 = True Then Epsilon = 216 / 24389: Kappa = 24389 / 27
    
    'Create correct matrix values
    Mxr = 0.64: Myr = 0.33: MXred = Mxr / Myr: MYred = 1: MZred = (1 - Mxr - Myr) / Myr
    Mxg = 0.3: Myg = 0.6: MXgreen = Mxg / Myg: MYgreen = 1: MZgreen = (1 - Mxg - Myg) / Myg
    Mxb = 0.15: Myb = 0.06: MXblue = Mxb / Myb: MYblue = 1: MZblue = (1 - Mxb - Myb) / Myb
    
    'fill in XYZ matrix
    XYZ(1, 1) = MXred: XYZ(1, 2) = MXgreen: XYZ(1, 3) = MXblue
    XYZ(2, 1) = MYred: XYZ(2, 2) = MYgreen: XYZ(2, 3) = MYblue
    XYZ(3, 1) = MZred: XYZ(3, 2) = MZgreen: XYZ(3, 3) = MZblue
    
    'fill in XYZw matrix
    XYZw(1, 1) = Xref
    XYZw(2, 1) = Yref
    XYZw(3, 1) = Zref
    
    'Create iXYZ and Srgb matrix
    iXYZ = WorksheetFunction.MInverse(XYZ)
    Srgb = WorksheetFunction.MMult(iXYZ, XYZw)
    
    'fill M matrix
    M(1, 1) = XYZ(1, 1) * Srgb(1, 1): M(1, 2) = XYZ(1, 2) * Srgb(2, 1): M(1, 3) = XYZ(1, 3) * Srgb(3, 1)
    M(2, 1) = XYZ(2, 1) * Srgb(1, 1): M(2, 2) = XYZ(2, 2) * Srgb(2, 1): M(2, 3) = XYZ(2, 3) * Srgb(3, 1)
    M(3, 1) = XYZ(3, 1) * Srgb(1, 1): M(3, 2) = XYZ(3, 2) * Srgb(2, 1): M(3, 3) = XYZ(3, 3) * Srgb(3, 1)
    
    'create iM Matrix
    iM = WorksheetFunction.MInverse(M)
    
    For CRow = 1 To InpRange.Rows.Count
        'Get LAB data
        L = InpRange.Cells(CRow, 1)
        a = InpRange.Cells(CRow, 2)
        b = InpRange.Cells(CRow, 3)
        
        'calculate fy, fx and fz
        fy = (L + 16) / 116
        fx = fy + a / 500
        fz = fy - b / 200
        
        'calculate Xr, Yr and Zr
        If L > Kappa * Epsilon Then Yr = ((L + 16) / 116) ^ 3 Else Yr = L / Kappa
        If fx ^ 3 > Epsilon Then Xr = fx ^ 3 Else Xr = (116 * fx - 16) / Kappa
        If fz ^ 3 > Epsilon Then Zr = fz ^ 3 Else Zr = (116 * fz - 16) / Kappa
        
        'Calculate X, Y and Z
        Xval = Xref * Xr
        Yval = Yref * Yr
        Zval = Zref * Zr
        
        'create XYZval matrix
        XYZval(1, 1) = Xval
        XYZval(2, 1) = Yval
        XYZval(3, 1) = Zval
        
        'get rgb matrix, get rt, gt, and bt
        rgbt = WorksheetFunction.MMult(iM, XYZval)
        rt = rgbt(1, 1)
        gt = rgbt(2, 1)
        bt = rgbt(3, 1)
        
        If rt < 0 Then rt = 0
        If gt < 0 Then gt = 0
        If bt < 0 Then bt = 0
        
        'Get R, G and B by converting rt, gt and bt
        
        If rt < 0.0031308 Then Rf = 12.92 * rt * 255 Else Rf = (1.055 * rt ^ (1 / 2.4) - 0.055) * 255
        If gt < 0.0031308 Then Gf = 12.92 * gt * 255 Else Gf = (1.055 * gt ^ (1 / 2.4) - 0.055) * 255
        If bt < 0.0031308 Then Bf = 12.92 * bt * 255 Else Bf = (1.055 * bt ^ (1 / 2.4) - 0.055) * 255
        
        If LABtoRGB.RGBCB1.value = True Then OutRange(CRow, 1).Resize(1, 3).Interior.Color = RGB(Rf, Gf, Bf)
        OutRange(CRow, 1).value = Round(Rf)
        OutRange(CRow, 2).value = Round(Gf)
        OutRange(CRow, 3).value = Round(Bf)
    Next CRow
    
End Sub

