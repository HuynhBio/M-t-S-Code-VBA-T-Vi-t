Attribute VB_Name = "Calc"
Sub OxygenSolubility(Wt As Double, UnitWt As String, Bp As Double, UnitBp As String, Sc As Double, UnitSc As String)
    'Doi don vi nhiet do sang Celsius
    If UnitWt = "degC" Or UnitWt = "degrees Celsius" Then
        Wt = Wt
    ElseIf UnitWt = "degF" Or UnitWt = "degrees Fahrenheit" Then
        Wt = (Wt - 32) / 1.8
    Else
        MsgBox ("Vui long chon don vi nhiet do la Celsius hoac Fahrenheit")
        Exit Sub
    End If
    'Check for input problems
    If Wt < 0 Or Wt > 40 Then
        MsgBox "Nhiet do nuoc phai nam trong khoang 0-40 do Celsius hoac 32-104 do Fahrenheit", vbOKOnly, "ERROR !!!"
        Exit Sub
    End If
    'Doi nhiet do sang do Kelvin
    Dim tk As String
    tk = Wt + 273.15
    'Tính gia tri oxygen solubility cua nuoc ngot o 1 atm va nhiet do nuoc dau vao
    Dim sat As Double
    sat = Exp(-139.34411 + (15757010000# + (-66423080# + (12438000000# - 862194900000# / tk) / tk) / tk) / tk)
    'Doi gia tri ap suat sang atmospheres
    If UnitBp = "atm" Or UnitBp = "atmospheres" Then
        Bp = Bp
    ElseIf UnitBp = "mmHg" Or UnitBp = "mm Hg" Then
        Bp = Bp / 760
    ElseIf UnitBp = "inHg" Or UnitBp = "inches Hg" Then
        Bp = Bp / 29.9213
    ElseIf UnitBp = "mbar" Or UnitBp = "milibars" Then
        Bp = Bp / 1013.25
    ElseIf UnitBp = "kPa" Or UnitBp = "kiloPascals" Then
        Bp = Bp / 101.325
    Else
        MsgBox "Vui long chon dung don vi ap suat !", vbOKOnly, "ERROR !!!"
        Exit Sub
    End If
    'Check for input problems
    If Bp < 0.5 Or Bp > 1.1 Then
        MsgBox "Ap suat phai nam tron khoang 0.5-1.1 atmospheres hoac 380-836 mmHg hoac 14.97-32.91 inHg hoac 507-1114 millibars hoac 51-112 kiloPascals", vbOKOnly, "ERROR !!!"
        End Sub
    End If
    'Hieu chinh oxygen solubility theo ap suat
    Dim u As Double, theta As Double
    If Bp <> 1 Then
        u = Exp(11.8571 + (-3840.7 - 216961 / tk) / tk)
        theta = 0.000975 - 0.00001426 * Wt + 0.00000006436 * Wt * Wt
        sat = sat * (Bp - u) * (1 - theta * Bp) / ((1 - u) * (1 - theta))
    End If
    'Doi Specific conductance sang salinity
    Dim sal As Double
    If UnitSc = "sl" Or UnitSc = "Salinity" Then
        sal = Sc
    ElseIf UnitSc = "sc" Or UnitSc = "Specific conductance" Then
        sal = 0.0005572 * Sc + 0.00000000202 * Sc * Sc
    Else
        MsgBox "Vui long nhap dung don vi nhiet do !", vbOKOnly, "ERROR !!!"
        Exit Sub
    End If
    'Check for input problems
    If sal < 0 Or sal > 40 Then
        MsgBox "Do man phai trong khoan 0-40 (o/oo) hoac Do dan dien trong khoang 0-59118 (uS/cm) !", vbOKOnly, "ERROR !!!"
        Exit Sub
    End If
    'Hieu chinh oxygen solubility theo do man
    If Sc > 0 Then
        sat = sat * Exp(-1 * sal * (0.017674 + (-10.754 + 2140.7 / tk) / tk))
    End If
End Sub
Sub PercentSaturation(ox As Double, sat As Double)
    'Gan DO = 0 neu gia tri nhap nho hon 0
    If ox < 0 Then
        ox = 0
    End If
    'tinh Percent saturation
    Dim psat As Double
    psat = 100 * ox / sat
End Sub
