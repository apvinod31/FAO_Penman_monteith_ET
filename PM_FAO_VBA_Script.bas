Attribute VB_Name = "Module1"
Sub PenmanMonteith_ET_DetailedOutput()

Dim i As Long
Dim lastRow As Long

Dim Tmax As Double, Tmin As Double
Dim Tmean As Double
Dim Rs As Double
Dim Td As Double
Dim u2 As Double
Dim P As Double

Dim es As Double, ea As Double
Dim vpd As Double
Dim delta As Double
Dim gamma As Double
Dim Rn As Double
Dim ET0 As Double

' Find last row
lastRow = Cells(Rows.Count, "A").End(xlUp).Row

' ===== Detailed Output Headings =====
Range("H1") = "Mean Air Temperature (°C)"
Range("I1") = "Mean Saturation Vapour Pressure (kPa)"
Range("J1") = "Actual Vapour Pressure (kPa)"
Range("K1") = "Vapour Pressure Deficit (kPa)"
Range("L1") = "Slope of Saturation Vapour Pressure Curve (kPa/°C)"
Range("M1") = "Psychrometric Constant (kPa/°C)"
Range("N1") = "Net Radiation (MJ/m²/day)"
Range("O1") = "Reference Evapotranspiration ET0 (mm/day)"

For i = 2 To lastRow

    Tmax = Cells(i, 2).Value
    Tmin = Cells(i, 3).Value
    P = Cells(i, 4).Value
    Rs = Cells(i, 5).Value
    Td = Cells(i, 6).Value
    u2 = Cells(i, 7).Value

    ' Mean Temperature
    Tmean = (Tmax + Tmin) / 2

    ' Saturation vapour pressure
    es = (0.6108 * Exp((17.27 * Tmax) / (Tmax + 237.3)) + _
          0.6108 * Exp((17.27 * Tmin) / (Tmin + 237.3))) / 2

    ' Actual vapour pressure
    ea = 0.6108 * Exp((17.27 * Td) / (Td + 237.3))

    ' Vapour pressure deficit
    vpd = es - ea

    ' Slope of vapour pressure curve
    delta = (4098 * (0.6108 * Exp((17.27 * Tmean) / (Tmean + 237.3)))) _
             / ((Tmean + 237.3) ^ 2)

    ' Psychrometric constant
    gamma = 0.000665 * P

    ' Net radiation (FAO simplified)
    Rn = 0.77 * Rs

    ' FAO-56 Penman-Monteith ET
    ET0 = (0.408 * delta * Rn + _
           gamma * (900 / (Tmean + 273)) * u2 * vpd) _
           / (delta + gamma * (1 + 0.34 * u2))

    ' ===== Write Outputs =====
    Cells(i, 8).Value = Tmean
    Cells(i, 9).Value = es
    Cells(i, 10).Value = ea
    Cells(i, 11).Value = vpd
    Cells(i, 12).Value = delta
    Cells(i, 13).Value = gamma
    Cells(i, 14).Value = Rn
    Cells(i, 15).Value = ET0

Next i

MsgBox "Daily FAO-56 Penman-Monteith ET calculation completed successfully."

End Sub

