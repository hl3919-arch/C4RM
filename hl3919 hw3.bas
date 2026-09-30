Attribute VB_Name = "Module1"
Option Explicit
Option Base 1

Function getBondPrice(ByVal y As Double, ByVal face As Double, ByVal couponRate As Double, ByVal m As Integer, Optional ByVal ppy As Integer = 1) As Double

Dim n As Integer
Dim t As Integer
Dim coupon As Double
Dim price As Double

n = m * ppy
coupon = face * couponRate / ppy
price = 0

For t = 1 To n
    price = price + coupon / (1 + y / ppy) ^ t
Next t

price = price + face / (1 + y / ppy) ^ n

getBondPrice = price

End Function


Option Explicit
Option Base 1

Function getBondDuration(ByVal y As Double, ByVal face As Double, _
                         ByVal couponRate As Double, ByVal m As Integer) As Double

Dim t As Integer
Dim cf As Double
Dim pv As Double
Dim price As Double
Dim weightedPV As Double

price = 0
weightedPV = 0

For t = 1 To m

    cf = face * couponRate

    If t = m Then
        cf = cf + face
    End If

    pv = cf / (1 + y) ^ t

    price = price + pv
    weightedPV = weightedPV + pv * t

Next t

getBondDuration = weightedPV / price

End Function

