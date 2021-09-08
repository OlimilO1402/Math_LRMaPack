VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Math"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public Class Math
Option Explicit

Private Sub Class_Initialize()
'#####  for Bit Shifting #####
'See ShL ShRz ShR
  InitFunctionsShift
End Sub

'Public Shared Function Abs(ByVal value As Decimal) As Decimal
'Public Shared Function Abs(ByVal value As Double) As Double
'Public Shared Function Abs(ByVal value As Integer) As Integer
'Public Shared Function Abs(ByVal value As Long) As Long
'Public Shared Function Abs(ByVal value As Short) As Short
'Public Shared Function Abs(ByVal value As Single) As Single
'Public Shared Function Abs(ByVal value As System.SByte) As System.SByte
Public Function AAbs(ByVal varValue As Variant) As Variant
  AAbs = VBA.Math.Abs(varValue)
End Function

'Public Shared Function Acos(ByVal d As Double) As Double
Public Function Acos(ByVal d As Double) As Double
  'Acos ist die Umkehrfunktion zur Cosinusfunktion
  'Cos(x) = y; ACos(y) = x
  'Set Acos = New cDouble
  Acos = (3.14159265358979 / 2) - Atn(d / (Sqr(1 - d ^ 2)))
End Function

'Public Shared Function Asin(ByVal d As Double) As Double
Public Function Asin(ByVal d As Double) As Double
  'Set Asin = New cDouble
  Asin = Atn(d / (Sqr(1 - d ^ 2)))
End Function

'Public Shared Function Atan(ByVal d As Double) As Double
Public Function Atan(ByVal d As Double) As Double
  'Set Atan = New cDouble
  Atan = Atn(d)
End Function
'Und was ist mit ACot ????? =Pi/2 - Atan(x)

'Public Shared Function Atan2(ByVal y As Double, ByVal x As Double) As Double
Public Function Atan2(ByVal y As Double, ByVal X As Double) As Double
  'Set Atan2 = New cDouble
  Atan2 = Atn(y / X)
End Function

'Public Shared Function BigMul(ByVal a As Integer, ByVal b As Integer) As Long
Public Function BigMul(ByVal A As Long, ByVal B As Long) As Variant 'As Long
  'vergiss es
  BigMul = A * B
End Function

'Public Shared Function Ceiling(ByVal a As Double) As Double
Public Function Ceiling(ByVal A As Double) As Double
  'Set Ceiling = New cDouble
  Ceiling = Int(A)
End Function

'Public Shared Function Cos(ByVal d As Double) As Double
Public Function Cos(ByVal d As Double) As Double
  'Set Cos = New cDouble
  Cos = VBA.Math.Cos(d)
End Function

'Public Shared Function Cosh(ByVal value As Double) As Double
Public Function Cosh(ByVal value As Double) As Double
  'Set Cosh = New cDouble
  Cosh = (VBA.Math.Exp(value) + VBA.Math.Exp(-value)) / 2
End Function

'Public Shared Function DivRem(ByVal a As Integer, ByVal b As Integer, ByRef result As Integer) As Integer
'Public Shared Function DivRem(ByVal a As Long, ByVal b As Long, ByRef result As Long) As Long
Public Function DivRem(ByVal A As Long, ByVal B As Long, ByRef result As Long) As Long
  'Set DivRem = New cInteger
End Function

'Public Shared Function Exp(ByVal d As Double) As Double
Public Function Exp(ByVal d As Double) As Double
  'Set Exp = New cDouble
  Exp = VBA.Math.Exp(d)
End Function

'Public Shared Function Floor(ByVal d As Double) As Double
Public Function Floor(ByVal d As Double) As Double
  'Set Floor = New cDouble
End Function

'Public Shared Function IEEERemainder(ByVal x As Double, ByVal y As Double) As Double
Public Function IEEERemainder(ByVal X As Double, ByVal y As Double) As Double
  'Set IEEERemainder = New cDouble
End Function

'Public Shared Function Log(ByVal d As Double) As Double
'Public Shared Function Log(ByVal a As Double, ByVal newBase As Double) As Double
Public Function Log(ByVal d As Double, ByVal newBase As Double) As Double
  'Set Log = New cDouble
  Log = VBA.Math.Log(d) / VBA.Math.Log(newBase)
End Function

'Public Shared Function Log10(ByVal d As Double) As Double
Public Function Log10(ByVal d As Double) As Double
  'Set Log10 = New cDouble
  Log10 = VBA.Math.Log(d) / VBA.Math.Log(10)
End Function

'Public Shared Function Max(ByVal val1 As Byte, ByVal val2 As Byte) As Byte
'Public Shared Function Max(ByVal val1 As Decimal, ByVal val2 As Decimal) As Decimal
'Public Shared Function Max(ByVal val1 As Double, ByVal val2 As Double) As Double
'Public Shared Function Max(ByVal val1 As Integer, ByVal val2 As Integer) As Integer
'Public Shared Function Max(ByVal val1 As Long, ByVal val2 As Long) As Long
'Public Shared Function Max(ByVal val1 As Short, ByVal val2 As Short) As Short
'Public Shared Function Max(ByVal val1 As Single, ByVal val2 As Single) As Single
'Public Shared Function Max(ByVal val1 As System.SByte, ByVal val2 As System.SByte) As System.SByte
'Public Shared Function Max(ByVal val1 As System.UInt16, ByVal val2 As System.UInt16) As System.UInt16
'Public Shared Function Max(ByVal val1 As System.UInt32, ByVal val2 As System.UInt32) As System.UInt32
'Public Shared Function Max(ByVal val1 As System.UInt64, ByVal val2 As System.UInt64) As System.UInt64
Public Function Max(ByVal val1 As Variant, ByVal val2 As Variant) As Variant
  If val1 > val2 Then Max = val1 Else Max = val2
End Function
Public Function MaxD(val1 As Double, val2 As Double) As Double
  If val1 > val2 Then MaxD = val1 Else MaxD = val2
End Function
Public Function MaxL(val1 As Long, val2 As Long) As Long
  If val1 > val2 Then MaxL = val1 Else MaxL = val2
End Function

'Public Shared Function Min(ByVal val1 As Byte, ByVal val2 As Byte) As Byte
'Public Shared Function Min(ByVal val1 As Decimal, ByVal val2 As Decimal) As Decimal
'Public Shared Function Min(ByVal val1 As Double, ByVal val2 As Double) As Double
'Public Shared Function Min(ByVal val1 As Integer, ByVal val2 As Integer) As Integer
'Public Shared Function Min(ByVal val1 As Long, ByVal val2 As Long) As Long
'Public Shared Function Min(ByVal val1 As Short, ByVal val2 As Short) As Short
'Public Shared Function Min(ByVal val1 As Single, ByVal val2 As Single) As Single
'Public Shared Function Min(ByVal val1 As System.SByte, ByVal val2 As System.SByte) As System.SByte
'Public Shared Function Min(ByVal val1 As System.UInt16, ByVal val2 As System.UInt16) As System.UInt16
'Public Shared Function Min(ByVal val1 As System.UInt32, ByVal val2 As System.UInt32) As System.UInt32
'Public Shared Function Min(ByVal val1 As System.UInt64, ByVal val2 As System.UInt64) As System.UInt64
Public Function Min(ByVal val1 As Variant, ByVal val2 As Variant) As Variant
  If val1 < val2 Then Min = val1 Else Min = val2
End Function
Public Function MinD(val1 As Double, val2 As Double) As Double
  If val1 < val2 Then MinD = val1 Else MinD = val2
End Function
Public Function MinL(val1 As Long, val2 As Long) As Long
  If val1 < val2 Then MinL = val1 Else MinL = val2
End Function

'Public Shared Function Pow(ByVal x As Double, ByVal y As Double) As Double
Public Function Pow(ByVal X As Double, ByVal y As Double) As Double
  'Set Pow = New cDouble
  Pow = X ^ y
End Function

Public Function Powr2(ByVal Exponent As Long) As Long
  Powr2 = Pow2(Exponent)
End Function
'Public Shared Function Round(ByVal d As Decimal) As Decimal
'Public Shared Function Round(ByVal d As Decimal, ByVal decimals As Integer) As Decimal
'Public Shared Function Round(ByVal a As Double) As Double
'Public Shared Function Round(ByVal value As Double, ByVal digits As Integer) As Double
Public Function Round(ByVal varValue As Variant, Optional ByVal decimaldigits As Long) As Variant
  Round = VBA.Math.Round(varValue, decimaldigits)
End Function

'Public Shared Function Sign(ByVal value As Decimal) As Integer
'Public Shared Function Sign(ByVal value As Double) As Integer
'Public Shared Function Sign(ByVal value As Integer) As Integer
'Public Shared Function Sign(ByVal value As Long) As Integer
'Public Shared Function Sign(ByVal value As Short) As Integer
'Public Shared Function Sign(ByVal value As Single) As Integer
'Public Shared Function Sign(ByVal value As System.SByte) As Integer
Public Function Sign(ByVal varValue As Variant) As Variant
  Sign = Sgn(varValue)
End Function

'Public Shared Function Sin(ByVal a As Double) As Double
Public Function Sin(ByVal A As Double) As Double
  'Set Sin = New cDouble
  Sin = VBA.Math.Sin(A)
End Function

'Public Shared Function Sinh(ByVal value As Double) As Double
Public Function Sinh(ByVal value As Double) As Double
  'Set Sinh = New cDouble
  Sinh = (VBA.Math.Exp(value) - VBA.Math.Exp(-value)) / 2
End Function

'Public Shared Function Sqrt(ByVal d As Double) As Double
Public Function Sqrt(ByVal d As Double) As Double
  'Set Sqrt = New cDouble
  Sqrt = VBA.Math.Sqr(d)
End Function

'Public Shared Function Tan(ByVal a As Double) As Double
Public Function Tan(ByVal A As Double) As Double
  'Set Tan = New cDouble
  Tan = VBA.Math.Tan(A)
End Function
'Und was ist mit Cot??? =Cos(a)/Sin(a)

'Public Shared Const E As Double = 2.7182818284590451
Public Property Get e() As Double
  'Set e = New cDouble
  e = 2.71828182845905
End Property

'Public Shared Const PI As Double = 3.1415926535897931
Public Property Get pI() As Double
  'Set Pi = New cDouble
  pI = 3.14159265358979
End Property

'#######  for Bit Shifting ##########

Public Function ShL(Shifting As Long, Shifter As Long) As Long
  ShL = ShiftLeft(Shifting, Shifter)
End Function
Public Function ShRz(Shifting As Long, Shifter As Long) As Long
  ShRz = ShiftRightZ(Shifting, Shifter)
End Function
Public Function ShR(Shifting As Long, Shifter As Long) As Long
  ShR = ShiftRight(Shifting, Shifter)
End Function

Public Sub Increment(ByRef LngVal As Long) 'As Long
  LngVal = LngVal + 1
End Sub

Public Sub Decrement(ByRef LngVal As Long) 'As Long
  LngVal = LngVal - 1
End Sub

