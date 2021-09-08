VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Random"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public Class Random
'     Member von: System
'in Axdll Instancing: 6 - GlobalMultiUse
Option Explicit
Private mSeed As Long

Private Sub Class_Initialize()
  mSeed = 1
  Randomize
End Sub

'Public Sub New(ByVal Seed As Integer)
Friend Sub New_(Optional ByVal Seed As Long)
'Seed: Eine Zahl, mit der ein Startwert für Folgen von Pseudozufallszahlen berechnet wird.
  mSeed = Seed
End Sub

'in ein Modul, z.B. ModConstructors
Public Function Random(Optional ByVal Seed As Long) As Random
  Set Random = New Random: Random.New_ Seed
End Function

'Public Overridable Sub NextBytes(ByVal buffer() As Byte)
Public Sub NextBytes(Buffer() As Byte)
' Füllt die Elemente eines angegebenen Bytearrays mit Zufallszahlen.
Dim i As Long
  For i = LBound(Buffer) To UBound(Buffer)
    Buffer(i) = 255 * Rnd(mSeed)
  Next
End Sub

'Public Overridable Function [Next]() As Integer
'Public Overridable Function [Next](ByVal maxValue As Integer) As Integer
'Rückgabe: Eine 32-Bit-Ganzzahl mit Vorzeichen, die größer oder gleich 0 und kleiner als
'      MaxValue ist, d. h., der Bereich der Rückgabewerte umfasst 0, aber nicht MaxValue.
'Public Overridable Function [Next](ByVal minValue As Integer, ByVal maxValue As Integer) As Integer
'Rückgabe: Eine 32-Bit-Ganzzahl mit Vorzeichen, die größer oder gleich minValue und kleiner als
'      MaxValue ist, d. h., der Bereich der Rückgabewerte umfasst minValue, aber nicht MaxValue.
'      Wenn minValue gleich maxValue ist, wird minValue zurückgegeben.
Public Function NNext(Optional ByVal max_minValue As Long, Optional ByVal MaxValue As Long) As Long 'cInteger
Dim TmpVal As Long
  'Set NNext = New cInteger
  If max_minValue > MaxValue Then
    TmpVal = MaxValue
    MaxValue = max_minValue
    max_minValue = TmpVal
  End If
  If max_minValue = 0 And MaxValue = 0 Then
    MaxValue = 2147483647
  End If
  NNext = CLng(max_minValue + (MaxValue - 1 - max_minValue) * Rnd(mSeed))
End Function

'Public Overridable Function NextDouble() As Double
Public Function NextDouble() As Double 'cDouble
  'Set NextDouble = New cDouble
  NextDouble = CDbl(Rnd(mSeed))
End Function

'Protected Overridable Function Sample() As Double
Private Function Sample() As Double
  Sample = CDbl(Rnd(mSeed))
End Function

