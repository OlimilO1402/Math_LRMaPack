VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IComparable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Public Overridable Function CompareTo(ByVal obj As Object) As Integer
Public Function CompareTo(aObj As Variant) As Long
'Zusammenfassung:
' Vergleicht die aktuelle Instanz mit einem anderen Objekt desselben Typs.
'Parameter:
' obj: Ein Objekt, das mit dieser Instanz verglichen werden soll.
'Rückgabewerte:
' Eine 32-Bit-Ganzzahl, die die relative Reihenfolge der verglichenen Objekte angibt.
' Der Rückgabewert hat folgende Bedeutung:
' Wert            Bedeutung
' Kleiner als 0   Diese Instanz ist kleiner als obj.
'   0             Diese Instanz ist gleich obj.
' Größer als 0    Diese Instanz ist größer als obj.

End Function
