Attribute VB_Name = "MCompare"
Option Explicit
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (pDst As Any, pSrc As Any, ByVal nBytes As Long)
Private Declare Function CompareStringA Lib "kernel32.dll" (ByVal Locale As Long, ByVal dwCmpFlags As Long, ByVal lpString1 As String, ByVal cchCount1 As Long, ByVal lpString2 As String, ByVal cchCount2 As Long) As Long
Private Declare Function CompareStringW Lib "kernel32.dll" (ByVal Locale As Long, ByVal dwCmpFlags As Long, ByVal lpString1 As String, ByVal cchCount1 As Long, ByVal lpString2 As String, ByVal cchCount2 As Long) As Long
Private Declare Function lstrcmp Lib "kernel32.dll" Alias "lstrcmpA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function lstrcmpi Lib "kernel32.dll" Alias "lstrcmpiA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
'Private Declare Sub RtlCompareMemory Lib "ntdll.dll" (ByRef Source1 As Any, ByRef Source2 As Any, ByRef Length As Long)
Private Declare Function RtlCompareMemory Lib "ntdll.dll" (ByRef Source1 As Any, ByRef Source2 As Any, ByRef Length As Long) As Long
Private Declare Function CompareMem Lib "ntdll.dll" Alias "RtlCompareMemory" (ByRef Source1 As Any, ByRef Source2 As Any, ByRef Length As Long) As Long

'Einen Speicherbereich vergleichen auf Gleichheit
Public Function IsEqualMemory(aPtr1 As Long, aPtr2 As Long, nBCount As Long) As Boolean
'gibt true zurück, wenn der Speicherbereich gleich ist
TryE: On Error GoTo CatchE
ReDim btemp1(0 To nBCount - 1) As Byte
ReDim btemp2(0 To nBCount - 1) As Byte
Dim pT1 As Long: pT1 = VarPtr(btemp1(0))
Dim pT2 As Long: pT2 = VarPtr(btemp2(0))
Dim i As Long
  Call RtlMoveMemory(ByVal pT1, ByVal aPtr1, nBCount)
  Call RtlMoveMemory(ByVal pT2, ByVal aPtr2, nBCount)
  For i = 0 To (nBCount - 1)
    If btemp1(i) <> btemp2(i) Then Exit Function
  Next
  IsEqualMemory = True
  Exit Function
CatchE:
  IsEqualMemory = False
End Function

'Einen Speicherbereich vergleichen
'geht für alle unsigned numerischen Typen, da von rechts verglichen wird
'also: Byte, nur positive Integer, nur positive Long
Public Function CompareUnsNum(aPtr1 As Long, aPtr2 As Long, nByteCount As Long) As Long
'Rückgabe, Bedingung
'      -1, wenn aPtr1 < aPtr2
'       0, wenn aPtr1 = aPtr2
'       1, wenn aPtr1 > aPtr2
TryE: On Error GoTo CatchE
ReDim btemp1(0 To nByteCount - 1) As Byte
ReDim btemp2(0 To nByteCount - 1) As Byte
Dim pT1 As Long: pT1 = VarPtr(btemp1(0))
Dim pT2 As Long: pT2 = VarPtr(btemp2(0))
Dim i As Long
  Call RtlMoveMemory(ByVal pT1, ByVal aPtr1, nByteCount)
  Call RtlMoveMemory(ByVal pT2, ByVal aPtr2, nByteCount)
  For i = (nByteCount - 1) To 0 Step -1
    If btemp1(i) < btemp2(i) Then
      CompareUnsNum = -1
      Exit Function
    ElseIf btemp1(i) > btemp2(i) Then
      CompareUnsNum = 1
      Exit Function
    End If
  Next
  Exit Function
CatchE:
  CompareUnsNum = 0
End Function

'geht nur für String, da von links verglichen wird
Public Function CompareStr(aStr1 As String, aStr2 As String) As Long
TryE: On Error GoTo CatchE
Dim StrPtr1 As Long: StrPtr1 = StrPtr(aStr1)
Dim StrPtr2 As Long: StrPtr2 = StrPtr(aStr2)
Dim L1 As Long: L1 = LenB(aStr1)
Dim L2 As Long: L2 = LenB(aStr2)
Dim nBCount As Long: If L1 < L2 Then nBCount = L1 Else nBCount = L2
ReDim btemp1(0 To nBCount - 1) As Byte
ReDim btemp2(0 To nBCount - 1) As Byte
Dim pT1 As Long: pT1 = VarPtr(btemp1(0))
Dim pT2 As Long: pT2 = VarPtr(btemp2(0))
Dim i As Long
  Call RtlMoveMemory(ByVal pT1, ByVal StrPtr1, nBCount)
  Call RtlMoveMemory(ByVal pT2, ByVal StrPtr2, nBCount)
  For i = 0 To (nBCount - 1)
    If btemp1(i) < btemp2(i) Then
      CompareStr = -1
      Exit Function
    ElseIf btemp1(i) > btemp2(i) Then
      CompareStr = 1
      Exit Function
    End If
  Next
  Exit Function
CatchE:
  CompareStr = 0
End Function
Private Function MinLng(LngVal1 As Long, LngVal2 As Long) As Long
  If LngVal1 < LngVal2 Then MinLng = LngVal1 Else MinLng = LngVal2
End Function

'allgemein, für alle Typen, wenn angegeben wird
'von welcher Seite her der Vergleich starten soll
'und ob IsSigned
'IsSigned = True:  auch  negative  Werte  können vorkommen
'IsSigned = False: keine negativen Werte
Public Function CompareMemory(aPtr1 As Long, aPtr2 As Long, nByteCount As Long, Optional FromLeft As Boolean, Optional IsSigned As Boolean) As Long
'Rückgabe, Bedingung
'      -1, wenn aPtr1 < aPtr2
'       0, wenn aPtr1 = aPtr2
'       1, wenn aPtr1 > aPtr2
TryE: On Error GoTo CatchE
ReDim btemp1(0 To nByteCount - 1) As Byte
ReDim btemp2(0 To nByteCount - 1) As Byte
Dim pT1 As Long: pT1 = VarPtr(btemp1(0))
Dim pT2 As Long: pT2 = VarPtr(btemp2(0))
Dim sgn1 As Long: sgn1 = 1
Dim sgn2 As Long: sgn2 = 1
Dim i As Long, st As Long, n As Long, stp As Long
  Call RtlMoveMemory(ByVal pT1, ByVal aPtr1, nByteCount)
  Call RtlMoveMemory(ByVal pT2, ByVal aPtr2, nByteCount)
  If FromLeft Then
    st = nByteCount - 1: n = 0: stp = -1
  Else
    st = 0: n = nByteCount - 1: stp = 1
  End If
  If IsSigned Then
    'jetzt rausfinden, ob val1 < 0 und val2 < 0
    If btemp1(nByteCount - 1) And &H80 Then sgn1 = -1
    If btemp2(nByteCount - 1) And &H80 Then sgn2 = -1
  End If
  For i = st To n Step stp
    If btemp1(i) < btemp2(i) Then
      CompareMemory = -1 * (sgn1 * sgn2)
      Exit Function
    ElseIf btemp1(i) > btemp2(i) Then
      CompareMemory = 1 * (sgn1 * sgn2)
      Exit Function
    End If
  Next
  Exit Function
CatchE:
  CompareMemory = 0
End Function

Public Function IsNegative(pVarPtr As Long, nByteCount As Long) As Boolean
ReDim ByteArr(0 To nByteCount - 1) As Byte
  Call RtlMoveMemory(ByteArr(0), ByVal pVarPtr, nByteCount)
  IsNegative = ByteArr(nByteCount - 1) And &H80
End Function

Public Function IsNegativeP(pVarPtr As Long, nByteCount As Long) As Boolean
Dim aB As Byte: Call RtlMoveMemory(aB, ByVal UnsignedAdd(pVarPtr, nByteCount - 1), 1)
  IsNegativeP = aB And &H80
End Function

