VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GaussElimination"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Siehe auch Robert Sedgewick Algorithmen Kap.37 Gaußsches Eliminationsverfahren
Private mA As Matrix
Private mN As Long
Private mX() As Double 'der Lösungsvektor X

'Halt wo ist B????????
Friend Sub New_(m As Matrix)
  'first check if it's square!!
  If m.Square Then
    Set mA = m.Clone
    mN = mA.Rows - 1
    ReDim mX(mN)
    Call Eliminate
    Call Substitute
  Else
    MsgBox "System is not sqare!"
  End If
End Sub

'Am besten für quadratische Syteme bei denen die meisten
'Elemente ungleich Null sind.
Private Sub Eliminate()
Dim i As Long, j As Long, k As Long, Max As Long
Dim t As Double
Dim A() As Double: Call AssignArray(ArrPtr(A()), mA.pArray)
  For i = 0 To mN
    Max = i
    For j = i + 1 To mN
      If VBA.Math.Abs(A(j, i)) > VBA.Math.Abs(A(Max, i)) Then Max = j
    Next
    For k = i To mN + 1
      'Swap
      t = A(i, k)
      A(i, k) = A(Max, k)
      A(Max, k) = t
    Next
    For j = i + 1 To mN
      For k = mN + 1 To i Step -1
        If A(i, i) > 0 Then
          A(j, k) = A(j, k) - A(i, k) * A(j, i) / A(i, i)
        End If
      Next
    Next
  Next
  Call ZeroSAPtr(ArrPtr(A()))
End Sub
Private Sub Substitute()
Dim i As Long, j As Long, k As Long
Dim t As Double
Dim A() As Double: Call AssignArray(ArrPtr(A()), mA.pArray)
  For j = mN To 0 Step -1
    t = 0
    For k = j + 1 To mN
      t = t + mA(j, k) * mX(k)
    Next
    If A(j, j) > 0 Then
      mX(j) = (mA(j, mN + 1) - t) / A(j, j)
    End If
  Next
  Call ZeroSAPtr(ArrPtr(A()))
End Sub
