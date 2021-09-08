Attribute VB_Name = "MNew"
Option Explicit

'die folgenden 8 Funktionen sind zwar ganz nett,
'aber werden nicht verwendet
Public Function New_Double(n As Long) As Double()
  ReDim New_Double(0 To n)
End Function
Public Sub New_DoubleS(DblArr() As Double, n As Long)
  ReDim DblArr(0 To n)
End Sub
Public Function New_Double2(n1 As Long, n2 As Long) As Double()
  ReDim New_Double2(0 To n1, 0 To n2)
End Function
Public Sub New_DoubleS2(DblArr() As Double, n1 As Long, n2 As Long)
  ReDim DblArr(0 To n1, 0 To n2)
End Sub

Public Function New_Integer(n As Long) As Long()
  ReDim New_Integer(0 To n)
End Function
Public Sub New_IntegerS(LngArr() As Long, n As Long)
  ReDim LngArr(0 To n)
End Sub
Public Function New_Integer2(n1 As Long, n2 As Long) As Long()
  ReDim New_Integer2(0 To n1, 0 To n2)
End Function
Public Sub New_IntegerS2(LngArr() As Long, n1 As Long, n2 As Long)
  ReDim LngArr(0 To n1, 0 To n2)
End Sub

Public Function CholeskyDecomposition(Value As Matrix) As CholeskyDecomposition
    Set CholeskyDecomposition = New CholeskyDecomposition: CholeskyDecomposition.New_ Value
End Function

Public Function EigenvalueDecomposition(Value As Matrix) As EigenvalueDecomposition
    Set EigenvalueDecomposition = New EigenvalueDecomposition:  EigenvalueDecomposition.New_ Value
End Function

Public Function LuDecomposition(Value As Matrix) As LuDecomposition
    Set LuDecomposition = New LuDecomposition:  LuDecomposition.New_ Value
End Function

Public Function Matrix(ByVal Rows As Long, ByVal Columns As Long, Optional ByVal Value) As Matrix
    Set Matrix = New Matrix: Matrix.New_ Rows, Columns, Value
End Function

Public Function QrDecomposition(Value As Matrix) As QrDecomposition
    Set QrDecomposition = New QrDecomposition: QrDecomposition.New_ Value
End Function

Public Function SingularValueDecomposition(Value As Matrix) As SingularValueDecomposition
    Set SingularValueDecomposition = New SingularValueDecomposition: SingularValueDecomposition.New_ Value
End Function
