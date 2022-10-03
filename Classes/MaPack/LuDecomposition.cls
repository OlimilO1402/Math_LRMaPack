VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LuDecomposition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Siehe auch Numerical Recipes
'Kap 2.3 LU Decomposition
'  L * U = A
'  (Lower triangular, Upper triangular)
'  Public Class LuDecomposition
Private mLU As Matrix
Private mPivotSign As Long
Private mPivotVector() As Long '()
''' <summary>Construct a LU decomposition.</summary>
Friend Sub New_(Value As Matrix)
  If Value Is Nothing Then
    'Throw New ArgumentNullException("value")
    MsgBox "ArgumentNull: " & "value"
  End If
  Set mLU = Value.Clone 'DirectCast(value.Clone(), Matrix)
  Dim lu() As Double: Call AssignArray(ArrPtr(lu()), mLU.pArray)
  Dim Rows As Long: Rows = Value.Rows
  Dim Columns As Long: Columns = Value.Columns
  'mPivotVector = New_Integer(rows - 1) ' {} '(rows-1) !!!!!!!!!!!! vorher Fehler(rows)
  ReDim mPivotVector(0 To Rows - 1)
  Dim i As Long, j As Long, k As Long
  For i = 0 To Rows - 1
    mPivotVector(i) = i
  Next
  mPivotSign = 1
  ReDim LUrowi(0 To Columns - 1) As Double ': Call New_DoubleS(LUcolj(), rows - 1) 'LUcolj() = New_Double(rows) ' {}
  ReDim LUcolj(0) As Double 'erzeugt einen SADesc
  Dim pvDataOld As Long: pvDataOld = VarPtr(LUcolj(0))
  Dim kmax As Long
  Dim s As Double
  Dim t As Double
  Dim v As Long
  Dim p As Long
  ' Outer loop.
  For j = 0 To Columns - 1
    ' Make a copy of the j-th column to localize references.
    'For i = 0 To rows - 1
    '  LUcolj(i) = lu(i, j)
    'Next
    'Achtung in VB6 ist es genau andersherum als in VB.NET:
      'hier wird ein Zeiger auf ein Spaltenarray ermittelt
      'Achtung den Gestohlenen ArrayPointer hinterher wieder freigeben!
    Call CopyColPtr(ArrPtr(LUcolj), VarPtr(lu(0, j)), Columns)
    For i = 0 To Rows - 1
      ' Apply previous transformations.
      'LUrowi() = lu(i, 0)
      ' Most of the time is spent in the following dot product.
      'Achtung in VB6 ist es genau andersherum als in VB.NET:
      'hier wird die Zeile kopiert
      For k = 0 To Columns - 1
        LUrowi(k) = lu(i, k)
      Next
      kmax = Math.MinL(i, j)
      s = 0
      For k = 0 To kmax - 1
        s = s + LUrowi(k) * LUcolj(k)
      Next
      LUcolj(i) = LUcolj(i) - s
      LUrowi(j) = LUcolj(i)
    Next
    ' Find pivot and exchange if necessary.
    p = j
    For i = j + 1 To Rows - 1
      If VBA.Math.Abs(LUcolj(i)) > VBA.Math.Abs(LUcolj(p)) Then
        p = i
      End If
    Next
    If p <> j Then
      For k = 0 To Columns - 1
        t = lu(p, k)
        lu(p, k) = lu(j, k)
        lu(j, k) = t
      Next
      v = mPivotVector(p)
      mPivotVector(p) = mPivotVector(j)
      mPivotVector(j) = v
      mPivotSign = -mPivotSign
    End If
    ' Compute multipliers.
    If j < Rows And lu(j, j) <> 0 Then
      For i = j + 1 To Rows - 1
        lu(i, j) = lu(i, j) / lu(j, j)
      Next
    End If
  Next
  Call CopyColPtr(ArrPtr(LUcolj()), pvDataOld, 1)
  Call MArray.ZeroSAPtr(ArrPtr(lu()))
End Sub
''' <summary>Returns if the matrix is non-singular.</summary>
'Public ReadOnly Property NonSingular() As Boolean
Public Property Get NonSingular() As Boolean
  'Get
  Dim j As Long
    For j = 0 To mLU.Columns - 1
      If mLU(j, j) = 0 Then
        'Return False
        NonSingular = False: Exit Property
      End If
    Next
    NonSingular = True 'Return True
  'End Get
End Property
''' <summary>Returns the determinant of the matrix.</summary>
'Public ReadOnly Property Determinant() As Double
Public Property Get Determinant() As Double
  'Get
    If mLU.Rows <> mLU.Columns Then
      'Throw New ArgumentException("Matrix must be square.")
      MsgBox "Matrix must be square."
    End If
    'Dim ddeterminant As Double = DirectCast(pivotSign, Double)
    'Dim ddeterminant As Double
    Determinant = CDbl(mPivotSign)
    Dim j As Long
    For j = 0 To mLU.Columns - 1
      Determinant = Determinant * mLU(j, j)
    Next
    'Return ddeterminant
    'Determinant = ddeterminant
  'End Get
End Property
''' <summary>Returns the lower triangular factor <c>L</c> with <c>A=LU</c>.</summary>
'Public ReadOnly Property LowerTriangularFactor() As Matrix
Public Property Get LowerTriangularFactor() As Matrix
  'Get
    Dim Rows As Long: Rows = mLU.Rows
    Dim Columns As Long: Columns = mLU.Columns
    Dim X As Matrix: Set X = MNew.Matrix(Rows, Columns)
    Dim i As Long, j As Long
    For i = 0 To Rows - 1
      For j = 0 To Columns - 1
        If i > j Then
          X(i, j) = mLU(i, j)
        ElseIf i = j Then
          X(i, j) = 1
        Else
          X(i, j) = 0
        End If
      Next
    Next
    'Return X
    Set LowerTriangularFactor = X
  'End Get
End Property
''' <summary>Returns the lower triangular factor <c>L</c> with <c>A=LU</c>.</summary>
'Public ReadOnly Property UpperTriangularFactor() As Matrix
Public Property Get UpperTriangularFactor() As Matrix
  'Get
    Dim Rows As Long: Rows = mLU.Rows
    Dim Columns As Long: Columns = mLU.Columns
    Dim X As Matrix: Set X = MNew.Matrix(Rows, Columns)
    Dim i As Long, j As Long
    For i = 0 To Rows - 1
      For j = 0 To Columns - 1
        If i <= j Then
          X(i, j) = mLU(i, j)
        Else
          X(i, j) = 0
        End If
      Next
    Next
    'Return X
    Set UpperTriangularFactor = X
  'End Get
End Property
''' <summary>Returns the pivot permuation vector.</summary>
'Public ReadOnly Property PivotPermutationVector() As Double()
Public Property Get PivotPermutationVector() As Double()
  'Get
    Dim Rows As Long: Rows = mLU.Rows
    ReDim p(0 To Rows - 1) As Double ': p = New_Double(Rows) ' {}
    Dim i As Long
    For i = 0 To Rows - 1
      'p(i) = DirectCast(Me.pivotVector(i), Double)
      p(i) = CDbl(mPivotVector(i))
    Next
    'Return p
    PivotPermutationVector = p()
  'End Get
End Property
''' <summary>Solves a set of equation systems of type <c>A * X = B</c>.</summary>
''' <param name="value">Right hand side matrix with as many rows as <c>A</c> and any number of columns.</param>
''' <returns>Matrix <c>X</c> so that <c>L * U * X = B</c>.</returns>
Public Function Solve(ByVal Value As Matrix) As Matrix
  If Value Is Nothing Then
    'Throw New ArgumentNullException("value")
    MsgBox "Solve ArgumentNull: " & "value"
  End If
  If Value.Rows <> mLU.Rows Then
    'Throw New ArgumentException("Invalid matrix dimensions.", "value")
    MsgBox "Argument " & "Invalid matrix dimensions. " & "value"
  End If
  If Not Me.NonSingular Then
    'Throw New InvalidOperationException("Matrix is singular")
    MsgBox "InvalidOperationException " & "Matrix is singular"
  End If
  ' Copy right hand side with pivoting
  Dim count As Long: count = Value.Columns
  Dim X As Matrix: Set X = Value.SubmatrixR(mPivotVector(), 0, count - 1)
  Dim xx() As Double: Call AssignArray(ArrPtr(xx), X.pArray)
  Dim Rows As Long: Rows = mLU.Rows
  Dim Columns As Long: Columns = mLU.Columns
  Dim lu() As Double: Call AssignArray(ArrPtr(lu()), mLU.pArray)
  Dim i As Long, j As Long, k As Long
  For k = 0 To Columns - 1
    ' Solve L*Y = B(piv,:)
    For i = k + 1 To Columns - 1
      For j = 0 To count - 1
        xx(i, j) = xx(i, j) - xx(k, j) * lu(i, k)
      Next
    Next
  Next
  For k = Columns - 1 To 0 Step -1
    ' Solve U*X = Y;
    For j = 0 To count - 1
      xx(k, j) = xx(k, j) / lu(k, k)
    Next
    For i = 0 To k - 1
      For j = 0 To count - 1
        xx(i, j) = xx(i, j) - xx(k, j) * lu(i, k)
      Next
    Next
  Next
  Call MArray.ZeroSAPtr(ArrPtr(lu()))
  Call MArray.ZeroSAPtr(ArrPtr(xx()))
  'Return X
  Set Solve = X
End Function
'  End Class

