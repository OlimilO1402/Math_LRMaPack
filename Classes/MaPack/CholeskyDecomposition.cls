VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CholeskyDecomposition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Public Class CholeskyDecomposition
Private mL As Matrix
Private m_symmetric As Boolean
Private m_positiveDefinite As Boolean
''' <summary>Construct a Cholesky Decomposition.</summary>
Friend Sub New_(ByVal Value As Matrix)
  If Value Is Nothing Then
    'Throw New ArgumentNullException("value")
    MsgBox "ArgumentNull: " & "value"
  End If
  If Not Value.Square Then
    'Throw New ArgumentException("Matrix is not square.", "value")
    MsgBox "Matrix is not square." & "value"
  End If
  Dim dimension As Long: dimension = Value.Rows
  'Dim cols As Long: cols = value.Columns
  Dim L As Matrix: Set L = MNew.Matrix(dimension, dimension)
  'Dim A() As Double: A() = Value.pArray
  Dim A() As Double: Call AssignArray(ArrPtr(A), Value.pArray)
  'Dim ll() As Double: ll() = L.pArray
  Dim ll() As Double: Call AssignArray(ArrPtr(ll), L.pArray)
  ReDim Lrowj(0 To dimension - 1) As Double '() = ll(j)
  ReDim Lrowk(0 To dimension - 1) As Double '() = ll(j)
  Dim d As Double ' = 0
  Dim s As Double
  m_positiveDefinite = True
  m_symmetric = True
  Dim i As Long, j As Long, k As Long
  For j = 0 To dimension - 1
    'Lrowj() = ll(j)
    'das geht leider in VB6 jetzt nicht, da muß man jetzt das Array kopieren
    For i = 0 To dimension
      Lrowj(i) = ll(i, j)
    Next
    d = 0
    For k = 0 To j - 1
      'Lrowk() = ll(k)
      For i = 0 To dimension
        Lrowk(i) = ll(i, k)
      Next
      s = 0
      For i = 0 To k - 1
        s = s + Lrowk(i) * Lrowj(i)
      Next
      s = (A(j, k) - s) / ll(k, k)
      Lrowj(k) = s
      d = d + s * s
      m_symmetric = m_symmetric And (A(k, j) = A(j, k))
    Next
    d = A(j, j) - d
    m_positiveDefinite = m_positiveDefinite And (d > 0)
    ll(j, j) = VBA.Math.Sqr(Math.MaxD(d, 0))
    For k = j + 1 To dimension - 1
      ll(j, k) = 0
    Next
  Next
End Sub
''' <summary>Returns <see langword="true"/> if the matrix is symmetric.</summary>
'Public ReadOnly Property Symmetric() As Boolean
Public Property Get Symmetric() As Boolean
  'Get
    Symmetric = m_symmetric
  'End Get
End Property
''' <summary>Returns <see langword="true"/> if the matrix is positive definite.</summary>
'Public ReadOnly Property PositiveDefinite() As Boolean
Public Property Get PositiveDefinite() As Boolean
  'Get
    PositiveDefinite = m_positiveDefinite
  'End Get
End Property
''' <summary>Returns the left triangular factor <c>L</c> so that <c>A = L * L'</c>.</summary>
'Public ReadOnly Property LeftTriangularFactor() As Matrix
Public Property Get LeftTriangularFactor() As Matrix
  'Get
    'Return Me.L
    Set LeftTriangularFactor = mL
  'End Get
End Property
''' <summary>Solves a set of equation systems of type <c>A * X = B</c>.</summary>
''' <param name="value">Right hand side matrix with as many rows as <c>A</c> and any number of columns.</param>
''' <returns>Matrix <c>X</c> so that <c>L * L' * X = B</c>.</returns>
''' <exception cref="T:System.ArgumentException">Matrix dimensions do not match.</exception>
''' <exception cref="T:System.InvalidOperationException">Matrix is not symmetrix and positive definite.</exception>
Public Function Solve(ByVal Value As Matrix) As Matrix
  If Value Is Nothing Then
    'Throw New ArgumentNullException("value")
  End If
  If Value.Rows <> mL.Rows Then
    'Throw New ArgumentException("Matrix dimensions do not match.")
  End If
  If Not m_symmetric Then
    'Throw New InvalidOperationException("Matrix is not symmetric.")
  End If
  If Not m_positiveDefinite Then
    'Throw New InvalidOperationException("Matrix is not positive definite.")
  End If
  Dim dimension As Long: dimension = mL.Rows
  Dim count As Long: count = Value.Columns
  Dim B As Matrix: Set B = Value.Clone 'DirectCast(value.Clone(), Matrix)
  Dim ll() As Double: Call AssignArray(ArrPtr(ll()), mL.pArray)
  Dim i As Long, j As Long, k As Long
  For k = 0 To mL.Rows - 1
    ' Solve L*Y = B;
    For i = k + 1 To dimension - 1
      For j = 0 To count - 1
        B(i, j) = B(i, j) - B(k, j) * ll(i, k)
      Next
    Next
    For j = 0 To count - 1
      B(k, j) = B(k, j) / ll(k, k)
    Next
  Next
  For k = dimension - 1 To 0 Step -1
    ' Solve L'*X = Y;
    For j = 0 To count - 1
      B(k, j) = B(k, j) / ll(k, k)
    Next
    For i = 0 To k - 1
      For j = 0 To count - 1
        B(i, j) = B(i, j) - B(k, j) * ll(k, i)
      Next
    Next
  Next
  Call MArray.ZeroSAPtr(ArrPtr(ll()))
  'Return B
  Set Solve = B
End Function
'End Class

