VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QrDecomposition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Siehe auch: Numerical Recipes
'   Kap.11.6 "The QR Algorithm for Real Hessenberg Matrices"
'  Public Class QrDecomposition
Private mQr As Matrix
Private mRdiag() As Double
''' <summary>Construct a QR decomposition.</summary>
Friend Sub New_(ByVal Value As Matrix)
  If Value Is Nothing Then
    'Throw New ArgumentNullException("value")
    MsgBox "QrDecomposition::New " & "ArgumentNull: " & "value"
  End If
  Set mQr = Value.Clone 'DirectCast(Value.Clone(), Matrix)
  Dim qr() As Double: Call AssignArray(ArrPtr(qr()), mQr.pArray)
  Dim m As Long: m = Value.Rows
  Dim n As Long: n = Value.Columns
  ReDim mRdiag(0 To n - 1) '= New_Double(n) '{}
  Dim i As Long, j As Long, k As Long
  Dim nrm As Double
  Dim s As Double
  For k = 0 To n - 1
    ' Compute 2-norm of k-th column without under/overflow.
    nrm = 0
    For i = k To m - 1
      nrm = Hypotenuse(nrm, qr(i, k))
    Next
    If nrm <> 0 Then
      ' Form k-th Householder vector.
      If qr(k, k) < 0 Then
        nrm = -nrm
      End If
      For i = k To m - 1
        qr(i, k) = qr(i, k) / nrm
      Next
      qr(k, k) = qr(k, k) + 1
      For j = k + 1 To n - 1
        ' Apply transformation to remaining columns.
        s = 0
        For i = k To m - 1
          s = s + qr(i, k) * qr(i, j)
        Next
        s = -s / qr(k, k)
        For i = k To m - 1
          qr(i, j) = qr(i, j) + s * qr(i, k)
        Next
      Next
    End If
    mRdiag(k) = -nrm
  Next
  Call ZeroSAPtr(ArrPtr(qr()))
End Sub
''' <summary>Least squares solution of <c>A * X = B</c></summary>
''' <param name="value">Right-hand-side matrix with as many rows as <c>A</c> and any number of columns.</param>
''' <returns>A matrix that minimized the two norm of <c>Q * R * X - B</c>.</returns>
''' <exception cref="T:System.ArgumentException">Matrix row dimensions must be the same.</exception>
''' <exception cref="T:System.InvalidOperationException">Matrix is rank deficient.</exception>
Public Function Solve(ByVal Value As Matrix) As Matrix
  If Value Is Nothing Then
    'Throw New ArgumentNullException("value")
    MsgBox "ArgumentNull: " & "value"
  End If
  If Value.Rows <> mQr.Rows Then
    'Throw New ArgumentException("Matrix row dimensions must agree.")
  End If
  If Not Me.FullRank Then
    'Throw New InvalidOperationException("Matrix is rank deficient.")
  End If
  ' Copy right hand side
  Dim count As Long: count = Value.Columns
  Dim X As Matrix: Set X = Value.Clone()
  Dim xx() As Double: Call AssignArray(ArrPtr(xx()), X.pArray)
  Dim m As Long: m = mQr.Rows
  Dim n As Long: n = mQr.Columns
  Dim qr() As Double: Call AssignArray(ArrPtr(qr()), mQr.pArray)
  Dim i As Long, j As Long, k As Long
  Dim s As Double
  For k = 0 To n - 1
    ' Compute Y = transpose(Q)*B
    For j = 0 To count - 1
      s = 0
      For i = k To m - 1
        s = s + qr(i, k) * X(i, j)
      Next
      s = -s / qr(k, k)
      For i = k To m - 1
        xx(i, j) = xx(i, j) + s * qr(i, k)
      Next
    Next
  Next
  For k = n - 1 To 0 Step -1
    ' Solve R*X = Y;
    For j = 0 To count - 1
      xx(k, j) = xx(k, j) / mRdiag(k)
    Next
    For i = 0 To k - 1
      For j = 0 To count - 1
        xx(i, j) = xx(i, j) - xx(k, j) * qr(i, k)
      Next
    Next
  Next
  Call ZeroSAPtr(ArrPtr(xx()))
  Call ZeroSAPtr(ArrPtr(qr()))
  Set Solve = X.Submatrix(0, n - 1, 0, count - 1)
  'Return X.Submatrix(0, n - 1, 0, count - 1)
End Function
''' <summary>Shows if the matrix <c>A</c> is of full rank.</summary>
''' <value>The value is <see langword="true"/> if <c>R</c>, and hence <c>A</c>, has full rank.</value>
'Public ReadOnly Property FullRank() As Boolean
Public Property Get FullRank() As Boolean
  'Get
    Dim Columns As Long: Columns = mQr.Columns
    Dim i As Long
    For i = 0 To Columns - 1
      If mRdiag(i) = 0 Then
        'Return False
        FullRank = False: Exit Property
      End If
    Next
    'Return True
    FullRank = True
  'End Get
End Property
''' <summary>Returns the upper triangular factor <c>R</c>.</summary>
'Public ReadOnly Property UpperTriangularFactor() As Matrix
Public Property Get UpperTriangularFactor() As Matrix
  'Get
    Dim n As Long: n = mQr.Columns
    Dim X As Matrix: Set X = MNew.Matrix(n, n)
    Dim xx() As Double: Call AssignArray(ArrPtr(xx()), X.pArray)
    Dim qr() As Double: Call AssignArray(ArrPtr(qr()), mQr.pArray)
    Dim i As Long, j As Long
    For i = 0 To n - 1
      For j = 0 To n - 1
        If i < j Then
          xx(i, j) = qr(i, j)
        ElseIf i = j Then
          xx(i, j) = mRdiag(i)
        Else
          xx(i, j) = 0
        End If
      Next
    Next
    Call ZeroSAPtr(ArrPtr(xx()))
    Call ZeroSAPtr(ArrPtr(qr()))
    'Return X
    Set UpperTriangularFactor = X
  'End Get
End Property
''' <summary>Returns the orthogonal factor <c>Q</c>.</summary>
'Public ReadOnly Property OrthogonalFactor() As Matrix
Public Property Get OrthogonalFactor() As Matrix
  'Get
    Dim X As Matrix: Set X = MNew.Matrix(mQr.Rows, mQr.Columns)
    Dim xx() As Double: Call AssignArray(ArrPtr(xx()), X.pArray)
    Dim qr() As Double: Call AssignArray(ArrPtr(qr()), mQr.pArray)
    Dim s As Double
    Dim i As Long, j As Long, k As Long
    For k = mQr.Columns - 1 To 0 Step -1
      For i = mQr.Rows - 1 To 0 Step -1 'To????
        xx(i, k) = 0
      Next
      xx(k, k) = 1
      For j = k To mQr.Columns - 1
        If qr(k, k) <> 0 Then
          s = 0
          For i = k To mQr.Rows - 1
            s = s + qr(i, k) * xx(i, j)
          Next
          s = -s / qr(k, k)
          For i = k To mQr.Rows - 1
            xx(i, j) = xx(i, j) + s * qr(i, k)
          Next
        End If
      Next
    Next
    'Return X
    Call ZeroSAPtr(ArrPtr(xx()))
    Call ZeroSAPtr(ArrPtr(qr()))
    Set OrthogonalFactor = X
  'End Get
End Property
'Private Shared Function Hypotenuse(ByVal a As Double, ByVal b As Double) As Double
Private Function Hypotenuse(ByVal A As Double, ByVal B As Double) As Double
Dim r As Double
  If VBA.Math.Abs(A) > VBA.Math.Abs(B) Then
    r = B / A
    'Return VBA.Math.Abs(a) * Math.Sqrt(1 + r * r)
    Hypotenuse = VBA.Math.Abs(A) * VBA.Math.Sqr(1 + r * r): Exit Function
  End If
  If B <> 0 Then
    r = A / B
    'Return VBA.Math.Abs(b) * Math.Sqrt(1 + r * r)
    Hypotenuse = VBA.Math.Abs(B) * VBA.Math.Sqr(1 + r * r): Exit Function
  End If
  'Return 0
  Hypotenuse = 0
End Function
'  End Class

