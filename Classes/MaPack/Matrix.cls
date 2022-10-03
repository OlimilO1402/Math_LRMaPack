VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Matrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'  Public Class Matrix
Private data() As Double '()()
Private m_rows As Long
Private m_columns As Long
'Private Random As New Random

''' <summary>Constructs an empty matrix of the given size.</summary>
''' <param name="rows">Number of rows.</param>
''' <param name="columns">Number of columns.</param>
Friend Sub New_(ByVal Rows As Long, ByVal Columns As Long, Optional ByVal Value)
  m_rows = Rows
  m_columns = Columns
  'Call New_DoubleS2(data, rows - 1, columns - 1) '() {}
  ReDim data(0 To Rows - 1, 0 To Columns - 1)
  Dim i As Long
'  For i = 0 To rows - 1
'    data(i) = New_Double(columns - 1) ' {}
'  Next
  If Not IsMissing(Value) Then
    Dim DD As Double: DD = CDbl(Value)
    For i = 0 To Rows - 1
      data(i, i) = DD
    Next
  End If
End Sub
''' <summary>Constructs a matrix of the given size and assigns a given value to all diagonal elements.</summary>
''' <param name="rows">Number of rows.</param>
''' <param name="columns">Number of columns.</param>
''' <param name="value">Value to assign to the diagnoal elements.</param>
'Public Sub NewD(ByVal rows As Long, ByVal columns As Long, ByVal value As Double)
'  m_rows = rows
'  m_columns = columns
'  'Me.data = New Double(rows)() {}
'  Dim i As Long
'  'For i = 0 To rows - 1
'  '  data(i) = New Double(columns) {}
'  'Next
'  For i = 0 To rows - 1
'    data(i, i) = value
'  Next
'End Sub
''' <summary>Constructs a matrix from the given array.</summary>
''' <param name="value">The array the matrix gets constructed from.</param>
'<CLSCompliant(False)>
Public Sub NewD(Value() As Double)
  m_rows = UBound(Value, 1) 'value.Length
  m_columns = UBound(Value, 2) 'value(0).Length
'  Dim i As Long
'  For i = 0 To m_rows - 1
'    If value(i).Length <> m_columns Then
'      'Throw New ArgumentException("Argument out of range.")
'      MsgBox "Argument out of range."
'    End If
'  Next
  'in .NET wird hier nicht komplett kopiert, sondern ein Zeiger auf das Array übergeben.
  'data() = Value()
  Call AssignArray(ArrPtr(data), ArrPtr(Value))
  'Achtung bei Verwendung dieser Sub muß in Terminate der Zeiger wieder genullt werden
End Sub
''' <summary>Determines weather two instances are equal.</summary>
'Public Overloads Overrides Function Equals(ByVal obj As Object) As Boolean
'Public Function Equals(ByVal obj As Object) As Boolean
'  Return Equals(Me, DirectCast(obj, Matrix))
'End Function
''' <summary>Determines weather two instances are equal.</summary>
'Public Overloads Shared Function Equals(ByVal left As Matrix, ByVal right As Matrix) As Boolean
Public Function Equals(ByVal left As Matrix, ByVal right As Matrix) As Boolean
  'If (DirectCast(left, Object)) = (DirectCast(right, Object)) Then
  If left Is right Then
    'Return True
    Equals = True: Exit Function
  End If
  'If ((DirectCast(left, Object)) Is Nothing) OrElse ((DirectCast(right, Object)) Is Nothing) Then
  If (left Is Nothing) Or (right Is Nothing) Then
    Equals = False: Exit Function ''Return False
  End If
  'If (left.Rows <> right.Rows) OrElse (left.Columns <> right.Columns) Then
  If (left.Rows <> right.Rows) Or (left.Columns <> right.Columns) Then
    Equals = False: Exit Function 'Return False
  End If
  Dim i As Long, j As Long
  For i = 0 To left.Rows - 1
    For j = 0 To left.Columns - 1
      If left(i, j) <> right(i, j) Then
        Equals = False: Exit Function 'Return False
      End If
    Next
  Next
  Equals = True: Exit Function 'Return True
End Function
''' <summary>Serves as a hash function for a particular type, suitable for use in hashing algorithms and data structures like a hash table.</summary>
'Public Overloads Overrides Function GetHashCode() As Long
Public Function GetHashCode() As Long
  'Return (Me.Rows + Me.Columns)
  GetHashCode = (Me.Rows + Me.Columns)
End Function
'Friend ReadOnly Property Array() As Double()()
Friend Property Get pArray() As Long 'Double() '()
  'Get
    'Return Me.data
    pArray = ArrPtr(data())
  'End Get
End Property
''' <summary>Returns the number of columns.</summary>
'Public ReadOnly Property Rows() As Long
Public Property Get Rows() As Long
  'Get
    'Return Me.m_rows
    Rows = m_rows
  'End Get
End Property
''' <summary>Returns the number of columns.</summary>
'Public ReadOnly Property Columns() As Long
Public Property Get Columns() As Long
  'Get
    'Return Me.m_columns
    Columns = m_columns
  'End Get
End Property
''' <summary>Return <see langword="true"/> if the matrix is a square matrix.</summary>
'Public ReadOnly Property Square() As Boolean
Public Property Get Square() As Boolean
  'Get
    'Return (m_rows = m_columns)
    Square = (m_rows = m_columns)
  'End Get
End Property
''' <summary>Returns <see langword="true"/> if the matrix is symmetric.</summary>
'Public ReadOnly Property Symmetric() As Boolean
Public Property Get Symmetric() As Boolean
  'Get
    If Me.Square Then
      Dim i As Long, j As Long
      For i = 0 To m_rows - 1
        For j = 0 To i
          If data(i, j) <> data(j, i) Then
            'Return False
            Symmetric = False: Exit Property
          End If
        Next
      Next
      'Return True
      Symmetric = True: Exit Property
    End If
    'Return False
    Symmetric = False
  'End Get
End Property
''' <summary>Access the value at the given location.</summary>
'Default Public Property Item(ByVal row As Long, ByVal column As Long) As Double
Public Property Get Item(ByVal row As Long, ByVal column As Long) As Double
Attribute Item.VB_UserMemId = 0
  'Get
  '  Return Me.data(row)(column)
  Item = data(row, column)
  'End Get
End Property
Public Property Let Item(ByVal row As Long, ByVal column As Long, Value As Double)
  'Set(ByVal Value As Double)
  '  Me.data(row)(column) = value
  data(row, column) = Value
  'End Set
End Property
''' <summary>Returns a sub matrix extracted from the current matrix.</summary>
''' <param name="startRow">Start row index</param>
''' <param name="endRow">End row index</param>
''' <param name="startColumn">Start column index</param>
''' <param name="endColumn">End column index</param>
Public Function Submatrix(ByVal startRow As Long, ByVal endRow As Long, ByVal startColumn As Long, ByVal endColumn As Long) As Matrix
  'If (startRow > endRow) OrElse (startColumn > endColumn) OrElse (startRow < 0) OrElse (startRow >= Me.m_rows) OrElse (endRow < 0) OrElse (endRow >= Me.m_rows) OrElse (startColumn < 0) OrElse (startColumn >= Me.m_columns) OrElse (endColumn < 0) OrElse (endColumn >= Me.m_columns) Then
  If (startRow > endRow) Or (startColumn > endColumn) Or (startRow < 0) Or (startRow >= m_rows) Or (endRow < 0) Or (endRow >= m_rows) Or (startColumn < 0) Or (startColumn >= m_columns) Or (endColumn < 0) Or (endColumn >= m_columns) Then
    'Throw New ArgumentException("Argument out of range.")
    MsgBox "Matrix::Submatrix Argument out of range."
  End If
'  Dim X As Matrix: Set X = New_Matrix(endRow - startRow + 1, endColumn - startColumn + 1)
  Dim X As Matrix: Set X = MNew.Matrix(endRow - startRow, endColumn - startColumn)
  'Dim xx() As Double: xx() = X.pArray
  Dim xx() As Double: Call AssignArray(ArrPtr(xx), X.pArray)
  Dim i As Long, j As Long
  For i = startRow To endRow
    For j = startColumn To endColumn
      xx(i - startRow, j - startColumn) = data(i, j)
    Next
  Next
  Call MArray.ZeroSAPtr(ArrPtr(xx()))
  'Return X
  Set Submatrix = X
End Function
''' <summary>Returns a sub matrix extracted from the current matrix.</summary>
''' <param name="rowIndexes">Array of row indices</param>
''' <param name="columnIndexes">Array of column indices</param>
'Public Function Submatrix(ByVal rowIndexes As Long(), ByVal columnIndexes As Long()) As Matrix
Public Function SubmatrixI(rowIndexes() As Long, columnIndexes() As Long) As Matrix
  Dim X As Matrix: Set X = MNew.Matrix(UBound(rowIndexes) + 1, UBound(columnIndexes) + 1)
  Dim xx() As Double: Call AssignArray(ArrPtr(xx), X.pArray) 'xx() = X.pArray
  Dim i As Long, j As Long
  For i = 0 To UBound(rowIndexes) '.Length - 1
    For j = 0 To UBound(columnIndexes) '.Length - 1
      If (rowIndexes(i) < 0) Or (rowIndexes(i) >= m_rows) Or (columnIndexes(j) < 0) Or (columnIndexes(j) >= m_columns) Then
        'Throw New ArgumentException("Argument out of range.")
        MsgBox "Argument out of range."
      End If
      xx(i, j) = data(rowIndexes(i), columnIndexes(j))
    Next
  Next
  'Return X
  Call MArray.ZeroSAPtr(ArrPtr(xx()))
  Set SubmatrixI = X
End Function
''' <summary>Returns a sub matrix extracted from the current matrix.</summary>
''' <param name="i0">Starttial row index</param>
''' <param name="i1">End row index</param>
''' <param name="c">Array of row indices</param>
'Public Function Submatrix(ByVal i0 As Long, ByVal i1 As Long, ByVal c As Long()) As Matrix
Public Function SubmatrixC(ByVal i0 As Long, ByVal i1 As Long, c() As Long) As Matrix
  If (i0 > i1) Or (i0 < 0) Or (i0 >= m_rows) Or (i1 < 0) Or (i1 >= m_rows) Then
    'Throw New ArgumentException("Argument out of range.")
    MsgBox "Argument out of range."
  End If
  Dim X As Matrix: Set X = MNew.Matrix(i1 - i0, UBound(c))
  'Dim xx() As Double: xx() = X.pArray
  Dim xx() As Double: Call AssignArray(ArrPtr(xx), X.pArray)
  Dim i As Long, j As Long
  For i = i0 To i1
    For j = 0 To UBound(c) '.Length - 1
      If (c(j) < 0) Or (c(j) >= m_columns) Then
        'Throw New ArgumentException("Argument out of range.")
        MsgBox "Argument out of range."
      End If
      xx(i - i0, j) = data(i, c(j))
    Next
  Next
  Call MArray.ZeroSAPtr(ArrPtr(xx()))
  'Return X
  Set SubmatrixC = X
End Function
''' <summary>Returns a sub matrix extracted from the current matrix.</summary>
''' <param name="r">Array of row indices</param>
''' <param name="j0">Start column index</param>
''' <param name="j1">End column index</param>
Public Function SubmatrixR(r() As Long, ByVal j0 As Long, ByVal j1 As Long) As Matrix
  If (j0 > j1) Or (j0 < 0) Or (j0 >= m_columns) Or (j1 < 0) Or (j1 >= m_columns) Then
    'Throw New ArgumentException("Argument out of range.")
    MsgBox "Argument out of range."
  End If
  Dim X As Matrix: Set X = MNew.Matrix(UBound(r) + 1, j1 - j0 + 1)
  Dim xx() As Double: Call AssignArray(ArrPtr(xx), X.pArray) 'xx() = X.pArray
  Dim i As Long, j As Long
  For i = 0 To UBound(r) '.Length - 1
    For j = j0 To j1
      If (r(i) < 0) Or (r(i) >= m_rows) Then
        'Throw New ArgumentException("Argument out of range.")
        MsgBox "Argument out of range."
      End If
      xx(i, j - j0) = data(r(i), j)
    Next
  Next
  Call MArray.ZeroSAPtr(ArrPtr(xx()))
  'Return X
  Set SubmatrixR = X
End Function
''' <summary>Creates a copy of the matrix.</summary>
Public Function Clone() As Matrix
  Dim X As Matrix: Set X = MNew.Matrix(m_rows, m_columns)
  Dim xx() As Double: Call AssignArray(ArrPtr(xx()), X.pArray) ''xx() = X.pArray
  Dim i As Long, j As Long
  For i = 0 To m_rows - 1
    For j = 0 To m_columns - 1
      xx(i, j) = data(i, j)
    Next
  Next
  'Return X
  Call MArray.ZeroSAPtr(ArrPtr(xx()))
  Set Clone = X
End Function
''' <summary>Returns the transposed matrix.</summary>
Public Function Transpose() As Matrix
  Dim X As Matrix: Set X = MNew.Matrix(m_columns, m_rows)
  'Dim xx() As Double: xx() = X.pArray
  Dim xx() As Double: Call AssignArray(ArrPtr(xx), X.pArray)
  Dim i As Long, j As Long
  For i = 0 To m_rows - 1
    For j = 0 To m_columns - 1
      xx(j, i) = data(i, j)
    Next
  Next
  'Return X
  Call MArray.ZeroSAPtr(ArrPtr(xx()))
  Set Transpose = X
End Function
''' <summary>Returns the One Norm for the matrix.</summary>
''' <value>The maximum column sum.</value>
'Public ReadOnly Property Norm1() As Double
Public Property Get Norm1() As Double
  'Get
    Dim f As Double ': f = 0
    Dim s As Double ': s = 0
    Dim i As Long, j As Long
    For j = 0 To m_columns - 1
      s = 0
      For i = 0 To m_rows - 1
        s = s + VBA.Math.Abs(data(i, j))
      Next
      f = Math.MaxD(f, s)
    Next
    'Return f
    Norm1 = f
  'End Get
End Property
''' <summary>Returns the Infinity Norm for the matrix.</summary>
''' <value>The maximum row sum.</value>
'Public ReadOnly Property InfinityNorm() As Double
Public Property Get InfinityNorm() As Double
  'Get
    Dim f As Double ' = 0
    Dim s As Double ' = 0
    Dim i As Long, j As Long
    For i = 0 To m_rows - 1
      s = 0
      For j = 0 To m_columns - 1
        s = s + VBA.Math.Abs(data(i, j))
      Next
      f = Math.MaxD(f, s)
    Next
    'Return f
    InfinityNorm = f
  'End Get
End Property
''' <summary>Returns the Frobenius Norm for the matrix.</summary>
''' <value>The square root of sum of squares of all elements.</value>
'Public ReadOnly Property FrobeniusNorm() As Double
Public Property Get FrobeniusNorm() As Double
  'Get
    'Dim f As Double ' = 0
    Dim i As Long, j As Long
    For i = 0 To m_rows - 1
      For j = 0 To m_columns - 1
        'f = Hypotenuse(f, data(i, j))
        FrobeniusNorm = Hypotenuse(FrobeniusNorm, data(i, j))
      Next
    Next
    'Return f
  '  FrobeniusNorm = f
  'End Get
End Property
''' <summary>Unary minus.</summary>
'Public Shared Function Negate(ByVal value As Matrix) As Matrix
Public Function Negate(Value As Matrix) As Matrix
  If Value Is Nothing Then
    'Throw New ArgumentNullException("value")
    MsgBox "ArgumentNull: " & "value"
  End If
  Dim Rows As Long: Rows = Value.Rows
  Dim Columns As Long: Columns = Value.Columns
  Dim dataV() As Double: Call AssignArray(ArrPtr(dataV()), Value.pArray)
  Dim X As Matrix: Set X = MNew.Matrix(Rows, Columns)
  'Dim xx() As Double: xx() = X.pArray
  Dim xx() As Double: Call AssignArray(ArrPtr(xx), X.pArray)
  Dim i As Long, j As Long
  For i = 0 To Rows - 1
    For j = 0 To Columns - 1
      xx(i, j) = -dataV(i, j)
    Next
  Next
  Call MArray.ZeroSAPtr(ArrPtr(xx()))
  Call MArray.ZeroSAPtr(ArrPtr(dataV()))
  'Return X
  Set Negate = X
End Function
''' <summary>Unary minus.</summary>
'Public Shared Function op_Dif(ByVal value As Matrix) As Matrix
Public Function op_Neg(Value As Matrix) As Matrix
  If Value Is Nothing Then
    'Throw New ArgumentNullException("value")
    MsgBox "ArgumentNull: " & "value"
  End If
  'Return Negate(value)
  Set op_Neg = Negate(Value)
End Function 'Operator
''' <summary>Matrix equality.</summary>
'Public Shared Function op_Equal(ByVal left As Matrix, ByVal right As Matrix) As Boolean
Public Function op_Equal(left As Matrix, right As Matrix) As Boolean
  'Return Equals(left, right)
  op_Equal = Equals(left, right)
End Function 'Operator
''' <summary>Matrix inequality.</summary>
'Public Shared Function op_Unequal(ByVal left As Matrix, ByVal right As Matrix) As Boolean
Public Function op_Unequal(left As Matrix, right As Matrix) As Boolean
  'Return Not Equals(left, right)
  op_Unequal = Not Equals(left, right)
End Function 'Operator
''' <summary>Matrix addition.</summary>
'Public Shared Function Add(ByVal left As Matrix, ByVal right As Matrix) As Matrix
Public Function Add(ByVal left As Matrix, ByVal right As Matrix) As Matrix
  If left Is Nothing Then
    'Throw New ArgumentNullException("left")
    MsgBox "left"
  End If
  If right Is Nothing Then
    'Throw New ArgumentNullException("right")
    MsgBox "right"
  End If
  Dim Rows As Long: Rows = left.Rows
  Dim Columns As Long: Columns = left.Columns
  'Dim data() As Double: data() = left.pArray
  Dim dataL() As Double: Call AssignArray(ArrPtr(dataL), left.pArray)
  If (Rows <> right.Rows) Or (Columns <> right.Columns) Then
    'Throw New ArgumentException("Matrix dimension do not match.")
    MsgBox "Matrix::Add " & "Matrix dimension do not match."
  End If
  Dim X As Matrix: Set X = MNew.Matrix(Rows, Columns)
  'Dim xx() As Double: xx() = X.pArray
  Dim xx() As Double: Call AssignArray(ArrPtr(xx()), X.pArray)
  Dim i As Long, j As Long
  For i = 0 To Rows - 1
    For j = 0 To Columns - 1
      xx(i, j) = dataL(i, j) + right(i, j)
    Next
  Next
  'Return X
  Call MArray.ZeroSAPtr(ArrPtr(xx()))
  Set Add = X
End Function
''' <summary>Matrix addition.</summary>
'Public Shared Function op_Plus(ByVal left As Matrix, ByVal right As Matrix) As Matrix
Public Function op_Plus(left As Matrix, right As Matrix) As Matrix
  If left Is Nothing Then
    'Throw New ArgumentNullException("left")
    MsgBox "ArgumentNull: " & "left"
  End If
  If right Is Nothing Then
    'Throw New ArgumentNullException("right")
    MsgBox "ArgumentNull: " & "right"
  End If
  Set op_Plus = Add(left, right)
End Function 'Operator
''' <summary>Matrix subtraction.</summary>
'Public Shared Function Subtract(ByVal left As Matrix, ByVal right As Matrix) As Matrix
Public Function Subtract(left As Matrix, right As Matrix) As Matrix
  If left Is Nothing Then
    'Throw New ArgumentNullException("left")
    MsgBox "Matrix::Subtract " & "ArgumentNull: " & "left"
  End If
  If right Is Nothing Then
    'Throw New ArgumentNullException("right")
    MsgBox "Matrix::Subtract " & "ArgumentNull: " & "right"
  End If
  Dim Rows As Long: Rows = left.Rows
  Dim Columns As Long: Columns = left.Columns
  'Dim data() As Double: data() = left.pArray
  Dim dataL() As Double: Call AssignArray(ArrPtr(dataL()), left.pArray)
  If (Rows <> right.Rows) Or (Columns <> right.Columns) Then
    'Throw New ArgumentException("Matrix dimension do not match.")
    MsgBox "Matrix::Subract " & "Matrix dimension do not match."
  End If
  Dim X As Matrix: Set X = MNew.Matrix(Rows, Columns)
'  Dim xx()  As Double: xx() = X.pArray
  Dim xx() As Double: Call AssignArray(ArrPtr(xx), X.pArray)
  Dim i As Long, j As Long
  For i = 0 To Rows - 1
    For j = 0 To Columns - 1
    'Warum eigentlich nicht dataR für right?
      xx(i, j) = dataL(i, j) - right(i, j)
    Next
  Next
  'Return X
  Call MArray.ZeroSAPtr(ArrPtr(dataL()))
  Call MArray.ZeroSAPtr(ArrPtr(xx()))
  Set Subtract = X
End Function
''' <summary>Matrix subtraction.</summary>
'Public Shared Function op_Dif(ByVal left As Matrix, ByVal right As Matrix) As Matrix
Public Function op_Dif(left As Matrix, right As Matrix) As Matrix
  If left Is Nothing Then
    'Throw New ArgumentNullException("left")
    MsgBox "ArgumentNull: " & "left"
  End If
  If right Is Nothing Then
    'Throw New ArgumentNullException("right")
    MsgBox "ArgumentNull: " & "right"
  End If
  'Return Subtract(left, right)
  Set op_Dif = Subtract(left, right)
End Function 'Operator
''' <summary>Matrix-scalar multiplication.</summary>
'Public Shared Function Multiply(ByVal left As Matrix, ByVal right As Double) As Matrix
Public Function MultiplyD(left As Matrix, right As Double) As Matrix
  If left Is Nothing Then
    'Throw New ArgumentNullException("left")
    MsgBox "ArgumentNull: " & "left"
  End If
  Dim Rows As Long: Rows = left.Rows
  Dim Columns As Long: Columns = left.Columns
  Dim dataL() As Double: Call AssignArray(ArrPtr(dataL()), left.pArray)
  'Dim data() As Double: dataL()= left.pArray
  Dim X As Matrix: Set X = MNew.Matrix(Rows, Columns)
'  Dim xx() As Double: xx() = X.pArray
  Dim xx() As Double: Call AssignArray(ArrPtr(xx), X.pArray)
  Dim i As Long, j As Long
  For i = 0 To Rows - 1
    For j = 0 To Columns - 1
      xx(i, j) = data(i, j) * right
    Next
  Next
  'Return X
  Call MArray.ZeroSAPtr(ArrPtr(xx()))
  Call MArray.ZeroSAPtr(ArrPtr(dataL()))
  Set MultiplyD = X
End Function
''' <summary>Matrix-scalar multiplication.</summary>
'Public Shared Function op_MulD(ByVal left As Matrix, ByVal right As Double) As Matrix
Public Function op_MulD(left As Matrix, right As Double) As Matrix
  If left Is Nothing Then
    'Throw New ArgumentNullException("left")
    MsgBox "ArgumentNull: " & "left"
  End If
  Set op_MulD = MultiplyD(left, right)
End Function 'Operator
''' <summary>Matrix-matrix multiplication.</summary>
'Public Shared Function Multiply(ByVal left As Matrix, ByVal right As Matrix) As Matrix
Public Function Multiply(left As Matrix, right As Matrix) As Matrix
  If left Is Nothing Then
    'Throw New ArgumentNullException("left")
    MsgBox "Matrix::Multiply " & "ArgumentNull: " & "left"
  End If
  If right Is Nothing Then
    'Throw New ArgumentNullException("right")
    MsgBox "Matrix::Multiply " & "ArgumentNull: " & "right"
  End If
  Dim Rows As Long: Rows = left.Rows
  'Dim data() As Double: data() = left.pArray
  Dim dataL() As Double: Call AssignArray(ArrPtr(dataL()), left.pArray)
  Dim dataR() As Double: Call AssignArray(ArrPtr(dataR()), right.pArray)
  If right.Rows <> left.Columns Then
    'Throw New ArgumentException("Matrix dimensions are not valid.")
    'Exit Function
    MsgBox "Matrix::Multiply " & "Matrix dimensions are not valid."
  End If
  Dim cols As Long: cols = right.Columns
  Dim X As Matrix: Set X = MNew.Matrix(Rows, cols)
  'Dim xx() As Double: Set xx() = X.pArray
  Dim xx() As Double: Call AssignArray(ArrPtr(xx), X.pArray)
  Dim size As Long: size = left.Columns
  ReDim column(0) As Double ': column() = New_Double(size) ' {}
  Dim pvDataOld As Long: pvDataOld = VarPtr(column(0))
  ReDim row(size - 1) As Double '()
  Dim i As Long, j As Long, k As Long
  Dim s As Double
  For j = 0 To cols - 1
    'For k = 0 To size - 1
    'column(k) = right(k, j)
    'Achtung hier umgedreht als in VB.NET
    Call CopyColPtr(ArrPtr(column()), VarPtr(dataR(0, j)), size)
    'Next
    For i = 0 To Rows - 1
      'Achtung hier umgedreht als in VB.NET
      For k = 0 To size - 1
        row(k) = dataL(i, k)
      Next
      s = 0
      For k = 0 To size - 1
        s = s + (row(k) * column(k))
      Next
      xx(i, j) = s
      'X(i, j) = s
    Next
  Next
  'Return X
  Call CopyColPtr(ArrPtr(column()), pvDataOld, 1)
  Call MArray.ZeroSAPtr(ArrPtr(dataL()))
  Call MArray.ZeroSAPtr(ArrPtr(dataR()))
  Call MArray.ZeroSAPtr(ArrPtr(xx()))
  Set Multiply = X
End Function
''' <summary>Matrix-matrix multiplication.</summary>
'Public Shared Function op_MulM(ByVal left As Matrix, ByVal right As Matrix) As Matrix
Public Function op_MulM(left As Matrix, right As Matrix) As Matrix
  If left Is Nothing Then
    'Throw New ArgumentNullException("left")
    MsgBox "ArgumentNull: " & "left"
  End If
  If right Is Nothing Then
    'Throw New ArgumentNullException("right")
    MsgBox "ArgumentNull: " & "right"
  End If
  'Return Multiply(left, right)
  Set op_MulM = Multiply(left, right)
End Function 'Operator
''' <summary>Returns the LHS solution vetor if the matrix is square or the least squares solution otherwise.</summary>
Public Function Solve(rightHandSide As Matrix) As Matrix
  'Return IIf((m_rows = m_columns), New LuDecomposition(Me).Solve(rightHandSide), New QrDecomposition(Me).Solve(rightHandSide))
  If m_rows = m_columns Then
    Set Solve = MNew.LuDecomposition(Me).Solve(rightHandSide)
  Else
    Set Solve = MNew.QrDecomposition(Me).Solve(rightHandSide)
  End If
End Function
''' <summary>Inverse of the matrix if matrix is square, pseudoinverse otherwise.</summary>
'Public ReadOnly Property Inverse() As Matrix
Public Property Get Inverse() As Matrix
  'Get
    'Return Me.Solve(Diagonal(m_rows, m_rows, 1))
    Set Inverse = Me.Solve(Diagonal(m_rows, m_rows, 1))
  'End Get
End Property
''' <summary>Determinant if matrix is square.</summary>
'Public ReadOnly Property Determinant() As Double
Public Property Get Determinant() As Double
  'Get
    'Return New LuDecomposition(Me).Determinant
    Determinant = MNew.LuDecomposition(Me).Determinant
  'End Get
End Property
''' <summary>Returns the trace of the matrix.</summary>
''' <returns>Sum of the diagonal elements.</returns>
'Public ReadOnly Property Trace() As Double
Public Property Get Trace() As Double
  'Get
    'Dim ttrace As Double: ttrace = 0
    Dim i As Long
    For i = 0 To Math.MinL(m_rows, m_columns) - 1
      Trace = Trace + data(i, i)
    Next
    'Return ttrace
    'Trace = ttrace
  'End Get
End Property
''' <summary>Returns a matrix filled with random values.</summary>
'Public Shared Function RandomM(ByVal rows As Long, ByVal columns As Long) As Matrix
Public Function RandomM(ByVal Rows As Long, ByVal Columns As Long) As Matrix
  Dim X As Matrix: Set X = MNew.Matrix(Rows, Columns)
'  Dim xx() As Double: xx() = X.pArray
  Dim xx() As Double: Call AssignArray(ArrPtr(xx), X.pArray)
  Dim i As Long, j As Long
  For i = 0 To Rows - 1
    For j = 0 To Columns - 1
      xx(i, j) = Random.NextDouble()
    Next
  Next
  'Return X
  Call MArray.ZeroSAPtr(ArrPtr(xx()))
  Set RandomM = X
End Function
''' <summary>Returns a diagonal matrix of the given size.</summary>
'Public Shared Function Diagonal(ByVal rows As Long, ByVal columns As Long, ByVal value As Double) As Matrix
Public Function Diagonal(ByVal Rows As Long, ByVal Columns As Long, ByVal Value As Double) As Matrix
  Dim X As Matrix: Set X = MNew.Matrix(Rows, Columns)
'  Dim xx() As Double: xx() = X.pArray
  Dim xx() As Double: Call AssignArray(ArrPtr(xx), X.pArray)
  Dim i As Long, j As Long
  For i = 0 To Rows - 1
    For j = 0 To Columns - 1
      If i = j Then
        xx(i, j) = Value
      Else
        xx(i, j) = 0
      End If
    Next
  Next
  'Return X
  Call MArray.ZeroSAPtr(ArrPtr(xx()))
  Set Diagonal = X
End Function
''' <summary>Returns the matrix in a textual form.</summary>
'Public Overloads Overrides Function ToString() As String
Public Function ToString() As String
  'Dim writer As New StringWriter ': set writer New_StringWriter(CultureInfo.InvariantCulture)
  Dim writer As String
  Dim i As Long, j As Long
  For i = 0 To m_rows - 1
    For j = 0 To m_columns - 1
      'writer.Write (CStr(Me.data(i, j)) & " ")
      'eigentlich gehört hier noch eine Abfrage nach data(i,j)<= (1/eps) rein
      'damit man nicht so häßliche Zahlen drin hat, wie: -1.73092876837311E-16
      writer = writer & CStr(data(i, j)) & " "
    Next
    'writer.WriteLine '()
    writer = writer & vbNewLine
  Next
  'Return writer.ToString()
  ToString = writer '.ToString
  'End Using
End Function
'Private Shared Function Hypotenuse(ByVal a As Double, ByVal b As Double) As Double
Private Function Hypotenuse(ByVal A As Double, ByVal B As Double) As Double
  Dim r As Double
  If VBA.Math.Abs(A) > VBA.Math.Abs(B) Then
    r = B / A
    'Return Math.Abs(a) * Math.Sqrt(1 + r * r)
    Hypotenuse = VBA.Math.Abs(A) * VBA.Math.Sqr(1 + r * r): Exit Function
  End If
  If B <> 0 Then
    r = A / B
    'Return Math.Abs(b) * Math.Sqrt(1 + r * r)
    Hypotenuse = VBA.Math.Abs(B) * VBA.Math.Sqr(1 + r * r): Exit Function
  End If
  'Return 0
  Hypotenuse = 0
End Function
'  End Class
