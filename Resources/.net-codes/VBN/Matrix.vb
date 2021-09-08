' ----------------------------------------------
' Lutz Roeder's Mapack for .NET, September 2000
' Adapted from Mapack for COM and Jama routines.
' http://www.aisto.com/roeder/dotnet
' ----------------------------------------------
Imports System
Imports System.IO
Imports System.Globalization
Namespace Mapack
	''' <summary>Matrix provides the fundamental operations of numerical linear algebra.</summary>
	Public Class Matrix
		Private data As Double()()
		Private m_rows As Integer
		Private m_columns As Integer
		Private Shared random As New Random()
		''' <summary>Constructs an empty matrix of the given size.</summary>
		''' <param name="rows">Number of rows.</param>
		''' <param name="columns">Number of columns.</param>
		Public Sub New(ByVal rows As Integer, ByVal columns As Integer)
			Me.m_rows = rows
			Me.m_columns = columns
      Me.data = New Double(rows)() {}
      Dim i As Integer
      For i = 0 To rows - 1
        Me.data(i) = New Double(columns) {}
      Next
		End Sub
		''' <summary>Constructs a matrix of the given size and assigns a given value to all diagonal elements.</summary>
		''' <param name="rows">Number of rows.</param>
		''' <param name="columns">Number of columns.</param>
		''' <param name="value">Value to assign to the diagnoal elements.</param>
		Public Sub New(ByVal rows As Integer, ByVal columns As Integer, ByVal value As Double)
			Me.m_rows = rows
			Me.m_columns = columns
      Me.data = New Double(rows)() {}
      Dim i As Integer
			For i = 0 To rows - 1
				data(i) = New Double(columns) {}
			Next
      For i = 0 To rows - 1
        data(i)(i) = value
      Next
		End Sub
		''' <summary>Constructs a matrix from the given array.</summary>
		''' <param name="value">The array the matrix gets constructed from.</param>
		<CLSCompliant(False)> _
		Public Sub New(ByVal value As Double()())
			Me.m_rows = value.Length
      Me.m_columns = value(0).Length
      Dim i As Integer
      For i = 0 To m_rows - 1
        If value(i).Length <> m_columns Then
          Throw New ArgumentException("Argument out of range.")
        End If
      Next
      Me.data = value
		End Sub
		''' <summary>Determines weather two instances are equal.</summary>
		Public Overloads Overrides Function Equals(ByVal obj As Object) As Boolean
			Return Equals(Me, DirectCast(obj, Matrix))
		End Function
		''' <summary>Determines weather two instances are equal.</summary>
    Public Overloads Shared Function Equals(ByVal left As Matrix, ByVal right As Matrix) As Boolean
      If (DirectCast(left, Object)) = (DirectCast(right, Object)) Then
        Return True
      End If
      If ((DirectCast(left, Object)) Is Nothing) OrElse ((DirectCast(right, Object)) Is Nothing) Then
        Return False
      End If
      If (left.Rows <> right.Rows) OrElse (left.Columns <> right.Columns) Then
        Return False
      End If
      Dim i As Integer, j As Integer
      For i = 0 To left.Rows - 1
        For j = 0 To left.Columns - 1
          If left(i, j) <> right(i, j) Then
            Return False
          End If
        Next
      Next
      Return True
    End Function
    ''' <summary>Serves as a hash function for a particular type, suitable for use in hashing algorithms and data structures like a hash table.</summary>
    Public Overloads Overrides Function GetHashCode() As Integer
      Return (Me.Rows + Me.Columns)
    End Function
    Friend ReadOnly Property Array() As Double()()
      Get
        Return Me.data
      End Get
    End Property
    ''' <summary>Returns the number of columns.</summary>
    Public ReadOnly Property Rows() As Integer
      Get
        Return Me.m_rows
      End Get
    End Property
    ''' <summary>Returns the number of columns.</summary>
    Public ReadOnly Property Columns() As Integer
      Get
        Return Me.m_columns
      End Get
    End Property
    ''' <summary>Return <see langword="true"/> if the matrix is a square matrix.</summary>
    Public ReadOnly Property Square() As Boolean
      Get
        Return (m_rows = m_columns)
      End Get
    End Property
    ''' <summary>Returns <see langword="true"/> if the matrix is symmetric.</summary>
    Public ReadOnly Property Symmetric() As Boolean
      Get
        If Me.Square Then
          For i As Integer = 0 To m_rows - 1
            For j As Integer = 0 To i
              If data(i)(j) <> data(j)(i) Then
                Return False
              End If
            Next
          Next
          Return True
        End If
        Return False
      End Get
    End Property
    ''' <summary>Access the value at the given location.</summary>
    Default Public Property Item(ByVal row As Integer, ByVal column As Integer) As Double
      Get
        Return Me.data(row)(column)
      End Get
      Set(ByVal Value As Double)
        Me.data(row)(column) = Value
      End Set
    End Property
    ''' <summary>Returns a sub matrix extracted from the current matrix.</summary>
    ''' <param name="startRow">Start row index</param>
    ''' <param name="endRow">End row index</param>
    ''' <param name="startColumn">Start column index</param>
    ''' <param name="endColumn">End column index</param>
    Public Function Submatrix(ByVal startRow As Integer, ByVal endRow As Integer, ByVal startColumn As Integer, ByVal endColumn As Integer) As Matrix
      If (startRow > endRow) OrElse (startColumn > endColumn) OrElse (startRow < 0) OrElse (startRow >= Me.m_rows) OrElse (endRow < 0) OrElse (endRow >= Me.m_rows) OrElse (startColumn < 0) OrElse (startColumn >= Me.m_columns) OrElse (endColumn < 0) OrElse (endColumn >= Me.m_columns) Then
        Throw New ArgumentException("Argument out of range.")
      End If
      Dim X As New Matrix(endRow - startRow + 1, endColumn - startColumn + 1)
      Dim xx As Double()() = X.Array
      Dim i As Integer
      For i = startRow To endRow
        For j As Integer = startColumn To endColumn
          xx(i - startRow)(j - startColumn) = data(i)(j)
        Next
      Next
      Return X
    End Function
    ''' <summary>Returns a sub matrix extracted from the current matrix.</summary>
    ''' <param name="rowIndexes">Array of row indices</param>
    ''' <param name="columnIndexes">Array of column indices</param>
    Public Function Submatrix(ByVal rowIndexes As Integer(), ByVal columnIndexes As Integer()) As Matrix
      Dim X As New Matrix(rowIndexes.Length, columnIndexes.Length)
      Dim xx As Double()() = X.Array
      Dim i As Integer, j As Integer
      For i = 0 To rowIndexes.Length - 1
        For j = 0 To columnIndexes.Length - 1
          If (rowIndexes(i) < 0) OrElse (rowIndexes(i) >= m_rows) OrElse (columnIndexes(j) < 0) OrElse (columnIndexes(j) >= m_columns) Then
            Throw New ArgumentException("Argument out of range.")
          End If
          xx(i)(j) = data(rowIndexes(i))(columnIndexes(j))
        Next
      Next
      Return X
    End Function
    ''' <summary>Returns a sub matrix extracted from the current matrix.</summary>
    ''' <param name="i0">Starttial row index</param>
    ''' <param name="i1">End row index</param>
    ''' <param name="c">Array of row indices</param>
    Public Function Submatrix(ByVal i0 As Integer, ByVal i1 As Integer, ByVal c As Integer()) As Matrix
      If (i0 > i1) OrElse (i0 < 0) OrElse (i0 >= Me.m_rows) OrElse (i1 < 0) OrElse (i1 >= Me.m_rows) Then
        Throw New ArgumentException("Argument out of range.")
      End If
      Dim X As New Matrix(i1 - i0 + 1, c.Length)
      Dim xx As Double()() = X.Array
      For i As Integer = i0 To i1
        For j As Integer = 0 To c.Length - 1
          If (c(j) < 0) OrElse (c(j) >= m_columns) Then
            Throw New ArgumentException("Argument out of range.")
          End If
          xx(i - i0)(j) = data(i)(c(j))
        Next
      Next
      Return X
    End Function
    ''' <summary>Returns a sub matrix extracted from the current matrix.</summary>
    ''' <param name="r">Array of row indices</param>
    ''' <param name="j0">Start column index</param>
    ''' <param name="j1">End column index</param>
    Public Function Submatrix(ByVal r As Integer(), ByVal j0 As Integer, ByVal j1 As Integer) As Matrix
      If (j0 > j1) OrElse (j0 < 0) OrElse (j0 >= m_columns) OrElse (j1 < 0) OrElse (j1 >= m_columns) Then
        Throw New ArgumentException("Argument out of range.")
      End If
      Dim X As New Matrix(r.Length, j1 - j0 + 1)
      Dim xx As Double()() = X.Array
      Dim i As Integer, j As Integer
      For i = 0 To r.Length - 1
        For j = j0 To j1
          If (r(i) < 0) OrElse (r(i) >= Me.m_rows) Then
            Throw New ArgumentException("Argument out of range.")
          End If
          xx(i)(j - j0) = data(r(i))(j)
        Next
      Next
      Return X
    End Function
    ''' <summary>Creates a copy of the matrix.</summary>
    Public Function Clone() As Matrix
      Dim X As New Matrix(m_rows, m_columns)
      Dim xx As Double()() = X.Array
      Dim i As Integer, j As Integer
      For i = 0 To m_rows - 1
        For j = 0 To m_columns - 1
          xx(i)(j) = data(i)(j)
        Next
      Next
      Return X
    End Function
    ''' <summary>Returns the transposed matrix.</summary>
    Public Function Transpose() As Matrix
      Dim X As New Matrix(m_columns, m_rows)
      Dim xx As Double()() = X.Array
      Dim i As Integer, j As Integer
      For i = 0 To m_rows - 1
        For j = 0 To m_columns - 1
          xx(j)(i) = data(i)(j)
        Next
      Next
      Return X
    End Function
    ''' <summary>Returns the One Norm for the matrix.</summary>
    ''' <value>The maximum column sum.</value>
    Public ReadOnly Property Norm1() As Double
      Get
        Dim f As Double = 0
        Dim s As Double = 0
        Dim i As Integer, j As Integer
        For j = 0 To m_columns - 1
          s = 0
          For i = 0 To m_rows - 1
            s = s + Math.Abs(data(i)(j))
          Next
          f = Math.Max(f, s)
        Next
        Return f
      End Get
    End Property
    ''' <summary>Returns the Infinity Norm for the matrix.</summary>
    ''' <value>The maximum row sum.</value>
    Public ReadOnly Property InfinityNorm() As Double
      Get
        Dim f As Double = 0
        Dim s As Double = 0
        Dim i As Integer, j As Integer
        For i = 0 To m_rows - 1
          s = 0
          For j = 0 To m_columns - 1
            s += Math.Abs(data(i)(j))
          Next
          f = Math.Max(f, s)
        Next
        Return f
      End Get
    End Property
    ''' <summary>Returns the Frobenius Norm for the matrix.</summary>
    ''' <value>The square root of sum of squares of all elements.</value>
    Public ReadOnly Property FrobeniusNorm() As Double
      Get
        Dim f As Double = 0
        Dim i As Integer, j As Integer
        For i = 0 To m_rows - 1
          For j = 0 To m_columns - 1
            f = Hypotenuse(f, data(i)(j))
          Next
        Next
        Return f
      End Get
    End Property
    ''' <summary>Unary minus.</summary>
    Public Shared Function Negate(ByVal value As Matrix) As Matrix
      If value Is Nothing Then
        Throw New ArgumentNullException("value")
      End If
      Dim rows As Integer = value.Rows
      Dim columns As Integer = value.Columns
      Dim data As Double()() = value.Array
      Dim X As New Matrix(rows, columns)
      Dim xx As Double()() = X.Array
      Dim i As Integer, j As Integer
      For i = 0 To rows - 1
        For j = 0 To columns - 1
          xx(i)(j) = -data(i)(j)
        Next
      Next
      Return X
    End Function
    ''' <summary>Unary minus.</summary>
    Public Shared Function op_Dif(ByVal value As Matrix) As Matrix
      If value Is Nothing Then
        Throw New ArgumentNullException("value")
      End If
      Return Negate(value)
    End Function 'Operator
    ''' <summary>Matrix equality.</summary>
    Public Shared Function op_Equal(ByVal left As Matrix, ByVal right As Matrix) As Boolean
      Return Equals(left, right)
    End Function 'Operator
    ''' <summary>Matrix inequality.</summary>
    Public Shared Function op_Unequal(ByVal left As Matrix, ByVal right As Matrix) As Boolean
      Return Not Equals(left, right)
    End Function 'Operator
    ''' <summary>Matrix addition.</summary>
    Public Shared Function Add(ByVal left As Matrix, ByVal right As Matrix) As Matrix
      If left Is Nothing Then
        Throw New ArgumentNullException("left")
      End If
      If right Is Nothing Then
        Throw New ArgumentNullException("right")
      End If
      Dim rows As Integer = left.Rows
      Dim columns As Integer = left.Columns
      Dim data As Double()() = left.Array
      If (rows <> right.Rows) OrElse (columns <> right.Columns) Then
        Throw New ArgumentException("Matrix dimension do not match.")
      End If
      Dim X As New Matrix(rows, columns)
      Dim xx As Double()() = X.Array
      Dim i As Integer, j As Integer
      For i = 0 To rows - 1
        For j = 0 To columns - 1
          xx(i)(j) = data(i)(j) + right(i, j)
        Next
      Next
      Return X
    End Function
    ''' <summary>Matrix addition.</summary>
    Public Shared Function op_Plus(ByVal left As Matrix, ByVal right As Matrix) As Matrix
      If left Is Nothing Then
        Throw New ArgumentNullException("left")
      End If
      If right Is Nothing Then
        Throw New ArgumentNullException("right")
      End If
      Return Add(left, right)
    End Function 'Operator
    ''' <summary>Matrix subtraction.</summary>
    Public Shared Function Subtract(ByVal left As Matrix, ByVal right As Matrix) As Matrix
      If left Is Nothing Then
        Throw New ArgumentNullException("left")
      End If
      If right Is Nothing Then
        Throw New ArgumentNullException("right")
      End If
      Dim rows As Integer = left.Rows
      Dim columns As Integer = left.Columns
      Dim data As Double()() = left.Array
      If (rows <> right.Rows) OrElse (columns <> right.Columns) Then
        Throw New ArgumentException("Matrix dimension do not match.")
      End If
      Dim X As New Matrix(rows, columns)
      Dim xx As Double()() = X.Array
      Dim i As Integer, j As Integer
      For i = 0 To rows - 1
        For j = 0 To columns - 1
          xx(i)(j) = data(i)(j) - right(i, j)
        Next
      Next
      Return X
    End Function
    ''' <summary>Matrix subtraction.</summary>
    Public Shared Function op_Dif(ByVal left As Matrix, ByVal right As Matrix) As Matrix
      If left Is Nothing Then
        Throw New ArgumentNullException("left")
      End If
      If right Is Nothing Then
        Throw New ArgumentNullException("right")
      End If
      Return Subtract(left, right)
    End Function 'Operator
    ''' <summary>Matrix-scalar multiplication.</summary>
    Public Shared Function Multiply(ByVal left As Matrix, ByVal right As Double) As Matrix
      If left Is Nothing Then
        Throw New ArgumentNullException("left")
      End If
      Dim rows As Integer = left.Rows
      Dim columns As Integer = left.Columns
      Dim data As Double()() = left.Array
      Dim X As New Matrix(rows, columns)
      Dim xx As Double()() = X.Array
      Dim i As Integer, j As Integer
      For i = 0 To rows - 1
        For j = 0 To columns - 1
          xx(i)(j) = data(i)(j) * right
        Next
      Next
      Return X
    End Function
    ''' <summary>Matrix-scalar multiplication.</summary>
    Public Shared Function op_MulD(ByVal left As Matrix, ByVal right As Double) As Matrix
      If left Is Nothing Then
        Throw New ArgumentNullException("left")
      End If
      Return Multiply(left, right)
    End Function 'Operator
    ''' <summary>Matrix-matrix multiplication.</summary>
    Public Shared Function Multiply(ByVal left As Matrix, ByVal right As Matrix) As Matrix
      If left Is Nothing Then
        Throw New ArgumentNullException("left")
      End If
      If right Is Nothing Then
        Throw New ArgumentNullException("right")
      End If
      Dim rows As Integer = left.Rows
      Dim data As Double()() = left.Array
      If right.Rows <> left.Columns Then
        'Throw New ArgumentException("Matrix dimensions are not valid.")
        Exit Function
      End If
      Dim columns As Integer = right.Columns
      Dim X As New Matrix(rows, columns)
      Dim xx As Double()() = X.Array
      Dim size As Integer = left.Columns
      Dim column As Double() = New Double(size) {}
      Dim i As Integer, j As Integer, k As Integer
      Dim row As Double()
      Dim s As Double
      For j = 0 To columns - 1
        For k = 0 To size - 1
          column(k) = right(k, j)
        Next
        For i = 0 To rows - 1
          row = data(i)
          s = 0
          For k = 0 To size - 1
            s = s + (row(k) * column(k))
          Next
          xx(i)(j) = s
        Next
      Next
      Return X
    End Function
    ''' <summary>Matrix-matrix multiplication.</summary>
    Public Shared Function op_MulM(ByVal left As Matrix, ByVal right As Matrix) As Matrix
      If left Is Nothing Then
        Throw New ArgumentNullException("left")
      End If
      If right Is Nothing Then
        Throw New ArgumentNullException("right")
      End If
      Return Multiply(left, right)
    End Function 'Operator
    ''' <summary>Returns the LHS solution vetor if the matrix is square or the least squares solution otherwise.</summary>
    Public Function Solve(ByVal rightHandSide As Matrix) As Matrix
      Return IIf((m_rows = m_columns), New LuDecomposition(Me).Solve(rightHandSide), New QrDecomposition(Me).Solve(rightHandSide))
    End Function
    ''' <summary>Inverse of the matrix if matrix is square, pseudoinverse otherwise.</summary>
    Public ReadOnly Property Inverse() As Matrix
      Get
        Return Me.Solve(Diagonal(m_rows, m_rows, 1))
      End Get
    End Property
    ''' <summary>Determinant if matrix is square.</summary>
    Public ReadOnly Property Determinant() As Double
      Get
        Return New LuDecomposition(Me).Determinant
      End Get
    End Property
    ''' <summary>Returns the trace of the matrix.</summary>
    ''' <returns>Sum of the diagonal elements.</returns>
    Public ReadOnly Property Trace() As Double
      Get
        Dim ttrace As Double = 0
        Dim i As Integer
        For i = 0 To Math.Min(m_rows, m_columns) - 1
          ttrace = ttrace + data(i)(i)
        Next
        Return ttrace
      End Get
    End Property
    ''' <summary>Returns a matrix filled with random values.</summary>
    Public Shared Function RandomM(ByVal rows As Integer, ByVal columns As Integer) As Matrix
      Dim X As New Matrix(rows, columns)
      Dim xx As Double()() = X.Array
      Dim i As Integer, j As Integer
      For i = 0 To rows - 1
        For j = 0 To columns - 1
          xx(i)(j) = random.NextDouble()
        Next
      Next
      Return X
    End Function
    ''' <summary>Returns a diagonal matrix of the given size.</summary>
    Public Shared Function Diagonal(ByVal rows As Integer, ByVal columns As Integer, ByVal value As Double) As Matrix
      Dim X As New Matrix(rows, columns)
      Dim xx As Double()() = X.Array
      Dim i As Integer, j As Integer
      For i = 0 To rows - 1
        For j = 0 To columns - 1
          If i = j Then
            xx(i)(j) = value
          Else
            xx(i)(j) = 0
          End If
        Next
      Next
      Return X
    End Function
    ''' <summary>Returns the matrix in a textual form.</summary>
    Public Overloads Overrides Function ToString() As String
      Dim writer As New StringWriter(CultureInfo.InvariantCulture)
      Dim i As Integer, j As Integer
      For i = 0 To m_rows - 1
        For j = 0 To m_columns - 1
          writer.Write(CStr(Me.data(i)(j)) & " ")
        Next
        writer.WriteLine()
      Next
      Return writer.ToString()
      'End Using
    End Function
    Private Shared Function Hypotenuse(ByVal a As Double, ByVal b As Double) As Double
      Dim r As Double
      If Math.Abs(a) > Math.Abs(b) Then
        r = b / a
        Return Math.Abs(a) * Math.Sqrt(1 + r * r)
      End If
      If b <> 0 Then
        r = a / b
        Return Math.Abs(b) * Math.Sqrt(1 + r * r)
      End If
      Return 0
    End Function
  End Class
End Namespace
