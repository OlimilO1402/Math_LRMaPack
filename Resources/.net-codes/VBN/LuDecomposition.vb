' ----------------------------------------------
' Lutz Roeder's Mapack for .NET, September 2000
' Adapted from Mapack for COM and Jama routines.
' http://www.aisto.com/roeder/dotnet
' ----------------------------------------------
Imports System
Namespace Mapack
	''' <summary>
	'''   LU decomposition of a rectangular matrix.
	''' </summary>
	''' <remarks>
	'''   For an m-by-n matrix <c>A</c> with m >= n, the LU decomposition is an m-by-n
	'''   unit lower triangular matrix <c>L</c>, an n-by-n upper triangular matrix <c>U</c>,
	'''   and a permutation vector <c>piv</c> of length m so that <c>A(piv)=L*U</c>.
	'''   If m &lt; n, then <c>L</c> is m-by-m and <c>U</c> is m-by-n.
	'''   The LU decompostion with pivoting always exists, even if the matrix is
	'''   singular, so the constructor will never fail.  The primary use of the
	'''   LU decomposition is in the solution of square systems of simultaneous
	'''   linear equations. This will fail if <see cref="NonSingular"/> returns <see langword="false"/>.
	''' </remarks>
	Public Class LuDecomposition
		Private LU As Matrix
		Private pivotSign As Integer
		Private pivotVector As Integer()
		''' <summary>Construct a LU decomposition.</summary>	
		Public Sub New(ByVal value As Matrix)
			If value Is Nothing Then
				Throw New ArgumentNullException("value")
			End If
			Me.LU = DirectCast(value.Clone(), Matrix)
      Dim lu As Double()() = Me.LU.Array
			Dim rows As Integer = value.Rows
			Dim columns As Integer = value.Columns
      pivotVector = New Integer(rows - 1) {} '(rows-1) !!!!!!!!!!!! vorher Fehler(rows)
      Dim i As Integer, j As Integer, k As Integer
      Dim p As Integer
      For i = 0 To rows - 1
        pivotVector(i) = i
      Next
      pivotSign = 1
      Dim LUrowi As Double()
      Dim LUcolj As Double() = New Double(rows) {}
      Dim kmax As Integer
      Dim s As Double
      For j = 0 To columns - 1
        ' Outer loop.
        For i = 0 To rows - 1
          ' Make a copy of the j-th column to localize references.
          LUcolj(i) = lu(i)(j)
        Next
        For i = 0 To rows - 1
          ' Apply previous transformations.
          LUrowi = lu(i)
          ' Most of the time is spent in the following dot product.
          kmax = Math.Min(i, j)
          s = 0
          For k = 0 To kmax - 1
            s = s + LUrowi(k) * LUcolj(k)
          Next
          LUcolj(i) = LUcolj(i) - s
          LUrowi(j) = LUcolj(i)
        Next
        ' Find pivot and exchange if necessary.
        p = j
        For i = j + 1 To rows - 1
          If Math.Abs(LUcolj(i)) > Math.Abs(LUcolj(p)) Then
            p = i
          End If
        Next
        If p <> j Then
          For k = 0 To columns - 1
            Dim t As Double = lu(p)(k)
            lu(p)(k) = lu(j)(k)
            lu(j)(k) = t
          Next
          Dim v As Integer = pivotVector(p)
          pivotVector(p) = pivotVector(j)
          pivotVector(j) = v
          pivotSign = -pivotSign
        End If
        ' Compute multipliers.
        If j < rows And lu(j)(j) <> 0 Then
          For i = j + 1 To rows - 1
            lu(i)(j) = lu(i)(j) / lu(j)(j)
          Next
        End If
      Next
		End Sub
		''' <summary>Returns if the matrix is non-singular.</summary>
		Public ReadOnly Property NonSingular() As Boolean
			Get
				For j As Integer = 0 To LU.Columns - 1
					If LU(j, j) = 0 Then
						Return False
					End If
				Next
				Return True
			End Get
		End Property
		''' <summary>Returns the determinant of the matrix.</summary>
		Public ReadOnly Property Determinant() As Double
			Get
				If LU.Rows <> LU.Columns Then
					Throw New ArgumentException("Matrix must be square.")
				End If
        'Dim ddeterminant As Double = DirectCast(pivotSign, Double)
        Dim ddeterminant As Double = CDbl(pivotSign)
				For j As Integer = 0 To LU.Columns - 1
          ddeterminant *= LU(j, j)
				Next
        Return ddeterminant
			End Get
		End Property
		''' <summary>Returns the lower triangular factor <c>L</c> with <c>A=LU</c>.</summary>
		Public ReadOnly Property LowerTriangularFactor() As Matrix
			Get
				Dim rows As Integer = LU.Rows
				Dim columns As Integer = LU.Columns
				Dim X As New Matrix(rows, columns)
				For i As Integer = 0 To rows - 1
					For j As Integer = 0 To columns - 1
						If i > j Then
							X(i, j) = LU(i, j)
ElseIf i = j Then
							X(i, j) = 1
						Else
							X(i, j) = 0
						End If
					Next
				Next
				Return X
			End Get
		End Property
		''' <summary>Returns the lower triangular factor <c>L</c> with <c>A=LU</c>.</summary>
		Public ReadOnly Property UpperTriangularFactor() As Matrix
			Get
				Dim rows As Integer = LU.Rows
				Dim columns As Integer = LU.Columns
				Dim X As New Matrix(rows, columns)
				For i As Integer = 0 To rows - 1
					For j As Integer = 0 To columns - 1
						If i <= j Then
							X(i, j) = LU(i, j)
						Else
							X(i, j) = 0
						End If
					Next
				Next
				Return X
			End Get
		End Property
		''' <summary>Returns the pivot permuation vector.</summary>
		Public ReadOnly Property PivotPermutationVector() As Double()
			Get
				Dim rows As Integer = LU.Rows
				Dim p As Double() = New Double(rows) {}
				For i As Integer = 0 To rows - 1
          'p(i) = DirectCast(Me.pivotVector(i), Double)
          p(i) = CDbl(Me.pivotVector(i))
				Next
				Return p
			End Get
		End Property
		''' <summary>Solves a set of equation systems of type <c>A * X = B</c>.</summary>
		''' <param name="value">Right hand side matrix with as many rows as <c>A</c> and any number of columns.</param>
		''' <returns>Matrix <c>X</c> so that <c>L * U * X = B</c>.</returns>
		Public Function Solve(ByVal value As Matrix) As Matrix
			If value Is Nothing Then
				Throw New ArgumentNullException("value")
			End If
			If value.Rows <> Me.LU.Rows Then
				Throw New ArgumentException("Invalid matrix dimensions.", "value")
			End If
			If Not Me.NonSingular Then
				Throw New InvalidOperationException("Matrix is singular")
			End If
			' Copy right hand side with pivoting
			Dim count As Integer = value.Columns
			Dim X As Matrix = value.Submatrix(pivotVector, 0, count - 1)
      Dim rows As Integer = Me.LU.Rows
      Dim columns As Integer = Me.LU.Columns
      Dim lu As Double()() = Me.LU.Array
			For k As Integer = 0 To columns - 1
				' Solve L*Y = B(piv,:)
				For i As Integer = k + 1 To columns - 1
					For j As Integer = 0 To count - 1
						X(i, j) -= X(k, j) * lu(i)(k)
					Next
				Next
			Next
			For k As Integer = columns - 1 To 0 Step -1
				' Solve U*X = Y;
				For j As Integer = 0 To count - 1
					X(k, j) /= lu(k)(k)
				Next
				For i As Integer = 0 To k - 1
					For j As Integer = 0 To count - 1
						X(i, j) -= X(k, j) * lu(i)(k)
					Next
				Next
			Next
			Return X
		End Function
	End Class
End Namespace
