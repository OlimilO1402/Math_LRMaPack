' ----------------------------------------------
' Lutz Roeder's Mapack for .NET, September 2000
' Adapted from Mapack for COM and Jama routines.
' http://www.aisto.com/roeder/dotnet
' ----------------------------------------------
Imports System
Namespace Mapack
	''' <summary>
	'''		Cholesky Decomposition of a symmetric, positive definite matrix.
	'''	</summary>
	''' <remarks>
	'''		For a symmetric, positive definite matrix <c>A</c>, the Cholesky decomposition is a
	'''		lower triangular matrix <c>L</c> so that <c>A = L * L'</c>.
	'''		If the matrix is not symmetric or positive definite, the constructor returns a partial 
	'''		decomposition and sets two internal variables that can be queried using the
	'''		<see cref="Symmetric"/> and <see cref="PositiveDefinite"/> properties.
	'''	</remarks>
	Public Class CholeskyDecomposition
		Private L As Matrix
		Private m_symmetric As Boolean
		Private m_positiveDefinite As Boolean
		''' <summary>Construct a Cholesky Decomposition.</summary>
		Public Sub New(ByVal value As Matrix)
			If value Is Nothing Then
				Throw New ArgumentNullException("value")
			End If
			If Not value.Square Then
				Throw New ArgumentException("Matrix is not square.", "value")
			End If
      Dim dimension As Integer = value.Rows
      Dim L As Matrix
			L = New Matrix(dimension, dimension)
			Dim a As Double()() = value.Array
      Dim ll As Double()() = L.Array
			Me.m_positiveDefinite = True
      Me.m_symmetric = True
      Dim i As Integer, j As Integer, k As Integer
      For j = 0 To dimension - 1
        Dim Lrowj As Double() = ll(j)
        Dim d As Double = 0
        For k = 0 To j - 1
          Dim Lrowk As Double() = ll(k)
          Dim s As Double = 0
          For i = 0 To k - 1
            s = s + Lrowk(i) * Lrowj(i)
          Next
          s = (a(j)(k) - s) / ll(k)(k)
          Lrowj(k) = s
          d = d + s * s
          Me.m_symmetric = Me.m_symmetric And (a(k)(j) = a(j)(k))
        Next
        d = a(j)(j) - d
        Me.m_positiveDefinite = Me.m_positiveDefinite And (d > 0)
        ll(j)(j) = Math.Sqrt(Math.Max(d, 0))
        For k = j + 1 To dimension - 1
          ll(j)(k) = 0
        Next
      Next
		End Sub
		''' <summary>Returns <see langword="true"/> if the matrix is symmetric.</summary>
		Public ReadOnly Property Symmetric() As Boolean
			Get
				Return Me.m_symmetric
			End Get
		End Property
		''' <summary>Returns <see langword="true"/> if the matrix is positive definite.</summary>
		Public ReadOnly Property PositiveDefinite() As Boolean
			Get
				Return Me.m_positiveDefinite
			End Get
		End Property
		''' <summary>Returns the left triangular factor <c>L</c> so that <c>A = L * L'</c>.</summary>
		Public ReadOnly Property LeftTriangularFactor() As Matrix
			Get
				Return Me.L
			End Get
		End Property
		''' <summary>Solves a set of equation systems of type <c>A * X = B</c>.</summary>
		''' <param name="value">Right hand side matrix with as many rows as <c>A</c> and any number of columns.</param>
		''' <returns>Matrix <c>X</c> so that <c>L * L' * X = B</c>.</returns>
		''' <exception cref="T:System.ArgumentException">Matrix dimensions do not match.</exception>
		''' <exception cref="T:System.InvalidOperationException">Matrix is not symmetrix and positive definite.</exception>
		Public Function Solve(ByVal value As Matrix) As Matrix
			If value Is Nothing Then
				Throw New ArgumentNullException("value")
			End If
      If value.Rows <> Me.L.Rows Then
        Throw New ArgumentException("Matrix dimensions do not match.")
      End If
      If Not Me.m_symmetric Then
        Throw New InvalidOperationException("Matrix is not symmetric.")
      End If
      If Not Me.m_positiveDefinite Then
        Throw New InvalidOperationException("Matrix is not positive definite.")
      End If
      Dim dimension As Integer = Me.L.Rows
      Dim count As Integer = value.Columns
      Dim B As Matrix = DirectCast(value.Clone(), Matrix)
      Dim ll As Double()() = L.Array
      Dim i As Integer, j As Integer, k As Integer
      For k = 0 To L.Rows - 1
        ' Solve L*Y = B;
        For i = k + 1 To dimension - 1
          For j = 0 To count - 1
            B(i, j) = B(i, j) - B(k, j) * ll(i)(k)
          Next
        Next
        For j = 0 To count - 1
          B(k, j) = B(k, j) / ll(k)(k)
        Next
      Next
      For k = dimension - 1 To 0 Step -1
        ' Solve L'*X = Y;
        For j = 0 To count - 1
          B(k, j) = B(k, j) / ll(k)(k)
        Next
        For i = 0 To k - 1
          For j = 0 To count - 1
            B(i, j) = B(i, j) - B(k, j) * ll(k)(i)
          Next
        Next
      Next
      Return B
		End Function
	End Class
End Namespace
