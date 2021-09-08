' ----------------------------------------------
' Lutz Roeder's Mapack for .NET, September 2000
' Adapted from Mapack for COM and Jama routines.
' http://www.aisto.com/roeder/dotnet
' ----------------------------------------------
Imports System
Namespace Mapack
	''' <summary>
	'''	  QR decomposition for a rectangular matrix.
	''' </summary>
	''' <remarks>
	'''   For an m-by-n matrix <c>A</c> with <c>m &gt;= n</c>, the QR decomposition is an m-by-n
	'''   orthogonal matrix <c>Q</c> and an n-by-n upper triangular 
	'''   matrix <c>R</c> so that <c>A = Q * R</c>.
	'''   The QR decompostion always exists, even if the matrix does not have
	'''   full rank, so the constructor will never fail.  The primary use of the
	'''   QR decomposition is in the least squares solution of nonsquare systems
	'''   of simultaneous linear equations.
	'''   This will fail if <see cref="FullRank"/> returns <see langword="false"/>.
	''' </remarks>
	Public Class QrDecomposition
		Private QR As Matrix
		Private Rdiag As Double()
		''' <summary>Construct a QR decomposition.</summary>	
		Public Sub New(ByVal value As Matrix)
			If value Is Nothing Then
				Throw New ArgumentNullException("value")
			End If
			Me.QR = DirectCast(value.Clone(), Matrix)
			Dim qr As Double()() = Me.QR.Array
			Dim m As Integer = value.Rows
			Dim n As Integer = value.Columns
			Me.Rdiag = New Double(n) {}
			For k As Integer = 0 To n - 1
				' Compute 2-norm of k-th column without under/overflow.
				Dim nrm As Double = 0
				For i As Integer = k To m - 1
					nrm = Hypotenuse(nrm, qr(i)(k))
				Next
				If nrm <> 0 Then
					' Form k-th Householder vector.
					If qr(k)(k) < 0 Then
						nrm = -nrm
					End If
					For i As Integer = k To m - 1
						qr(i)(k) /= nrm
					Next
					qr(k)(k) += 1
					For j As Integer = k + 1 To n - 1
						' Apply transformation to remaining columns.
						Dim s As Double = 0
						For i As Integer = k To m - 1
							s += qr(i)(k) * qr(i)(j)
						Next
						s = -s / qr(k)(k)
						For i As Integer = k To m - 1
							qr(i)(j) += s * qr(i)(k)
						Next
					Next
				End If
				Me.Rdiag(k) = -nrm
			Next
		End Sub
		''' <summary>Least squares solution of <c>A * X = B</c></summary>
		''' <param name="value">Right-hand-side matrix with as many rows as <c>A</c> and any number of columns.</param>
		''' <returns>A matrix that minimized the two norm of <c>Q * R * X - B</c>.</returns>
		''' <exception cref="T:System.ArgumentException">Matrix row dimensions must be the same.</exception>
		''' <exception cref="T:System.InvalidOperationException">Matrix is rank deficient.</exception>
		Public Function Solve(ByVal value As Matrix) As Matrix
			If value Is Nothing Then
				Throw New ArgumentNullException("value")
			End If
      If value.Rows <> Me.QR.Rows Then
        Throw New ArgumentException("Matrix row dimensions must agree.")
      End If
      If Not Me.FullRank Then
        Throw New InvalidOperationException("Matrix is rank deficient.")
      End If
      ' Copy right hand side
      Dim count As Integer = value.Columns
      Dim X As Matrix = value.Clone()
      Dim m As Integer = Me.QR.Rows
      Dim n As Integer = Me.QR.Columns
      Dim qr As Double()() = Me.QR.Array
      For k As Integer = 0 To n - 1
        ' Compute Y = transpose(Q)*B
        For j As Integer = 0 To count - 1
          Dim s As Double = 0
          For i As Integer = k To m - 1
            s += qr(i)(k) * X(i, j)
          Next
          s = -s / qr(k)(k)
          For i As Integer = k To m - 1
            X(i, j) += s * qr(i)(k)
          Next
        Next
      Next
      For k As Integer = n - 1 To 0 Step -1
        ' Solve R*X = Y;
        For j As Integer = 0 To count - 1
          X(k, j) /= Rdiag(k)
        Next
        For i As Integer = 0 To k - 1
          For j As Integer = 0 To count - 1
            X(i, j) -= X(k, j) * qr(i)(k)
          Next
        Next
      Next
      Return X.Submatrix(0, n - 1, 0, count - 1)
		End Function
		''' <summary>Shows if the matrix <c>A</c> is of full rank.</summary>
		''' <value>The value is <see langword="true"/> if <c>R</c>, and hence <c>A</c>, has full rank.</value>
		Public ReadOnly Property FullRank() As Boolean
			Get
        Dim columns As Integer = Me.QR.Columns
        Dim i As Integer
        For i = 0 To columns - 1
          If Me.Rdiag(i) = 0 Then
            Return False
          End If
        Next
        Return True
			End Get
		End Property
		''' <summary>Returns the upper triangular factor <c>R</c>.</summary>
		Public ReadOnly Property UpperTriangularFactor() As Matrix
			Get
				Dim n As Integer = Me.QR.Columns
				Dim X As New Matrix(n, n)
        Dim xx As Double()() = X.Array
        Dim qr As Double()() = Me.QR.Array
				For i As Integer = 0 To n - 1
					For j As Integer = 0 To n - 1
						If i < j Then
              xx(i)(j) = qr(i)(j)
            ElseIf i = j Then
              xx(i)(j) = Rdiag(i)
            Else
              xx(i)(j) = 0
            End If
          Next
				Next
				Return X
			End Get
		End Property
		''' <summary>Returns the orthogonal factor <c>Q</c>.</summary>
		Public ReadOnly Property OrthogonalFactor() As Matrix
			Get
        Dim X As New Matrix(Me.QR.Rows, Me.QR.Columns)
        Dim xx As Double()() = X.Array
        Dim qr As Double()() = Me.QR.Array
        For k As Integer = Me.QR.Columns - 1 To 0 Step -1
          For i As Integer = Me.QR.Rows - 1 To 0 Step -1 'To????
            xx(i)(k) = 0
          Next
          xx(k)(k) = 1
          For j As Integer = k To Me.QR.Columns - 1
            If qr(k)(k) <> 0 Then
              Dim s As Double = 0
              For i As Integer = k To Me.QR.Rows - 1
                s += qr(i)(k) * xx(i)(j)
              Next
              s = -s / qr(k)(k)
              For i As Integer = k To Me.QR.Rows - 1
                xx(i)(j) += s * qr(i)(k)
              Next
            End If
          Next
        Next
        Return X
			End Get
		End Property
		Private Shared Function Hypotenuse(ByVal a As Double, ByVal b As Double) As Double
			If Math.Abs(a) > Math.Abs(b) Then
				Dim r As Double = b / a
				Return Math.Abs(a) * Math.Sqrt(1 + r * r)
			End If
			If b <> 0 Then
				Dim r As Double = a / b
				Return Math.Abs(b) * Math.Sqrt(1 + r * r)
			End If
			Return 0
		End Function
	End Class
End Namespace
