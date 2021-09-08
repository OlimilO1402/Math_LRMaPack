' ----------------------------------------------
' Lutz Roeder's Mapack for .NET, September 2000
' Adapted from Mapack for COM and Jama routines.
' http://www.aisto.com/roeder/dotnet
' ----------------------------------------------
Imports System
Namespace Mapack
	''' <summary>
	''' 	Singular Value Decomposition for a rectangular matrix.
	''' </summary>
	''' <remarks>
	'''	  For an m-by-n matrix <c>A</c> with <c>m >= n</c>, the singular value decomposition is
	'''   an m-by-n orthogonal matrix <c>U</c>, an n-by-n diagonal matrix <c>S</c>, and
	'''   an n-by-n orthogonal matrix <c>V</c> so that <c>A = U * S * V'</c>.
	'''   The singular values, <c>sigma[k] = S[k,k]</c>, are ordered so that
	'''   <c>sigma[0] >= sigma[1] >= ... >= sigma[n-1]</c>.
	'''   The singular value decompostion always exists, so the constructor will
	'''   never fail. The matrix condition number and the effective numerical
	'''   rank can be computed from this decomposition.
	''' </remarks>
	Public Class SingularValueDecomposition
		Private U As Matrix
		Private V As Matrix
		Private s As Double()
		' singular values
		Private m As Integer
		Private n As Integer
		''' <summary>Construct singular value decomposition.</summary>
		Public Sub New(ByVal value As Matrix)
			If value Is Nothing Then
				Throw New ArgumentNullException("value")
			End If
			Dim copy As Matrix = DirectCast(value.Clone(), Matrix)
			Dim a As Double()() = copy.Array
			m = value.Rows
			n = value.Columns
			Dim nu As Integer = Math.Min(m, n)
			s = New Double(Math.Min(m + 1, n)) {}
      Me.U = New Matrix(m, nu)
      Me.V = New Matrix(n, n)
      Dim u As Double()() = Me.U.Array
      Dim v As Double()() = Me.V.Array
			Dim e As Double() = New Double(n) {}
			Dim work As Double() = New Double(m) {}
			Dim wantu As Boolean = True
			Dim wantv As Boolean = True
			' Reduce A to bidiagonal form, storing the diagonal elements in s and the super-diagonal elements in e.
			Dim nct As Integer = Math.Min(m - 1, n)
			Dim nrt As Integer = Math.Max(0, Math.Min(n - 2, m))
			For k As Integer = 0 To Math.Max(nct, nrt) - 1
				If k < nct Then
					' Compute the transformation for the k-th column and place the k-th diagonal in s[k].
					' Compute 2-norm of k-th column without under/overflow.
					s(k) = 0
					For i As Integer = k To m - 1
						s(k) = Hypotenuse(s(k), a(i)(k))
					Next
					If s(k) <> 0 Then
						If a(k)(k) < 0 Then
							s(k) = -s(k)
						End If
						For i As Integer = k To m - 1
							a(i)(k) /= s(k)
						Next
						a(k)(k) += 1
					End If
					s(k) = -s(k)
				End If
				For j As Integer = k + 1 To n - 1
					If (k < nct) And (s(k) <> 0) Then
						' Apply the transformation.
						Dim t As Double = 0
						For i As Integer = k To m - 1
							t += a(i)(k) * a(i)(j)
						Next
						t = -t / a(k)(k)
						For i As Integer = k To m - 1
							a(i)(j) += t * a(i)(k)
						Next
					End If
					' Place the k-th row of A into e for the subsequent calculation of the row transformation.
					e(j) = a(k)(j)
				Next
				If wantu And (k < nct) Then
					For i As Integer = k To m - 1
						u(i)(k) = a(i)(k)
						' Place the transformation in U for subsequent back
						' multiplication.
					Next
				End If
				If k < nrt Then
					' Compute the k-th row transformation and place the k-th super-diagonal in e[k].
					' Compute 2-norm without under/overflow.
					e(k) = 0
					For i As Integer = k + 1 To n - 1
						e(k) = Hypotenuse(e(k), e(i))
					Next
					If e(k) <> 0 Then
						If e(k + 1) < 0 Then
							e(k) = -e(k)
						End If
						For i As Integer = k + 1 To n - 1
							e(i) /= e(k)
						Next
						e(k + 1) += 1
					End If
					e(k) = -e(k)
					If (k + 1 < m) And (e(k) <> 0) Then
						For i As Integer = k + 1 To m - 1
							work(i) = 0
						Next
						For j As Integer = k + 1 To n - 1
							For i As Integer = k + 1 To m - 1
								work(i) += e(j) * a(i)(j)
							Next
						Next
						For j As Integer = k + 1 To n - 1
							' Apply the transformation.
							Dim t As Double = -e(j) / e(k + 1)
							For i As Integer = k + 1 To m - 1
								a(i)(j) += t * work(i)
							Next
						Next
					End If
					If wantv Then
						For i As Integer = k + 1 To n - 1
							v(i)(k) = e(i)
							' Place the transformation in V for subsequent back multiplication.
						Next
					End If
				End If
			Next
			' Set up the final bidiagonal matrix or order p.
			Dim p As Integer = Math.Min(n, m + 1)
			If nct < n Then
				s(nct) = a(nct)(nct)
			End If
			If m < p Then
				s(p - 1) = 0
			End If
			If nrt + 1 < p Then
				e(nrt) = a(nrt)(p - 1)
			End If
			e(p - 1) = 0
			' If required, generate U.
			If wantu Then
				For j As Integer = nct To nu - 1
					For i As Integer = 0 To m - 1
						u(i)(j) = 0
					Next
					u(j)(j) = 1
				Next
				For k As Integer = nct - 1 To 0 Step -1
					If s(k) <> 0 Then
						For j As Integer = k + 1 To nu - 1
							Dim t As Double = 0
							For i As Integer = k To m - 1
								t += u(i)(k) * u(i)(j)
							Next
							t = -t / u(k)(k)
							For i As Integer = k To m - 1
								u(i)(j) += t * u(i)(k)
							Next
						Next
						For i As Integer = k To m - 1
							u(i)(k) = -u(i)(k)
						Next
						u(k)(k) = 1 + u(k)(k)
						For i As Integer = 0 To k - 2
							u(i)(k) = 0
						Next
					Else
						For i As Integer = 0 To m - 1
							u(i)(k) = 0
						Next
						u(k)(k) = 1
					End If
				Next
			End If
			' If required, generate V.
			If wantv Then
				For k As Integer = n - 1 To 0 Step -1
					If (k < nrt) And (e(k) <> 0) Then
						For j As Integer = k + 1 To nu - 1
							Dim t As Double = 0
							For i As Integer = k + 1 To n - 1
								t += v(i)(k) * v(i)(j)
							Next
							t = -t / v(k + 1)(k)
							For i As Integer = k + 1 To n - 1
								v(i)(j) += t * v(i)(k)
							Next
						Next
					End If
					For i As Integer = 0 To n - 1
						v(i)(k) = 0
					Next
					v(k)(k) = 1
				Next
			End If
			' Main iteration loop for the singular values.
			Dim pp As Integer = p - 1
			Dim iter As Integer = 0
			Dim eps As Double = Math.Pow(2, -52)
			While p > 0
				Dim k As Integer, kase As Integer
				For k = p - 2 To -1 Step -1
					' Here is where a test for too many iterations would go.
					' This section of the program inspects for
					' negligible elements in the s and e arrays.  On
					' completion the variables kase and k are set as follows.
					' kase = 1     if s(p) and e[k-1] are negligible and k<p
					' kase = 2     if s(k) is negligible and k<p
					' kase = 3     if e[k-1] is negligible, k<p, and s(k), ..., s(p) are not negligible (qr step).
					' kase = 4     if e(p-1) is negligible (convergence).
					If k = -1 Then
						Exit For
					End If
					If Math.Abs(e(k)) <= eps * (Math.Abs(s(k)) + Math.Abs(s(k + 1))) Then
						e(k) = 0
						Exit For
					End If
				Next
				If k = p - 2 Then
					kase = 4
				Else
					Dim ks As Integer
					For ks = p - 1 To k Step -1
						If ks = k Then
							Exit For
						End If
            Dim t As Double ': t = (IIf(ks <> p, Math.Abs(e(ks)), 0)) + (IIf(ks <> k + 1, Math.Abs(e(ks - 1)), 0))
            If ks <> p Then
              If ks <> k + 1 Then
                t = Math.Abs(e(ks - 1))
              Else
                t = 0
              End If
              t = Math.Abs(e(ks)) + t
            End If
            If Math.Abs(s(ks)) <= eps * t Then
              s(ks) = 0
              Exit For
            End If
          Next
					If ks = k Then
						kase = 3
          ElseIf ks = p - 1 Then
            kase = 1
          Else
            kase = 2
            k = ks
          End If
				End If
				k += 1
				' Perform the task indicated by kase.
				Select Case kase
					Case 1
						' Deflate negligible s(p).
						Dim f As Double = e(p - 2)
						e(p - 2) = 0
						For j As Integer = p - 2 To k Step -1
							Dim t As Double = Hypotenuse(s(j), f)
							Dim cs As Double = s(j) / t
							Dim sn As Double = f / t
							s(j) = t
							If j <> k Then
								f = -sn * e(j - 1)
								e(j - 1) = cs * e(j - 1)
							End If
							If wantv Then
								For i As Integer = 0 To n - 1
									t = cs * v(i)(j) + sn * v(i)(p - 1)
									v(i)(p - 1) = -sn * v(i)(j) + cs * v(i)(p - 1)
									v(i)(j) = t
								Next
							End If
						Next
						Exit Select
					Case 2
						' Split at negligible s(k).
						Dim f As Double = e(k - 1)
						e(k - 1) = 0
						For j As Integer = k To p - 1
							Dim t As Double = Hypotenuse(s(j), f)
							Dim cs As Double = s(j) / t
							Dim sn As Double = f / t
							s(j) = t
							f = -sn * e(j)
							e(j) = cs * e(j)
							If wantu Then
								For i As Integer = 0 To m - 1
									t = cs * u(i)(j) + sn * u(i)(k - 1)
									u(i)(k - 1) = -sn * u(i)(j) + cs * u(i)(k - 1)
									u(i)(j) = t
								Next
							End If
						Next
						Exit Select
					Case 3
						' Perform one qr step.
						' Calculate the shift.
						Dim scale As Double = Math.Max(Math.Max(Math.Max(Math.Max(Math.Abs(s(p - 1)), Math.Abs(s(p - 2))), Math.Abs(e(p - 2))), Math.Abs(s(k))), Math.Abs(e(k)))
						Dim sp As Double = s(p - 1) / scale
						Dim spm1 As Double = s(p - 2) / scale
						Dim epm1 As Double = e(p - 2) / scale
						Dim sk As Double = s(k) / scale
						Dim ek As Double = e(k) / scale
						Dim b As Double = ((spm1 + sp) * (spm1 - sp) + epm1 * epm1) / 2
						Dim c As Double = (sp * epm1) * (sp * epm1)
						Dim shift As Double = 0
						If (b <> 0) Or (c <> 0) Then
							shift = Math.Sqrt(b * b + c)
							If b < 0 Then
								shift = -shift
							End If
							shift = c / (b + shift)
						End If
						Dim f As Double = (sk + sp) * (sk - sp) + shift
						Dim g As Double = sk * ek
						For j As Integer = k To p - 2
							' Chase zeros.
							Dim t As Double = Hypotenuse(f, g)
							Dim cs As Double = f / t
							Dim sn As Double = g / t
							If j <> k Then
								e(j - 1) = t
							End If
							f = cs * s(j) + sn * e(j)
							e(j) = cs * e(j) - sn * s(j)
							g = sn * s(j + 1)
							s(j + 1) = cs * s(j + 1)
							If wantv Then
								For i As Integer = 0 To n - 1
									t = cs * v(i)(j) + sn * v(i)(j + 1)
									v(i)(j + 1) = -sn * v(i)(j) + cs * v(i)(j + 1)
									v(i)(j) = t
								Next
							End If
							t = Hypotenuse(f, g)
							cs = f / t
							sn = g / t
							s(j) = t
							f = cs * e(j) + sn * s(j + 1)
							s(j + 1) = -sn * e(j) + cs * s(j + 1)
							g = sn * e(j + 1)
							e(j + 1) = cs * e(j + 1)
							If wantu AndAlso (j < m - 1) Then
								For i As Integer = 0 To m - 1
									t = cs * u(i)(j) + sn * u(i)(j + 1)
									u(i)(j + 1) = -sn * u(i)(j) + cs * u(i)(j + 1)
									u(i)(j) = t
								Next
							End If
						Next
						e(p - 2) = f
						iter = iter + 1
						Exit Select
					Case 4
						' Convergence.
						' Make the singular values positive.
						If s(k) <= 0 Then
							s(k) = (IIf(s(k) < 0,-s(k),0))
							If wantv Then
								For i As Integer = 0 To pp
									v(i)(k) = -v(i)(k)
								Next
							End If
						End If
						' Order the singular values.
						While k < pp
							If s(k) >= s(k + 1) Then
                Exit While 'Do
							End If
							Dim t As Double = s(k)
							s(k) = s(k + 1)
							s(k + 1) = t
							If wantv AndAlso (k < n - 1) Then
								For i As Integer = 0 To n - 1
									t = v(i)(k + 1)
									v(i)(k + 1) = v(i)(k)
									v(i)(k) = t
								Next
							End If
							If wantu AndAlso (k < m - 1) Then
								For i As Integer = 0 To m - 1
									t = u(i)(k + 1)
									u(i)(k + 1) = u(i)(k)
									u(i)(k) = t
								Next
							End If
							k += 1
						End While
						iter = 0
						p -= 1
						Exit Select
				End Select
			End While
		End Sub
		''' <summary>Returns the condition number <c>max(S) / min(S)</c>.</summary>
		Public ReadOnly Property Condition() As Double
			Get
				Return s(0) / s(Math.Min(m, n) - 1)
			End Get
		End Property
		''' <summary>Returns the Two norm.</summary>
		Public ReadOnly Property Norm2() As Double
			Get
				Return s(0)
			End Get
		End Property
		''' <summary>Returns the effective numerical matrix rank.</summary>
		''' <value>Number of non-negligible singular values.</value>
		Public ReadOnly Property Rank() As Integer
			Get
				Dim eps As Double = Math.Pow(2, -52)
				Dim tol As Double = Math.Max(m, n) * s(0) * eps
				Dim r As Integer = 0
				For i As Integer = 0 To s.Length - 1
					If s(i) > tol Then
						r += 1
					End If
				Next
				Return r
			End Get
		End Property
		''' <summary>Return the one-dimensional array of singular values.</summary>		
		Public ReadOnly Property Diagonal() As Double()
			Get
				Return Me.s
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
