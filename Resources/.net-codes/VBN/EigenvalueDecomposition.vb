' ----------------------------------------------
' Lutz Roeder's Mapack for .NET, September 2000
' Adapted from Mapack for COM and Jama routines.
' http://www.aisto.com/roeder/dotnet
' ----------------------------------------------
Imports System
Namespace Mapack
	''' <summary>
	''' Determines the eigenvalues and eigenvectors of a real square matrix.
	''' </summary>
	''' <remarks>
	''' If <c>A</c> is symmetric, then <c>A = V * D * V'</c> and <c>A = V * V'</c>
	''' where the eigenvalue matrix <c>D</c> is diagonal and the eigenvector matrix <c>V</c> is orthogonal.
	''' If <c>A</c> is not symmetric, the eigenvalue matrix <c>D</c> is block diagonal
	''' with the real eigenvalues in 1-by-1 blocks and any complex eigenvalues,
	''' <c>lambda+i*mu</c>, in 2-by-2 blocks, <c>[lambda, mu; -mu, lambda]</c>.
	''' The columns of <c>V</c> represent the eigenvectors in the sense that <c>A * V = V * D</c>.
	''' The matrix V may be badly conditioned, or even singular, so the validity of the equation
	''' <c>A=V*D*inverse(V)</c> depends upon the condition of <c>V</c>.
	''' </remarks>
	Public Class EigenvalueDecomposition
		Private n As Integer
    ' matrix dimension
    Private d As Double(), e As Double()
		' storage of eigenvalues.
		Private V As Matrix
		' storage of eigenvectors.
		Private H As Matrix
		' storage of nonsymmetric Hessenberg form.
		Private ort As Double()
		' storage for nonsymmetric algorithm.
		Private cdivr As Double, cdivi As Double
		Private symmetric As Boolean
		''' <summary>Construct an eigenvalue decomposition.</summary>
    Public Sub New(ByVal value As Matrix)
      If value Is Nothing Then
        Throw New ArgumentNullException("value")
      End If
      If value.Rows <> value.Columns Then
        Throw New ArgumentException("Matrix is not a square matrix.", "value")
      End If
      n = value.Columns
      V = New Matrix(n, n)
      d = New Double(n) {}
      e = New Double(n) {}
      ' Check for symmetry.
      Dim i As Integer, j As Integer
      Me.symmetric = value.Symmetric
      If Me.symmetric Then
        For i = 0 To n - 1
          For j = 0 To n - 1
            V(i, j) = value(i, j)
          Next
        Next
        ' Tridiagonalize.
        Me.tred2()
        ' Diagonalize.
        Me.tql2()
      Else
        H = New Matrix(n, n)
        ort = New Double(n) {}
        For j = 0 To n - 1
          For i = 0 To n - 1
            H(i, j) = value(i, j)
          Next
        Next
        ' Reduce to Hessenberg form.
        Me.orthes()
        ' Reduce Hessenberg to real Schur form.
        Me.hqr2()
      End If
    End Sub
    Private Sub tred2()
      ' Symmetric Householder reduction to tridiagonal form.
      ' This is derived from the Algol procedures tred2 by Bowdler, Martin, Reinsch, and Wilkinson, 
      ' Handbook for Auto. Comp., Vol.ii-Linear Algebra, and the corresponding Fortran subroutine in EISPACK.
      Dim i As Integer, j As Integer, k As Integer
      Dim scale As Double = 0
      Dim f As Double, g As Double, h1 As Double, hh As Double
      For j = 0 To n - 1
        d(j) = V(n - 1, j)
      Next
      For i = n - 1 To 1 Step -1
        ' Householder reduction to tridiagonal form.
        ' Scale to avoid under/overflow.
        scale = 0
        h1 = 0
        For k = 0 To i - 1
          scale = scale + Math.Abs(d(k))
        Next
        If scale = 0 Then
          e(i) = d(i - 1)
          For j = 0 To i - 1
            d(j) = V(i - 1, j)
            V(i, j) = 0
            V(j, i) = 0
          Next
        Else
          For k = 0 To i - 1
            ' Generate Householder vector.
            d(k) = d(k) / scale
            h1 = h1 + d(k) * d(k)
          Next
          f = d(i - 1)
          g = Math.Sqrt(h1)
          If f > 0 Then g = -g
          e(i) = scale * g
          h1 = h1 - f * g
          d(i - 1) = f - g
          For j = 0 To i - 1
            e(j) = 0
          Next
          For j = 0 To i - 1
            ' Apply similarity transformation to remaining columns.
            f = d(j)
            V(j, i) = f
            g = e(j) + V(j, j) * f
            For k = j + 1 To i - 1
              g = g + V(k, j) * d(k)
              e(k) = e(k) + V(k, j) * f
            Next
            e(j) = g
          Next
          f = 0
          For j = 0 To i - 1
            e(j) = e(j) / h1
            f = f + e(j) * d(j)
          Next
          hh = f / (h1 + h1)
          For j = 0 To i - 1
            e(j) = e(j) - hh * d(j)
          Next
          For j = 0 To i - 1
            f = d(j)
            g = e(j)
            For k = j To i - 1
              V(k, j) = V(k, j) - (f * e(k) + g * d(k))
            Next
            d(j) = V(i - 1, j)
            V(i, j) = 0
          Next
        End If
        d(i) = h1
      Next
      For i = 0 To n - 2
        ' Accumulate transformations.
        V(n - 1, i) = V(i, i)
        V(i, i) = 1
        h1 = d(i + 1)
        If h1 <> 0 Then
          For k = 0 To i
            d(k) = V(k, i + 1) / h1
          Next
          For j = 0 To i
            g = 0
            For k = 0 To i
              g = g + V(k, i + 1) * V(k, j)
            Next
            For k = 0 To i
              V(k, j) = V(k, j) - g * d(k)
            Next
          Next
        End If
        For k = 0 To i
          V(k, i + 1) = 0
        Next
      Next
      For j = 0 To n - 1
        d(j) = V(n - 1, j)
        V(n - 1, j) = 0
      Next
      V(n - 1, n - 1) = 1
      e(0) = 0
    End Sub
    Private Sub tql2()
      Dim i As Integer, j As Integer, k As Integer, l As Integer, m As Integer
      Dim iter As Integer
      Dim g As Double, h As Double, p As Double, r As Double
      Dim dl1 As Double, el1 As Double
      Dim c As Double, c2 As Double, c3 As Double
      Dim s As Double, s2 As Double
      For i = 1 To n - 1
        e(i - 1) = e(i)
      Next
      ' Symmetric tridiagonal QL algorithm.
      ' This is derived from the Algol procedures tql2, by Bowdler, Martin, Reinsch, and Wilkinson, 
      ' Handbook for Auto. Comp., Vol.ii-Linear Algebra, and the corresponding Fortran subroutine in EISPACK.
      e(n - 1) = 0
      Dim f As Double = 0
      Dim tst1 As Double = 0
      Dim eps As Double = Math.Pow(2, -52)
      For l = 0 To n - 1
        ' Find small subdiagonal element.
        tst1 = Math.Max(tst1, Math.Abs(d(l)) + Math.Abs(e(l)))
        m = l
        While m < n
          If Math.Abs(e(m)) <= eps * tst1 Then
            Exit While 'Do
          End If
          m += 1
        End While
        ' If m == l, d[l] is an eigenvalue, otherwise, iterate.
        If m > l Then
          iter = 0
          Do
            iter = iter + 1
            ' (Could check iteration count here.)
            ' Compute implicit shift
            g = d(l)
            p = (d(l + 1) - g) / (2 * e(l))
            r = Hypotenuse(p, 1)
            If p < 0 Then
              r = -r
            End If
            d(l) = e(l) / (p + r)
            d(l + 1) = e(l) * (p + r)
            dl1 = d(l + 1)
            h = g - d(l)
            For i = l + 2 To n - 1
              d(i) -= h
            Next
            f = f + h
            ' Implicit QL transformation.
            p = d(m)
            c = 1
            c2 = c
            c3 = c
            el1 = e(l + 1)
            s = 0
            s2 = 0
            For i = m - 1 To l Step -1
              c3 = c2
              c2 = c
              s2 = s
              g = c * e(i)
              h = c * p
              r = Hypotenuse(p, e(i))
              e(i + 1) = s * r
              s = e(i) / r
              c = p / r
              p = c * d(i) - s * g
              d(i + 1) = h + s * (c * g + s * d(i))
              For k = 0 To n - 1
                ' Accumulate transformation.
                h = V(k, i + 1)
                V(k, i + 1) = s * V(k, i) + c * h
                V(k, i) = c * V(k, i) - s * h
              Next
            Next
            p = -s * s2 * c3 * el1 * e(l) / dl1
            e(l) = s * p
            ' Check for convergence.
            d(l) = c * p
          Loop While Math.Abs(e(l)) > eps * tst1
        End If
        d(l) = d(l) + f
        e(l) = 0
      Next
      For i = 0 To n - 2
        ' Sort eigenvalues and corresponding vectors.
        k = i
        p = d(i)
        For j = i + 1 To n - 1
          If d(j) < p Then
            k = j
            p = d(j)
          End If
        Next
        If k <> i Then
          d(k) = d(i)
          d(i) = p
          For j = 0 To n - 1
            p = V(j, i)
            V(j, i) = V(j, k)
            V(j, k) = p
          Next
        End If
      Next
    End Sub
    Private Sub orthes()
      ' Nonsymmetric reduction to Hessenberg form.
      ' This is derived from the Algol procedures orthes and ortran, by Martin and Wilkinson, 
      ' Handbook for Auto. Comp., Vol.ii-Linear Algebra, and the corresponding Fortran subroutines in EISPACK.
      Dim i As Integer, j As Integer, m As Integer
      Dim low As Integer = 0
      Dim high As Integer = n - 1
      Dim scale As Double
      Dim f As Double, g As Double, h As Double
      For m = low + 1 To high - 1
        ' Scale column.
        scale = 0
        For i = m To high
          scale = scale + Math.Abs(Me.H(i, m - 1))
        Next
        If scale <> 0 Then
          ' Compute Householder transformation.
          h = 0
          For i = high To m Step -1
            ort(i) = Me.H(i, m - 1) / scale
            h = h + ort(i) * ort(i)
          Next
          g = Math.Sqrt(h)
          If ort(m) > 0 Then
            g = -g
          End If
          h = h - ort(m) * g
          ort(m) = ort(m) - g
          For j = m To n - 1
            ' Apply Householder similarity transformation
            ' H = (I - u * u' / h) * H * (I - u * u') / h)
            f = 0
            For i = high To m Step -1
              f += ort(i) * Me.H(i, j)
            Next
            f = f / h
            For i = m To high
              Me.H(i, j) = Me.H(i, j) - f * ort(i)
            Next
          Next
          For i = 0 To high
            f = 0
            For j = high To m Step -1
              f += ort(j) * Me.H(i, j)
            Next
            f = f / h
            For j = m To high
              Me.H(i, j) -= f * ort(j)
            Next
          Next
          ort(m) = scale * ort(m)
          Me.H(m, m - 1) = scale * g
        End If
      Next
      For i = 0 To n - 1
        For j = 0 To n - 1
          V(i, j) = (IIf(i = j, 1, 0))
        Next
      Next
      For m = high - 1 To low + 1 Step -1
        ' Accumulate transformations (Algol's ortran).
        If Me.H(m, m - 1) <> 0 Then
          For i = m + 1 To high
            ort(i) = Me.H(i, m - 1)
          Next
          For j = m To high
            g = 0
            For i = m To high
              g = g + ort(i) * V(i, j)
            Next
            ' Double division avoids possible underflow.
            g = (g / ort(m)) / Me.H(m, m - 1)
            For i = m To high
              V(i, j) = V(i, j) + g * ort(i)
            Next
          Next
        End If
      Next
    End Sub
    Private Sub cdiv(ByVal xr As Double, ByVal xi As Double, ByVal yr As Double, ByVal yi As Double)
      ' Complex scalar division.
      Dim r As Double
      Dim d As Double
      If Math.Abs(yr) > Math.Abs(yi) Then
        r = yi / yr
        d = yr + r * yi
        cdivr = (xr + r * xi) / d
        cdivi = (xi - r * xr) / d
      Else
        r = yr / yi
        d = yi + r * yr
        cdivr = (r * xr + xi) / d
        cdivi = (r * xi - xr) / d
      End If
    End Sub
    Private Sub hqr2()
      ' Nonsymmetric reduction from Hessenberg to real Schur form.   
      ' This is derived from the Algol procedure hqr2, by Martin and Wilkinson, Handbook for Auto. Comp.,
      ' Vol.ii-Linear Algebra, and the corresponding  Fortran subroutine in EISPACK.
      Dim i As Integer, j As Integer, k As Integer
      Dim l As Integer, m As Integer
      Dim nn As Integer = Me.n
      Dim n As Integer = nn - 1
      Dim low As Integer = 0
      Dim high As Integer = nn - 1
      Dim eps As Double = Math.Pow(2, -52)
      Dim exshift As Double = 0
      Dim p As Double = 0
      Dim q As Double = 0
      Dim r As Double = 0
      Dim s As Double = 0
      Dim z As Double = 0
      Dim t As Double
      Dim w As Double
      Dim x As Double
      Dim y As Double
      Dim iter As Integer
      Dim notlast As Boolean
      Dim ra As Double, sa As Double, vr As Double, vi As Double
      ' Store roots isolated by balanc and compute matrix norm
      Dim norm As Double = 0
      For i = 0 To nn - 1
        If i < low Or i > high Then
          d(i) = H(i, i)
          e(i) = 0
        End If
        For j = Math.Max(i - 1, 0) To nn - 1
          norm = norm + Math.Abs(H(i, j))
        Next
      Next
      ' Outer loop over eigenvalue index
      iter = 0
      While n >= low
        ' Look for single small sub-diagonal element
        l = n
        While l > low
          s = Math.Abs(H(l - 1, l - 1)) + Math.Abs(H(l, l))
          If s = 0 Then
            s = norm
          End If
          If Math.Abs(H(l, l - 1)) < eps * s Then
            Exit While 'Do
          End If
          l -= 1
        End While
        ' Check for convergence
        If l = n Then
          ' One root found
          H(n, n) = H(n, n) + exshift
          d(n) = H(n, n)
          e(n) = 0
          n -= 1
          iter = 0
        ElseIf l = n - 1 Then
          ' Two roots found
          w = H(n, n - 1) * H(n - 1, n)
          p = (H(n - 1, n - 1) - H(n, n)) / 2
          q = p * p + w
          z = Math.Sqrt(Math.Abs(q))
          H(n, n) = H(n, n) + exshift
          H(n - 1, n - 1) = H(n - 1, n - 1) + exshift
          x = H(n, n)
          If q >= 0 Then
            ' Real pair
            z = IIf((p >= 0), (p + z), (p - z))
            d(n - 1) = x + z
            d(n) = d(n - 1)
            If z <> 0 Then
              d(n) = x - w / z
            End If
            e(n - 1) = 0
            e(n) = 0
            x = H(n, n - 1)
            s = Math.Abs(x) + Math.Abs(z)
            p = x / s
            q = z / s
            r = Math.Sqrt(p * p + q * q)
            p = p / r
            q = q / r
            For j = n - 1 To nn - 1
              ' Row modification
              z = H(n - 1, j)
              H(n - 1, j) = q * z + p * H(n, j)
              H(n, j) = q * H(n, j) - p * z
            Next
            For i = 0 To n
              ' Column modification
              z = H(i, n - 1)
              H(i, n - 1) = q * z + p * H(i, n)
              H(i, n) = q * H(i, n) - p * z
            Next
            For i = low To high
              ' Accumulate transformations
              z = V(i, n - 1)
              V(i, n - 1) = q * z + p * V(i, n)
              V(i, n) = q * V(i, n) - p * z
            Next
          Else
            ' Complex pair
            d(n - 1) = x + p
            d(n) = x + p
            e(n - 1) = z
            e(n) = -z
          End If
          n = n - 2
          iter = 0
        Else
          ' No convergence yet	 
          ' Form shift
          x = H(n, n)
          y = 0
          w = 0
          If l < n Then
            y = H(n - 1, n - 1)
            w = H(n, n - 1) * H(n - 1, n)
          End If
          ' Wilkinson's original ad hoc shift
          If iter = 10 Then
            exshift = exshift + x
            For i = low To n
              H(i, i) = H(i, i) - x
            Next
            s = Math.Abs(H(n, n - 1)) + Math.Abs(H(n - 1, n - 2))
            x = 0.75 * s
            y = x
            w = -0.4375 * s * s
          End If
          ' MATLAB's new ad hoc shift
          If iter = 30 Then
            s = (y - x) / 2
            s = s * s + w
            If s > 0 Then
              s = Math.Sqrt(s)
              If y < x Then
                s = -s
              End If
              s = x - w / ((y - x) / 2 + s)
              For i = low To n
                H(i, i) -= s
              Next
              exshift = exshift + s
              w = 0.964
              y = w
              x = y
            End If
          End If
          iter = iter + 1
          ' Look for two consecutive small sub-diagonal elements
          m = n - 2
          While m >= l
            z = H(m, m)
            r = x - z
            s = y - z
            p = (r * s - w) / H(m + 1, m) + H(m, m + 1)
            q = H(m + 1, m + 1) - z - r - s
            r = H(m + 2, m + 1)
            s = Math.Abs(p) + Math.Abs(q) + Math.Abs(r)
            p = p / s
            q = q / s
            r = r / s
            If m = l Then
              Exit While 'Do
            End If
            If Math.Abs(H(m, m - 1)) * (Math.Abs(q) + Math.Abs(r)) < eps * (Math.Abs(p) * (Math.Abs(H(m - 1, m - 1)) + Math.Abs(z) + Math.Abs(H(m + 1, m + 1)))) Then
              Exit While 'Do
            End If
            m = m - 1
          End While
          For i = m + 2 To n
            H(i, i - 2) = 0
            If i > m + 2 Then
              H(i, i - 3) = 0
            End If
          Next
          For k = m To n - 1
            ' Double QR step involving rows l:n and columns m:n
            notlast = (k <> n - 1)
            If k <> m Then
              p = H(k, k - 1)
              q = H(k + 1, k - 1)
              r = (IIf(notlast, H(k + 2, k - 1), 0))
              x = Math.Abs(p) + Math.Abs(q) + Math.Abs(r)
              If x <> 0 Then
                p = p / x
                q = q / x
                r = r / x
              End If
            End If
            If x = 0 Then
              Exit For
            End If
            s = Math.Sqrt(p * p + q * q + r * r)
            If p < 0 Then
              s = -s
            End If
            If s <> 0 Then
              If k <> m Then
                H(k, k - 1) = -s * x
              ElseIf l <> m Then
                H(k, k - 1) = -H(k, k - 1)
              End If
              p = p + s
              x = p / s
              y = q / s
              z = r / s
              q = q / p
              r = r / p
              For j = k To nn - 1
                ' Row modification
                p = H(k, j) + q * H(k + 1, j)
                If notlast Then
                  p = p + r * H(k + 2, j)
                  H(k + 2, j) = H(k + 2, j) - p * z
                End If
                H(k, j) = H(k, j) - p * x
                H(k + 1, j) = H(k + 1, j) - p * y
              Next
              For i = 0 To Math.Min(n, k + 3)
                ' Column modification
                p = x * H(i, k) + y * H(i, k + 1)
                If notlast Then
                  p = p + z * H(i, k + 2)
                  H(i, k + 2) = H(i, k + 2) - p * r
                End If
                H(i, k) = H(i, k) - p
                H(i, k + 1) = H(i, k + 1) - p * q
              Next
              For i = low To high
                ' Accumulate transformations
                p = x * V(i, k) + y * V(i, k + 1)
                If notlast Then
                  p = p + z * V(i, k + 2)
                  V(i, k + 2) = V(i, k + 2) - p * r
                End If
                V(i, k) = V(i, k) - p
                V(i, k + 1) = V(i, k + 1) - p * q
              Next
            End If
          Next
        End If
      End While
      ' Backsubstitute to find vectors of upper triangular form
      If norm = 0 Then
        Return
      End If
      For n = nn - 1 To 0 Step -1
        p = d(n)
        q = e(n)
        ' Real vector
        If q = 0 Then
          l = n
          H(n, n) = 1
          For i = n - 1 To 0 Step -1
            w = H(i, i) - p
            r = 0
            For j = l To n
              r = r + H(i, j) * H(j, n)
            Next
            If e(i) < 0 Then
              z = w
              s = r
            Else
              l = i
              If e(i) = 0 Then
                H(i, n) = IIf((w <> 0), (-r / w), (-r / (eps * norm)))
              Else
                ' Solve real equations
                x = H(i, i + 1)
                y = H(i + 1, i)
                q = (d(i) - p) * (d(i) - p) + e(i) * e(i)
                t = (x * s - z * r) / q
                H(i, n) = t
                H(i + 1, n) = IIf((Math.Abs(x) > Math.Abs(z)), ((-r - w * t) / x), ((-s - y * t) / z))
              End If
              ' Overflow control
              t = Math.Abs(H(i, n))
              If (eps * t) * t > 1 Then
                For j = i To n
                  H(j, n) = H(j, n) / t
                Next
              End If
            End If
          Next
        ElseIf q < 0 Then
          ' Complex vector
          l = n - 1
          ' Last vector component imaginary so matrix is triangular
          If Math.Abs(H(n, n - 1)) > Math.Abs(H(n - 1, n)) Then
            H(n - 1, n - 1) = q / H(n, n - 1)
            H(n - 1, n) = -(H(n, n) - p) / H(n, n - 1)
          Else
            cdiv(0, -H(n - 1, n), H(n - 1, n - 1) - p, q)
            H(n - 1, n - 1) = cdivr
            H(n - 1, n) = cdivi
          End If
          H(n, n - 1) = 0
          H(n, n) = 1
          For i = n - 2 To 0 Step -1
            ra = 0
            sa = 0
            For j = l To n
              ra = ra + H(i, j) * H(j, n - 1)
              sa = sa + H(i, j) * H(j, n)
            Next
            w = H(i, i) - p
            If e(i) < 0 Then
              z = w
              r = ra
              s = sa
            Else
              l = i
              If e(i) = 0 Then
                cdiv(-ra, -sa, w, q)
                H(i, n - 1) = cdivr
                H(i, n) = cdivi
              Else
                ' Solve complex equations
                x = H(i, i + 1)
                y = H(i + 1, i)
                vr = (d(i) - p) * (d(i) - p) + e(i) * e(i) - q * q
                vi = (d(i) - p) * 2 * q
                If vr = 0 And vi = 0 Then
                  vr = eps * norm * (Math.Abs(w) + Math.Abs(q) + Math.Abs(x) + Math.Abs(y) + Math.Abs(z))
                End If
                cdiv(x * r - z * ra + q * sa, x * s - z * sa - q * ra, vr, vi)
                H(i, n - 1) = cdivr
                H(i, n) = cdivi
                If Math.Abs(x) > (Math.Abs(z) + Math.Abs(q)) Then
                  H(i + 1, n - 1) = (-ra - w * H(i, n - 1) + q * H(i, n)) / x
                  H(i + 1, n) = (-sa - w * H(i, n) - q * H(i, n - 1)) / x
                Else
                  cdiv(-r - y * H(i, n - 1), -s - y * H(i, n), z, q)
                  H(i + 1, n - 1) = cdivr
                  H(i + 1, n) = cdivi
                End If
              End If
              ' Overflow control
              t = Math.Max(Math.Abs(H(i, n - 1)), Math.Abs(H(i, n)))
              If (eps * t) * t > 1 Then
                For j = i To n
                  H(j, n - 1) = H(j, n - 1) / t
                  H(j, n) = H(j, n) / t
                Next
              End If
            End If
          Next
        End If
      Next
      For i = 0 To nn - 1
        If i < low Or i > high Then
          For j = i To nn - 1
            V(i, j) = H(i, j)
          Next
        End If
      Next
      For j = nn - 1 To low Step -1
        For i = low To high
          ' Vectors of isolated roots
          ' Back transformation to get eigenvectors of original matrix
          z = 0
          For k = low To Math.Min(j, high)
            z = z + V(i, k) * H(k, j)
          Next
          V(i, j) = z
        Next
      Next
    End Sub
    ''' <summary>Returns the real parts of the eigenvalues.</summary>
    Public ReadOnly Property RealEigenvalues() As Double()
      Get
        Return Me.d
      End Get
    End Property
    ''' <summary>Returns the imaginary parts of the eigenvalues.</summary>	
    Public ReadOnly Property ImaginaryEigenvalues() As Double()
      Get
        Return Me.e
      End Get
    End Property
    ''' <summary>Returns the eigenvector matrix.</summary>
    Public ReadOnly Property EigenvectorMatrix() As Matrix
      Get
        Return Me.V
      End Get
    End Property
    ''' <summary>Returns the block diagonal eigenvalue matrix.</summary>
    Public ReadOnly Property DiagonalMatrix() As Matrix
      Get
        Dim X As New Matrix(n, n)
        Dim xx As Double()() = X.Array
        For i As Integer = 0 To n - 1
          For j As Integer = 0 To n - 1
            xx(i)(j) = 0
          Next
          xx(i)(i) = d(i)
          If e(i) > 0 Then
            xx(i)(i + 1) = e(i)
          ElseIf e(i) < 0 Then
            xx(i)(i - 1) = e(i)
          End If
        Next
        DiagonalMatrix = X
      End Get
    End Property
    Private Shared Function Hypotenuse(ByVal a As Double, ByVal b As Double) As Double
      Dim r As Double
      If Math.Abs(a) > Math.Abs(b) Then
        r = b / a
        Hypotenuse = Math.Abs(a) * Math.Sqrt(1 + r * r) : Exit Function
      End If
      If b <> 0 Then
        r = a / b
        Hypotenuse = Math.Abs(b) * Math.Sqrt(1 + r * r) : Exit Function
      End If
      Hypotenuse = 0
    End Function
  End Class
End Namespace
