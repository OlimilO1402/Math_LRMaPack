VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EigenvalueDecomposition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'  Public Class EigenvalueDecomposition
Private mN As Long 'Value.Columns
' matrix dimension
Private d() As Double
Private e() As Double
' storage of eigenvalues.
Private mV As Matrix
' storage of eigenvectors.
Private mH As Matrix
' storage of nonsymmetric Hessenberg form.
Private ort() As Double
' storage for nonsymmetric algorithm.
Private cdivr As Double, cdivi As Double
Private mSymmetric As Boolean

''' <summary>Construct an eigenvalue decomposition.</summary>
Friend Sub New_(Value As Matrix)
    If Value Is Nothing Then
        'Throw New ArgumentNullException("value")
        MsgBox "EigenvalueDecomposition::NewC ArgumentNull: " & "value"
    End If
    If Value.Rows <> Value.Columns Then
        'Throw New ArgumentException("Matrix is not a square matrix.", "value")
        MsgBox "EigenvalueDecomposition::NewC Matrix value is not a square matrix."
    End If
    mN = Value.Columns
    Set mV = MNew.Matrix(mN, mN)
    ReDim d(0 To mN - 1) 'd = New_Double(n - 1) '{}
    ReDim e(0 To mN - 1) 'e = New_Double(n - 1) '{}
    ' Check for symmetry.
    Dim i As Long, j As Long
    mSymmetric = Value.Symmetric
    If mSymmetric Then
        'versteh ich nicht, also doch einfach clonen
        For i = 0 To mN - 1
            For j = 0 To mN - 1
                mV(i, j) = Value(i, j)
            Next
        Next
        ' Tridiagonalize.
        Call tred2
        ' Diagonalize.
        Call tql2
    Else
        Set mH = MNew.Matrix(mN, mN)
        ReDim ort(0 To mN - 1) 'ort() = New_Double(n - 1) ' {}
        For j = 0 To mN - 1
            For i = 0 To mN - 1
                mH(i, j) = Value(i, j)
            Next
        Next
        ' Reduce to Hessenberg form.
        Call orthes
        ' Reduce Hessenberg to real Schur form.
        Call hqr2
    End If
End Sub
Private Sub tred2()
'Siehe auch Numerical Recipes Kap 11.2 "Reduction of a Symmetric Matrix to Tridiagonal Form"
  ' Symmetric Householder reduction to tridiagonal form.
  ' This is derived from the Algol procedures tred2 by Bowdler, Martin, Reinsch, and Wilkinson,
  ' Handbook for Auto. Comp., Vol.ii-Linear Algebra, and the corresponding Fortran subroutine in EISPACK.
  Dim i As Long, j As Long, k As Long
  Dim sscale As Double ' = 0
  Dim f As Double, g As Double, h1 As Double, hh As Double
  For j = 0 To mN - 1
    d(j) = mV(mN - 1, j)
  Next
  For i = mN - 1 To 1 Step -1
    ' Householder reduction to tridiagonal form.
    ' Scale to avoid under/overflow.
    sscale = 0
    h1 = 0
    For k = 0 To i - 1
      sscale = sscale + VBA.Math.Abs(d(k))
    Next
    If sscale = 0 Then
      e(i) = d(i - 1)
      For j = 0 To i - 1
        d(j) = mV(i - 1, j)
        mV(i, j) = 0
        mV(j, i) = 0
      Next
    Else
      For k = 0 To i - 1
        ' Generate Householder vector.
        d(k) = d(k) / sscale
        h1 = h1 + d(k) * d(k)
      Next
      f = d(i - 1)
      g = VBA.Math.Sqr(h1)
      If f > 0 Then g = -g
      e(i) = sscale * g
      h1 = h1 - f * g
      d(i - 1) = f - g
      For j = 0 To i - 1
        e(j) = 0
      Next
      For j = 0 To i - 1
        ' Apply similarity transformation to remaining columns.
        f = d(j)
        mV(j, i) = f
        g = e(j) + mV(j, j) * f
        For k = j + 1 To i - 1
          g = g + mV(k, j) * d(k)
          e(k) = e(k) + mV(k, j) * f
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
          mV(k, j) = mV(k, j) - (f * e(k) + g * d(k))
        Next
        d(j) = mV(i - 1, j)
        mV(i, j) = 0
      Next
    End If
    d(i) = h1
  Next
  For i = 0 To mN - 2
    ' Accumulate transformations.
    mV(mN - 1, i) = mV(i, i)
    mV(i, i) = 1
    h1 = d(i + 1)
    If h1 <> 0 Then
      For k = 0 To i
        d(k) = mV(k, i + 1) / h1
      Next
      For j = 0 To i
        g = 0
        For k = 0 To i
          g = g + mV(k, i + 1) * mV(k, j)
        Next
        For k = 0 To i
          mV(k, j) = mV(k, j) - g * d(k)
        Next
      Next
    End If
    For k = 0 To i
      mV(k, i + 1) = 0
    Next
  Next
  For j = 0 To mN - 1
    d(j) = mV(mN - 1, j)
    mV(mN - 1, j) = 0
  Next
  mV(mN - 1, mN - 1) = 1
  e(0) = 0
End Sub
Private Sub tql2()
  Dim i As Long, j As Long, k As Long, L As Long, m As Long
  Dim iter As Long
  Dim g As Double, h As Double, p As Double, r As Double
  Dim dl1 As Double, el1 As Double
  Dim c As Double, c2 As Double, c3 As Double
  Dim s As Double, s2 As Double
  For i = 1 To mN - 1
    e(i - 1) = e(i)
  Next
  ' Symmetric tridiagonal QL algorithm.
  ' This is derived from the Algol procedures tql2, by Bowdler, Martin, Reinsch, and Wilkinson,
  ' Handbook for Auto. Comp., Vol.ii-Linear Algebra, and the corresponding Fortran subroutine in EISPACK.
  e(mN - 1) = 0
  Dim f As Double ' = 0
  Dim tst1 As Double ' = 0
  Dim eps As Double: eps = Math.Pow(2, -52)
  For L = 0 To mN - 1
    ' Find small subdiagonal element.
    tst1 = Math.MaxD(tst1, VBA.Math.Abs(d(L)) + VBA.Math.Abs(e(L)))
    m = L
    Do While m < mN
      If VBA.Math.Abs(e(m)) <= eps * tst1 Then
        'Exit While 'Do
        Exit Do
      End If
      m = m + 1
    Loop 'End While
    ' If m == l, d[l] is an eigenvalue, otherwise, iterate.
    If m > L Then
      iter = 0
      Do
        iter = iter + 1
        ' (Could check iteration count here.)
        ' Compute implicit shift
        g = d(L)
        p = (d(L + 1) - g) / (2 * e(L))
        r = Hypotenuse(p, 1)
        If p < 0 Then
          r = -r
        End If
        d(L) = e(L) / (p + r)
        d(L + 1) = e(L) * (p + r)
        dl1 = d(L + 1)
        h = g - d(L)
        For i = L + 2 To mN - 1
          d(i) = d(i) - h
        Next
        f = f + h
        ' Implicit QL transformation.
        p = d(m)
        c = 1
        c2 = c
        c3 = c
        el1 = e(L + 1)
        s = 0
        s2 = 0
        For i = m - 1 To L Step -1
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
          For k = 0 To mN - 1
            ' Accumulate transformation.
            h = mV(k, i + 1)
            mV(k, i + 1) = s * mV(k, i) + c * h
            mV(k, i) = c * mV(k, i) - s * h
          Next
        Next
        p = -s * s2 * c3 * el1 * e(L) / dl1
        e(L) = s * p
        ' Check for convergence.
        d(L) = c * p
      Loop While VBA.Math.Abs(e(L)) > eps * tst1
    End If
    d(L) = d(L) + f
    e(L) = 0
  Next
  For i = 0 To mN - 2
    ' Sort eigenvalues and corresponding vectors.
    k = i
    p = d(i)
    For j = i + 1 To mN - 1
      If d(j) < p Then
        k = j
        p = d(j)
      End If
    Next
    If k <> i Then
      d(k) = d(i)
      d(i) = p
      For j = 0 To mN - 1
        p = mV(j, i)
        mV(j, i) = mV(j, k)
        mV(j, k) = p
      Next
    End If
  Next
End Sub
Private Sub orthes()
  ' Nonsymmetric reduction to Hessenberg form.
  ' This is derived from the Algol procedures orthes and ortran, by Martin and Wilkinson,
  ' Handbook for Auto. Comp., Vol.ii-Linear Algebra, and the corresponding Fortran subroutines in EISPACK.
  Dim i As Long, j As Long, m As Long
  Dim low As Long ' = 0
  Dim high As Long: high = mN - 1
  Dim sscale As Double
  Dim f As Double, g As Double, hh As Double
  For m = low + 1 To high - 1
    ' Scale column.
    sscale = 0
    For i = m To high
      sscale = sscale + VBA.Math.Abs(mH(i, m - 1))
    Next
    If sscale <> 0 Then
      ' Compute Householder transformation.
      hh = 0
      For i = high To m Step -1
        ort(i) = mH(i, m - 1) / sscale
        hh = hh + ort(i) * ort(i)
      Next
      g = VBA.Math.Sqr(hh)
      If ort(m) > 0 Then
        g = -g
      End If
      hh = hh - ort(m) * g
      ort(m) = ort(m) - g
      For j = m To mN - 1
        ' Apply Householder similarity transformation
        ' H = (I - u * u' / h) * H * (I - u * u') / h)
        f = 0
        For i = high To m Step -1
          f = f + ort(i) * mH(i, j)
        Next
        f = f / hh
        For i = m To high
          mH(i, j) = mH(i, j) - f * ort(i)
        Next
      Next
      For i = 0 To high
        f = 0
        For j = high To m Step -1
          f = f + ort(j) * mH(i, j)
        Next
        f = f / hh
        For j = m To high
          mH(i, j) = mH(i, j) - f * ort(j)
        Next
      Next
      ort(m) = sscale * ort(m)
      mH(m, m - 1) = sscale * g
    End If
  Next
  For i = 0 To mN - 1
    For j = 0 To mN - 1
      If i = j Then '(IIf(i = j, 1, 0))
        mV(i, j) = 1
      Else
        mV(i, j) = 0
      End If
    Next
  Next
  For m = high - 1 To low + 1 Step -1
    ' Accumulate transformations (Algol's ortran).
    If mH(m, m - 1) <> 0 Then
      For i = m + 1 To high
        ort(i) = mH(i, m - 1)
      Next
      For j = m To high
        g = 0
        For i = m To high
          g = g + ort(i) * mV(i, j)
        Next
        ' Double division avoids possible underflow.
        g = (g / ort(m)) / mH(m, m - 1)
        For i = m To high
          mV(i, j) = mV(i, j) + g * ort(i)
        Next
      Next
    End If
  Next
End Sub
Private Sub cdiv(ByVal xr As Double, ByVal xi As Double, ByVal yr As Double, ByVal yi As Double)
  ' Complex scalar division.
  Dim r As Double
  Dim d As Double
  If VBA.Math.Abs(yr) > VBA.Math.Abs(yi) Then
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
  Dim i As Long, j As Long, k As Long
  Dim L As Long, m As Long
  Dim nn As Long: nn = mN
  Dim n As Long: n = nn - 1
  Dim low As Long ': low = 0
  Dim high As Long: high = nn - 1
  Dim eps As Double: eps = Math.Pow(2, -52)
  Dim exshift As Double ' = 0
  Dim p As Double, q As Double, r As Double, s As Double, t As Double
  Dim w As Double, X As Double, y As Double, z As Double ' = 0
  Dim iter As Long
  Dim notlast As Boolean
  Dim ra As Double, SA As Double, vr As Double, vi As Double
  ' Store roots isolated by balanc and compute matrix norm
  Dim norm As Double ' = 0
  For i = 0 To nn - 1
    If i < low Or i > high Then
      d(i) = mH(i, i)
      e(i) = 0
    End If
    For j = Math.MaxL(i - 1, 0) To nn - 1
      norm = norm + VBA.Math.Abs(mH(i, j))
    Next
  Next
  ' Outer loop over eigenvalue index
  iter = 0
  While n >= low
    ' Look for single small sub-diagonal element
    L = n
    Do While L > low
      s = VBA.Math.Abs(mH(L - 1, L - 1)) + VBA.Math.Abs(mH(L, L))
      If s = 0 Then
        s = norm
      End If
      If VBA.Math.Abs(mH(L, L - 1)) < eps * s Then
        Exit Do 'Exit While 'Do
      End If
      L = L - 1
    Loop 'End While
    ' Check for convergence
    If L = n Then
      ' One root found
      mH(n, n) = mH(n, n) + exshift
      d(n) = mH(n, n)
      e(n) = 0
      n = n - 1
      iter = 0
    ElseIf L = n - 1 Then
      ' Two roots found
      w = mH(n, n - 1) * mH(n - 1, n)
      p = (mH(n - 1, n - 1) - mH(n, n)) / 2
      q = p * p + w
      z = VBA.Math.Sqr(VBA.Math.Abs(q))
      mH(n, n) = mH(n, n) + exshift
      mH(n - 1, n - 1) = mH(n - 1, n - 1) + exshift
      X = mH(n, n)
      If q >= 0 Then
        ' Real pair
        z = IIf((p >= 0), (p + z), (p - z))
        d(n - 1) = X + z
        d(n) = d(n - 1)
        If z <> 0 Then
          d(n) = X - w / z
        End If
        e(n - 1) = 0
        e(n) = 0
        X = mH(n, n - 1)
        s = VBA.Math.Abs(X) + VBA.Math.Abs(z)
        p = X / s
        q = z / s
        r = VBA.Math.Sqr(p * p + q * q)
        p = p / r
        q = q / r
        For j = n - 1 To nn - 1
          ' Row modification
          z = mH(n - 1, j)
          mH(n - 1, j) = q * z + p * mH(n, j)
          mH(n, j) = q * mH(n, j) - p * z
        Next
        For i = 0 To n
          ' Column modification
          z = mH(i, n - 1)
          mH(i, n - 1) = q * z + p * mH(i, n)
          mH(i, n) = q * mH(i, n) - p * z
        Next
        For i = low To high
          ' Accumulate transformations
          z = mV(i, n - 1)
          mV(i, n - 1) = q * z + p * mV(i, n)
          mV(i, n) = q * mV(i, n) - p * z
        Next
      Else
        ' Complex pair
        d(n - 1) = X + p
        d(n) = X + p
        e(n - 1) = z
        e(n) = -z
      End If
      n = n - 2
      iter = 0
    Else
      ' No convergence yet
      ' Form shift
      X = mH(n, n)
      y = 0
      w = 0
      If L < n Then
        y = mH(n - 1, n - 1)
        w = mH(n, n - 1) * mH(n - 1, n)
      End If
      ' Wilkinson's original ad hoc shift
      If iter = 10 Then
        exshift = exshift + X
        For i = low To n
          mH(i, i) = mH(i, i) - X
        Next
        s = VBA.Math.Abs(mH(n, n - 1)) + VBA.Math.Abs(mH(n - 1, n - 2))
        X = 0.75 * s
        y = X
        w = -0.4375 * s * s
      End If
      ' MATLAB's new ad hoc shift
      If iter = 30 Then
        s = (y - X) / 2
        s = s * s + w
        If s > 0 Then
          s = VBA.Math.Sqr(s)
          If y < X Then
            s = -s
          End If
          s = X - w / ((y - X) / 2 + s)
          For i = low To n
            mH(i, i) = mH(i, i) - s
          Next
          exshift = exshift + s
          w = 0.964
          y = w
          X = y
        End If
      End If
      iter = iter + 1
      ' Look for two consecutive small sub-diagonal elements
      m = n - 2
      'While m >= L
      Do While m >= L
        z = mH(m, m)
        r = X - z
        s = y - z
        p = (r * s - w) / mH(m + 1, m) + mH(m, m + 1)
        q = mH(m + 1, m + 1) - z - r - s
        r = mH(m + 2, m + 1)
        s = VBA.Math.Abs(p) + VBA.Math.Abs(q) + VBA.Math.Abs(r)
        p = p / s
        q = q / s
        r = r / s
        If m = L Then
          Exit Do 'Exit While 'Do
        End If
        If VBA.Math.Abs(mH(m, m - 1)) * (VBA.Math.Abs(q) + VBA.Math.Abs(r)) < eps * (VBA.Math.Abs(p) * (VBA.Math.Abs(mH(m - 1, m - 1)) + VBA.Math.Abs(z) + VBA.Math.Abs(mH(m + 1, m + 1)))) Then
          Exit Do 'Exit While 'Do
        End If
        m = m - 1
      Loop 'Wend 'End While
      For i = m + 2 To n
        mH(i, i - 2) = 0
        If i > m + 2 Then
          mH(i, i - 3) = 0
        End If
      Next
      For k = m To n - 1
        ' Double QR step involving rows l:n and columns m:n
        notlast = (k <> n - 1)
        If k <> m Then
          p = mH(k, k - 1)
          q = mH(k + 1, k - 1)
          r = (IIf(notlast, mH(k + 2, k - 1), 0))
          X = VBA.Math.Abs(p) + VBA.Math.Abs(q) + VBA.Math.Abs(r)
          If X <> 0 Then
            p = p / X
            q = q / X
            r = r / X
          End If
        End If
        If X = 0 Then
          Exit For
        End If
        s = VBA.Math.Sqr(p * p + q * q + r * r)
        If p < 0 Then
          s = -s
        End If
        If s <> 0 Then
          If k <> m Then
            mH(k, k - 1) = -s * X
          ElseIf L <> m Then
            mH(k, k - 1) = -mH(k, k - 1)
          End If
          p = p + s
          X = p / s
          y = q / s
          z = r / s
          q = q / p
          r = r / p
          For j = k To nn - 1
            ' Row modification
            p = mH(k, j) + q * mH(k + 1, j)
            If notlast Then
              p = p + r * mH(k + 2, j)
              mH(k + 2, j) = mH(k + 2, j) - p * z
            End If
            mH(k, j) = mH(k, j) - p * X
            mH(k + 1, j) = mH(k + 1, j) - p * y
          Next
          For i = 0 To Math.MinL(n, k + 3)
            ' Column modification
            p = X * mH(i, k) + y * mH(i, k + 1)
            If notlast Then
              p = p + z * mH(i, k + 2)
              mH(i, k + 2) = mH(i, k + 2) - p * r
            End If
            mH(i, k) = mH(i, k) - p
            mH(i, k + 1) = mH(i, k + 1) - p * q
          Next
          For i = low To high
            ' Accumulate transformations
            p = X * mV(i, k) + y * mV(i, k + 1)
            If notlast Then
              p = p + z * mV(i, k + 2)
              mV(i, k + 2) = mV(i, k + 2) - p * r
            End If
            mV(i, k) = mV(i, k) - p
            mV(i, k + 1) = mV(i, k + 1) - p * q
          Next
        End If
      Next
    End If
  Wend 'End While
  ' Backsubstitute to find vectors of upper triangular form
  If norm = 0 Then
    Return
  End If
  For n = nn - 1 To 0 Step -1
    p = d(n)
    q = e(n)
    ' Real vector
    If q = 0 Then
      L = n
      mH(n, n) = 1
      For i = n - 1 To 0 Step -1
        w = mH(i, i) - p
        r = 0
        For j = L To n
          r = r + mH(i, j) * mH(j, n)
        Next
        If e(i) < 0 Then
          z = w
          s = r
        Else
          L = i
          If e(i) = 0 Then
            mH(i, n) = IIf((w <> 0), (-r / w), (-r / (eps * norm)))
          Else
            ' Solve real equations
            X = mH(i, i + 1)
            y = mH(i + 1, i)
            q = (d(i) - p) * (d(i) - p) + e(i) * e(i)
            t = (X * s - z * r) / q
            mH(i, n) = t
            'IIf((VBA.Math.Abs(X) > VBA.Math.Abs(z)), ((-r - w * T) / X), ((-s - y * T) / z))
            If VBA.Math.Abs(X) > VBA.Math.Abs(z) Then
              mH(i + 1, n) = ((-r - w * t) / X)
            Else
              mH(i + 1, n) = ((-s - y * t) / z)
            End If
          End If
          ' Overflow control
          t = VBA.Math.Abs(mH(i, n))
          If (eps * t) * t > 1 Then
            For j = i To n
              mH(j, n) = mH(j, n) / t
            Next
          End If
        End If
      Next
    ElseIf q < 0 Then
      ' Complex vector
      L = n - 1
      ' Last vector component imaginary so matrix is triangular
      If VBA.Math.Abs(mH(n, n - 1)) > VBA.Math.Abs(mH(n - 1, n)) Then
        mH(n - 1, n - 1) = q / mH(n, n - 1)
        mH(n - 1, n) = -(mH(n, n) - p) / mH(n, n - 1)
      Else
        Call cdiv(0, -mH(n - 1, n), mH(n - 1, n - 1) - p, q)
        mH(n - 1, n - 1) = cdivr
        mH(n - 1, n) = cdivi
      End If
      mH(n, n - 1) = 0
      mH(n, n) = 1
      For i = n - 2 To 0 Step -1
        ra = 0
        SA = 0
        For j = L To n
          ra = ra + mH(i, j) * mH(j, n - 1)
          SA = SA + mH(i, j) * mH(j, n)
        Next
        w = mH(i, i) - p
        If e(i) < 0 Then
          z = w
          r = ra
          s = SA
        Else
          L = i
          If e(i) = 0 Then
            Call cdiv(-ra, -SA, w, q)
            mH(i, n - 1) = cdivr
            mH(i, n) = cdivi
          Else
            ' Solve complex equations
            X = mH(i, i + 1)
            y = mH(i + 1, i)
            vr = (d(i) - p) * (d(i) - p) + e(i) * e(i) - q * q
            vi = (d(i) - p) * 2 * q
            If vr = 0 And vi = 0 Then
              vr = eps * norm * (VBA.Math.Abs(w) + VBA.Math.Abs(q) + VBA.Math.Abs(X) + VBA.Math.Abs(y) + VBA.Math.Abs(z))
            End If
            Call cdiv(X * r - z * ra + q * SA, X * s - z * SA - q * ra, vr, vi)
            mH(i, n - 1) = cdivr
            mH(i, n) = cdivi
            If VBA.Math.Abs(X) > (VBA.Math.Abs(z) + VBA.Math.Abs(q)) Then
              mH(i + 1, n - 1) = (-ra - w * mH(i, n - 1) + q * mH(i, n)) / X
              mH(i + 1, n) = (-SA - w * mH(i, n) - q * mH(i, n - 1)) / X
            Else
              Call cdiv(-r - y * mH(i, n - 1), -s - y * mH(i, n), z, q)
              mH(i + 1, n - 1) = cdivr
              mH(i + 1, n) = cdivi
            End If
          End If
          ' Overflow control
          t = Math.MaxL(VBA.Math.Abs(mH(i, n - 1)), VBA.Math.Abs(mH(i, n)))
          If (eps * t) * t > 1 Then
            For j = i To n
              mH(j, n - 1) = mH(j, n - 1) / t
              mH(j, n) = mH(j, n) / t
            Next
          End If
        End If
      Next
    End If
  Next
  For i = 0 To nn - 1
    If i < low Or i > high Then
      For j = i To nn - 1
        mV(i, j) = mH(i, j)
      Next
    End If
  Next
  For j = nn - 1 To low Step -1
    For i = low To high
      ' Vectors of isolated roots
      ' Back transformation to get eigenvectors of original matrix
      z = 0
      For k = low To Math.MinL(j, high)
        z = z + mV(i, k) * mH(k, j)
      Next
      mV(i, j) = z
    Next
  Next
End Sub
''' <summary>Returns the real parts of the eigenvalues.</summary>
'Public ReadOnly Property RealEigenvalues() As Double()
Public Property Get RealEigenvalues() As Double()
  'Get
    'Return Me.d
    RealEigenvalues = d()
  'End Get
End Property
''' <summary>Returns the imaginary parts of the eigenvalues.</summary>
'Public ReadOnly Property ImaginaryEigenvalues() As Double()
Public Property Get ImaginaryEigenvalues() As Double()
  'Get
    'Return Me.e
    ImaginaryEigenvalues = e()
  'End Get
End Property
''' <summary>Returns the eigenvector matrix.</summary>
'Public ReadOnly Property EigenvectorMatrix() As Matrix
Public Property Get EigenvectorMatrix() As Matrix
  'Get
    'Return Me.V
    Set EigenvectorMatrix = mV
  'End Get
End Property
''' <summary>Returns the block diagonal eigenvalue matrix.</summary>
'Public ReadOnly Property DiagonalMatrix() As Matrix
Public Property Get DiagonalMatrix() As Matrix
  'Get
    Dim X As Matrix: Set X = MNew.Matrix(mN, mN)
    'Dim xx() As Double: xx() = X.pArray
    Dim xx() As Double: Call AssignArray(ArrPtr(xx), X.pArray)
    Dim i As Long, j As Long
    For i = 0 To mN - 1
      For j = 0 To mN - 1
        xx(i, j) = 0
      Next
      xx(i, i) = d(i)
      If e(i) > 0 Then
        xx(i, i + 1) = e(i)
      ElseIf e(i) < 0 Then
        xx(i, i - 1) = e(i)
      End If
    Next
    Call MArray.ZeroSAPtr(ArrPtr(xx()))
    Set DiagonalMatrix = X
  'End Get
End Property
'Private Shared Function Hypotenuse(ByVal a As Double, ByVal b As Double) As Double
Private Function Hypotenuse(ByVal A As Double, ByVal B As Double) As Double
  Dim r As Double
  If VBA.Math.Abs(A) > VBA.Math.Abs(B) Then
    r = B / A
    Hypotenuse = VBA.Math.Abs(A) * VBA.Math.Sqr(1 + r * r): Exit Function
  End If
  If B <> 0 Then
    r = A / B
    Hypotenuse = VBA.Math.Abs(B) * VBA.Math.Sqr(1 + r * r): Exit Function
  End If
  Hypotenuse = 0
End Function
'  End Class

