VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SingularValueDecomposition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Siehe auch Numerical Recipes
'Kap. 2.9 Singular Value Decomposition
'Also for More or for Fewer Equations than unknowns
'  Public Class SingularValueDecomposition
Private mU As Matrix
Private mV As Matrix
Private mS() As Double '()
' singular values
Private m As Long 'rows
Private n As Long 'columns
''' <summary>Construct singular value decomposition.</summary>
Friend Sub New_(ByVal Value As Matrix)
  If Value Is Nothing Then
    'Throw New ArgumentNullException("value")
    MsgBox "SingularValueDecomposition::New " & "ArgumentNull: " & "value"
  End If
  Dim Copy As Matrix: Set Copy = Value.Clone 'DirectCast(value.Clone(), Matrix)
  Dim A() As Double: Call AssignArray(ArrPtr(A()), Copy.pArray)
  m = Value.Rows
  n = Value.Columns
  Dim nu As Long: nu = Math.MinL(m, n)
  'mS = New_Double(Math.MinL(m + 1, n)) '{}
  ReDim mS(0 To Math.MinL(m + 1, n))
  Set mU = MNew.Matrix(m, nu)
  Set mV = MNew.Matrix(n, n)
  Dim u() As Double: Call AssignArray(ArrPtr(u()), mU.pArray)
  Dim v() As Double: Call AssignArray(ArrPtr(v()), mV.pArray)
  ReDim e(0 To n - 1) As Double ': e = New_Double(n) ' {}
  ReDim work(0 To m - 1) As Double ': work = New_Double(m) '{}
  Dim wantu As Boolean: wantu = True
  Dim wantv As Boolean: wantv = True
  ' Reduce A to bidiagonal form, storing the diagonal elements in s and the super-diagonal elements in e.
  Dim nct As Long: nct = Math.MinL(m - 1, n)
  Dim nrt As Long: nrt = Math.MaxL(0, Math.MinL(n - 2, m))
  Dim i As Long, j As Long, k As Long
  Dim t As Double, f As Double, g As Double
  Dim cs As Double '= mS(j) / t
  Dim sn As Double '= f / t
  Dim sp As Double '= mS(p - 1) / scale
  Dim spm1 As Double '= mS(p - 2) / scale
  Dim epm1 As Double '= e(p - 2) / scale
  Dim sk As Double '= mS(k) / scale
  Dim ek As Double '= e(k) / scale
  Dim B As Double '= ((spm1 + sp) * (spm1 - sp) + epm1 * epm1) / 2
  Dim c As Double '= (sp * epm1) * (sp * epm1)
  Dim shift As Double '= 0
  For k = 0 To Math.MaxL(nct, nrt) - 1
    If k < nct Then
      ' Compute the transformation for the k-th column and place the k-th diagonal in s[k].
      ' Compute 2-norm of k-th column without under/overflow.
      mS(k) = 0
      For i = k To m - 1
        mS(k) = Hypotenuse(mS(k), A(i, k))
      Next
      If mS(k) <> 0 Then
        If A(k, k) < 0 Then
          mS(k) = -mS(k)
        End If
        For i = k To m - 1
          A(i, k) = A(i, k) / mS(k)
        Next
        A(k, k) = A(k, k) + 1
      End If
      mS(k) = -mS(k)
    End If
    For j = k + 1 To n - 1
      If (k < nct) And (mS(k) <> 0) Then
        ' Apply the transformation.
        t = 0
        For i = k To m - 1
          t = t + A(i, k) * A(i, j)
        Next
        t = -t / A(k, k)
        For i = k To m - 1
          A(i, j) = A(i, j) + t * A(i, k)
        Next
      End If
      ' Place the k-th row of A into e for the subsequent calculation of the row transformation.
      e(j) = A(k, j)
    Next
    If wantu And (k < nct) Then
      For i = k To m - 1
        u(i, k) = A(i, k)
        ' Place the transformation in U for subsequent back
        ' multiplication.
      Next
    End If
    If k < nrt Then
      ' Compute the k-th row transformation and place the k-th super-diagonal in e[k].
      ' Compute 2-norm without under/overflow.
      e(k) = 0
      For i = k + 1 To n - 1
        e(k) = Hypotenuse(e(k), e(i))
      Next
      If e(k) <> 0 Then
        If e(k + 1) < 0 Then
          e(k) = -e(k)
        End If
        For i = k + 1 To n - 1
          e(i) = e(i) / e(k)
        Next
        e(k + 1) = e(k + 1) + 1
      End If
      e(k) = -e(k)
      If (k + 1 < m) And (e(k) <> 0) Then
        For i = k + 1 To m - 1
          work(i) = 0
        Next
        For j = k + 1 To n - 1
          For i = k + 1 To m - 1
            work(i) = work(i) + e(j) * A(i, j)
          Next
        Next
        For j = k + 1 To n - 1
          ' Apply the transformation.
          t = -e(j) / e(k + 1)
          For i = k + 1 To m - 1
            A(i, j) = A(i, j) + t * work(i)
          Next
        Next
      End If
      If wantv Then
        For i = k + 1 To n - 1
          v(i, k) = e(i)
          ' Place the transformation in V for subsequent back multiplication.
        Next
      End If
    End If
  Next
  ' Set up the final bidiagonal matrix or order p.
  Dim p As Long: p = Math.MinL(n, m + 1)
  If nct < n Then
    mS(nct) = A(nct, nct)
  End If
  If m < p Then
    mS(p - 1) = 0
  End If
  If nrt + 1 < p Then
    e(nrt) = A(nrt, p - 1)
  End If
  e(p - 1) = 0
  ' If required, generate U.
  If wantu Then
    For j = nct To nu - 1
      For i = 0 To m - 1
        u(i, j) = 0
      Next
      u(j, j) = 1
    Next
    For k = nct - 1 To 0 Step -1
      If mS(k) <> 0 Then
        For j = k + 1 To nu - 1
          t = 0
          For i = k To m - 1
            t = t + u(i, k) * u(i, j)
          Next
          t = -t / u(k, k)
          For i = k To m - 1
            u(i, j) = u(i, j) + t * u(i, k)
          Next
        Next
        For i = k To m - 1
          u(i, k) = -u(i, k)
        Next
        u(k, k) = 1 + u(k, k)
        For i = 0 To k - 2
          u(i, k) = 0
        Next
      Else
        For i = 0 To m - 1
          u(i, k) = 0
        Next
        u(k, k) = 1
      End If
    Next
  End If
  ' If required, generate V.
  If wantv Then
    For k = n - 1 To 0 Step -1
      If (k < nrt) And (e(k) <> 0) Then
        For j = k + 1 To nu - 1
          t = 0
          For i = k + 1 To n - 1
            t = t + v(i, k) * v(i, j)
          Next
          t = -t / v(k + 1, k)
          For i = k + 1 To n - 1
            v(i, j) = v(i, j) + t * v(i, k)
          Next
        Next
      End If
      For i = 0 To n - 1
        v(i, k) = 0
      Next
      v(k, k) = 1
    Next
  End If
  ' Main iteration loop for the singular values.
  Dim pp As Long: pp = p - 1
  Dim iter As Long: iter = 0
  Dim eps As Double: eps = Math.Pow(2, -52)
  Dim kase As Long
  While p > 0
    For k = p - 2 To -1 Step -1
      ' Here is where a test for too many iterations would go.
      ' This section of the program inspects for
      ' negligible elements in the s and e arrays.  On
      ' completion the variables kase and k are set as follows.
      ' kase = 1     if mS(p) and e[k-1] are negligible and k<p
      ' kase = 2     if mS(k) is negligible and k<p
      ' kase = 3     if e[k-1] is negligible, k<p, and mS(k), ..., mS(p) are not negligible (qr step).
      ' kase = 4     if e(p-1) is negligible (convergence).
      If k = -1 Then
        Exit For
      End If
      If VBA.Math.Abs(e(k)) <= eps * (VBA.Math.Abs(mS(k)) + VBA.Math.Abs(mS(k + 1))) Then
        e(k) = 0
        Exit For
      End If
    Next
    If k = p - 2 Then
      kase = 4
    Else
      Dim ks As Long
      For ks = p - 1 To k Step -1
        If ks = k Then
          Exit For
        End If
        't = (IIf(ks <> p, VBA.Math.Abs(e(ks)), 0)) + (IIf(ks <> k + 1, VBA.Math.Abs(e(ks - 1)), 0))
        If ks <> p Then
          If ks <> k + 1 Then
            t = VBA.Math.Abs(e(ks - 1))
          Else
            t = 0
          End If
          t = VBA.Math.Abs(e(ks)) + t
        End If
        If VBA.Math.Abs(mS(ks)) <= eps * t Then
          mS(ks) = 0
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
    k = k + 1
    ' Perform the task indicated by kase.
    Select Case kase
      Case 1
        ' Deflate negligible mS(p).
        f = e(p - 2)
        e(p - 2) = 0
        For j = p - 2 To k Step -1
          t = Hypotenuse(mS(j), f)
          cs = mS(j) / t
          sn = f / t
          mS(j) = t
          If j <> k Then
            f = -sn * e(j - 1)
            e(j - 1) = cs * e(j - 1)
          End If
          If wantv Then
            For i = 0 To n - 1
              t = cs * v(i, j) + sn * v(i, p - 1)
              v(i, p - 1) = -sn * v(i, j) + cs * v(i, p - 1)
              v(i, j) = t
            Next
          End If
        Next
        'Exit Select
      Case 2
        ' Split at negligible mS(k).
        f = e(k - 1)
        e(k - 1) = 0
        For j = k To p - 1
          t = Hypotenuse(mS(j), f)
          cs = mS(j) / t
          sn = f / t
          mS(j) = t
          f = -sn * e(j)
          e(j) = cs * e(j)
          If wantu Then
            For i = 0 To m - 1
              t = cs * u(i, j) + sn * u(i, k - 1)
              u(i, k - 1) = -sn * u(i, j) + cs * u(i, k - 1)
              u(i, j) = t
            Next
          End If
        Next
        'Exit Select
      Case 3
        ' Perform one qr step.
        ' Calculate the shift.
        Dim sscale As Double: sscale = Math.MaxD(Math.MaxD(Math.MaxD(Math.MaxD(VBA.Math.Abs(mS(p - 1)), VBA.Math.Abs(mS(p - 2))), VBA.Math.Abs(e(p - 2))), VBA.Math.Abs(mS(k))), VBA.Math.Abs(e(k)))
        sp = mS(p - 1) / sscale
        spm1 = mS(p - 2) / sscale
        epm1 = e(p - 2) / sscale
        sk = mS(k) / sscale
        ek = e(k) / sscale
        B = ((spm1 + sp) * (spm1 - sp) + epm1 * epm1) / 2
        c = (sp * epm1) * (sp * epm1)
        shift = 0
        If (B <> 0) Or (c <> 0) Then
          shift = VBA.Math.Sqr(B * B + c)
          If B < 0 Then
            shift = -shift
          End If
          shift = c / (B + shift)
        End If
        f = (sk + sp) * (sk - sp) + shift
        g = sk * ek
        For j = k To p - 2
          ' Chase zeros.
          t = Hypotenuse(f, g)
          cs = f / t
          sn = g / t
          If j <> k Then
            e(j - 1) = t
          End If
          f = cs * mS(j) + sn * e(j)
          e(j) = cs * e(j) - sn * mS(j)
          g = sn * mS(j + 1)
          mS(j + 1) = cs * mS(j + 1)
          If wantv Then
            For i = 0 To n - 1
              t = cs * v(i, j) + sn * v(i, j + 1)
              v(i, j + 1) = -sn * v(i, j) + cs * v(i, j + 1)
              v(i, j) = t
            Next
          End If
          t = Hypotenuse(f, g)
          cs = f / t
          sn = g / t
          mS(j) = t
          f = cs * e(j) + sn * mS(j + 1)
          mS(j + 1) = -sn * e(j) + cs * mS(j + 1)
          g = sn * e(j + 1)
          e(j + 1) = cs * e(j + 1)
          If wantu And (j < m - 1) Then
            For i = 0 To m - 1
              t = cs * u(i, j) + sn * u(i, j + 1)
              u(i, j + 1) = -sn * u(i, j) + cs * u(i, j + 1)
              u(i, j) = t
            Next
          End If
        Next
        e(p - 2) = f
        iter = iter + 1
        'Exit Select
      Case 4
        ' Convergence.
        ' Make the singular values positive.
        If mS(k) <= 0 Then
          mS(k) = (IIf(mS(k) < 0, -mS(k), 0))
          If wantv Then
            For i = 0 To pp
              v(i, k) = -v(i, k)
            Next
          End If
        End If
        ' Order the singular values.
        Do While k < pp
          If mS(k) >= mS(k + 1) Then
            Exit Do 'While 'Do
          End If
          t = mS(k)
          mS(k) = mS(k + 1)
          mS(k + 1) = t
          If wantv And (k < n - 1) Then
            For i = 0 To n - 1
              t = v(i, k + 1)
              v(i, k + 1) = v(i, k)
              v(i, k) = t
            Next
          End If
          If wantu And (k < m - 1) Then
            For i = 0 To m - 1
              t = u(i, k + 1)
              u(i, k + 1) = u(i, k)
              u(i, k) = t
            Next
          End If
          k = k + 1
        Loop 'End While
        iter = 0
        p = p - 1
        'Exit Select
    End Select
  Wend 'End While
  Call ZeroSAPtr(ArrPtr(A()))
  Call ZeroSAPtr(ArrPtr(u()))
  Call ZeroSAPtr(ArrPtr(v()))
End Sub
''' <summary>Returns the condition number <c>max(S) / min(S)</c>.</summary>
'Public ReadOnly Property Condition() As Double
Public Property Get Condition() As Double
Dim ss As Double
  'Get
    ss = mS(Math.MinL(m, n) - 1)
    If ss <> 0 Then
      Condition = mS(0) / ss
    End If
  'End Get
End Property
''' <summary>Returns the Two norm.</summary>
'Public ReadOnly Property Norm2() As Double
Public Property Get Norm2() As Double
  'Get
    Norm2 = mS(0)
  'End Get
End Property
''' <summary>Returns the effective numerical matrix rank.</summary>
''' <value>Number of non-negligible singular values.</value>
'Public ReadOnly Property Rank() As Integer
Public Property Get Rank() As Long
  'Get
    Dim eps As Double: eps = Math.Pow(2, -52)
    Dim tol As Double: tol = Math.MaxL(m, n) * mS(0) * eps
    Dim r As Long ' = 0
    Dim i As Long
    For i = 0 To UBound(mS) '.Length - 1
      If mS(i) > tol Then
        r = r + 1
      End If
    Next
    Rank = r
  'End Get
End Property
''' <summary>Return the one-dimensional array of singular values.</summary>
'Public ReadOnly Property Diagonal() As Double()
Public Property Get Diagonal() As Double()
  'Get
  Diagonal = mS()
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

