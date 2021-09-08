Attribute VB_Name = "MMain"
Option Explicit
Public Console As New Console
Public Math As New Math
Public Matrix As New Matrix
Public Random As New Random

'Class Example
Public Sub Main() 'ByVal args As String())
  'InitFunctionsShift
  Dim A As Matrix: Set A = MNew.Matrix(3, 3)
  A(0, 0) = 2
  A(0, 1) = 1
  A(0, 2) = 2
  A(1, 0) = 1
  A(1, 1) = 4
  A(1, 2) = 0
  A(2, 0) = 2
  A(2, 1) = 0
  A(2, 2) = 8
  
'  A(0, 0) = 1
'  A(0, 1) = 3
'  A(0, 2) = -4
'  A(1, 0) = 0
'  A(1, 1) = 2
'  A(1, 2) = -2
'  A(2, 0) = -1
'  A(2, 1) = -2
'  A(2, 2) = 5

  Call Console.WriteLine("A = ")
  Call Console.WriteLine(A.ToString())
  Call Console.WriteLine("A.Determinant = " & CStr(A.Determinant))
  Call Console.WriteLine("A.Trace = " & CStr(A.Trace))
  Call Console.WriteLine("A.Norm1 = " & CStr(A.Norm1))
  Call Console.WriteLine("A.NormInfinite = " & CStr(A.InfinityNorm))
  Call Console.WriteLine("A.NormFrobenius = " & CStr(A.FrobeniusNorm))
  Dim svg As SingularValueDecomposition: Set svg = MNew.SingularValueDecomposition(A)
  Call Console.WriteLine("A.Norm2 = " & CStr(svg.Norm2))
  Call Console.WriteLine("A.Condition = " & CStr(svg.Condition))
  Call Console.WriteLine("A.Rank = " & CStr(svg.Rank))
  Call Console.WriteLine("")
  Call Console.WriteLine("A.Transpose = ")
  Call Console.WriteLine(A.Transpose().ToString())
  Call Console.WriteLine("A.Inverse = ")
  Call Console.WriteLine(A.Inverse.ToString())
  Dim i As Matrix: Set i = Matrix.op_MulM(A, A.Inverse)
  Call Console.WriteLine("I = A * A.Inverse = ")
  Call Console.WriteLine(i.ToString())
  Dim B As Matrix: Set B = MNew.Matrix(3, 3)
  Call Console.WriteLine("B = ")
  B(0, 0) = 2
  B(0, 1) = 0
  B(0, 2) = 0
  B(1, 0) = 1
  B(1, 1) = 0
  B(1, 2) = 0
  B(2, 0) = 2
  B(2, 1) = 0
  B(2, 2) = 0
  
'  B(0, 0) = 8
'  B(0, 1) = 0
'  B(0, 2) = 0
'  B(1, 0) = 6
'  B(1, 1) = 0
'  B(1, 2) = 0
'  B(2, 0) = -1
'  B(2, 1) = 0
'  B(2, 2) = 0

  Call Console.WriteLine(B.ToString())
  Dim X As Matrix: Set X = A.Solve(B)
  Call Console.WriteLine("A.Solve(B)")
  Call Console.WriteLine(X.ToString())
  Dim t As Matrix: Set t = A.op_MulM(A, X)
  Call Console.WriteLine("A * A.Solve(B) = B = ")
  Call Console.WriteLine(t.ToString())
  Call Console.WriteLine("A = V * D * V")
  Dim eigen As EigenvalueDecomposition: Set eigen = MNew.EigenvalueDecomposition(A)
  Call Console.WriteLine("D = ")
  Call Console.WriteLine(eigen.DiagonalMatrix.ToString())
  Call Console.WriteLine("lambda = ")
  Dim eigenvalue 'As Double
  For Each eigenvalue In eigen.RealEigenvalues
    Call Console.WriteLine(CStr(eigenvalue)) '.ToString())
  Next
  Call Console.WriteLine("")
  Call Console.WriteLine("V = ")
  Call Console.WriteLine(eigen.EigenvectorMatrix)
  Call Console.WriteLine("V * D * V' = ")
  Call Console.WriteLine(Matrix.op_MulM(eigen.EigenvectorMatrix, Matrix.op_MulM(eigen.DiagonalMatrix, eigen.EigenvectorMatrix.Transpose)))
  Call Console.WriteLine("A * V = ")
  Call Console.WriteLine(Matrix.op_MulM(A, eigen.EigenvectorMatrix))
  Call Console.WriteLine("V * D = ")
  Call Console.WriteLine(Matrix.op_MulM(eigen.EigenvectorMatrix, eigen.DiagonalMatrix))
  Call Console.WriteLine("Exit? j/n")
  'einen Userbreak einbauen:
  Dim s As String
  Do While LCase(s) <> "j"
    s = Console.ReadLine
  Loop
End Sub

'erzeugt folgende Ausgabe:
'A =
'2 1 2
'1 4 0
'2 0 8
'
'A.Determinant = 40
'A.Trace = 14
'A.Norm1 = 10
'A.NormInfinite = 10
'A.NormFrobenius = 9.69535971483266
'A.Norm2 = 8.62422264025397
'A.Condition = 7.98847395632229
'A.Rank = 3
'
'A.Transpose =
'2 1 2
'1 4 0
'2 0 8
'
'A.Inverse =
'0.8 -0.2 -0.2
'-0.2 0.3 0.05
'-0.2 0.05 0.175
'
'I = A * A.Inverse =
'1 1.38777878078145E-17 0
'0 1 0
'0 0 1
'
'B =
'2 0 0
'1 0 0
'2 0 0
'
'A.Solve(B)
'1 0 0
'0 0 0
'0 0 0
'
'A * A.Solve(B) = B =
'2 0 0
'1 0 0
'2 0 0
'
'A = V * D * V
'D =
'1.0795832454869 0 0
'0 4.29619411425913 0
'0 0 8.62422264025397
'
'lambda =
'1.0795832454869
'4.29619411425913
'8.62422264025397
'0
'
'V =
'-0.912578082576527 0.280716539220287 0.297320479969233
'0.312482141860836 0.947745163412308 0.0642963159647745
'0.263735007572021 -0.151582749139746 0.952610369429297
'
'V * D * V' =
'2 1 2
'1 4 -1.11022302462516E-16
'2 -1.11022302462516E-16 8
'
'A * V =
'-0.985204008148176 1.20601274357339 2.56415801476184
'0.337350484866818 4.07169719286952 0.554505743828331
'0.284723895423112 -0.651228914677396 8.21552391537285
'
'V * D =
'-0.985204008148176 1.20601274357339 2.56415801476183
'0.337350484866819 4.07169719286952 0.554505743828331
'0.284723895423113 -0.651228914677396 8.21552391537284

'Beispiel für ein nicht-symmtrisches GLS
'A:             X:     B:
' 1   3  -4      1      8
' 0   2  -2  *   5  =   6
'-1  -2   5      2     -1
'erzeugt folgende Ausgabe:
'A =
'1 3 -4
'0 2 -2
'-1 -2 5
'
'A.Determinant = 4
'A.Trace = 8
'A.Norm1 = 11
'A.NormInfinite = 8
'A.NormFrobenius = 8
'A.Norm2 = 7.89216377178265
'A.Condition = 19.372587018313
'A.Rank = 3
'
'
'A.Transpose =
'1 0 -1
'3 2 -2
'-4 -2 5
'
'A.Inverse =
'1.5 -1.75 0.5
'0.5 0.25 0.5
'0.5 -0.25 0.5
'
'I = A * A.Inverse =
'1 0 0
'0 1 0
'0 0 1
'
'B =
'8 0 0
'6 0 0
'-1 0 0
'
'A.Solve (B)
'1 0 0
'5 0 0
'2 0 0
'
'A * A.Solve(B) = B =
'8 0 0
'6 0 0
'-1 0 0
'
'A = v * d * v
'A =
'1 3 -4
'0 2 -2
'-1 -2 5
'
'A.Determinant = 4
'A.Trace = 8
'A.Norm1 = 11
'A.NormInfinite = 8
'A.NormFrobenius = 8
'A.Norm2 = 7.89216377178265
'A.Condition = 19.372587018313
'A.Rank = 3
'
'
'A.Transpose =
'1 0 -1
'3 2 -2
'-4 -2 5
'
'A.Inverse =
'1.5 -1.75 0.5
'0.5 0.25 0.5
'0.5 -0.25 0.5
'
'I = A * A.Inverse =
'1 0 0
'0 1 0
'0 0 1
'
'B =
'8 0 0
'6 0 0
'-1 0 0
'
'A.Solve (B)
'1 0 0
'5 0 0
'2 0 0
'
'A * A.Solve(B) = B =
'8 0 0
'6 0 0
'-1 0 0
'
'A = v * d * v
'D =
'0.622309847626099 0.452604769461901 0
'-0.452604769461901 0.622309847626099 0
'0 0 6.7553803047478
'
'lambda =
'0.622309847626099
'0.622309847626099
'6.7553803047478
'
'
'V =
'1.65951554168962 -0.9623262868625 -0.757614666183905
'1.09427889468601 0.506967017691488 -0.348528811672204
'0.868516473663277 0.101583810497771 0.828693523331577
'
'V * D * V' =
'6.16759295858358 3.46765697375799 -2.9505369586871
'1.75285247247937 1.72571773920889 -1.47659554369603
'-3.8597078262471 -1.17864746546669 5.11498536922197
'
'A * V =
'1.46828633109454 0.152239524220881 -5.11797519452683
'0.451524842045466 0.810766414387435 -2.35444467000756
'0.494509037254744 0.456311303968377 5.5981399061862
'
'V * D =
'1.46828633109454 0.152239524220881 -5.11797519452683
'0.451524842045464 0.810766414387433 -2.35444467000756
'0.494509037254746 0.456311303968377 5.5981399061862
'
'
'
'2. Beispiel für ein nicht-symmtrisches GLS
'A:             X:     B:
' 1   1   1             6
' 2   1   3  *      =  12
' 3   1   3            14
'erzeugt folgende Ausgabe:

