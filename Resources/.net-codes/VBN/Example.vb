Imports System
'Imports Mapack
Namespace Mapack
  Class Example
    Public Shared Sub Main(ByVal args As String())
      Dim A As New Matrix(3, 3)
      A(0, 0) = 2
      A(0, 1) = 1
      A(0, 2) = 2
      A(1, 0) = 1
      A(1, 1) = 4
      A(1, 2) = 0
      A(2, 0) = 2
      A(2, 1) = 0
      A(2, 2) = 8
      Console.WriteLine("A = ")
      Console.WriteLine(A.ToString())
      Console.WriteLine("A.Determinant = " & CStr(A.Determinant))
      Console.WriteLine("A.Trace = " & CStr(A.Trace))
      Console.WriteLine("A.Norm1 = " & CStr(A.Norm1))
      Console.WriteLine("A.NormInfinite = " & CStr(A.InfinityNorm))
      Console.WriteLine("A.NormFrobenius = " & CStr(A.FrobeniusNorm))
      Dim svg As New SingularValueDecomposition(A)
      Console.WriteLine("A.Norm2 = " & CStr(svg.Norm2))
      Console.WriteLine("A.Condition = " & CStr(svg.Condition))
      Console.WriteLine("A.Rank = " & CStr(svg.Rank))
      Console.WriteLine()
      Console.WriteLine("A.Transpose = ")
      Console.WriteLine(A.Transpose().ToString())
      Console.WriteLine("A.Inverse = ")
      Console.WriteLine(A.Inverse.ToString())
      Dim I As Matrix : I = Matrix.op_MulM(A, A.Inverse)
      Console.WriteLine("I = A * A.Inverse = ")
      Console.WriteLine(I.ToString())
      Dim B As New Matrix(3, 3)
      Console.WriteLine("B = ")
      B(0, 0) = 2
      B(0, 1) = 0
      B(0, 2) = 0
      B(1, 0) = 1
      B(1, 1) = 0
      B(1, 2) = 0
      B(2, 0) = 2
      B(2, 1) = 0
      B(2, 2) = 0
      Console.WriteLine(B.ToString())
      Dim X As Matrix = A.Solve(B)
      Console.WriteLine("A.Solve(B)")
      Console.WriteLine(X.ToString())
      Dim T As Matrix : T = A.op_MulM(A, X)
      Console.WriteLine("A * A.Solve(B) = B = ")
      Console.WriteLine(T.ToString())
      Console.WriteLine("A = V * D * V")
      Dim eigen As New EigenvalueDecomposition(A)
      Console.WriteLine("D = ")
      Console.WriteLine(eigen.DiagonalMatrix.ToString())
      Console.WriteLine("lambda = ")
      For Each eigenvalue As Double In eigen.RealEigenvalues
        Console.WriteLine(eigenvalue.ToString())
      Next
      Console.WriteLine()
      Console.WriteLine("V = ")
      Console.WriteLine(eigen.EigenvectorMatrix)
      Console.WriteLine("V * D * V' = ")
      Console.WriteLine(Matrix.op_MulM(eigen.EigenvectorMatrix, (Matrix.op_MulM(eigen.DiagonalMatrix, eigen.EigenvectorMatrix.Transpose()))))
      Console.WriteLine("A * V = ")
      Console.WriteLine(Matrix.op_MulM(A, eigen.EigenvectorMatrix))
      Console.WriteLine("V * D = ")
      Console.WriteLine(Matrix.op_MulM(eigen.EigenvectorMatrix, eigen.DiagonalMatrix))
    End Sub
  End Class
End Namespace
