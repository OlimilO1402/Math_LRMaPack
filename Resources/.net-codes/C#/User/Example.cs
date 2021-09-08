
	using System;
	using Mapack;
	
	class Example
	{
		public static void Main(String[] args)
		{
			Matrix A = new Matrix(3, 3);
			A[0,0] = 2.0; A[0,1] = 1.0; A[0,2] = 2.0;
			A[1,0] = 1.0; A[1,1] = 4.0; A[1,2] = 0.0;
			A[2,0] = 2.0; A[2,1] = 0.0; A[2,2] = 8.0;

			Console.WriteLine("A = ");
			Console.WriteLine(A.ToString());
	
			Console.WriteLine("A.Determinant = " + A.Determinant);
			Console.WriteLine("A.Trace = " + A.Trace);
			Console.WriteLine("A.Norm1 = " + A.Norm1);
			Console.WriteLine("A.NormInfinite = " + A.InfinityNorm);
			Console.WriteLine("A.NormFrobenius = " + A.FrobeniusNorm);
	
			SingularValueDecomposition svg = new SingularValueDecomposition(A);
			Console.WriteLine("A.Norm2 = " + svg.Norm2);
			Console.WriteLine("A.Condition = " + svg.Condition);
			Console.WriteLine("A.Rank = " + svg.Rank);
			Console.WriteLine();
	
			Console.WriteLine("A.Transpose = ");
			Console.WriteLine(A.Transpose().ToString());
	
			Console.WriteLine("A.Inverse = ");
			Console.WriteLine(A.Inverse.ToString());
	
			Matrix I = A * A.Inverse; 
			Console.WriteLine("I = A * A.Inverse = ");
			Console.WriteLine(I.ToString());
	
			Matrix B = new Matrix(3, 3);
			
			Console.WriteLine("B = ");
			B[0, 0] = 2.0; B[0, 1] = 0.0; B[0, 2] = 0.0;
			B[1, 0] = 1.0; B[1, 1] = 0.0; B[1, 2] = 0.0;
			B[2, 0] = 2.0; B[2, 1] = 0.0; B[2, 2] = 0.0;
	
			Console.WriteLine(B.ToString());
			
			Matrix X = A.Solve(B);
	
			Console.WriteLine("A.Solve(B)");
			Console.WriteLine(X.ToString());
			
			Matrix T = A * X;
			Console.WriteLine("A * A.Solve(B) = B = ");
			Console.WriteLine(T.ToString());
	
			Console.WriteLine("A = V * D * V");
	
			EigenvalueDecomposition eigen = new EigenvalueDecomposition(A);
	
			Console.WriteLine("D = ");
			Console.WriteLine(eigen.DiagonalMatrix.ToString());
	
			Console.WriteLine("lambda = ");
			foreach (double eigenvalue in eigen.RealEigenvalues)
			{
				Console.WriteLine(eigenvalue.ToString());
			}
			Console.WriteLine();
	
			Console.WriteLine("V = ");
			Console.WriteLine(eigen.EigenvectorMatrix);
		 
			Console.WriteLine("V * D * V' = ");
			Console.WriteLine(eigen.EigenvectorMatrix * (eigen.DiagonalMatrix * eigen.EigenvectorMatrix.Transpose()));
			
			Console.WriteLine("A * V = ");
			Console.WriteLine(A * eigen.EigenvectorMatrix);
			
			Console.WriteLine("V * D = ");
			Console.WriteLine(eigen.EigenvectorMatrix * eigen.DiagonalMatrix);
		}
	}
