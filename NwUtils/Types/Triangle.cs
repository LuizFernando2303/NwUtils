using Autodesk.Navisworks.Api;
using System;

namespace NwUtils.Types
{
    public class Triangle
    {
        public Vector3D A { get; }
        public Vector3D B { get; }
        public Vector3D C { get; }
        public Vector3D Normal { get; }
        public Triangle(Vector3D a, Vector3D b, Vector3D c)
        {
            A = a;
            B = b;
            C = c;
            Normal = (B.Subtract(A)).Cross(C.Subtract(A)).Normalize();
        }

        public double Area()
        {
            double abLength = B.Subtract(A).Length;
            double acLength = C.Subtract(A).Length;
            double bcLength = C.Subtract(B).Length;

            double s = (abLength + acLength + bcLength) / 2.0;

            double area = Math.Sqrt(s * (s - abLength) * (s - acLength) * (s - bcLength));

            return area;
        }
    }
}
