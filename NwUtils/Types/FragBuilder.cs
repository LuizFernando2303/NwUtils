using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Interop.ComApi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;

namespace NwUtils.Types
{
    public class FragBuilder : InwSimplePrimitivesCB
    {
        private double[] _matrixElements;
        private readonly List<Vector3D> _vertices = new List<Vector3D>();
        private readonly List<Triangle> _triangles = new List<Triangle>();

        public IEnumerable<Vector3D> Vertices => _vertices;
        public IEnumerable<Triangle> Triangles => _triangles;

        private const double Tolerance = 1e-6;

        public void ConfigureTransform(InwLTransform3f3 transform)
        {
            Array localToWorldMatrix = (Array)(object)transform.Matrix;
            _matrixElements = ToArray<double>(localToWorldMatrix);
        }

        public void Line(InwSimpleVertex v1, InwSimpleVertex v2)
        {
            AddOrGetVertex(TransformVertex(v1));
            AddOrGetVertex(TransformVertex(v2));
        }

        public void Point(InwSimpleVertex v1)
        {
            AddOrGetVertex(TransformVertex(v1));
        }

        public void SnapPoint(InwSimpleVertex v1)
        {
            AddOrGetVertex(TransformVertex(v1));
        }

        public void Triangle(InwSimpleVertex v1, InwSimpleVertex v2, InwSimpleVertex v3)
        {
            var vert1 = AddOrGetVertex(TransformVertex(v1));
            var vert2 = AddOrGetVertex(TransformVertex(v2));
            var vert3 = AddOrGetVertex(TransformVertex(v3));

            _triangles.Add(new Triangle(vert1, vert2, vert3));
        }

        private Vector3D TransformVertex(InwSimpleVertex vertex)
        {
            Array coords = (Array)(object)vertex.coord;
            float x = (float)coords.GetValue(1);
            float y = (float)coords.GetValue(2);
            float z = (float)coords.GetValue(3);

            float w = (float)(_matrixElements[3] * x + _matrixElements[7] * y + _matrixElements[11] * z + _matrixElements[15]);

            float tx = (float)(_matrixElements[0] * x + _matrixElements[4] * y + _matrixElements[8] * z + _matrixElements[12]) / w;
            float ty = (float)(_matrixElements[1] * x + _matrixElements[5] * y + _matrixElements[9] * z + _matrixElements[13]) / w;
            float tz = (float)(_matrixElements[2] * x + _matrixElements[6] * y + _matrixElements[10] * z + _matrixElements[14]) / w;

            return new Vector3D(tx, ty, tz);
        }

        private Vector3D AddOrGetVertex(Vector3D vertex)
        {
            var existing = _vertices.FirstOrDefault(v =>
                Math.Abs(v.X - vertex.X) < Tolerance &&
                Math.Abs(v.Y - vertex.Y) < Tolerance &&
                Math.Abs(v.Z - vertex.Z) < Tolerance);

            if (existing != default(Vector3D))
                return existing;

            _vertices.Add(vertex);
            return vertex;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private T[] ToArray<T>(Array arr)
        {
            T[] result = new T[arr.Length];
            Array.Copy(arr, result, result.Length);
            return result;
        }
    }
}
