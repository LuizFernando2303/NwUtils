using Autodesk.Navisworks.Api;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NwUtils.Types
{
    public class Mesh
    {
        public readonly IEnumerable<Vector3D> Vertices;
        public readonly IEnumerable<Triangle> Triangles;
        static readonly Random RandomShared = new Random();

        public Mesh(IEnumerable<Vector3D> vertices, IEnumerable<Triangle> triangles)
        {
            this.Vertices = vertices;
            this.Triangles = triangles;
        }

        public double surfaceArea
        {
            get
            {
                double surfaceArea = 0;
                foreach (var t in Triangles)
                {
                    surfaceArea += t.Area();
                }
                return surfaceArea;
            }
        }

        public Vector3D Location
        {
            get
            {
                double avgX = Vertices.Select(v => v.X).Average();
                double avgY = Vertices.Select(v => v.Y).Average();
                double avgZ = Vertices.Select(v => v.Z).Average();

                return new Vector3D(avgX, avgY, avgZ);
            }
        }

        public string ToObj()
        {
            StringBuilder sb = new StringBuilder();
            foreach (var v in Vertices)
            {
                sb.AppendLine($"v {v.X} {v.Y} {v.Z}");
            }
            foreach (var t in Triangles)
            {
                int v1Index = Vertices.ToList().IndexOf(t.A) + 1;
                int v2Index = Vertices.ToList().IndexOf(t.B) + 1;
                int v3Index = Vertices.ToList().IndexOf(t.C) + 1;
                sb.AppendLine($"f {v1Index} {v2Index} {v3Index}");
            }
            return sb.ToString();
        }

        public static Mesh FromObj(string obj)
        {
            Mesh mesh = null;
            List<Vector3D> vertices = new List<Vector3D>();
            List<Triangle> triangles = new List<Triangle>();

            using (StringReader reader = new StringReader(obj))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    if (line.StartsWith("v "))
                    {
                        var parts = line.Split(' ');
                        double x = double.Parse(parts[1]);
                        double y = double.Parse(parts[2]);
                        double z = double.Parse(parts[3]);
                        vertices.Add(new Vector3D(x, y, z));
                    }
                    else if (line.StartsWith("f "))
                    {
                        var parts = line.Split(' ');
                        int v1Index = int.Parse(parts[1]) - 1;
                        int v2Index = int.Parse(parts[2]) - 1;
                        int v3Index = int.Parse(parts[3]) - 1;
                        triangles.Add(new Triangle(vertices[v1Index], vertices[v2Index], vertices[v3Index]));
                    }
                }
            }

            mesh = new Mesh(vertices, triangles);
            return mesh;
        }
    }
}
