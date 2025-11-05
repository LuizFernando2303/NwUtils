using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.ComApi;
using Autodesk.Navisworks.Api.Interop.ComApi;
using NwUtils.Types;
using System.Collections.Generic;
using System.Linq;

namespace NwUtils.Geometry
{
    public static class GeometryFunctions
    {
        /// <summary>
        /// Gets all meshes from a ModelItem
        /// </summary>
        public static IEnumerable<Mesh> GetAllMeshes(ModelItem model, bool OnlyFacingUp = false)
        {
            var collection = new ModelItemCollection();
            collection.Add(model);

            var selection = ComApiBridge.ToInwOpSelection(collection);

            var result = new List<Mesh>();

            foreach (InwOaPath path in selection.Paths())
            {
                foreach (InwOaFragment3 fragment in path.Fragments())
                {
                    var meshBuilder = new FragBuilder();
                    meshBuilder.ConfigureTransform((InwLTransform3f3)fragment.GetLocalToWorldMatrix());
                    fragment.GenerateSimplePrimitives(nwEVertexProperty.eNORMAL, meshBuilder);

                    result.Add(new Mesh(meshBuilder.Vertices.ToList(), meshBuilder.Triangles.ToList()));
                }
            }

            return result;
        }
    }
}
