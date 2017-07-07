using System;
using System.Runtime.InteropServices;

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using SolidWorks.Interop.swconst;
using SolidWorksTools;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Newtonsoft.Json;

namespace Voxel8SolidworksAddin {
    [ComVisible(true)]
    public abstract class Disposable : IDisposable {
        bool disposed = false;

        public void Dispose() {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        void Dispose(bool disposing) {
            if (disposed)
                return;

            if (disposing)
                DisposeManaged();

            DisposeUnmanaged();

            disposed = true;
        }

        protected virtual void DisposeManaged() { }
        protected virtual void DisposeUnmanaged() { }
    }

    /// <summary>
    /// Summary description for Voxel8SolidworksAddin.
    /// </summary>
    [Guid("5d9937aa-9f20-4fc5-8bad-4efaa94d5a42"), ComVisible(true)]
    [SwAddin(
        Description = "Addin for exporting to Three.JS JSON format.",
        Title = "SwToJson",
        LoadAtStartup = true
        )]
    public class SwAddin : Disposable, ISwAddin {
        ISldWorks iSwApp;
        int addinID;

        public ISldWorks SwApp => iSwApp;

        #region SolidWorks Registration
        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type t) {
            #region Get Custom Attribute: SwAddinAttribute
            SwAddinAttribute SWattr = null;
            Type type = typeof(SwAddin);

            foreach (System.Attribute attr in type.GetCustomAttributes(false)) {
                if (attr is SwAddinAttribute) {
                    SWattr = attr as SwAddinAttribute;
                    break;
                }
            }

            #endregion

            try {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                Microsoft.Win32.RegistryKey addinkey = hklm.CreateSubKey(keyname);
                addinkey.SetValue(null, 0);

                addinkey.SetValue("Description", SWattr.Description);
                addinkey.SetValue("Title", SWattr.Title);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                addinkey = hkcu.CreateSubKey(keyname);
                addinkey.SetValue(null, Convert.ToInt32(SWattr.LoadAtStartup), Microsoft.Win32.RegistryValueKind.DWord);
            }
            catch (NullReferenceException nl) {
                Console.WriteLine("There was a problem registering this dll: SWattr is null. \n\"" + nl.Message + "\"");
                System.Windows.Forms.MessageBox.Show("There was a problem registering this dll: SWattr is null.\n\"" + nl.Message + "\"");
            }
            catch (Exception e) {
                Console.WriteLine(e.Message);

                System.Windows.Forms.MessageBox.Show("There was a problem registering the function: \n\"" + e.Message + "\"");
            }
        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type t) {
            try {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                hklm.DeleteSubKey(keyname);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                hkcu.DeleteSubKey(keyname);
            }
            catch (NullReferenceException nl) {
                Console.WriteLine("There was a problem unregistering this dll: " + nl.Message);
                System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + nl.Message + "\"");
            }
            catch (Exception e) {
                Console.WriteLine("There was a problem unregistering this dll: " + e.Message);
                System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + e.Message + "\"");
            }
        }

        #endregion

        public bool ConnectToSW(object ThisSW, int cookie) {
            iSwApp = (ISldWorks)ThisSW;
            addinID = cookie;

            //Setup callbacks
            iSwApp.SetAddinCallbackInfo(0, this, addinID);

            iSwApp.AddFileSaveAsItem2(cookie, "ExportToJSON", "Three.JS JSON Format", "json", (int)swDocumentTypes_e.swDocASSEMBLY);
            iSwApp.AddFileSaveAsItem2(cookie, "ExportToJSON", "Three.JS JSON Format", "json", (int)swDocumentTypes_e.swDocPART);

            return true;
        }

        public bool DisconnectFromSW() {
            iSwApp.RemoveFileSaveAsItem2(addinID, "ExportToJSON", "Three.JS JSON Format", "json", (int)swDocumentTypes_e.swDocASSEMBLY);
            iSwApp.RemoveFileSaveAsItem2(addinID, "ExportToJSON", "Three.JS JSON Format", "json", (int)swDocumentTypes_e.swDocPART);

            Marshal.ReleaseComObject(iSwApp);
            iSwApp = null;
            //The addin _must_ call GC.Collect() here in order to retrieve all managed code pointers 
            GC.Collect();
            GC.WaitForPendingFinalizers();

            GC.Collect();
            GC.WaitForPendingFinalizers();

            return true;
        }

        static long ColorArrayToColor(double[] colorArray) {
            return ((long)(colorArray[0] * 255) << 16) | ((long)(colorArray[1] * 255) << 8) | ((long)(colorArray[2] * 255));
        }

        public void ExportToJSON(string fileName) {
            var activeDoc = ((IModelDoc2)SwApp.ActiveDoc);

            var maxId = 0;
            var partIDs = new Dictionary<Tuple<object, string, long?>, int>();

            if (activeDoc is IAssemblyDoc) {
                partIDs[new Tuple<object, string, long?>(activeDoc.ConfigurationManager.ActiveConfiguration.GetRootComponent3(true), activeDoc.ConfigurationManager.ActiveConfiguration.Name, null)] = maxId++;

                foreach (Component2 component in (object[]) ((IAssemblyDoc) activeDoc).GetComponents(true)) {
                    AddIds(component, partIDs, ref maxId);
                }
            }
            else {
                partIDs[new Tuple<object, string, long?>(activeDoc, activeDoc.ConfigurationManager.ActiveConfiguration.Name, null)] = maxId++;
            }

            var modelDocList = new Tuple<object, string, long?>[maxId];

            foreach (var pair in partIDs)
                modelDocList[pair.Value] = pair.Key;

            var partTessellations = new List<PartTessellation>();

            foreach (var modelDoc in modelDocList) {
                var totalTessellation = new PartTessellation();

                var partDoc = modelDoc.Item1 as IPartDoc;

                if (partDoc != null) {
                    string oldConfiguration = null;
                    if (modelDoc.Item2 != ((IModelDoc2) modelDoc.Item1).ConfigurationManager.ActiveConfiguration.Name) {
                        oldConfiguration = ((IModelDoc2)modelDoc.Item1).ConfigurationManager.ActiveConfiguration.Name;
                        ((IModelDoc2)modelDoc.Item1).ShowConfiguration2(modelDoc.Item2);
                    }

                    var bodies = (object[])partDoc.GetBodies2((int)swBodyType_e.swAllBodies, true);
                    if (bodies != null) {
                        foreach (IBody2 body in bodies) {
                            var tessellation = (ITessellation) body.GetTessellation(null);
                            tessellation.NeedEdgeFinMap = true;
                            tessellation.NeedFaceFacetMap = true;
                            tessellation.NeedVertexNormal = true;
                            tessellation.NeedVertexParams = false;
                            tessellation.ImprovedQuality = false;
                            tessellation.CurveChordAngleTolerance = 0.6;
                            tessellation.CurveChordTolerance = 1;
                            tessellation.SurfacePlaneAngleTolerance = 0.6;
                            tessellation.SurfacePlaneTolerance = 1;

                            tessellation.Tessellate();

                            var bodyTessellation = GetBodyMesh(body, partDoc, tessellation, modelDoc.Item3);

                            totalTessellation.Meshes.Add(bodyTessellation);

                            var edges = (object[]) body.GetEdges();
                            if (edges != null) {
                                foreach (var edge in edges) {
                                    var polyline = GetEdgePolyline(tessellation, edge);
                                    totalTessellation.Lines.Add(new BodyTessellation {
                                        VertexPositions = new List<double[]>(polyline.Select(v => (double[]) tessellation.GetVertexPoint(v)))
                                    });
                                }
                            }
                        }
                    }

                    if (oldConfiguration != null)
                        ((IModelDoc2)modelDoc.Item1).ShowConfiguration2(oldConfiguration);
                }

                var assemblyDoc = modelDoc.Item1 as IComponent2;

                if (assemblyDoc != null) {
                    foreach (Component2 component in (object[])assemblyDoc.GetChildren()) {
                        if (component.IsHidden(true))
                            continue;

                        var tuple = GetTuple(component);

                        int id;
                        if (!partIDs.TryGetValue(tuple, out id))
                            continue;

                        var transformation = component.Transform2;

                        if (assemblyDoc.Transform2 != null)
                            transformation = (MathTransform) transformation.Multiply(assemblyDoc.Transform2.Inverse());

                        var threeJsComponent = new ThreeJsComponent {
                            Id = id,
                            Transformation = transformation
                        };

                        totalTessellation.Components.Add(threeJsComponent);
                    }
                }

                partTessellations.Add(totalTessellation);
            }

            StringBuilder sb = new StringBuilder();
            StringWriter sw = new StringWriter(sb);

            using (JsonWriter writer = new JsonTextWriter(sw)) {
                writer.Formatting = Formatting.Indented;

                writer.WriteStartArray();

                foreach (var tessellation in partTessellations)
                    tessellation.ToJson(writer);

                writer.WriteEndArray();
            }

            var result = sb.ToString();

            // Strip the trailing 'w' or 'r' and any leading and trailing white space
            fileName = fileName.Substring(0, fileName.Length - 1).Trim(' ');

            // The filename contains the full path followed by just the filename. We need to strip this out.
            var searchEnd = fileName.Length - 1;
            while (true) {
                var spaceIndex = fileName.LastIndexOf(' ', searchEnd, searchEnd + 1);
                if (spaceIndex == -1)
                    break;
                if (fileName.Substring(spaceIndex + 1) ==
                    fileName.Substring(spaceIndex - (fileName.Length - spaceIndex - 1), fileName.Length - spaceIndex - 1)) {
                    fileName = fileName.Substring(0, spaceIndex);
                    break;
                }
                else
                    searchEnd = spaceIndex - 1;
            }

            File.WriteAllText(fileName, result);
        }

        private static BodyTessellation GetBodyMesh(IBody2 body, IPartDoc partDoc, ITessellation tessellation, long? overrideColor) {
            var bodyTessellation = new BodyTessellation();

            if (overrideColor.HasValue) {
                bodyTessellation.FaceColors.Add(overrideColor.Value);
            }
            else {
                double[] bodyColorArray = null;
                var features = (object[]) body.GetFeatures();
                if (features != null) {
                    foreach (IFeature feature in features.Reverse()) {
                        bodyColorArray =
                            (double[]) feature.GetMaterialPropertyValues2(
                                (int) swInConfigurationOpts_e.swThisConfiguration, null);
                        if (bodyColorArray[0] == -1
                        ) // All -1s are returned by features that don't assign color.
                            bodyColorArray = null;

                        if (bodyColorArray != null)
                            break;
                    }
                }

                if (bodyColorArray == null)
                    bodyColorArray = (double[]) body.MaterialPropertyValues2;

                if (bodyColorArray == null)
                    bodyColorArray = (double[]) partDoc.MaterialPropertyValues;

                var bodyColor = ColorArrayToColor(bodyColorArray);

                bodyTessellation.FaceColors.Add(bodyColor);
            }

            var coloredFaces = new Dictionary<IFace2, long>();
            var faceCount = 0;
            foreach (IFace2 face in (object[]) body.GetFaces()) {
                faceCount++;
                var colorArray = (double[]) face.MaterialPropertyValues;
                if (colorArray != null)
                    coloredFaces[face] = ColorArrayToColor(colorArray);
            }

            if (coloredFaces.Count < faceCount) {
                for (var i = 0; i < tessellation.GetVertexCount(); i++) {
                    bodyTessellation.VertexPositions.Add((double[]) tessellation.GetVertexPoint(i));
                    bodyTessellation.VertexNormals.Add((double[]) tessellation.GetVertexNormal(i));
                }

                foreach (IFace2 face in (object[]) body.GetFaces()) {
                    if (coloredFaces.ContainsKey(face))
                        continue;

                    foreach (var facet in (int[]) tessellation.GetFaceFacets(face)) {
                        var vertexIndices = new List<int>();
                        foreach (var fin in (int[]) tessellation.GetFacetFins(facet)) {
                            vertexIndices.Add(((int[]) tessellation.GetFinVertices(fin))[0]);
                        }

                        bodyTessellation.Faces.Add(new FaceStruct {
                            Color = 0,
                            Vertex1 = vertexIndices[0],
                            Vertex2 = vertexIndices[1],
                            Vertex3 = vertexIndices[2],
                        });
                    }
                }
            }

            foreach (var pair in coloredFaces) {
                var colorIndex = bodyTessellation.FaceColors.IndexOf(pair.Value);
                if (colorIndex == -1) {
                    bodyTessellation.FaceColors.Add(pair.Value);
                    colorIndex = bodyTessellation.FaceColors.Count - 1;
                }

                foreach (var facet in (int[]) tessellation.GetFaceFacets(pair.Key)) {
                    var vertexIndices = new List<int>();
                    foreach (var fin in (int[]) tessellation.GetFacetFins(facet)) {
                        vertexIndices.Add(((int[]) tessellation.GetFinVertices(fin))[0]);
                    }

                    bodyTessellation.Faces.Add(
                        new FaceStruct {
                            Color = colorIndex,
                            Vertex1 = bodyTessellation.VertexPositions.Count,
                            Vertex2 = bodyTessellation.VertexPositions.Count + 1,
                            Vertex3 = bodyTessellation.VertexPositions.Count + 2
                        });

                    bodyTessellation.VertexPositions.Add(
                        (double[]) tessellation.GetVertexPoint(vertexIndices[0]));
                    bodyTessellation.VertexPositions.Add(
                        (double[]) tessellation.GetVertexPoint(vertexIndices[1]));
                    bodyTessellation.VertexPositions.Add(
                        (double[]) tessellation.GetVertexPoint(vertexIndices[2]));

                    bodyTessellation.VertexNormals.Add(
                        (double[]) tessellation.GetVertexNormal(vertexIndices[0]));
                    bodyTessellation.VertexNormals.Add(
                        (double[]) tessellation.GetVertexNormal(vertexIndices[1]));
                    bodyTessellation.VertexNormals.Add(
                        (double[]) tessellation.GetVertexNormal(vertexIndices[2]));
                }
            }
            return bodyTessellation;
        }

        private static List<int> GetEdgePolyline(ITessellation tessellation, object edge) {
            var fins = (int[])tessellation.GetEdgeFins(edge);

            var finVertices = new Dictionary<int, List<int>>();
            int minOpenVertex = -1;
            int minVertex = int.MaxValue;

            foreach (var fin in fins) {
                var finEnds = (int[])tessellation.GetFinVertices(fin);
                var min = Math.Min(finEnds[0], finEnds[1]);
                var max = Math.Max(finEnds[0], finEnds[1]);

                if (min < minVertex)
                    minVertex = min;

                List<int> list;
                if (!finVertices.TryGetValue(min, out list)) {
                    list = new List<int> { max };
                    finVertices[min] = list;

                    if (minOpenVertex == -1 || minOpenVertex > min)
                        minOpenVertex = min;
                }
                else {
                    list.Add(max);
                    if (minOpenVertex == min)
                        minOpenVertex = -1;
                }

                if (!finVertices.TryGetValue(max, out list)) {
                    list = new List<int> { min };
                    finVertices[max] = list;

                    if (minOpenVertex == -1 || minOpenVertex > max)
                        minOpenVertex = max;
                }
                else {
                    list.Add(min);
                    if (minOpenVertex == max)
                        minOpenVertex = -1;
                }
            }

            var polyline = new List<int>();

            var visited = new HashSet<int>();

            var startVertex = minOpenVertex == -1 ? minVertex : minOpenVertex;

            while (visited.Count < finVertices.Count && startVertex != -1) {
                polyline.Add(startVertex);
                visited.Add(startVertex);

                var next = finVertices[startVertex];

                startVertex = -1;
                for (var i = 0; i < next.Count; i++) {
                    if (!visited.Contains(next[i]))
                        startVertex = next[i];
                }
            }

            if (minOpenVertex == -1)
                polyline.Add(polyline[0]);

            return polyline;
        }

        Tuple<object, string, long?> GetTuple(Component2 component) {
            var modelDoc = (IModelDoc2)component.GetModelDoc2();

            if (modelDoc is IAssemblyDoc)
                return new Tuple<object, string, long?>(component, component.ReferencedConfiguration, null);

            var componentColor = (double[]) component.GetMaterialPropertyValues2((int)swInConfigurationOpts_e.swThisConfiguration, null);
            if (componentColor != null && componentColor[0] != -1)
                return new Tuple<object, string, long?>(modelDoc, component.ReferencedConfiguration, ColorArrayToColor(componentColor));
            return new Tuple<object, string, long?>(modelDoc, component.ReferencedConfiguration, null);
        }

        void AddIds(Component2 component, Dictionary<Tuple<object, string, long?>, int> ids, ref int id) {
            if (component.IsHidden(true))
                return;

            var modelDoc = (IModelDoc2)component.GetModelDoc2();

            var tuple = GetTuple(component);
            
            if (ids.ContainsKey(tuple))
                return;

            ids[tuple] = id++;

            var assembly = modelDoc as IAssemblyDoc;

            if (assembly != null) {
                foreach (Component2 c in (object[]) component.GetChildren()) {
                    AddIds(c, ids, ref id);
                }
            }
        }

        class PartTessellation {
            public List<BodyTessellation> Lines = new List<BodyTessellation>();
            public List<BodyTessellation> Meshes = new List<BodyTessellation>();
            public List<ThreeJsComponent> Components = new List<ThreeJsComponent>();

            public void Add(PartTessellation other) {
                Lines.AddRange(other.Lines);
                Meshes.AddRange(other.Meshes);
                Components.AddRange(other.Components);
            }

            public void ToJson(JsonWriter writer) {
                writer.WriteStartObject();

                if (Lines.Count > 0) {
                    writer.WritePropertyName("lines");
                    writer.WriteStartArray();

                    foreach (var line in Lines) {
                        line.ToJson(writer);
                    }

                    writer.WriteEndArray();
                }

                if (Meshes.Count > 0) {
                    writer.WritePropertyName("meshes");
                    writer.WriteStartArray();

                    foreach (var mesh in Meshes) {
                        mesh.ToJson(writer);
                    }

                    writer.WriteEndArray();
                }

                if (Components.Count > 0) {
                    writer.WritePropertyName("components");
                    writer.WriteStartArray();

                    foreach (var component in Components) {
                        component.ToJson(writer);
                    }

                    writer.WriteEndArray();
                }

                writer.WriteEndObject();
            }
        }

        class ThreeJsComponent {
            public int Id { get; set; }
            public MathTransform Transformation { get; set; }

            public void ToJson(JsonWriter writer) {
                writer.WriteStartObject();

                writer.WritePropertyName("id");
                writer.WriteValue(Id);

                writer.WritePropertyName("transformation");
                writer.WriteStartArray();

                var data = ((double[])Transformation.ArrayData);

                var values = new double[16];

                // Solidworks array is transpose of rotation matrix,
                // Then translation vector, then scale value.
                values[0] = data[0];
                values[1] = data[3];
                values[2] = data[6];
                values[3] = data[9];
                values[4] = data[1];
                values[5] = data[4];
                values[6] = data[7];
                values[7] = data[10];
                values[8] = data[2];
                values[9] = data[5];
                values[10] = data[8];
                values[11] = data[11];
                values[15] = data[12];

                for (var i = 0; i < 16; i++) {
                    writer.WriteValue(values[i]);
                }

                writer.WriteEndArray();

                writer.WriteEndObject();
            }
        }

        class BodyTessellation {
            public List<double[]> VertexPositions = new List<double[]>();
            public List<double[]> VertexNormals = new List<double[]>();
            public List<long> FaceColors = new List<long>();
            public List<FaceStruct> Faces = new List<FaceStruct>();

            public void Add(BodyTessellation other) {
                int vertexOffset = VertexPositions.Count;
                int faceColorOffset = FaceColors.Count;

                VertexPositions.AddRange(other.VertexPositions);
                VertexNormals.AddRange(other.VertexNormals);
                FaceColors.AddRange(other.FaceColors);
                Faces.AddRange(other.Faces.Select(f => new FaceStruct {
                    Vertex1 = f.Vertex1 + vertexOffset,
                    Vertex2 = f.Vertex2 + vertexOffset,
                    Vertex3 = f.Vertex3 + vertexOffset,
                    Color = f.Color + faceColorOffset
                }));
            }

            public void ToJson(JsonWriter writer) {
                writer.Formatting = Formatting.Indented;

                writer.WriteStartObject();

                writer.WritePropertyName("metadata");
                writer.WriteStartObject();
                writer.WritePropertyName("formatVersion");
                writer.WriteValue(3);
                writer.WriteEndObject();

                writer.WritePropertyName("vertices");
                writer.WriteStartArray();
                foreach (var vertex in VertexPositions)
                    writer.WriteRawValue(string.Format("{0:0.0#######},{1:0.0#######},{2:0.0#######}", vertex[0], vertex[1], vertex[2]));
                writer.WriteEndArray();

                writer.WritePropertyName("normals");
                writer.WriteStartArray();
                foreach (var normal in VertexNormals)
                    writer.WriteRawValue(string.Format("{0:0.0#######},{1:0.0#######},{2:0.0#######}", normal[0], normal[1], normal[2]));
                writer.WriteEndArray();

                writer.WritePropertyName("colors");
                writer.WriteStartArray();
                foreach (var color in FaceColors)
                    writer.WriteValue(color);
                writer.WriteEndArray();

                writer.WritePropertyName("faces");
                writer.WriteStartArray();
                foreach (var face in Faces)
                    writer.WriteRawValue(face.ToString());
                writer.WriteEndArray();

                writer.WriteEndObject();
            }
        }


        class FaceStruct {
            public int Vertex1;
            public int Vertex2;
            public int Vertex3;
            public int Color;

            [Flags]
            enum FaceType {
                Triangle = 0,
                Quad = 1,
                Material = 2,
                UV = 4,
                VertexUV = 8,
                Normal = 16,
                VertexNormal = 32,
                Color = 64,
                VertexColor = 128
            };

            public override string ToString() {
                var faceType = FaceType.Triangle | FaceType.VertexNormal | FaceType.Color;

                return string.Format("{0}, {1},{2},{3}, {4},{5},{6}, {7}", (int)faceType, Vertex1, Vertex2, Vertex3, Vertex1, Vertex2, Vertex3, Color);
            }
        }
    }
}
