using SolidWorks.Interop.dsgnchk;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Temp
{
    // ------------------------ enum'ы ------------------------

    public enum DocumentType
    {
        DRAWING,
        PART,
        ASSEMBLY
    }

    public enum DefaultPlaneName
    {
        TOP,
        FRONT,
        RIGHT,
        TEST
    }

    public enum HoleType
    {
        CUT_THROUGH = swEndConditions_e.swEndCondThroughAll,
        DISTANCE = swEndConditions_e.swEndCondBlind,
        OFFSET = swEndConditions_e.swEndCondOffsetFromSurface
    }

    public enum ObjectType
    {
        EDGE = swSelectType_e.swSelEDGES,
        FACE = swSelectType_e.swSelFACES,
        VERTEX = swSelectType_e.swSelVERTICES,
        PLANE = swSelectType_e.swSelDATUMPLANES,
        AXIS = swSelectType_e.swSelDATUMAXES,
        SKETCH_SEGMENT = swSelectType_e.swSelSKETCHSEGS
    }

    // ========================================================
    //                 КЛАСС SWDrawer
    // ========================================================
    public class SWDrawer
    {
        public SldWorks app;
        public ModelDoc2 model;
        public PartDoc part;
        public AssemblyDoc assembly;

        public SketchManager skMng;
        public FeatureManager ftMng;
        public SelectionMgr selMng;

        // ---------- init: подключиться к SW + активной детали ----------
        public bool init()
        {
            try
            {
                // 1. Получаем/запускаем SolidWorks
                try
                {
                    app = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                }
                catch
                {
                    app = new SldWorks();
                    app.FrameState = (int)swWindowState_e.swWindowMaximized;
                    app.Visible = true;
                }

                // 2. Пробуем привязаться к активному документу (если открыт)
                model = app.IActiveDoc2 as ModelDoc2;

                if (model != null)
                {
                    // если это деталь
                    part = model as PartDoc;
                    // если это сборка
                    assembly = model as AssemblyDoc;

                    skMng = model.SketchManager;
                    ftMng = model.FeatureManager;
                    selMng = model.SelectionManager;

                    model.SetUnits(
                        (short)swLengthUnit_e.swMM,
                        (short)swFractionDisplay_e.swDECIMAL,
                        0, 0, false);
                }

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Не удалось подключиться к SolidWorks: {ex.Message}",
                    "SWDrawer.init", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        // ---------- создание нового документа ----------
        public void newProject(DocumentType documentType, int drawingTemplateNum = 10)
        {
            if (app == null)
            {
                MessageBox.Show("SolidWorks не инициализирован. Сначала вызови init().");
                return;
            }

            switch (documentType)
            {
                case DocumentType.DRAWING:
                    app.NewDrawing(drawingTemplateNum);
                    break;
                case DocumentType.PART:
                    app.NewPart();
                    break;
                case DocumentType.ASSEMBLY:
                    app.NewAssembly();
                    break;
                default:
                    MessageBox.Show("Неизвестный тип документа.");
                    return;
            }

            model = (ModelDoc2)app.IActiveDoc2;
            if (model == null)
            {
                MessageBox.Show("Не удалось создать документ.");
                return;
            }

            model.SetUnits((short)swLengthUnit_e.swMM,
                           (short)swFractionDisplay_e.swDECIMAL,
                           0, 0, false);

            part = model as PartDoc;
            assembly = model as AssemblyDoc;

            skMng = model.SketchManager;
            ftMng = model.FeatureManager;
            selMng = model.SelectionManager;
        }

        // ---------- подключение к уже открытой детали ----------
        public bool connectToOpenedPart()
        {
            if (app == null && !init())
                return false;

            model = app.IActiveDoc2 as ModelDoc2;
            if (model == null)
            {
                app.SendMsgToUser("Нет активного документа в SolidWorks.");
                return false;
            }

            if (!(model is PartDoc))
            {
                app.SendMsgToUser("Активный документ не является деталью (PartDoc).");
                return false;
            }

            part = (PartDoc)model;

            skMng = model.SketchManager;
            ftMng = model.FeatureManager;
            selMng = model.SelectionManager;

            model.SetUnits(
                (short)swLengthUnit_e.swMM,
                (short)swFractionDisplay_e.swDECIMAL,
                0, 0, false);

            app.SendMsgToUser("Подключен к открытой детали.");
            return true;
        }

        // ---------- подключение к уже открытой сборке ----------
        public bool connectToOpenedAssembly()
        {
            if (app == null && !init())
                return false;

            model = app.IActiveDoc2 as ModelDoc2;
            if (model == null)
            {
                app.SendMsgToUser("Нет активного документа в SolidWorks.");
                return false;
            }

            if (!(model is AssemblyDoc))
            {
                app.SendMsgToUser("Активный документ не является сборкой.");
                return false;
            }

            assembly = (AssemblyDoc)model;
            ftMng = model.FeatureManager;
            selMng = model.SelectionManager;

            model.SetUnits(
                (short)swLengthUnit_e.swMM,
                (short)swFractionDisplay_e.swDECIMAL,
                0, 0, false);

            app.SendMsgToUser("Подключен к открытой сборке.");
            return true;
        }

        // ----------- дальше твой исходный функционал -----------
        // (оставил без изменений, кроме того, что он опирается на уже инициализированные
        //  model/skMng/ftMng/selMng)

        public SketchSegment createLineByCoords(double bX, double bY, double bZ,
                                                double eX, double eY, double eZ)
        {
            return skMng.CreateLine(bX, bY, bZ, eX, eY, eZ);
        }

        public SketchSegment createCircleByPoint(double cX, double cY, double cZ,
                                                 double rX, double rY, double rZ)
        {
            return skMng.CreateCircle(cX, cY, cZ, rX, rY, rZ);
        }

        public SketchSegment createCircleByRadius(double cX, double cY, double cZ, double radius)
        {
            return skMng.CreateCircleByRadius(cX, cY, cZ, radius);
        }

        public void createRectangle(double bX, double bY, double bZ,
                                    double eX, double eY, double eZ)
        {
            skMng.CreateCornerRectangle(bX, bY, bZ, eX, eY, eZ);
        }

        public void createCenterRectangle(double cX, double cY, double cZ,
                                          double eX, double eY, double eZ)
        {
            skMng.CreateCenterRectangle(cX, cY, cZ, eX, eY, eZ);
        }

        public void trim(SketchSegment segment, double X, double Y, double Z, int option = 0)
        {
            model.ClearSelection2(true);
            segment.Select(true);
            skMng.SketchTrim(option, X, Y, Z);
            segment.Select(false);
        }

        public void selectDefaultPlane(DefaultPlaneName planeName)
        {
            switch (planeName)
            {
                case DefaultPlaneName.TOP:
                    model.Extension.SelectByID2("СВЕРХУ", "PLANE", 0, 0, 0, false, 0, null, 0);
                    model.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
                    break;
                case DefaultPlaneName.FRONT:
                    model.Extension.SelectByID2("СПЕРЕДИ", "PLANE", 0, 0, 0, false, 0, null, 0);
                    model.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
                    break;
                case DefaultPlaneName.RIGHT:
                    model.Extension.SelectByID2("СПРАВА", "PLANE", 0, 0, 0, false, 0, null, 0);
                    model.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
                    break;
                default:
                    app.SendMsgToUser("Не удалось получить плоскость " + planeName);
                    break;
            }
        }

        public void insertSketch(bool start)
        {
            skMng.InsertSketch(start);
        }

        public void selectSketchByNumber(int number)
        {
            model.ClearSelection2(true);
            model.Extension.SelectByID2("Sketch" + number, "SKETCH", 0, 0, 0, false, 0, null, 0);
            model.Extension.SelectByID2("Эскиз" + number, "SKETCH", 0, 0, 0, false, 0, null, 0);
        }

        public void fastCube()
        {
            selectDefaultPlane(DefaultPlaneName.TOP);
            insertSketch(true);
            createCenterRectangle(0, 0, 0, 0.5, 0.5, 0);
            extrude(1);
        }

        public Body2[] getAllBodies()
        {
            if (part == null) return new Body2[0];

            object bodiesObj = part.GetBodies2((int)swBodyType_e.swSolidBody, true);
            if (bodiesObj == null) return new Body2[0];

            object[] objArray = bodiesObj as object[];
            if (objArray == null || objArray.Length == 0) return new Body2[0];

            return objArray.Cast<Body2>().ToArray();
        }

        public Face2[] getAllFaces(Body2 body)
        {
            object[] facesObj = body.GetFaces();
            object[] objArray = facesObj as object[];
            return objArray.Cast<Face2>().ToArray();
        }

        public Edge[] getAllEdges(Face2 face)
        {
            object[] edgesObj = face.GetEdges();
            object[] objArray = edgesObj as object[];
            return objArray.Cast<Edge>().ToArray();
        }

        public void SelectFaceByIndex(Body2 body, int faceIndex)
        {
            Face2[] faces = getAllFaces(body);
            Face2 face = faces[faceIndex];
            Entity faceEnt = (Entity)face;
            faceEnt.Select2(false, 0);
        }

        public void SelectEdgeByIndex(Face2 face, int edgeIndex)
        {
            Edge[] edges = getAllEdges(face);
            Entity edge = (Entity)edges[edgeIndex];
            edge.Select2(false, 0);
        }

        public void viewBodyFaces(Body2 body)
        {
            Face2[] faces = getAllFaces(body);

            for (int j = 0; j < faces.Length; j++)
            {
                SelectFaceByIndex(body, j);
                app.RunCommand(169, "");
                app.SendMsgToUser("Тело: " + body.Name + "\nГрань: " + j);
            }
        }

        public void viewFaceEdges(int faceNumber)
        {
            Body2[] bodies = getAllBodies();
            Face2 face = getAllFaces(bodies[0])[faceNumber];

            SelectFaceByIndex(bodies[0], faceNumber);
            app.RunCommand(169, "");
            model.ClearSelection2(true);

            Edge[] edges = getAllEdges(face);
            for (int i = 0; i < edges.Length; i++)
            {
                Entity edge = (Entity)edges[i];
                edge.Select2(false, 0);
                app.SendMsgToUser("Грань: " + faceNumber + "\nРебро: " + i);
            }
        }

        public void extrude(double extrusionLength, bool changeDirection = false)
        {
            ftMng.FeatureExtrusion2(
                true, false, changeDirection,
                (int)swEndConditions_e.swEndCondBlind,
                0, extrusionLength, 0, false, false,
                false, false, 0, 0, false, false, false, false,
                true, true, true,
                0, 0, false);
        }

        public void cutHole(HoleType typeOfHole, bool changeDirection = false, double depth = 0)
        {
            ftMng.FeatureCut4(
                true, false, changeDirection,
                (int)typeOfHole, 0,
                depth, 0.0,
                false, false, false, false,
                0.0, 0.0,
                false, false, false, false, false,
                true, true, true, true, false,
                0, 0, false, false);
        }

        public void createChamfers(int faceNum, double distance1, double distance2, params int[] edgeNums)
        {
            model.ClearSelection2(true);

            Body2[] bodies = getAllBodies();
            Face2 face = getAllFaces(bodies[0])[faceNum];
            Edge[] edges = getAllEdges(face);

            for (int i = 0; i < edgeNums.Length; i++)
            {
                Entity edgeEnt = (Entity)edges[edgeNums[i]];
                edgeEnt.Select2(true, 0);
            }

            ftMng.InsertFeatureChamfer(
                (int)swFeatureChamferOption_e.swFeatureChamferTangentPropagation,
                (int)swChamferType_e.swChamferDistanceDistance,
                distance1, 0, distance2, 0, 0, 0);
        }

        public void deleteLastFeatures(int numberOfElements)
        {
            if (model == null || numberOfElements <= 0)
                return;

            List<Feature> featureList = new List<Feature>();
            Feature feat = model.FirstFeature();

            while (feat != null)
            {
                featureList.Add(feat);
                feat = feat.GetNextFeature();
            }

            int count = featureList.Count;
            for (int i = count - 1; i >= count - numberOfElements && i >= 0; i--)
            {
                Feature f = featureList[i];
                if (f != null && f.Select2(false, -1))
                {
                    model.Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Absorbed);
                }
            }
        }

        public List<Body2> GetAllBodiesFromAssembly()
        {
            var result = new List<Body2>();
            if (assembly == null) return result;

            object[] comps = (object[])assembly.GetComponents(true);
            foreach (object o in comps)
            {
                Component2 comp = o as Component2;
                if (comp == null) continue;

                ModelDoc2 m = comp.GetModelDoc2();
                if (m == null) continue;

                PartDoc p = m as PartDoc;
                if (p == null) continue;

                object bodiesObj = p.GetBodies2((int)swBodyType_e.swSolidBody, true);
                if (bodiesObj == null) continue;

                foreach (Body2 b in (object[])bodiesObj)
                    result.Add(b);
            }
            return result;
        }

        public List<object> GetSelectedObjects()
        {
            var result = new List<object>();
            if (model == null) return result;

            int count = selMng.GetSelectedObjectCount();
            for (int i = 1; i <= count; i++)
            {
                object obj = selMng.GetSelectedObject6(i, -1);
                if (obj != null) result.Add(obj);
            }
            return result;
        }
    }

    // ========================================================
    //            КЛАСС HolesArrayCutter (правка ctor)
    // ========================================================

    enum Directions
    {
        OX,
        OY,
        OZ
    }

    public class HolesArrayCutter
    {
        public class Point
        {
            public double x;
            public double y;
            public double z;

            public Point() { }
            public Point(double x, double y, double z)
            {
                this.x = x; this.y = y; this.z = z;
            }

            public void set(double x, double y, double z)
            {
                this.x = x; this.y = y; this.z = z;
            }
        }

        SWDrawer drawer;

        Point bottomLeftPoint = new Point();
        Point topRightPoint = new Point();

        private int verticiesNum;
        private List<Point> circlesCenters;
        private double radius;

        private double offsetX;
        private double offsetY;

        int verticalDirection;

        // --------- ВАЖНО: конструктор с инициализацией SWDrawer ---------
        public HolesArrayCutter(SWDrawer drawer)
        {
            // если вдруг передали null – создаём свой
            this.drawer = drawer ?? new SWDrawer();

            // если ещё нет приложения – поднимаем SW и инициализируем
            if (this.drawer.app == null)
            {
                if (!this.drawer.init())
                    throw new InvalidOperationException("Не удалось инициализировать SolidWorks в HolesArrayCutter.");
            }

            // если нет модели или FeatureManager'а – пробуем подцепиться к активной детали
            if (this.drawer.model == null || this.drawer.ftMng == null)
            {
                if (!this.drawer.connectToOpenedPart())
                    throw new InvalidOperationException("Открой деталь в SolidWorks и запусти ещё раз (HolesArrayCutter).");
            }
        }

        // --------- дальше твой исходный код HolesArrayCutter без изменений ---------

        public void cutHoles(int rowsNum, int columsNum, int verticiesNum,
                             double angle = 0.0, double offset = 0.0, bool reverse = false)
        {
            List<object> temp = drawer.GetSelectedObjects();
            List<Vertex> selected = new List<Vertex>();

            foreach (object obj in temp)
            {
                if (obj is Vertex v) selected.Add(v);
            }

            if (selected == null || selected.Count < 2)
            {
                drawer.app.SendMsgToUser("Не удалось получить массив выбранных вершин!");
                return;
            }

            Vertex bottom = selected[0];
            Vertex top = selected[1];

            setPoints(bottom, top);
            setCircles(rowsNum, columsNum);

            if (offset == 0.0)
            {
                drawer.selectDefaultPlane(DefaultPlaneName.TOP);
            }
            else
            {
                drawer.model.Extension.SelectByID2("СВЕРХУ", "PLANE", 0, 0, 0, false, 0, null, 0);
                drawer.model.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, false, 0, null, 0);

                drawer.model.Extension.SelectByID2("СВЕРХУ", "PLANE", 0, 0, 0, false, 1, null, 0);
                drawer.model.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, false, 1, null, 0);

                RefPlane planeRef = drawer.ftMng.InsertRefPlane(
                    (int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_Parallel, 0.0,
                    (int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_Distance, offset,
                    0, 0.0);

                Entity plane = (Entity)planeRef;
                drawer.model.ClearSelection2(true);
                plane.Select(true);
            }

            drawer.insertSketch(true);

            foreach (Point p in circlesCenters)
            {
                drawer.skMng.CreatePolygon(
                    p.x, -p.z, 0,
                    p.x + radius * Math.Cos(angle * Math.PI / 180.0),
                    -p.z + radius * Math.Sin(angle * Math.PI / 180.0),
                    0, verticiesNum, false);
            }

            if (offset == 0.0)
            {
                drawer.cutHole(HoleType.CUT_THROUGH, true);
            }
            else
            {
                double depth = Math.Abs(top.GetPoint()[2] - bottom.GetPoint()[2]) - 2 * offset;
                drawer.app.SendMsgToUser("Глубина: " + depth);
                drawer.cutHole(HoleType.DISTANCE, true, depth);
            }
        }

        private void setPoints(Vertex bottom, Vertex top)
        {
            bottomLeftPoint.set(bottom.GetPoint()[0], bottom.GetPoint()[1], bottom.GetPoint()[2]);
            topRightPoint.set(top.GetPoint()[0], top.GetPoint()[1], top.GetPoint()[2]);
        }

        private void setCircles(int rowsNum, int columsNum, double offset = 0)
        {
            if (rowsNum <= 0) throw new ArgumentException("rowsNum must be > 0");
            if (columsNum <= 0) throw new ArgumentException("columsNum must be > 0");

            circlesCenters = new List<Point>();

            double minX = Math.Min(bottomLeftPoint.x, topRightPoint.x);
            double maxX = Math.Max(bottomLeftPoint.x, topRightPoint.x);
            double minZ = Math.Min(bottomLeftPoint.z, topRightPoint.z);
            double maxZ = Math.Max(bottomLeftPoint.z, topRightPoint.z);

            double totalSpaceX = maxX - minX;
            double totalSpaceZ = maxZ - minZ;

            if (totalSpaceX <= 0 || totalSpaceZ <= 0)
                throw new InvalidOperationException("Invalid rectangle size.");

            double minGap = offset > 0 ? offset : 0.02 * Math.Min(totalSpaceX, totalSpaceZ);
            if (minGap < 1e-9) minGap = 1e-9;

            double radiusX = (totalSpaceX - (columsNum + 1) * minGap) / (2.0 * columsNum);
            double radiusZ = (totalSpaceZ - (rowsNum + 1) * minGap) / (2.0 * rowsNum);

            if (radiusX <= 0 || radiusZ <= 0)
                throw new InvalidOperationException("Not enough space for circles");

            radius = Math.Min(radiusX, radiusZ);

            offsetX = (totalSpaceX - 2.0 * radius * columsNum) / (columsNum + 1);
            offsetY = (totalSpaceZ - 2.0 * radius * rowsNum) / (rowsNum + 1);

            if (offsetX < 0) offsetX = 0;
            if (offsetY < 0) offsetY = 0;

            double startX = minX + offsetX + radius;
            double startZ = minZ + offsetY + radius;

            for (int r = 0; r < rowsNum; r++)
            {
                for (int c = 0; c < columsNum; c++)
                {
                    double cx = startX + c * (2.0 * radius + offsetX);
                    double cz = startZ + r * (2.0 * radius + offsetY);

                    circlesCenters.Add(new Point(cx, bottomLeftPoint.y, cz));
                }
            }
        }
    }
}
