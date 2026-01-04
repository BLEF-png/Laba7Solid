using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;

namespace Lab5_Kaluzhny
{
    public class Point3D
    {
        public double X, Y, Z;
        public Point3D(double x, double y, double z) { X = x; Y = y; Z = z; }
    }

    public class Iteration
    {
        private const string PlaneName = "RegionTopPlane";
        private const string PointsSketchName = "RegionTopPtsSketch";

        private Feature _topPlaneFeat;

        private Point3D[] pointsPlane;
        private Point3D pointCenter;
        private Point3D topCenter;      // центр верхней грани

        private double lx, ly, lz;
        private double minX, maxX, minY, maxY, minZ, maxZ;

        public Point3D PointCenter => pointCenter;  // центр всего параллелепипеда
        public Point3D TopCenter => topCenter;      // центр верхней грани

        public double Lx => lx;
        public double Ly => ly;
        public double Lz => lz;

        public double MinX => minX; public double MaxX => maxX;
        public double MinY => minY; public double MaxY => maxY;
        public double MinZ => minZ; public double MaxZ => maxZ;

        public Point3D[] PointsPlane => pointsPlane;

        public Iteration(Point3D two, Point3D one)
        {
            pointsPlane = new Point3D[8];
            RecomputeBox(two, one);
        }

        public void RecomputeBox(Point3D two, Point3D one)
        {
            // p3 - одна вершина, p5 - противоположная
            pointsPlane[3] = one;
            pointsPlane[5] = two;

            pointsPlane[0] = new Point3D(pointsPlane[5].X, pointsPlane[3].Y, pointsPlane[3].Z);
            pointsPlane[1] = new Point3D(pointsPlane[5].X, pointsPlane[3].Y, pointsPlane[5].Z);
            pointsPlane[2] = new Point3D(pointsPlane[3].X, pointsPlane[3].Y, pointsPlane[5].Z);
            pointsPlane[4] = new Point3D(pointsPlane[5].X, pointsPlane[5].Y, pointsPlane[3].Z);
            pointsPlane[6] = new Point3D(pointsPlane[3].X, pointsPlane[5].Y, pointsPlane[5].Z);
            pointsPlane[7] = new Point3D(pointsPlane[3].X, pointsPlane[5].Y, pointsPlane[3].Z);

            minX = Math.Min(pointsPlane[3].X, pointsPlane[5].X);
            maxX = Math.Max(pointsPlane[3].X, pointsPlane[5].X);
            minY = Math.Min(pointsPlane[3].Y, pointsPlane[5].Y);
            maxY = Math.Max(pointsPlane[3].Y, pointsPlane[5].Y);
            minZ = Math.Min(pointsPlane[3].Z, pointsPlane[5].Z);
            maxZ = Math.Max(pointsPlane[3].Z, pointsPlane[5].Z);

            lx = maxX - minX;
            ly = maxY - minY;
            lz = maxZ - minZ;

            // центр всего бокса
            pointCenter = new Point3D(
                (minX + maxX) / 2.0,
                (minY + maxY) / 2.0,
                (minZ + maxZ) / 2.0);

            // центр ВЕРХНЕЙ грани: X и Z по середине, Y = maxY
            topCenter = new Point3D(
                (minX + maxX) / 2.0,
                maxY,
                (minZ + maxZ) / 2.0);
        }

        // ---------- работа с плоскостью верхней грани ----------

        public void CreateTopPlane(ModelDoc2 md)
        {
            if (md == null) return;
            if (_topPlaneFeat != null) return;

            var exist = FindFeatureByName(md, PlaneName);
            if (exist != null) { _topPlaneFeat = exist; return; }

            SafeDeleteFeature(md, ref _topPlaneFeat);
            SafeDeleteByName(md, PointsSketchName);

            // 3 точки на верхней грани (Y = maxY)
            md.SketchManager.Insert3DSketch(true);
            SketchPoint pA = md.SketchManager.CreatePoint(minX, maxY, minZ);
            SketchPoint pB = md.SketchManager.CreatePoint(maxX, maxY, minZ);
            SketchPoint pC = md.SketchManager.CreatePoint(minX, maxY, maxZ);
            md.Insert3DSketch2(true);

            Feature skFeat = FindLastFeatureByType(md, "3DProfileFeature");
            if (skFeat == null) skFeat = FindLastFeatureByType(md, "3DSketch");
            if (skFeat != null) { try { skFeat.Name = PointsSketchName; } catch { } }

            pA.Select(true);
            pB.Select(true);
            pC.Select(true);
            var rp = md.CreatePlaneThru3Points3(true) as RefPlane;
            md.ClearSelection2(true);

            if (rp == null) throw new Exception("Не удалось создать RegionTopPlane.");
            _topPlaneFeat = rp as Feature;
            try { _topPlaneFeat.Name = PlaneName; } catch { }
        }

        public bool SelectTopPlane(ModelDoc2 md)
        {
            if (md == null) return false;

            if (_topPlaneFeat != null)
            {
                md.ClearSelection2(true);
                _topPlaneFeat.Select2(false, -1);
                return true;
            }

            var f = FindFeatureByName(md, PlaneName);
            if (f != null)
            {
                _topPlaneFeat = f;
                md.ClearSelection2(true);
                _topPlaneFeat.Select2(false, -1);
                return true;
            }

            return false;
        }

        /// <summary>
        /// Удалить вспомогательную плоскость и 3D-эскиз области.
        /// Используем для кнопки «Удалить вырез», чтобы полностью очистить область.
        /// </summary>
        public void DeleteRegionGeometry(ModelDoc2 md)
        {
            if (md == null) return;

            // 3D-эскиз с точками
            SafeDeleteByName(md, PointsSketchName);

            // плоскость
            SafeDeleteFeature(md, ref _topPlaneFeat);
        }

        // ---------- helpers ----------

        public static Feature FindLastFeatureByType(ModelDoc2 md, string typeName2)
        {
            int n = md.GetFeatureCount();
            for (int i = n; i >= 1; i--)
            {
                Feature f = md.FeatureByPositionReverse(i);
                if (f != null && f.GetTypeName2() == typeName2)
                    return f;
            }
            return null;
        }

        private static Feature FindFeatureByName(ModelDoc2 md, string name)
        {
            Feature f = md.FirstFeature();
            while (f != null)
            {
                if (string.Equals(f.Name, name, StringComparison.Ordinal))
                    return f;
                f = f.GetNextFeature();
            }
            return null;
        }

        private static void SafeDeleteByName(ModelDoc2 md, string featName)
        {
            Feature f = FindFeatureByName(md, featName);
            if (f == null) return;
            try
            {
                md.ClearSelection2(true);
                f.Select2(false, -1);
                md.Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Absorbed);
                md.ClearSelection2(true);
            }
            catch { }
        }

        private static void SafeDeleteFeature(ModelDoc2 md, ref Feature feat)
        {
            if (feat == null) return;
            try
            {
                md.ClearSelection2(true);
                feat.Select2(false, -1);
                md.Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Absorbed);
                md.ClearSelection2(true);
            }
            catch { }
            finally { feat = null; }
        }

        /// <summary>
        /// Прямой вызов FeatureCut2 "как в отрисовщике": режем по активному эскизу.
        /// size – глубина в метрах, flip – поменять направление (вниз/вверх).
        /// </summary>
        public Feature featureCut(ModelDoc2 md, double size, bool flip = true,
            swEndConditions_e mode = swEndConditions_e.swEndCondBlind)
        {
            return md.FeatureManager.FeatureCut2(
                true,           // использовать активный эскиз
                flip,           // поменять направление
                false,          // не оба направления
                (int)mode,      // конец 1
                (int)mode,      // конец 2
                size,           // глубина 1
                0,              // глубина 2
                false, false,   // черновые углы
                false, false,   // черновые наружу
                0, 0,           // углы
                false, false, false, false,
                false, false, false, false,
                false, false);
        }

        // обёртки для старого кода
        public void createPlane(ModelDoc2 md) => CreateTopPlane(md);
        public bool SelectWorkPlane(ModelDoc2 md) => SelectTopPlane(md);
    }
}
