using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using Temp;  // для SWDrawer и HoleType

namespace Lab5_Kaluzhny
{
    public class Honeycomb
    {
        public Point3D centerSketch;          // центр первой соты на верхней грани
        public double sideLength;            // расстояние от центра до вершины (радиус)

        private readonly Iteration iteration;
        private readonly double sqrt3 = Math.Sqrt(3.0);

        private double distanceBetweenCombs;

        private readonly List<Point3D> allCenters = new List<Point3D>();

        // сколько СОТ реально было построено в предыдущей итерации
        private int _lastCellsCount = 0;

        public Honeycomb(Iteration it, Point3D _)
        {
            iteration = it ?? throw new ArgumentNullException(nameof(it));
            // центр верхней грани области
            centerSketch = iteration.TopCenter;
        }

        private Point3D RoundPoint(Point3D p)
        {
            p.X = Math.Round(p.X, 12);
            p.Y = Math.Round(p.Y, 12);
            p.Z = Math.Round(p.Z, 12);
            return p;
        }

        private bool Contains(List<Point3D> list, Point3D p)
        {
            const double eps = 1e-9;
            foreach (var q in list)
            {
                if (Math.Abs(q.X - p.X) < eps &&
                    Math.Abs(q.Y - p.Y) < eps &&
                    Math.Abs(q.Z - p.Z) < eps)
                    return true;
            }
            return false;
        }

        /// <summary>
        /// Построение сот для итерации n.
        /// Теперь: каждая сота = отдельный эскиз + отдельный вырез.
        /// </summary>
        public void Cal(ModelDoc2 md, int n, bool useFixedGap, double dist, bool isFirst, double size)
        {
            if (md == null) return;

            // --- 0. Откат предыдущих сот ---
            if (isFirst)
            {
                // первый прогон — сносим всё «хвостом», оставляя базу и служебную геометрию
                Remover.RemoveFeature(md);
            }
            else
            {
                // удаляем соты предыдущей итерации:
                // StepRemove за один вызов удаляет (вырез + эскиз),
                // поэтому вызываем его столько раз, сколько СОТ реально было построено.
                for (int k = 0; k < _lastCellsCount; k++)
                    Remover.StepRemove(md);
            }

            // --- 1. Считаем центры сот для текущей итерации ---

            sideLength = size;
            distanceBetweenCombs = useFixedGap ? 0.05 : dist;

            allCenters.Clear();
            allCenters.Add(centerSketch);

            // сколько центров по формуле
            int needed = ((6 * (n - 1) + 2) * n) / 2 - (n - 1);

            int waveStart = 0;
            while (allCenters.Count < needed)
            {
                int waveEnd = allCenters.Count;
                for (int i = waveStart; i < waveEnd; i++)
                    GenerateNeighbors(allCenters[i]);

                if (waveEnd == allCenters.Count) break; // новых не появилось — всё
                waveStart = waveEnd;
            }

            // сюда будем считать СКОЛЬКО сот реально нарисовали (внутри области)
            int drawnCount = 0;

            // --- 2. Плоскость RegionTopPlane ---

            if (!iteration.SelectTopPlane(md))
                iteration.CreateTopPlane(md);

            // отрисовщик для вырезов
            SWDrawer drawer = new SWDrawer();
            drawer.init();

            // --- 3. Для каждой соты: свой эскиз + свой вырез ---

            foreach (var c in allCenters)
            {
                // сота должна целиком помещаться в область
                if (!HexInsideArea(c, sideLength))
                    continue;

                // выбираем плоскость
                md.ClearSelection2(true);
                iteration.SelectTopPlane(md);

                // создаём эскиз
                md.SketchManager.InsertSketch(true);

                // координаты в эскизе: Xsketch = X, Ysketch = -Z
                double x0 = c.X;
                double y0 = -c.Z;
                double xR = c.X + sideLength;
                double yR = -c.Z;

                md.SketchManager.CreatePolygon(
                    x0, y0, 0.0,
                    xR, yR, 0.0,
                    6, false);

                // сразу вырезаем на глубину всей области
                drawer.cutHole(HoleType.DISTANCE, true, iteration.Ly);

                md.ClearSelection2(true);
                drawnCount++;
            }

            // важно: запоминаем именно КОЛИЧЕСТВО ПОСТРОЕННЫХ сот,
            // а не количество центров (часть центров могла оказаться вне области)
            _lastCellsCount = drawnCount;

            md.ClearSelection2(true);
        }

        /// <summary>
        /// Генерация 6 соседних центров вокруг c (по XZ, Y = MaxY).
        /// </summary>
        private void GenerateNeighbors(Point3D c)
        {
            double y = iteration.MaxY; // верхняя грань по глобальному Y

            double centerStepZ = sideLength * sqrt3 + distanceBetweenCombs;
            double xHalfShift = (3 * sideLength + distanceBetweenCombs * sqrt3) / 2.0;
            double zHalfShift = (sideLength * sqrt3 + distanceBetweenCombs) / 2.0;

            var pts = new[]
            {
                new Point3D(c.X,              y, c.Z + centerStepZ),
                new Point3D(c.X,              y, c.Z - centerStepZ),
                new Point3D(c.X - xHalfShift, y, c.Z + zHalfShift),
                new Point3D(c.X - xHalfShift, y, c.Z - zHalfShift),
                new Point3D(c.X + xHalfShift, y, c.Z + zHalfShift),
                new Point3D(c.X + xHalfShift, y, c.Z - zHalfShift),
            };

            foreach (var p in pts)
            {
                var pp = RoundPoint(p);
                if (!Contains(allCenters, pp))
                    allCenters.Add(pp);
            }
        }

        /// <summary>Проверка, что шестиугольник целиком в прямоугольнике верхней грани.</summary>
        private bool HexInsideArea(Point3D c, double a)
        {
            double minX = iteration.MinX;
            double maxX = iteration.MaxX;
            double minZ = iteration.MinZ;
            double maxZ = iteration.MaxZ;

            double h = Math.Sqrt(3.0) / 2.0 * a;

            return (c.X - a >= minX) && (c.X + a <= maxX) &&
                   (c.Z - h >= minZ) && (c.Z + h <= maxZ);
        }

        /// <summary>
        /// На всякий случай: поиск последнего эскиза (ProfileFeature) в дереве фич.
        /// </summary>
        private static Feature FindLastProfileFeature(ModelDoc2 md)
        {
            Feature result = null;
            Feature f = md.FirstFeature();
            while (f != null)
            {
                if (string.Equals(f.GetTypeName(), "ProfileFeature", StringComparison.Ordinal))
                    result = f;
                f = f.GetNextFeature();
            }
            return result;
        }
    }
}
