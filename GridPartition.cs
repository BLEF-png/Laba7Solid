using System;
using System.Collections.Generic;
using System.Linq;
using SolidWorks.Interop.sldworks;
using Temp; // SWDrawer, HoleType

namespace Lab5_Kaluzhny
{
    public enum WeightMode
    {
        CenterQuadratic,
        ShiftedCenter,
        Gradient,
        Ring,
        Stripes
    }

    public class WeightOptions
    {
        public WeightMode Mode;
        public double P1;
        public double P2;
        public double P3;
    }

    public class GridElement
    {
        public int I;
        public int J;
        public Point3D Center;
        public double Weight;
        public double Radius;
    }

    /// <summary>
    /// Разбиение области интегрирования на прямоугольную сетку + веса узлов.
    /// </summary>
    public class GridPartition
    {
        private readonly Iteration _iteration;
        public List<GridElement> Elements { get; } = new List<GridElement>();

        public GridPartition(Iteration iteration)
        {
            _iteration = iteration ?? throw new ArgumentNullException(nameof(iteration));
        }

        // --------------------------- построение сетки ---------------------------

        public void BuildGrid(int nx, int ny)
        {
            if (nx <= 0 || ny <= 0) throw new ArgumentException("nx, ny must be > 0");

            Elements.Clear();

            double dx = _iteration.Lx / nx;
            double dz = _iteration.Lz / ny;

            double minX = _iteration.MinX;
            double minZ = _iteration.MinZ;
            double y = _iteration.MaxY; // верхняя грань

            for (int j = 0; j < ny; j++)
            {
                for (int i = 0; i < nx; i++)
                {
                    double cx = minX + (i + 0.5) * dx;
                    double cz = minZ + (j + 0.5) * dz;

                    Elements.Add(new GridElement
                    {
                        I = i,
                        J = j,
                        Center = new Point3D(cx, y, cz),
                        Weight = 1.0
                    });
                }
            }
        }

        // --------------------------- вычисление весов ---------------------------

        private static double Clamp01(double t) =>
            t < 0 ? 0 : (t > 1 ? 1 : t);

        public void ComputeWeights(WeightOptions opt)
        {
            if (Elements.Count == 0) return;
            if (opt == null) throw new ArgumentNullException(nameof(opt));

            double minX = _iteration.MinX;
            double maxX = _iteration.MaxX;
            double minZ = _iteration.MinZ;
            double maxZ = _iteration.MaxZ;
            double cxGeom = _iteration.TopCenter.X;
            double czGeom = _iteration.TopCenter.Z;
            double hx = _iteration.Lx / 2.0;
            double hz = _iteration.Lz / 2.0;

            foreach (var el in Elements)
            {
                double w = 1.0;

                switch (opt.Mode)
                {
                    // ---------- 1) CenterQuadratic ----------
                    case WeightMode.CenterQuadratic:
                        {
                            double rx = hx > 0 ? (el.Center.X - cxGeom) / hx : 0.0;
                            double rz = hz > 0 ? (el.Center.Z - czGeom) / hz : 0.0;
                            double r = Math.Sqrt(rx * rx + rz * rz);
                            double rClamped = Math.Min(r, 1.0);
                            double core = 1.0 - rClamped * rClamped; // 1..0
                            w = 1.0 + core;                         // 1..2
                            break;
                        }

                    // ---------- 2) ShiftedCenter ----------
                    case WeightMode.ShiftedCenter:
                        {
                            double fx = Clamp01(opt.P1); // 0..1
                            double fz = Clamp01(opt.P2); // 0..1

                            double cx = minX + fx * (maxX - minX);
                            double cz = minZ + fz * (maxZ - minZ);

                            double rx = (maxX > minX) ? (el.Center.X - cx) / (maxX - minX) : 0.0;
                            double rz = (maxZ > minZ) ? (el.Center.Z - cz) / (maxZ - minZ) : 0.0;

                            double r = Math.Sqrt(rx * rx + rz * rz);
                            double rClamped = Math.Min(r, 1.0);
                            double core = 1.0 - rClamped * rClamped;
                            w = 1.0 + core;
                            break;
                        }

                    // ---------- 3) Gradient ----------
                    case WeightMode.Gradient:
                        {
                            int dir = (int)Math.Round(opt.P1); // 0/1/2
                            double tx = (maxX > minX) ? (el.Center.X - minX) / (maxX - minX) : 0.0;
                            double tz = (maxZ > minZ) ? (el.Center.Z - minZ) / (maxZ - minZ) : 0.0;
                            double t;

                            if (dir == 1)      // по Z
                                t = tz;
                            else if (dir == 2) // диагональ
                                t = 0.5 * (tx + tz);
                            else               // по X
                                t = tx;

                            t = Clamp01(t);
                            w = 1.0 + t;       // 1..2
                            break;
                        }

                    // ---------- 4) Ring ----------
                    case WeightMode.Ring:
                        {
                            double rx = hx > 0 ? (el.Center.X - cxGeom) / hx : 0.0;
                            double rz = hz > 0 ? (el.Center.Z - czGeom) / hz : 0.0;
                            double rNorm = Math.Sqrt(rx * rx + rz * rz); // 0..~1.4
                            rNorm = Math.Min(rNorm, 1.0);

                            double r0 = Clamp01(opt.P1);
                            double sigma = opt.P2 > 0 ? opt.P2 : 0.2;
                            sigma = Math.Max(0.05, Math.Min(1.0, sigma));

                            double value = Math.Exp(-Math.Pow(rNorm - r0, 2.0) / (2.0 * sigma * sigma));
                            w = 1.0 + value;  // 1..2
                            break;
                        }

                    // ---------- 5) Stripes ----------
                    case WeightMode.Stripes:
                        {
                            int bands = (int)Math.Round(opt.P1);
                            if (bands < 1) bands = 3;
                            int orient = (int)Math.Round(opt.P2);

                            double tx = (maxX > minX) ? (el.Center.X - minX) / (maxX - minX) : 0.0;
                            double tz = (maxZ > minZ) ? (el.Center.Z - minZ) / (maxZ - minZ) : 0.0;
                            double t = (orient == 1) ? tz : tx;

                            t = Clamp01(t);

                            double arg = 2.0 * Math.PI * bands * t;
                            double stripe = 0.5 * (1.0 + Math.Sin(arg)); // 0..1
                            w = 1.0 + stripe;                            // 1..2
                            break;
                        }

                    default:
                        w = 1.0;
                        break;
                }

                el.Weight = w;
            }
        }

        // --------------------------- перевод весов в радиусы ---------------------------

        public void AssignRadii(double rMin, double rMax)
        {
            if (Elements.Count == 0) return;

            if (rMax <= 0)
                throw new ArgumentException("rMax must be > 0");

            if (rMin <= 0 || rMin >= rMax)
                rMin = 0.3 * rMax; // авто-диапазон, если пользователь не задал нормально

            double minW = Elements.Min(e => e.Weight);
            double maxW = Elements.Max(e => e.Weight);

            if (Math.Abs(maxW - minW) < 1e-9)
            {
                double rMid = 0.5 * (rMin + rMax);
                foreach (var el in Elements)
                    el.Radius = rMid;
                return;
            }

            foreach (var el in Elements)
            {
                double t = (el.Weight - minW) / (maxW - minW); // 0..1
                t = Clamp01(t);
                el.Radius = rMin + t * (rMax - rMin);
            }
        }

        // --------------------------- построение сот по сетке ---------------------------

        private bool HexInsideArea(Point3D c, double a)
        {
            double minX = _iteration.MinX;
            double maxX = _iteration.MaxX;
            double minZ = _iteration.MinZ;
            double maxZ = _iteration.MaxZ;

            double h = Math.Sqrt(3.0) / 2.0 * a;

            return (c.X - a >= minX) && (c.X + a <= maxX) &&
                   (c.Z - h >= minZ) && (c.Z + h <= maxZ);
        }

        /// <summary>
        /// Построение сот по сетке. ТЕПЕРЬ:
        /// для КАЖДОГО элемента – отдельный эскиз и отдельный вырез.
        /// </summary>
        public void BuildHoneycombByGrid(ModelDoc2 md)
        {
            if (md == null) return;
            if (Elements.Count == 0) return;

            // убеждаемся, что плоскость есть
            if (!_iteration.SelectTopPlane(md))
                _iteration.CreateTopPlane(md);

            SWDrawer drawer = new SWDrawer();
            drawer.init();

            foreach (var el in Elements)
            {
                double a = el.Radius;
                if (a <= 1e-8) continue;
                if (!HexInsideArea(el.Center, a)) continue;

                // выбираем плоскость
                md.ClearSelection2(true);
                _iteration.SelectTopPlane(md);

                // создаём эскиз
                md.SketchManager.InsertSketch(true);

                double x0 = el.Center.X;
                double y0 = -el.Center.Z;
                double xR = x0 + a;
                double yR = y0;

                md.SketchManager.CreatePolygon(
                    x0, y0, 0.0,
                    xR, yR, 0.0,
                    6, false);

                // сразу вырезаем
                drawer.cutHole(HoleType.DISTANCE, true, _iteration.Ly);

                md.ClearSelection2(true);
            }

            md.ClearSelection2(true);
        }
    }
}
