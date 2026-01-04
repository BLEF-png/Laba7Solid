using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Globalization;

namespace Lab5_Kaluzhny
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        // ----------------------------- SW поля -----------------------------
        private SldWorks swApp;
        private ModelDoc2 swModel;
        private SketchManager swSketchManager;
        private SelectionMgr swSelMgr;

        // твои классы
        public Iteration iteration;
        public Honeycomb honeycomb;

        // --- для разбиения области ---
        private GridPartition _gridPartition;
        private int _gridNx = 1;
        private int _gridNy = 1;
        private WeightOptions _weightOptions = new WeightOptions
        {
            Mode = WeightMode.CenterQuadratic,
            P1 = 0.5,
            P2 = 0.5,
            P3 = 0.0
        };

        // итерационные параметры (старый режим)
        public double Dmin, Dmax, Dstep;
        public double Rmin, Rmax, Rstep;
        public int start, end, step;
        public double Dmin1, Dmax1;
        public double Rmin1, Rmax1;

        double precision = 1000; // мм -> м

        // флаг: временно отключаем автопересчёт при заполнении таблицы
        private bool _suppressGridEvents = false;

        // ----------------------------- SW bootstrap -----------------------------

        private bool GetSolidworks()
        {
            try
            {
                swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                swApp.Visible = true;
            }
            catch
            {
                try
                {
                    swApp = Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application")) as SldWorks;
                    swApp.Visible = true;
                }
                catch
                {
                    MessageBox.Show("Не удалось запустить SolidWorks.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }

            if (swApp.ActiveDoc is ModelDoc2 ad)
            {
                swModel = ad;
                swModel.SetUnits((short)swLengthUnit_e.swMM, (short)swFractionDisplay_e.swDECIMAL, 0, 0, false);
                swSketchManager = (SketchManager)swModel.SketchManager;
                swSelMgr = (SelectionMgr)swModel.SelectionManager;
            }
            return true;
        }

        private void btnConnectPart_Click(object sender, EventArgs e)
        {
            try
            {
                if (!GetSolidworks()) return;

                using (var ofd = new OpenFileDialog())
                {
                    ofd.Title = "Выберите деталь SolidWorks";
                    ofd.Filter = "SolidWorks Part (*.sldprt)|*.sldprt|Все файлы|*.*";
                    ofd.CheckFileExists = true;
                    ofd.Multiselect = false;

                    if (ofd.ShowDialog(this) != DialogResult.OK) return;

                    int errs = 0, warns = 0;

                    int opts = (int)swOpenDocOptions_e.swOpenDocOptions_Silent
                             | (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel;

                    var doc = swApp.OpenDoc6(
                        ofd.FileName,
                        (int)swDocumentTypes_e.swDocPART,
                        opts,
                        "", ref errs, ref warns) as ModelDoc2;

                    if (doc == null)
                    {
                        MessageBox.Show($"Не удалось открыть файл.\nSW error={errs}, warn={warns}",
                            "Ошибка открытия", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    int actErr = 0;
                    swApp.ActivateDoc2(doc.GetTitle(), true, ref actErr);

                    swModel = doc;
                    swModel.SetUnits((short)swLengthUnit_e.swMM, (short)swFractionDisplay_e.swDECIMAL, 0, 0, false);
                    swSketchManager = (SketchManager)swModel.SketchManager;
                    swSelMgr = (SelectionMgr)swModel.SelectionManager;

                    Text = $"Итерационная модель — {swModel.GetTitle()}";
                }
            }
            catch (COMException comx)
            {
                MessageBox.Show("COM ошибка: " + comx.Message, "SolidWorks", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message, "Подключение к детали", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            GetSolidworks();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            SetupGridForManualInput();
            GetSolidworks();
            try { this.btnConnectPart.Click += btnConnectPart_Click; } catch { }

            // дефолтные точки области, если пользователь ничего не ввёл
            if (tbXMin != null)
            {
                if (string.IsNullOrWhiteSpace(tbXMin.Text)) tbXMin.Text = "-40";
                if (string.IsNullOrWhiteSpace(tbYMin.Text)) tbYMin.Text = "0";
                if (string.IsNullOrWhiteSpace(tbZMin.Text)) tbZMin.Text = "-42";

                if (string.IsNullOrWhiteSpace(tbXMax.Text)) tbXMax.Text = "40";
                if (string.IsNullOrWhiteSpace(tbYMax.Text)) tbYMax.Text = "20";
                if (string.IsNullOrWhiteSpace(tbZMax.Text)) tbZMax.Text = "18";
            }
        }

        // ----------------------------- Область + старый режим итераций -----------------------------

        private void button2_Click_1(object sender, EventArgs e)
        {
            if (swModel == null)
            {
                MessageBox.Show("Сначала подключите деталь.", "Нет модели",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // читаем две диагональные точки из текстбоксов (в мм)
            if (!TryReadPointFromUI(tbXMin, tbYMin, tbZMin, out Point3D pMinMm) ||
                !TryReadPointFromUI(tbXMax, tbYMax, tbZMax, out Point3D pMaxMm))
            {
                MessageBox.Show("Некорректно заданы координаты двух точек области.\r\n" +
                                "Проверьте Xmin, Ymin, Zmin и Xmax, Ymax, Zmax.",
                                "Рассчитать область", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // переводим из мм в метры
            var pointMin = new Point3D(pMinMm.X / precision, pMinMm.Y / precision, pMinMm.Z / precision);
            var pointMax = new Point3D(pMaxMm.X / precision, pMaxMm.Y / precision, pMaxMm.Z / precision);

            // создаём / пересчитываем область
            iteration = new Iteration(pointMax, pointMin);
            iteration.createPlane(swModel);
            honeycomb = new Honeycomb(iteration, iteration.PointCenter);

            // заполняем таблицу 8-ю вершинами
            _suppressGridEvents = true;
            try
            {
                dataGridView1.Rows.Clear();
                foreach (var p in iteration.PointsPlane)
                {
                    int r = dataGridView1.Rows.Add();
                    var row = dataGridView1.Rows[r];
                    row.Cells["NColumn"].Value = r + 1;
                    row.Cells["XColumn"].Value = p.X * precision;
                    row.Cells["YColumn"].Value = p.Y * precision;
                    row.Cells["ZColumn"].Value = p.Z * precision;
                }
            }
            finally
            {
                _suppressGridEvents = false;
            }

            label1.Text = $"lx = {iteration.Lx * precision}\r\n" +
                          $"ly = {iteration.Ly * precision}\r\n" +
                          $"lz = {iteration.Lz * precision}";
        }

        private bool TryReadPointFromUI(TextBox tbX, TextBox tbY, TextBox tbZ, out Point3D ptMm)
        {
            ptMm = null;
            if (tbX == null || tbY == null || tbZ == null)
                return false;

            double x, y, z;
            if (!TryParseTextBoxMm(tbX, out x)) return false;
            if (!TryParseTextBoxMm(tbY, out y)) return false;
            if (!TryParseTextBoxMm(tbZ, out z)) return false;

            ptMm = new Point3D(x, y, z);
            return true;
        }

        private bool TryParseTextBoxMm(TextBox tb, out double value)
        {
            value = 0;
            if (tb == null) return false;

            string s = (tb.Text ?? "").Trim().Replace(',', '.');
            if (!double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out value))
                return false;

            return true;
        }

        // ----------------------------- Старый блок параметров итераций -----------------------------

        private void button3_Click_1(object sender, EventArgs e)
        {
            Dmin = Convert.ToDouble(DminTB.Text) / precision;
            Dmax = Convert.ToDouble(DmaxTB.Text) / precision;
            Dstep = Convert.ToDouble(DstepTB.Text) / precision;

            start = Convert.ToInt32(startTB.Text);
            end = Convert.ToInt32(endTB.Text);
            step = Convert.ToInt32(stepTB.Text);

            Dmin1 = Convert.ToDouble(DminTB.Text) / precision;
            Dmax1 = Convert.ToDouble(DmaxTB.Text) / precision;
            Rmin1 = Convert.ToDouble(tbRMin.Text) / precision;
            Rmax1 = Convert.ToDouble(tbRMax.Text) / precision;

            Rmin = Convert.ToDouble(tbRMin.Text) / precision;
            Rmax = Convert.ToDouble(tbRMax.Text) / precision;
            Rstep = Convert.ToDouble(tbRStep.Text) / precision;

            tbParameters.Text = "";
            for (int i = start; i <= end; i += step)
            {
                var dMaxLocal = checkBoxD.Checked ? 0.05 : Dmax;
                int numberOfCumbs = ((6 * (i - 1) + 2) * i) / 2 - (i - 1);
                int h = 2 * i - 1;

                tbParameters.Text += $"Итерация: {i}\r\n";
                tbParameters.Text += $"Количество всех элементов: {numberOfCumbs}\r\n";
                tbParameters.Text += $"Длина максимального ряда: {h}\r\n";
                tbParameters.Text += $"Размер: {Rmax * precision}\r\n";
                tbParameters.Text += $"Расстояние: {dMaxLocal * precision}\r\n";
                tbParameters.Text += "-----------------------------------------------------\r\n";

                Rmax -= Rstep;
                Dmax -= Dstep;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (swModel == null)
            {
                MessageBox.Show("Сначала подключите деталь (.sldprt).", "Нет модели",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (iteration == null)
            {
                MessageBox.Show("Сначала нажмите «Рассчитать область», чтобы создать область и плоскость.", "Нет области",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (honeycomb == null)
                honeycomb = new Honeycomb(iteration, iteration.PointCenter);

            if (string.IsNullOrWhiteSpace(DmaxTB.Text)) DmaxTB.Text = "70";
            if (string.IsNullOrWhiteSpace(tbRMax.Text)) tbRMax.Text = "200";
            if (string.IsNullOrWhiteSpace(DstepTB.Text)) DstepTB.Text = "10";
            if (string.IsNullOrWhiteSpace(tbRStep.Text)) tbRStep.Text = "50";
            if (string.IsNullOrWhiteSpace(startTB.Text)) startTB.Text = "1";
            if (string.IsNullOrWhiteSpace(endTB.Text)) endTB.Text = "1";
            if (string.IsNullOrWhiteSpace(stepTB.Text)) stepTB.Text = "1";

            double dStepLocal, rStepLocal;
            double.TryParse(DminTB.Text.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out Dmin);
            double.TryParse(DmaxTB.Text.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out Dmax);
            double.TryParse(DstepTB.Text.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out dStepLocal);

            int.TryParse(startTB.Text, out start);
            int.TryParse(endTB.Text, out end);
            int.TryParse(stepTB.Text, out step);
            if (step <= 0) step = 1;

            double.TryParse(tbRMin.Text.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out Rmin);
            double.TryParse(tbRMax.Text.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out Rmax);
            double.TryParse(tbRStep.Text.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out rStepLocal);

            Dmin /= precision; Dmax /= precision; Dstep = dStepLocal / precision;
            Rmin /= precision; Rmax /= precision; Rstep = rStepLocal / precision;

            Dmin1 = Dmin; Dmax1 = Dmax;
            Rmin1 = Rmin; Rmax1 = Rmax;

            bool isFirst = true;
            for (int i = start; i <= end; i += step)
            {
                double dist = checkBoxD.Checked ? 0.05 : Dmax1;
                double size = Rmax1;

                if (!iteration.SelectWorkPlane(swModel))
                    iteration.createPlane(swModel);

                honeycomb.Cal(swModel, i, checkBoxD.Checked, dist, isFirst, size);
                isFirst = false;

                if (Dmin1 <= Dmax1) Dmax1 -= Dstep;
                if (Rmin1 <= Rmax1) Rmax1 -= Rstep;
            }

            swModel.ForceRebuild3(false);
        }

        // ----------------------------- ВЕСА / СЕТКА -----------------------------

        private WeightOptions BuildWeightOptionsFromUI()
        {
            var opt = new WeightOptions();

            string modeText = cbWeightMode?.SelectedItem?.ToString() ?? "Center";

            switch (modeText)
            {
                case "Center":
                    opt.Mode = WeightMode.CenterQuadratic; break;
                case "Shifted center":
                    opt.Mode = WeightMode.ShiftedCenter; break;
                case "Gradient":
                    opt.Mode = WeightMode.Gradient; break;
                case "Ring":
                    opt.Mode = WeightMode.Ring; break;
                case "Stripes":
                    opt.Mode = WeightMode.Stripes; break;
                default:
                    opt.Mode = WeightMode.CenterQuadratic; break;
            }

            opt.P1 = ParseDoubleOrDefault(tbModeP1?.Text, 0.5);
            opt.P2 = ParseDoubleOrDefault(tbModeP2?.Text, 0.5);

            return opt;
        }

        private void btnSplitRegion_Click_1(object sender, EventArgs e)
        {
            if (iteration == null)
            {
                MessageBox.Show("Сначала нажмите «Рассчитать область», чтобы задать область интегрирования.",
                    "Нет области", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (!int.TryParse(tbGridNx.Text, out _gridNx) || _gridNx <= 0)
            {
                MessageBox.Show("Некорректный Nx (по X).", "Размер сетки", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!int.TryParse(tbGridNy.Text, out _gridNy) || _gridNy <= 0)
            {
                MessageBox.Show("Некорректный Ny (по Z).", "Размер сетки", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _gridPartition = new GridPartition(iteration);
            _gridPartition.BuildGrid(_gridNx, _gridNy);

            _weightOptions = BuildWeightOptionsFromUI();
            _gridPartition.ComputeWeights(_weightOptions);

            tbParameters.Text = "";
            foreach (var el in _gridPartition.Elements)
            {
                tbParameters.AppendText(
                    $"i={el.I}, j={el.J}, " +
                    $"X={el.Center.X * precision:0.###} мм, " +
                    $"Z={el.Center.Z * precision:0.###} мм, " +
                    $"W={el.Weight:0.000}\r\n");
            }

            MessageBox.Show($"Разбиение выполнено: {_gridNx} × {_gridNy} элементов.",
                "Сетка", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnBuildByGrid_Click_1(object sender, EventArgs e)
        {
            if (swModel == null)
            {
                MessageBox.Show("Сначала подключите деталь (.sldprt).",
                    "Нет модели", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (iteration == null)
            {
                MessageBox.Show("Сначала нажмите «Рассчитать область», чтобы задать область интегрирования.",
                    "Нет области", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (_gridPartition == null || _gridPartition.Elements.Count == 0)
            {
                MessageBox.Show("Сначала нажмите «Разбить на элементы».",
                    "Нет сетки", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            _weightOptions = BuildWeightOptionsFromUI();
            _gridPartition.ComputeWeights(_weightOptions);

            double rMinMm, rMaxMm;
            bool okMin = double.TryParse(tbRMin.Text.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out rMinMm);
            bool okMax = double.TryParse(tbRMax.Text.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out rMaxMm);

            double rMin, rMax;

            if (okMin && okMax && rMaxMm > 0)
            {
                rMin = rMinMm / precision;
                rMax = rMaxMm / precision;
            }
            else
            {
                double dx = iteration.Lx / _gridNx;
                double dz = iteration.Lz / _gridNy;
                double maxByX = dx / 2.0;
                double maxByZ = dz / Math.Sqrt(3.0);
                rMax = Math.Min(maxByX, maxByZ);
                rMin = 0.3 * rMax;
            }

            _gridPartition.AssignRadii(rMin, rMax);

            Remover.RemoveFeature(swModel);
            _gridPartition.BuildHoneycombByGrid(swModel);
            swModel.ForceRebuild3(false);
        }

        // ----------------------------- КНОПКА «СПРАВКА» -----------------------------

        private void btnHelp_Click_1(object sender, EventArgs e)
        {
            string msg =
@"СПРАВКА

--- Задание области интегрирования ---

1. Ввести координаты двух диагональных точек области в миллиметрах:
   Xmin, Ymin, Zmin и Xmax, Ymax, Zmax (минимальная и максимальная точка параллелепипеда).
2. Нажать «Рассчитать область»:
   – по двум точкам автоматически вычисляются остальные 6 вершин области,
   – создаётся вспомогательная плоскость на верхней грани (RegionTopPlane),
   – таблица сверху заполняется всеми 8 вершинами.

--- Построение по итерациям ---

1. Открыть SolidWorks (кнопка «Открыть SolidWorks» при необходимости).
2. Загрузить деталь (.sldprt).
3. Задать две точки области и нажать «Рассчитать область».
4. В блоке «Итерации / Расстояние (D) / Размер (R)» задать параметры и нажать «Задать параметры».
5. Нажать кнопку «Нарисовать по итерациям» – соты будут построены по классическому
   итерационному алгоритму (без разбиения на элементы и весов).

--- Построение по весам (сетка) ---

1. Задать область (две точки) и нажать «Рассчитать область».
2. В блоке «Режим построения сот» выбрать режим:
   - Center        – максимум веса в геометрическом центре, квадратичный спад к краям.
   - Shifted center – максимум веса в смещённом центре.
   - Gradient      – линейный градиент веса.
   - Ring          – кольцевое распределение, максимум по радиусу.
   - Stripes       – полосы (чередующиеся зоны веса).

3. Параметры режимов:
   • Center
       Параметр 1 и 2 не используются.

   • Shifted center
       Параметр 1: смещение центра по X (0 – левый край, 1 – правый край, 0.5 – середина).
       Параметр 2: смещение центра по Z (0 – ближний край, 1 – дальний, 0.5 – середина).

   • Gradient
       Параметр 1:
           0 – градиент вдоль X (слева направо),
           1 – градиент вдоль Z (от ближнего к дальнему краю),
           2 – градиент по диагонали.
       Параметр 2 не используется.

   • Ring
       Параметр 1: радиус кольца r0 в относительных координатах (0 – центр, 1 – край).
       Параметр 2: ширина кольца (рекомендуется 0.1–0.3).

   • Stripes
       Параметр 1: количество полос (целое число, например 3 или 5).
       Параметр 2:
           0 – полосы меняются по X (полосы идут поперёк X),
           1 – полосы меняются по Z.

4. В «Размер сетки» указать количество ячеек по X и Z (Nx × Ny).
5. Нажать «Разбить на элементы» – будет построена сетка, посчитаны узлы и их веса
   (они отображаются в левом текстовом поле).
6. В блоке «Размер (R)» задать минимальный и максимальный размер сот (Rmin, Rmax) в мм.
   Если значения не заданы или одинаковы – диапазон размеров подбирается автоматически.
7. Нажать «Построить» – будет построен один эскиз и один вырез с учётом сетки и весов.

--- Кнопка «Удалить вырез» ---

Удаляет все построенные сотовые вырезы, эскизы этих вырезов и вспомогательную плоскость
области интегрирования. Удобно использовать для повторных тестов режимов без перезапуска
программы и без повторного открытия детали.

--- Кнопка «Выйти» ---

Закрывает все документы SolidWorks без сохранения и завершает работу SolidWorks и программы.";

            MessageBox.Show(msg, "Справка", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // ----------------------------- КНОПКА «Удалить вырез» -----------------------------

        private void btnDeleteCut_Click_1(object sender, EventArgs e)
        {
            if (swModel == null)
            {
                MessageBox.Show("Нет активной модели для очистки.", "Удалить вырез",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                Remover.RemoveFeature(swModel);

                if (iteration != null)
                    iteration.DeleteRegionGeometry(swModel);

                swModel.ForceRebuild3(false);
            }
            catch (COMException ex)
            {
                MessageBox.Show("COM-ошибка при удалении выреза: " + ex.Message,
                    "Удалить вырез", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при удалении выреза: " + ex.Message,
                    "Удалить вырез", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ----------------------------- КНОПКА «Выйти» -----------------------------

        private void btnExit_Click_1(object sender, EventArgs e)
        {
            try
            {
                if (swApp != null)
                {
                    swApp.CloseAllDocuments(true); // без сохранения
                    swApp.ExitApp();
                    swApp = null;
                    swModel = null;
                }
            }
            catch (COMException ex)
            {
                MessageBox.Show("COM-ошибка при завершении SolidWorks: " + ex.Message,
                    "Выйти", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при завершении SolidWorks: " + ex.Message,
                    "Выйти", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Close();
            }
        }

        // ----------------------------- Таблица / пересчёт области -----------------------------

        private void SetupGridForManualInput()
        {
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.EditMode = DataGridViewEditMode.EditOnEnter;

            if (dataGridView1.Rows.Count < 8)
            {
                dataGridView1.Rows.Clear();
                for (int i = 1; i <= 8; i++)
                {
                    int row = dataGridView1.Rows.Add();
                    dataGridView1.Rows[row].Cells["NColumn"].Value = i;
                }
            }

            dataGridView1.CellValidating += (s, e) =>
            {
                var col = dataGridView1.Columns[e.ColumnIndex].Name;
                if (col == "XColumn" || col == "YColumn" || col == "ZColumn")
                {
                    var txt = (e.FormattedValue ?? "").ToString().Trim().Replace(',', '.');
                    double tmp;
                    if (!double.TryParse(txt, NumberStyles.Float, CultureInfo.InvariantCulture, out tmp))
                    {
                        e.Cancel = true;
                        dataGridView1.Rows[e.RowIndex].ErrorText = "Введите число (мм)";
                    }
                    else dataGridView1.Rows[e.RowIndex].ErrorText = null;
                }
            };

            dataGridView1.CellEndEdit += DataGridView1_CellEndEdit;
            dataGridView1.CellValueChanged += DataGridView1_CellValueChanged;
        }

        private void DataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (_suppressGridEvents) return;
            dataGridView1.EndEdit();
            TryRebuildAreaFromGrid();
        }

        private void DataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (_suppressGridEvents) return;
            TryRebuildAreaFromGrid();
        }

        private void TryRebuildAreaFromGrid()
        {
            if (swModel == null || swApp == null || iteration == null) return;

            var pts = new List<(double X, double Y, double Z)>();

            foreach (DataGridViewRow r in dataGridView1.Rows)
            {
                if (r.IsNewRow) continue;

                double xmm, ymm, zmm;
                if (TryParseCell(r.Cells["XColumn"].Value, out xmm) &&
                    TryParseCell(r.Cells["YColumn"].Value, out ymm) &&
                    TryParseCell(r.Cells["ZColumn"].Value, out zmm))
                {
                    pts.Add((xmm / precision, ymm / precision, zmm / precision));
                }
            }

            if (pts.Count == 0) return;

            double minX = pts.Min(p => p.X), maxX = pts.Max(p => p.X);
            double minY = pts.Min(p => p.Y), maxY = pts.Max(p => p.Y);
            double minZ = pts.Min(p => p.Z), maxZ = pts.Max(p => p.Z);

            var p1 = new Point3D(minX, minY, minZ);
            var p2 = new Point3D(maxX, maxY, maxZ);

            iteration.RecomputeBox(p2, p1);
            honeycomb = new Honeycomb(iteration, iteration.PointCenter);

            if (label1 != null)
                label1.Text = $"lx = {iteration.Lx * precision}\r\n" +
                              $"ly = {iteration.Ly * precision}\r\n" +
                              $"lz = {iteration.Lz * precision}";
        }

        private static bool TryParseCell(object v, out double d)
        {
            var s = (v ?? "").ToString().Trim().Replace(',', '.');
            return double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out d);
        }

        private static double ParseDoubleOrDefault(string text, double def)
        {
            if (string.IsNullOrWhiteSpace(text)) return def;
            double val;
            if (double.TryParse(text.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out val))
                return val;
            return def;
        }
    }
}
