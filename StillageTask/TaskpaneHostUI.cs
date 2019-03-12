using OfficeOpenXml;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Ortogo.SolidWorks.StillageTask
{
    [ProgId(TaskpaneIntegration.SWTASKPANE_PROGID)]
    public partial class TaskpaneHostUI : UserControl
    {

        public SldWorks SwApp { get; set; }
        public TaskpaneHostUI()
        {
            InitializeComponent();
        }

        private void TaskpaneHostUI_Load(object sender, System.EventArgs e)
        {

        }

        private void Build_Click(object sender, System.EventArgs e)
        {
            var eq = new SolidworksEquations(@"F:\frame\eq.txt");
            var frame = new Frame();

            var res = $"Высота стелажа, м:{GlobalScope.SH}, Ширина стелажа, м: {GlobalScope.SW}{System.Environment.NewLine}";
            var traversa = new TraversaCut().Select();
            if (traversa == null)
            {
                ResultCalc.Text = "Не удалось подобрать траверсу, примение другие параметры или расширьте библиотеку";
            }
            else
            {
                res += $"Подобранная траверса: {traversa}{System.Environment.NewLine}";
                try
                {
                    var support = new SupportElement().Select();
                    if (support == null)
                    {
                        ResultCalc.Text = "Не удалось подобрать стойку, примение другие параметры или расширьте библиотеку";
                        return;
                    }
                    var hConn = frame.GetConnection("СГ", support);
                    var dConn = frame.GetConnection("СД", support);

                    res += $"Подобранная стойка: {support}{System.Environment.NewLine}" +
                        $"Горизонтальная связь: {hConn}{System.Environment.NewLine}" +
                        $"Диагональная связь {dConn}{System.Environment.NewLine}";
                    ResultCalc.Text = res;

                    var countTraversa = 2 * GlobalScope.NL;
                    var sumLenTraversa = countTraversa * traversa.Length;
                    var countConn = hConn.Count + dConn.Count;
                    var sumLenConn = hConn.Count * hConn.LK + dConn.Count * dConn.LK + (countConn - 1) * GlobalScope.TUSHM;
                    res += $"Суммарная длина профиля траверс, м:{sumLenTraversa}{System.Environment.NewLine}" +
                        $"С учетом припуска на распил(t={GlobalScope.TUSHM}), м: {Math.Round(sumLenTraversa + (GlobalScope.TUSHM * sumLenTraversa), 3)}{System.Environment.NewLine}" +
                        $"Количество резов {countTraversa - 1}{System.Environment.NewLine}";
                    res += $"Суммарная длина профиля связей с учетом припуска на распил, м: {sumLenConn}{System.Environment.NewLine}" +
                        $"Количество резов {countConn - 1}{System.Environment.NewLine}";
                    ResultCalc.Text = res;

                    eq.Values["NL"].Value = GlobalScope.NL;
                    eq.Values["K"].Value = GlobalScope.K * 1000;
                    eq.Values["L"].Value = GlobalScope.L * 1000;
                    eq.Values["H"].Value = GlobalScope.SH * 1000;
                    eq.Values["NS"].Value = GlobalScope.NS;
                    eq.Values["W"].Value = GlobalScope.SW * 1000;
                    eq.Values["G"].Value = GlobalScope.G * 1000;
                    eq.Values["tkh"].Value = traversa.HSize;
                    eq.Values["tkb"].Value = traversa.BSize;
                    eq.Values["SA"].Value = support.A;
                    eq.Values["SB"].Value = support.B;
                    eq.Save();
                }
                catch (Exception ex)
                {
                    ResultCalc.Text += "Ошибка: " + ex.Message;
                }
            }

            SwApp = new SldWorks();
            var doc = OpenAssembly(@"F:\frame\рама.SLDASM");
            doc.ForceRebuild3(false);
        }

        public ModelDoc2 OpenAssembly(string DocumentPath)
        {
            return Open(DocumentPath, swDocumentTypes_e.swDocASSEMBLY);
        }



        public ModelDoc2 Open(string DocumentPath, swDocumentTypes_e Type)
        {
            var Errors = 0;
            var Warnings = 0;
            var doc = SwApp.OpenDoc6(DocumentPath, (int)Type, 0, "", ref Errors, ref Warnings);
            if (swFileLoadError_e.swFileWithSameTitleAlreadyOpen.Equals(Errors)
                || swFileLoadWarning_e.swFileLoadWarning_AlreadyOpen.Equals(Warnings)
                )
            {
                doc = SwApp.ActivateDoc(Path.GetFileName(DocumentPath));

            }
            doc = SwApp.ActiveDoc;
            return doc;
        }

        private void Calculate_Click(object sender, System.EventArgs e)
        {
            UpdateGlobalScope();
            Calculate();
        }

        private void InputDigestHandler(object sender, KeyPressEventArgs e)
        {
            var key = e.KeyChar;
            var numbers = "0123456789,.";
            var tb = (TextBox)sender;

            if (key >= 26)
            {

                if (numbers.IndexOf(key) >= 0)
                {
                    if (tb.Text.IndexOf(',') >= 0 && (key == ',' || key == '.'))
                    {
                        e.Handled = true;
                    }
                    else if (key == '.')
                    {
                        e.KeyChar = ',';
                        base.OnKeyPress(e);
                    }
                }
                else
                {
                    e.Handled = true;
                }

            }
            else
            {
                base.OnKeyPress(e);
            }

        }

        public void UpdateGlobalScope()
        {
            GlobalScope.QI = double.Parse(QI.Text);
            GlobalScope.NL = double.Parse(NLI.Text);
            GlobalScope.K = double.Parse(KI.Text);
            GlobalScope.L = double.Parse(LI.Text) / 1000;
            GlobalScope.G = double.Parse(GI.Text) / 1000;
            GlobalScope.NS = double.Parse(NSI.Text);

            GlobalScope.SH = GlobalScope.K * (GlobalScope.NL + 1) + GlobalScope.K / 2;
            GlobalScope.SW = GlobalScope.NS * GlobalScope.L;
        }


        public Calculated Calculate()
        {
            var c = default(Calculated);

            c.frame = new Frame();

            var res = $"Высота стелажа, м:{GlobalScope.SH}, Ширина стелажа, м: {GlobalScope.SW}{System.Environment.NewLine}";
            c.traversa = new TraversaCut().Select();
            if (c.traversa == null)
            {
                ResultCalc.Text = "Не удалось подобрать траверсу, примение другие параметры или расширьте библиотеку";
                c.success = false;
            }
            else
            {
                res += $"Подобранная траверса: {c.traversa}{System.Environment.NewLine}";
                try
                {
                    c.supportType = new SupportElement().Select();
                    if (c.supportType == null)
                    {
                        ResultCalc.Text = "Не удалось подобрать стойку, примение другие параметры или расширьте библиотеку";
                        c.success = false;
                        return c;
                    }
                    c.hConn = c.frame.GetConnection("СГ", c.supportType);
                    c.dConn = c.frame.GetConnection("СД", c.supportType);

                    res += $"Подобранная стойка: {c.supportType}{System.Environment.NewLine}" +
                        $"Горизонтальная связь: {c.hConn}{System.Environment.NewLine}" +
                        $"Диагональная связь {c.dConn}{System.Environment.NewLine}";
                    ResultCalc.Text = res;

                    c.countTraversa = 2 * GlobalScope.NL;
                    c.sumLenTraversa = c.countTraversa * c.traversa.Length;
                    c.countConn = c.hConn.Count + c.dConn.Count;
                    c.sumLenConn = c.hConn.Count * c.hConn.LK + c.dConn.Count * c.dConn.LK + (c.countConn - 1) * GlobalScope.TUSHM;
                    res += $"Суммарная длина профиля траверс, м:{c.sumLenTraversa}{System.Environment.NewLine}" +
                        $"С учетом припуска на распил(t={GlobalScope.TUSHM}), м: {Math.Round(c.sumLenTraversa + (GlobalScope.TUSHM * c.sumLenTraversa), 3)}{System.Environment.NewLine}" +
                        $"Количество резов {c.countTraversa - 1}{System.Environment.NewLine}";
                    res += $"Суммарная длина профиля связей с учетом припуска на распил, м: {c.sumLenConn}{System.Environment.NewLine}" +
                        $"Количество резов {c.countConn - 1}{System.Environment.NewLine}";
                    ResultCalc.Text = res;
                }
                catch (Exception ex)
                {
                    ResultCalc.Text += "Ошибка: " + ex.Message;
                    c.success = false;
                }
            }
            c.success = true;
            return c;
        }

        private void GenerateSpec_Click(object sender, EventArgs e)
        {
            ExportToExcel();
        }

        public void ExportToExcel()
        {


            UpdateGlobalScope();
            var c = Calculate();


            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Книга Excel (*.xlsx)|*.xlsx",
                DefaultExt = "xlsx"
            };

            var desctopPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.DesktopDirectory);
            saveFileDialog.InitialDirectory = desctopPath;
            saveFileDialog.FileName = $"Р.{GlobalScope.SH * 1000}.{GlobalScope.G}.{c.supportType}.xlsx".Replace(",", ".");

            var dialogRes = saveFileDialog.ShowDialog();
            if (dialogRes.Equals(DialogResult.Cancel))
                return;

            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Спецификация");
                ws.Cells["A1"].Value = $"Комплектация рамы Р.{GlobalScope.SH * 1000}.{GlobalScope.G}.{c.supportType}";
                ws.Cells["B1"].Value = $"Тип траверсы:";
                ws.Cells["B2"].Value = $"{c.traversa}";
                ws.Cells["B3"].Value = $"Количество траверс";
                ws.Cells["B4"].Value = $"{c.countTraversa}";
                ws.Cells["B5"].Value = $"Длина материала";
                ws.Cells["B6"].Value = $"{Math.Round(c.sumLenTraversa + (GlobalScope.TUSHM * c.sumLenTraversa), 3)}";
                ws.Cells["C1"].Value = $"Тип стойки";
                ws.Cells["C2"].Value = $"{c.supportType}";
                ws.Cells["C3"].Value = $"Длина";
                ws.Cells["C4"].Value = $"{4 * GlobalScope.SH}";
                ws.Cells["D1"].Value = $"Связь";
                ws.Cells["D2"].Value = $"{c.hConn}";
                ws.Cells["D3"].Value = $"Количество";
                ws.Cells["D4"].Value = $"{c.hConn.Count}";
                ws.Cells["D5"].Value = $"Длина";
                ws.Cells["D6"].Value = $"{(c.hConn.Count - 1) * c.hConn.LK}";
                ws.Cells["E1"].Value = $"Связь";
                ws.Cells["E2"].Value = $"{c.dConn}";
                ws.Cells["E3"].Value = $"Количество";
                ws.Cells["E4"].Value = $"{c.dConn.Count}";
                ws.Cells["E5"].Value = $"Длина";
                ws.Cells["E6"].Value = $"{(c.dConn.Count - 1) * c.dConn.LK}";

                p.SaveAs(new FileInfo(saveFileDialog.FileName));
                MessageBox.Show("Успешно сохранено");
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }

    public struct Calculated
    {
        public Frame frame { get; set; }
        public TraversaCut.Traversa traversa { get; set; }
        public SupportElement.Type supportType { get; set; }
        public Frame.Connection hConn { get; set; }
        public Frame.Connection dConn { get; set; }

        public double countTraversa { get; set; }
        public double sumLenTraversa { get; set; }
        public int countConn { get; set; }
        public double sumLenConn { get; set; }

        public bool success;
    }
}
