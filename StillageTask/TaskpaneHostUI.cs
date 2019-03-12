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
            BuildFrame();
        }

        public void BuildFrame()
        {
            UpdateGlobalScope();
            var c = Calculator.Calculate();

            if (!c.success)
            {
                ResultCalc.Text = c.errorMessage;
                return;
            }

            ResultCalc.Text = c.message;

            var eq = new SolidworksEquations(@"F:\frame\eq.txt");

            eq.Values["NL"].Value = GlobalScope.NL;
            eq.Values["K"].Value = GlobalScope.K * 1000;
            eq.Values["L"].Value = GlobalScope.L * 1000;
            eq.Values["H"].Value = GlobalScope.SH * 1000;
            eq.Values["NS"].Value = GlobalScope.NS;
            eq.Values["W"].Value = GlobalScope.SW * 1000;
            eq.Values["G"].Value = GlobalScope.G * 1000;
            eq.Values["tkh"].Value = c.traversa.HSize;
            eq.Values["tkb"].Value = c.traversa.BSize;
            eq.Values["SA"].Value = c.supportType.A;
            eq.Values["SB"].Value = c.supportType.B;
            eq.Save();

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
            var c = Calculator.Calculate();

            if (!c.success)
            {
                ResultCalc.Text = c.errorMessage;
                return;
            }

            ResultCalc.Text = c.message;
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

        private void GenerateSpec_Click(object sender, EventArgs e)
        {
            ExportToExcel();
        }

        public void ExportToExcel()
        {


            UpdateGlobalScope();
            var c = Calculator.Calculate();

            if (!c.success)
            {
                ResultCalc.Text = c.errorMessage;
                return;
            }

            ResultCalc.Text = c.message;

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
}
