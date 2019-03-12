using System;

namespace Ortogo.SolidWorks.ShelvingTask
{
    public struct Calculator
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
        public string message;
        public string errorMessage;

        public static Calculator Calculate()
        {
            var c = default(Calculator);

            c.frame = new Frame();

            
            c.traversa = new TraversaCut().Select();
            if (c.traversa == null)
            {
                c.errorMessage = "Не удалось подобрать траверсу, примение другие параметры или расширьте библиотеку";
                c.success = false;
            }
            else
            {
                c.message += $"Подобранная траверса: {c.traversa}{System.Environment.NewLine}";
                c.message = $"Высота стелажа, м:{GlobalScope.SH}, Ширина стеллажа, м: {GlobalScope.SW}{System.Environment.NewLine}";
                try
                {
                    c.supportType = new SupportElement().Select();
                    if (c.supportType == null)
                    {
                        c.errorMessage = "Не удалось подобрать стойку, примение другие параметры или расширьте библиотеку";
                        c.success = false;
                        return c;
                    }
                    c.hConn = c.frame.GetConnection("СГ", c.supportType);
                    c.dConn = c.frame.GetConnection("СД", c.supportType);

                    c.message += $"Подобранная стойка: {c.supportType}{System.Environment.NewLine}" +
                        $"Горизонтальная связь: {c.hConn}{System.Environment.NewLine}" +
                        $"Диагональная связь {c.dConn}{System.Environment.NewLine}";

                    c.countTraversa = 2 * GlobalScope.NL;
                    c.sumLenTraversa = c.countTraversa * c.traversa.Length;
                    c.countConn = c.hConn.Count + c.dConn.Count;
                    c.sumLenConn = c.hConn.Count * c.hConn.LK + c.dConn.Count * c.dConn.LK + (c.countConn - 1) * GlobalScope.TUSHM;
                    c.message += $"Суммарная длина профиля траверс, м:{c.sumLenTraversa}{System.Environment.NewLine}" +
                        $"С учетом припуска на распил(t={GlobalScope.TUSHM}), м: {Math.Round(c.sumLenTraversa + (GlobalScope.TUSHM * c.sumLenTraversa), 3)}{System.Environment.NewLine}" +
                        $"Количество резов {c.countTraversa - 1}{System.Environment.NewLine}";
                    c.message += $"Суммарная длина профиля связей с учетом припуска на распил, м: {c.sumLenConn}{System.Environment.NewLine}" +
                        $"Количество резов {c.countConn - 1}{System.Environment.NewLine}";
                }
                catch (Exception ex)
                {
                    c.errorMessage += "Ошибка: " + ex.Message;
                    c.success = false;
                }
            }
            c.success = true;
            return c;
        }
    }
}
