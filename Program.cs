using System;
using NetOffice.ExcelApi;

namespace NuGet
{
    class Program
    {
        static void Main(string[] args)
        {
            LerExcel();
        }
        static void CriarExcel()
        {
            Application ex = new Application();
            ex.Workbooks.Add();
            ex.Cells[1, 1].Value = "Ford";
            ex.Cells[1, 2].Value = "Fiesta";
            ex.Cells[1, 3].Value = "1.8";

            ex.Cells[2, 1].Value = "Nintendo";
            ex.Cells[2, 2].Value = "Switch";
            ex.Cells[2, 3].Value = "2017";
            ex.ActiveWorkbook.SaveAs(@"C:\Users\43692939876\Desktop\Projetos\NuGet\teste.xls"); //sem o caminho, ele salva por default nos "MEUS DOCUMENTOS"
            ex.Quit();
            ex.Dispose();
        }
        static void LerExcel()
        {
            Application ex = new Application();
            ex.Workbooks.Open(@"C:\Users\43692939876\Desktop\Projetos\NuGet\teste.xls");
            string valor = ex.Cells[2, 1].Value.ToString();
            System.Console.WriteLine(valor);
            ex.Quit();
            ex.Dispose();
        }
    }
}
