using System;
using NetOffice.ExcelApi;

namespace aula2610
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel excel = new Excel(Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            excel.CriarExcel();
            excel.dispose();

        }

    }
    class Excel
    {
        static string caminho;
        public Excel(string caminhorecebido)
        {
            caminho = caminhorecebido;
        }
    
        public void CriarExcel()
        {
            Application ex = new Application();
            ex.Workbooks.Add();
            ex.Cells[1, 1].Value = "Ford";
            ex.ActiveWorkbook.SaveAs(caminho + "\\arquivo.xlsx");
            ex.Quit();
        }

        public void dispose()
        {
            this.dispose();
        }
    }
}
