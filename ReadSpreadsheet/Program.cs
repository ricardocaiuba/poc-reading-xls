using System;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace ReadSpreadsheet
{
    class Program
    {
        static void Main(string[] args)
        {
            var xls = new XLWorkbook(@"D:\Plano Contábil.xlsx");
            var planilha = xls.Worksheets.First(w => w.Name == "Planilha1");
            var totalLinhas = planilha.Rows().Count();

            // primeira linha é o cabecalho
            for (int l = 2; l <= totalLinhas; l++)
            {
                var conta = planilha.Cell($"A{l}").Value.ToString();
                var descricao = planilha.Cell($"B{l}").Value.ToString();
                var reduzido = planilha.Cell($"C{l}").Value.ToString();
                var grau = planilha.Cell($"D{l}").Value.ToString();
                Console.WriteLine($"{conta} | {descricao} | {reduzido} | {grau}");
            }
        }
    }
}
