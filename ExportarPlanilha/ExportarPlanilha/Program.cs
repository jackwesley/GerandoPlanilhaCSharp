using System;
using ClosedXML.Excel;
namespace ExportarPlanilha
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Gerando Arquivo");

            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Planilha 1");

            //Titulo do relatorio
            ws.Cell("B2").Value = "Exemplo de Relatorio";
            var range = ws.Range("B2:I2");
            range.Merge().Style.Font.SetBold().Font.FontSize = 20;

            //Cabeçalho do relatorio
            ws.Cell("B3").Value = "Titulo 1";
            ws.Cell("C3").Value = "Titulo 2";
            ws.Cell("D3").Value = "Titulo 3";
            ws.Cell("E3").Value = "Titulo 4";
            ws.Cell("F3").Value = "Titulo 5";
            ws.Cell("G3").Value = "Titulo 6";
            ws.Cell("H3").Value = "Titulo 7";
            ws.Cell("I3").Value = "Subtotal";


            //Corpo do Relatorio
            var linha = 4;

            for(int i = 0; i<20; i++)
            {
                ws.Cell("B" + linha.ToString()).Value = "B" + i.ToString();
                ws.Cell("C" + linha.ToString()).Value = "C" + i.ToString();
                ws.Cell("D" + linha.ToString()).Value = "D" + i.ToString();
                ws.Cell("E" + linha.ToString()).Value = "E" + i.ToString();
                ws.Cell("F" + linha.ToString()).Value = "F" + i.ToString();
                ws.Cell("G" + linha.ToString()).Value = "G" + i.ToString();
                ws.Cell("H" + linha.ToString()).Value = "H" + i.ToString();
                ws.Cell("I" + linha.ToString()).Value = string.Format("{0:F2}", i = linha);
                linha++;
            }
            //ajusta numeração da linha
            linha--;

            //Crio formatação do Tipo "Money" para o nosso subtotal
            ws.Range("I4:I" + linha.ToString()).Style.NumberFormat.Format = "R$ #,#.##00";

            //criação da tabela para ativar os filtros
            range = ws.Range("B3:I" + linha.ToString());
            range.CreateTable();

            //Ajusto no tamanho da coluna com o conteudo da mesma
            ws.Columns("2-9").AdjustToContents();

            //Salvar o arquivo em disco
            wb.SaveAs("C:/Users/Jack/Desktop/teste_table.xlsx");

            //liberar objetos
            wb.Dispose();

            Console.WriteLine("Finalizado");
            Console.ReadKey();
        }
    }
}
