using ConsoleBVApp.Models;
using IronXL;
using Newtonsoft.Json;

namespace ConsoleBVApp
{
    public class Program
    {
        static void Main(string[] args)
        {
            string json = File.ReadAllText(@"C:\Users\kelvy\OneDrive\Área de Trabalho\BV\outubro 22.txt");
            var fatura = JsonConvert.DeserializeObject<Fatura>(json);

            // ref: https://ironsoftware.com/csharp/excel/docs/
            var workbook = WorkBook.Create(ExcelFileFormat.XLSX);

            var sheet = workbook.CreateWorkSheet("Transações");

            sheet["A1"].Value = "Data";
            sheet["B1"].Value = "Nome";
            sheet["C1"].Value = "Valor";

            var index = 2;

            foreach (var item in fatura.detalhesFaturaCartoes)
            {
                var pagamentoEfetuado = item.movimentacoesNacionais.FirstOrDefault(m => m.nomeEstabelecimento.Trim() == "PAGAMENTO EFETUADO");
                item.movimentacoesNacionais.Remove(pagamentoEfetuado);

                foreach (var movNacional in item.movimentacoesNacionais)
                {
                    var date = DateTime.Parse(movNacional.dataMovimentacao);
                    sheet[$"A{index}"].Value = date.Date;
                    sheet[$"B{index}"].Value = movNacional.nomeEstabelecimento;

                    var sheetValue = sheet[$"C{index}"];

                    var sinal = movNacional.sinal == "-" ? "+" : "-";

                    var value = $"{sinal}{movNacional.valorAbsolutoMovimentacaoReal}";
                    sheetValue.Value = double.Parse(value);

                    index++;
                }
            }

            sheet.AutoSizeColumn(0);
            sheet.AutoSizeColumn(1);
            sheet.AutoSizeColumn(2);

            workbook.SaveAs($@"C:\Users\kelvy\OneDrive\Área de Trabalho\BV\Faturas\fatura_{Guid.NewGuid()}.xlsx");
        }
    }
}