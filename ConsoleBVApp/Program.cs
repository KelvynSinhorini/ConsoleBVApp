using ConsoleBVApp.Models;
using IronXL;
using IronXL.Styles;
using Newtonsoft.Json;
using System.Text.RegularExpressions;

namespace ConsoleBVApp
{
    public class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Converter txt para excel\n");

            FileInfo file = GetFile();
            string json = File.ReadAllText($@"C:\Users\kelvy\OneDrive\Área de Trabalho\BV\{file.Name}");

            var fatura = JsonConvert.DeserializeObject<Fatura>(json);

            // ref: https://ironsoftware.com/csharp/excel/docs/
            var workbook = WorkBook.Create(ExcelFileFormat.XLSX);

            var sheet = workbook.CreateWorkSheet("Transações");
            SetHeader(sheet);

            var transactionsOfLastMonth = ReadLastMonthExcel().ToList();

            var movimentacoesNacionais = fatura.detalhesFaturaCartoes.SelectMany(d => d.movimentacoesNacionais).ToList();

            var pagamentoEfetuado = movimentacoesNacionais.FirstOrDefault(m => m.nomeEstabelecimento.Trim() == "PAGAMENTO EFETUADO");
            movimentacoesNacionais.Remove(pagamentoEfetuado);

            var index = 2;
            foreach (var movimentacao in movimentacoesNacionais.OrderBy(m => m.valorAbsolutoMovimentacaoReal))
            {
                var transaction = new Transaction(movimentacao.dataMovimentacao, movimentacao.nomeEstabelecimento, movimentacao.valorAbsolutoMovimentacaoReal, movimentacao.numeroParcelaMovimentacao, movimentacao.quantidadeTotalParcelas, movimentacao.sinal);

                SetTransactionPerson(transactionsOfLastMonth, transaction);
                SetCellsValue(sheet, index, transaction);

                index++;
            }

            var borderType = IronXL.Styles.BorderType.Thin;
            var borderColor = "#000000";

            SetTotalValueAndStyle(sheet, index, borderType, borderColor);
            
            SetStyleInAllCells(sheet, index, borderType, borderColor);

            var dateCollumn = sheet[$"A1:A{index + 1}"];
            dateCollumn.Style.HorizontalAlignment = IronXL.Styles.HorizontalAlignment.Left;

            var fileName = $"fatura_{file.Name.Replace(".txt", "")}_{Guid.NewGuid()}";
            workbook.SaveAs($@"C:\Users\kelvy\OneDrive\Área de Trabalho\BV\Faturas\{fileName}.xlsx");

            Console.WriteLine("Pressione qualquer tecla para fechar!");
            Console.ReadLine();
        }

        private static void SetTransactionPerson(List<Transaction> transactionsOfLastMonth, Transaction transaction)
        {
            var lastTransaction = transactionsOfLastMonth.FirstOrDefault(t => t.Equals(transaction));
            transaction.Person = lastTransaction?.Person;
        }

        private static void SetCellsValue(WorkSheet sheet, int index, Transaction transaction)
        {
            // Data Item
            sheet[$"A{index}"].DateTimeValue = transaction.Date;
            sheet[$"A{index}"].FormatString = "dd/MM/yyyy";

            // Descrição item
            sheet[$"B{index}"].Value = transaction.NameTotalAmountAndNumberParcel;

            // Valor
            var sheetValue = sheet[$"C{index}"];
            sheetValue.Value = transaction.Value;
            sheetValue.FormatString = IronXL.Formatting.BuiltinFormats.Currency2;

            // Pessoa
            sheet[$"E{index}"].Value = transaction.Person ?? "";
        }

        private static void SetStyleInAllCells(WorkSheet sheet, int index, BorderType borderType, string borderColor)
        {
            sheet.AutoSizeColumn(0);
            sheet.AutoSizeColumn(1);
            sheet.AutoSizeColumn(2);

            var allCells = sheet[$"A1:C{index - 1}"];

            allCells.Style.BottomBorder.Type = borderType;
            allCells.Style.BottomBorder.SetColor(borderColor);

            allCells.Style.LeftBorder.Type = borderType;
            allCells.Style.LeftBorder.SetColor(borderColor);

            allCells.Style.TopBorder.Type = borderType;
            allCells.Style.TopBorder.SetColor(borderColor);

            allCells.Style.RightBorder.Type = borderType;
            allCells.Style.RightBorder.SetColor(borderColor);
        }

        private static void SetTotalValueAndStyle(WorkSheet sheet, int index, BorderType borderType, string borderColor)
        {
            var totalValue = sheet[$"C{index}"];
            totalValue.Formula = $"=SUM(C2:C{index - 1})";
            totalValue.FormatString = IronXL.Formatting.BuiltinFormats.Currency2;

            totalValue.Style.BottomBorder.Type = borderType;
            totalValue.Style.BottomBorder.SetColor(borderColor);

            totalValue.Style.LeftBorder.Type = borderType;
            totalValue.Style.LeftBorder.SetColor(borderColor);

            totalValue.Style.TopBorder.Type = borderType;
            totalValue.Style.TopBorder.SetColor(borderColor);

            totalValue.Style.RightBorder.Type = borderType;
            totalValue.Style.RightBorder.SetColor(borderColor);
        }

        private static void SetHeader(WorkSheet sheet)
        {
            sheet["A1"].Value = "Data";
            sheet["A1"].Style.HorizontalAlignment = IronXL.Styles.HorizontalAlignment.Center;

            sheet["B1"].Value = "Nome";

            sheet["C1"].Value = "Valor";
            sheet["C1"].Style.HorizontalAlignment = IronXL.Styles.HorizontalAlignment.Center;

            sheet["E1"].Value = "Pessoa";
        }

        private static FileInfo GetFile()
        {
            var directoryInfo = new DirectoryInfo(@"C:\Users\kelvy\OneDrive\Área de Trabalho\BV");

            var files = directoryInfo.GetFiles("*.*");

            for (int i = 0; i < files.Length; i++)
            {
                Console.WriteLine($"[{i}] {files[i].Name}");
            }

            Console.WriteLine("\nDigite qual arquivo deseja exportar.\nEx: 1");

            var fileNumberText = Console.ReadLine();
            var validateNumberBetweenZeroAndThousand = "([0-9]|[1-9][0-9]|[1-9][0-9][0-9])";

            if (string.IsNullOrWhiteSpace(fileNumberText) || !Regex.IsMatch(validateNumberBetweenZeroAndThousand, fileNumberText))
                Console.WriteLine("\nDigite um caracter válido de 0 a 1000!");

            var fileNumber = int.Parse(fileNumberText);
            var file = files[fileNumber];
            return file;
        }

        private static IEnumerable<Transaction> ReadLastMonthExcel()
        {
            var workbook = WorkBook.Load(@$"C:\Users\kelvy\OneDrive\Área de Trabalho\BV\Faturas\fatura_outubro.xlsx");
            var sheet = workbook.WorkSheets.First(x => x.Name == "Transações");

            for(int i = 1; i < sheet.Rows.Length; i++)
            {
                var row = sheet.Rows[i];

                var date = row.Columns[0].DateTimeValue;
                var name = row.Columns[1].StringValue;
                var value = row.Columns[2].DecimalValue;
                var person = row.Columns[3].StringValue;

                yield return new Transaction(date, name, value, person);
            }
        }
    }
}