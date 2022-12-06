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
            // TODO FIX: Quando tem uma transação em dupla que umas das pessoas nao tenha nenhuma outra transação, na segunda sheet ele vai dar um valor total diferente.

            Console.WriteLine("Converter txt para excel\n");

            FileInfo file = GetFile();
            string json = File.ReadAllText($@"C:\Users\kelvy\OneDrive\Área de Trabalho\BV\{file.Name}");

            var fatura = JsonConvert.DeserializeObject<Fatura>(json);

            // ref: https://ironsoftware.com/csharp/excel/docs/
            var workBook = WorkBook.Create(ExcelFileFormat.XLSX);

            var sheet = workBook.CreateWorkSheet("Transações");
            SetHeader(sheet);

            var transactionsOfLastMonth = ReadLastMonthExcel().ToList();

            var movimentacoesNacionais = fatura.detalhesFaturaCartoes.SelectMany(d => d.movimentacoesNacionais).ToList();

            var pagamentoEfetuado = movimentacoesNacionais.FirstOrDefault(m => m.nomeEstabelecimento.Trim() == "PAGAMENTO EFETUADO");
            movimentacoesNacionais.Remove(pagamentoEfetuado);

            var transactions = new List<Transaction>();

            var index = 2;
            foreach (var movimentacao in movimentacoesNacionais.OrderBy(m => m.valorAbsolutoMovimentacaoReal))
            {
                var transaction = new Transaction(movimentacao.dataMovimentacao, movimentacao.nomeEstabelecimento, movimentacao.valorAbsolutoMovimentacaoReal, movimentacao.numeroParcelaMovimentacao, movimentacao.quantidadeTotalParcelas, movimentacao.sinal);

                SetTransactionPerson(ref transaction, transactionsOfLastMonth);
                SetCellsValue(sheet, index, transaction);

                transactions.Add(transaction);

                index++;
            }

            CreateSheetTransactionByPerson(workBook, transactions);

            var borderType = IronXL.Styles.BorderType.Thin;
            var borderColor = "#000000";

            SetTotalValueAndStyle(sheet, index, borderType, borderColor);
            
            SetStyleInAllCells(sheet, index, borderType, borderColor);

            var dateCollumn = sheet[$"A1:A{index + 1}"];
            dateCollumn.Style.HorizontalAlignment = IronXL.Styles.HorizontalAlignment.Left;

            var fileName = $"fatura_{file.Name.Replace(".txt", "")}_{Guid.NewGuid()}";
            workBook.SaveAs($@"C:\Users\kelvy\OneDrive\Área de Trabalho\BV\Faturas\{fileName}.xlsx");

            Console.WriteLine("Pressione qualquer tecla para fechar!");
            Console.ReadLine();
        }

        private static void SetTransactionPerson(ref Transaction transaction, List<Transaction> transactionsOfLastMonth)
        {
            var transactionWithoutRef = transaction;
            var lastTransaction = transactionsOfLastMonth.FirstOrDefault(t => t.Equals(transactionWithoutRef));
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
            sheet[$"D{index}"].Value = transaction.Person ?? "";
        }

        private static void SetHeader(WorkSheet sheet)
        {
            sheet["A1"].Value = "Data";
            sheet["A1"].Style.HorizontalAlignment = IronXL.Styles.HorizontalAlignment.Center;

            sheet["B1"].Value = "Nome";

            sheet["C1"].Value = "Valor";
            sheet["C1"].Style.HorizontalAlignment = IronXL.Styles.HorizontalAlignment.Center;

            sheet["D1"].Value = "Pessoa";
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
            var workbook = WorkBook.Load(@$"C:\Users\kelvy\OneDrive\Área de Trabalho\BV\Faturas\fatura_dezembro 22_7226fa20-a9ea-4cef-b164-a348152c006b.xlsx");
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

        private static void CreateSheetTransactionByPerson(WorkBook workBook, List<Transaction> transactions)
        {
            var sheet = workBook.CreateWorkSheet("Transações Pessoas");

            var transactionsWithoutPerson = transactions.Where(t => string.IsNullOrWhiteSpace(t.Person)).ToList();
            var transactionsWithPerson = transactions.Where(t => !string.IsNullOrWhiteSpace(t.Person)).OrderBy(t => !t.Person.Contains(",")).ThenBy(t => t.Person).ToList();
            var values = new List<decimal>();

            var indexRow = 1;

            sheet[$"A{indexRow}"].StringValue = "";
            indexRow++;

            sheet[$"A{indexRow}"].StringValue = "Nome";
            sheet[$"B{indexRow}"].StringValue = "Valor";
            sheet[$"C{indexRow}"].StringValue = "Pessoas";
            indexRow++;

            for (int i = 0; i < transactionsWithoutPerson.Count; i++)
            {
                var transaction = transactionsWithoutPerson[i];
                SetCellsValueInSecondSheet(ref indexRow, sheet, i, transaction);

                indexRow++;
            }

            sheet[$"A{indexRow}"].StringValue = "Total";
            var totalValue = sheet[$"B{indexRow}"];
            totalValue.Formula = $"=SUM(B{indexRow - transactionsWithoutPerson.Count}:B{indexRow - 1})";
            totalValue.FormatString = IronXL.Formatting.BuiltinFormats.Currency2;
            values.Add(totalValue.DecimalValue);

            indexRow++;
            indexRow++;

            var persons = transactionsWithPerson.Where(t => !t.Person.Contains(",")).GroupBy(t => t.Person).Select(t => t.Key).ToList();

            foreach(var person in persons)
            {
                sheet[$"A{indexRow}"].StringValue = person;
                indexRow++;

                sheet[$"A{indexRow}"].StringValue = "Nome";
                sheet[$"B{indexRow}"].StringValue = "Valor";
                sheet[$"C{indexRow}"].StringValue = "Pessoas";
                indexRow++;

                var transactionWithSamePerson = transactionsWithPerson.Where(t => t.Person.Equals(person)).ToList();

                for (int i = 0; i < transactionWithSamePerson.Where(t => !t.Person.Contains(",")).Count(); i++)
                {
                    var transaction = transactionWithSamePerson[i];

                    SetCellsValueInSecondSheet(ref indexRow, sheet, i, transaction, false);

                    indexRow++;
                }

                var transactionsWithMoreOnePerson = transactionsWithPerson.Where(t => t.Person.Contains(person) && t.Person.Contains(",")).ToList();

                foreach(var transactionWithMoreOnePerson in transactionsWithMoreOnePerson)
                {
                    SetCellsValueInSecondSheet(ref indexRow, sheet, 1, transactionWithMoreOnePerson, true);

                    indexRow++;
                }

                sheet[$"A{indexRow}"].StringValue = "Total";

                var transactionsCount = transactionWithSamePerson.Count + transactionsWithMoreOnePerson.Count;
                totalValue = sheet[$"B{indexRow}"];
                totalValue.Formula = $"=SUM(B{indexRow - transactionsCount}:B{indexRow - 1})";
                totalValue.FormatString = IronXL.Formatting.BuiltinFormats.Currency2;
                values.Add(totalValue.DecimalValue);

                indexRow++;
                indexRow++;
            }

            indexRow++;

            sheet[$"A{indexRow}"].StringValue = "Total";
            sheet[$"B{indexRow}"].DecimalValue = values.Sum();
            sheet[$"B{indexRow}"].FormatString = IronXL.Formatting.BuiltinFormats.Currency2;

            sheet.AutoSizeColumn(0);
            sheet.AutoSizeColumn(1);
            sheet.AutoSizeColumn(2);
        }

        private static void SetCellsValueInSecondSheet(ref int indexRow, WorkSheet sheet, int index, Transaction transaction, bool divideValue = false)
        {
            sheet[$"A{indexRow}"].StringValue = transaction.NameTotalAmountAndNumberParcel;

            var sheetValue = sheet[$"B{indexRow}"];
            sheetValue.DecimalValue = divideValue ? (transaction.Value / transaction.Person.Split(",").Count()) :  transaction.Value;
            sheetValue.FormatString = IronXL.Formatting.BuiltinFormats.Currency2;

            sheet[$"C{indexRow}"].StringValue = transaction?.Person;
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
    }
}