namespace ConsoleBVApp.Models
{
    public class Transaction
    {
        public Transaction(DateTime? date, string name, decimal value, string person)
        {
            Date = date;
            Name = ClearName(name);
            Value = value;
            Person = person;
        }

        public Transaction(DateTime? date, string name, decimal value, int? totalAmountParcel, int? numberParcelMovement, string? sinal)
        {
            Date = date;
            Name = ClearName(name);
            TotalAmountParcel = totalAmountParcel;
            NumberParcelMovement = numberParcelMovement;
            Sinal = sinal == "-" ? "+" : "-";
            Value = decimal.Parse($"{Sinal}{value}");
        }

        public DateTime? Date { get; set; }
        public string Name { get; set; }
        public decimal Value { get; set; }
        public string? Person { get; set; }
        public int? TotalAmountParcel { get; set; }
        public int? NumberParcelMovement { get; set; }
        public string NameTotalAmountAndNumberParcel { get => $"{Name} {NumberParcelMovement}/{TotalAmountParcel}"; }
        public string Sinal { get; set; }

        public override bool Equals(object? obj)
        {
            var transaction = obj as Transaction;

            if (transaction == null) return false;

            return Name == transaction.Name && ValueIsApproximate(transaction.Value);
        }

        private bool ValueIsApproximate(decimal value)
        {
            return value >= (Value - 1) && value <= (Value + 1);
        }

        private string ClearName(string name)
        {
            name = name.Trim();

            if (name.Contains("/"))
            {
                var lastEmptyCharIndex = name.LastIndexOf(" ");
                return name.Substring(0, lastEmptyCharIndex);
            }
            else
            {
                return name;
            }
        }
    }
}
