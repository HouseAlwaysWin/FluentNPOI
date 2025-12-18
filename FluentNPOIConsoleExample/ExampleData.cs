using System;

namespace FluentNPOIConsoleExample
{
    public class ExampleData
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public DateTime DateOfBirth { get; set; }
        public bool IsActive { get; set; }
        public double Score { get; set; }
        public decimal Amount { get; set; }
        public string Notes { get; set; }
        public object MaybeNull { get; set; }

        public ExampleData() { }

        public ExampleData(int id, string name, DateTime dateOfBirth, bool? isActive = null, double? score = null, decimal? amount = null, string notes = null, object maybeNull = null)
        {
            ID = id;
            Name = name;
            DateOfBirth = dateOfBirth;
            IsActive = isActive ?? (id % 2 == 0);
            Score = score ?? (id * 12.5d);
            Amount = amount ?? (id * 1000.75m);
            Notes = notes ?? ((name?.Length ?? 0) > 10 ? "LongName" : "Short");
            MaybeNull = maybeNull ?? (id % 3 == 0 ? DBNull.Value : (object)"OK");
        }
    }
}
