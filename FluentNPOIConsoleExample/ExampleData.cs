using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FluentNPOIConsoleExample
{
    public class ExampleData
    {
        public readonly int ID;
        public readonly string Name;
        public readonly DateTime DateOfBirth;
        public readonly bool IsActive;
        public readonly double Score;
        public readonly decimal Amount;
        public readonly string Notes;
        public readonly object MaybeNull;

        public ExampleData(int id, string name, DateTime dateOfBirth, bool? isActive = null, double? score = null, decimal? amount = null, string notes = null, object maybeNull = null)
        {
            ID = id;
            Name = name;
            DateOfBirth = dateOfBirth;

            // 若有傳入則使用傳入值，否則使用預設運算邏輯
            IsActive = isActive ?? (id % 2 == 0);
            Score = score ?? (id * 12.5d);
            Amount = amount ?? (id * 1000.75m);
            Notes = notes ?? ((name?.Length ?? 0) > 10 ? "LongName" : "Short");
            MaybeNull = maybeNull ?? (id % 3 == 0 ? DBNull.Value : (object)"OK");
        }

    }
}
