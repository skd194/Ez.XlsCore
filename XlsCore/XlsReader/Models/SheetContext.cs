namespace Ez.XlsCore
{
    public class SheetContext
    {
        public SheetContext(string id, int number, string name)
        {
            Number = number;
            Name = name;
            Id = id;
        }
        public int Number { get; }
        public string Name { get;  }
        internal string Id { get;  }
    }
}