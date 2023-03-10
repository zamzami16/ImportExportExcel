using System;

namespace AxataPOS.ImportExport.Utils
{
    public class Formula
    {
        public Formula(string text) : this(row => text) { }

        public Formula(Func<uint, string> rowText)
        {
            RowText = rowText;
        }

        internal Func<uint, string> RowText { get; }

    }
}
