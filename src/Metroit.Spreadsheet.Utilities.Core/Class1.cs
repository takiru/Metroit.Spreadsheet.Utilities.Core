using Metroit.Spreadsheet.Utilities.Core.Mapping;
using System;
using System.Collections.Generic;
using System.Text;

namespace Metroit.Spreadsheet.Utilities.Core
{
    class Class1
    {
        [CellBorder(BorderPosition.EdgeLeft, LineStyle.Dash, BorderWeightType.Hairline, "Red")]
        [CellBorder(BorderPosition.EdgeRight, LineStyle.Dash, BorderWeightType.Hairline, "Red")]
        [CellBorder(BorderPosition.EdgeLeft, LineStyle.Dash, BorderWeightType.Hairline, "Red")]
        [CellOutputMap(Column = 1)]
        public int MyProperty { get; set; }
    }
}
