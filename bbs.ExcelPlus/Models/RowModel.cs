using System.Collections.Generic;

namespace bbs.ExcelPlus
{
    public class RowModel
    {
        public int RowId { get; set; }
        public List<CellModel> Cells { get; set; } = new List<CellModel>();
    }
}
