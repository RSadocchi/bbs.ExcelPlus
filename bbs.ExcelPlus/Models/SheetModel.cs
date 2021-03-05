using System.Collections.Generic;

namespace bbs.ExcelPlus
{
    public class SheetModel
    {
        public int SheetId { get; set; }
        public string SheetName { get; set; }
        public List<RowModel> Rows { get; set; } = new List<RowModel>();
    }
}
