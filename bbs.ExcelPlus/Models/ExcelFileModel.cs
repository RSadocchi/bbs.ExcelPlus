using System.Collections.Generic;

namespace bbs.ExcelPlus
{
    public class ExcelFileModel
    {
        public string Name { get; set; }
        public List<SheetModel> Sheets { get; set; } = new List<SheetModel>();
    }
}
