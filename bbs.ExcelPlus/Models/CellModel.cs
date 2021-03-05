using System;

namespace bbs.ExcelPlus
{
    public class CellModel
    {
        public int RowId { get; set; }
        public string ColumnName { get; set; }
        public int CellId { get; set; }
        public string CellValue { get; set; }
        public decimal? CellValueNumber { get; set; } = null;
        public DateTime? CellValueDateTime { get; set; } = null;
        public bool? CellValueBoolean { get; set; }
    }
}
