using System.IO;

namespace bbs.ExcelPlus
{
    public class ExcelFileInfo
    {
        public string FileName { get; set; }
        public string CompletePath { get; set; }
        public string ContentType { get; set; } = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        public byte[] Bytes { get; set; }
        public FileInfo FileInfo { get; set; }
    }
}
