namespace ExcelWorkbookPivotTable.Models.ResponseModel
{
    public class UserResponse
    {
        public bool IsCreated { get; set; }
        public string Message { get; set; }
        public string FileName { get; set; }
        public string FilePath { get; set; }
        public byte[] FileByte { get; set; }

    }
}
