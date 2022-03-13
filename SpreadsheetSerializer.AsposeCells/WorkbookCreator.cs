namespace SpreadsheetSerializer.AsposeCells
{
    public class WorkbookCreator
    {
        public static AsposeWorkbook CreateWorkbookWithFilePath(string filePath)
        {
            return new AsposeWorkbook().WithFilePath(filePath);
        }
    }
}
