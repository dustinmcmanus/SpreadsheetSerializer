namespace SpreadsheetSerializer.AsposeCells
{
    internal class WorkbookCreator
    {
        protected WorkbookCreator()
        {
        }

        public static AsposeWorkbook CreateWorkbookWithFilePath(string filePath)
        {
            return new AsposeWorkbook().WithFilePath(filePath);
        }
    }
}
