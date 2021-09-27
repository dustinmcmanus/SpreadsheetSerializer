namespace SpreadsheetSerializer.AsposeCells
{
    public class WorkbookCreator
    {
        public static AsposeWorkbook CreateWorkbookWithName(string workbookName)
        {
            return new AsposeWorkbook().WithWorkbookName(workbookName);
        }
    }
}
