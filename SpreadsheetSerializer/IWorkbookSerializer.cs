namespace SpreadsheetSerializer
{
    public interface IWorkbookSerializer<T>
    {
        void Serialize(T workbookClass);

        T Deserialize(string workbookName = "");
    }
}
