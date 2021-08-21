namespace SpreadsheetSerializer.AsposeCells
{
    public static class License
    {
        public static Aspose.Cells.License AsposeLicense { get; set; } = new Aspose.Cells.License();

        public static void SetLicense(string filePath)
        {
            AsposeLicense.SetLicense(filePath); // "Aspose.Total.lic"
        }
    }
}
