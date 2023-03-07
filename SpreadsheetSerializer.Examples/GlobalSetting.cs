namespace SpreadsheetSerializer.Examples
{
    public class GlobalSetting
    {
        [Order] public long Id { get; set; }
        [Order] public string SettingName { get; set; }
        [Order] public string Value { get; set; }
    }
}
