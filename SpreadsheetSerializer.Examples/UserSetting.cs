namespace SpreadsheetSerializer.Examples
{
    public class UserSetting
    {
        [Order]
        public long Id { get; set; }

        [Order]
        public long UserId { get; set; }

        [Order]
        public string SettingName { get; set; }

        [Order]
        public string Value { get; set; }
    }
}
