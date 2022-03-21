namespace SpreadsheetSerializer.Examples
{
    public class User
    {
        [Order] public long Id { get; set; }
        [Order] public string Name { get; set; }
        [Order] public string EmailAddress { get; set; }
    }
}
