using System.Collections.Generic;

namespace SpreadsheetSerializer.Examples
{
    class Program
    {
        static void Main(string[] args)
        {
            var users = new List<User>();

            var user1 = new User();
            user1.Id = 1;
            user1.Name = "Bob";
            user1.EmailAddress = "bob@example.com";
            users.Add(user1);

            var user2 = new User();
            user2.Id = 2;
            user2.Name = "Jill";
            user2.EmailAddress = "jill@example.com";
            users.Add(user2);


            var userSettings = new List<UserSetting>();

            var userSetting1 = new UserSetting();
            userSetting1.Id = 1;
            userSetting1.UserId = 1;
            userSetting1.SettingName = "Theme";
            userSetting1.Value = "Dark";
            userSettings.Add(userSetting1);

            var userSetting2 = new UserSetting();
            userSetting2.Id = 2;
            userSetting2.UserId = 2;
            userSetting2.SettingName = "Theme";
            userSetting2.Value = "Default";
            userSettings.Add(userSetting2);

            var userSetting3 = new UserSetting();
            userSetting3.Id = 3;
            userSetting3.UserId = 1;
            userSetting3.SettingName = "Zoom";
            userSetting3.Value = "150%";
            userSettings.Add(userSetting3);

            SpreadsheetSerializer.AsposeCells.License.SetLicense("Aspose.Total.lic");

            // use cases
            // from JSON files to class with List properties
            // from JSON file to list of class

            // From Excel Workbook with multiple tabs to Class with properties having Lists
            // from Excel Workbook with one tab to List<class>

            // from class with list of properties to JSON files
            // from List<class> to json file

            // from class with list of properties to Excel workbook with multiple tabs for each property
            // from List<class> to Excel workbook with workbook and tab named after class

            // indirect:
            // from JSON file to Excel
            // from JSON files to Excel
            // from Excel to JSON files
            // from Excel to Json File

            ////////////////////////////////////////////////

            var myWorkbookExample = new ExampleWorkbook();
            myWorkbookExample.Users = users;
            myWorkbookExample.UserSettings = userSettings;

            // TODO: allow name attributes
            // TODO: allow a different JsonConveter for each sheet in a workbook
            // TODO: test decimals and other types

            // See SpreadsheetSerializer.Tests.Unit project for examples and instructions, especially SerializerTestsDataSource1.cs
        }
    }
}
