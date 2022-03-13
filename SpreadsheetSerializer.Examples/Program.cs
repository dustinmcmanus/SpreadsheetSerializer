using SpreadsheetSerializer.AsposeCells;
using System;
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

            //var userSetting1 = new UserSetting { Id = 1, UserId = 1, SettingName = "Theme", Value = "Dark" };
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

            //using (var workbookCreator = new WorkbookCreator("MyWorkbook3"))
            //{



            //var workbook = new WorkbookSerializer().WithFilePath("MyWorkbook4");

            //var sheet1 = new WorksheetSerializer<User>(users).WithWorksheetName("myUsers1");


            //var workbook = new ExcelWorkbookSerializer().WithFilePath("MyWorkbook4");

            //var sheet1 = new ExcelWorksheetSerializer<User>(users).WithWorksheetName("myUsers1");
            ////sheet1.workbook = workbookCreator;

            //var sheet2 = new ExcelWorksheetSerializer<UserSetting>(userSettings).WithWorksheetName("myUsersSettings2");
            ////sheet2.workbook = workbookCreator;

            //workbook.WorkSheetSerializers.Add(sheet1); // WithTabName?
            //workbook.WorkSheetSerializers.Add(sheet2); // WithTabName?
            //workbook.Serialize();


            //var sheet3 = new WorksheetSerializer();

            //var wb3 = new WorkbookSerializer<List<User>>();



            ////}

            //ISpreadsheetSerializer<User> spreadsheetCreator = new AsposeCells.WorkbookSerializer<User>();
            ////ISpreadsheetSerializer<User> spreadsheetCreator = new AsposeCells.ExcelWorksheetSerializer<User>().WithFilePath("MyWorkbook");

            //spreadsheetCreator.Serialize(users, "users");











            ////////////////////////////////////////////////

            var myWorkbookExample = new ExampleWorkbook();
            myWorkbookExample.Users = users;
            myWorkbookExample.UserSettings = userSettings;

            // TODO: allow name attributes
            // TODO: restore single sheet API
            // TODO: allow a different JsonConveter for each sheet in a workbook
            // TODO: test decimals and other types
            //var myWorkbookSerializer = new WorkbookSerializer<ExampleWorkbook>().WithFilePath("myNewWorkbookName4");

            //myWorkbookSerializer.Serialize(myWorkbookExample);


            //List<User> myUsers;
            //var worksheetDeserializer = new WorksheetDeserializer("Users");


            //// new Aspose.Cells.Workbook(stream)
            //using (var workbook = WorkbookRetriever.GetWorkbookFromFilePath("myNewWorkbookName4.xlsx"))
            //{
            //    myUsers = (List<User>)worksheetDeserializer.DeserializeWithNullStrings(typeof(User), workbook);
            //}

            //var myWorkbookClass = new ExampleWorkbook();
            //var wbDeserializer = new WorkbookDeserializer<ExampleWorkbook>().WithJsonConverter(new JsonConverterWithNullStrings());


            var wbDeserializer = new WorkbookDeserializer<ExampleWorkbook>();
            //var myWorkbookExample = wbDeserializer.Deserialize("MyNewWorkbook.xlsx");

            JsonDirectorySerializer.JsonDirectorySerializer.Serialize(users, "../../../../SpreadsheetSerializer.Tests.Unit/DataSource1");
            //SpreadsheetSerializer.Examples.TextFileSerializer.SerializeToTextFiles(myWorkbookExample, "../../../../SpreadsheetSerializer.Tests.Unit/DataSource1");
            Console.WriteLine("Hello World!");
        }
    }
}
