using System.Collections.Generic;
using System.Windows;

namespace SpreadsheetSerializer.Examples
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }


        public void Run()
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

            SpreadsheetSerializer.AsposeCells.License.SetLicense("Aspose.Total.lic");
            ISpreadsheetSerializer<User> spreadsheetCreator = new AsposeCells.ExcelWorksheetSerializer<User>().WithWorkbookName("MyWorkbook");
            spreadsheetCreator.Serialize(users, "users");
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            Run();
        }
    }

}
