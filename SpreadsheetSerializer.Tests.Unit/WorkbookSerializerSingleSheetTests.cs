using FluentAssertions;
using JsonDirectorySerializer;
using NUnit.Framework;
using SpreadsheetSerializer.AsposeCells;
using SpreadsheetSerializer.Examples;
using System.Collections.Generic;
using System.IO;

namespace SpreadsheetSerializer.Tests.Unit
{
    [TestFixture()]
    public class WorkbookSerializerSingleSheetTests
    {
        private const string DirectoryPath = @"..\..\..\DataSourceSingleSheet\";

        private static List<User> GetUsers()
        {
            var users = new List<User>();

            var user1 = new User();
            user1.Id = 1;
            user1.Name = "Amy";
            user1.EmailAddress = "amy@example.com";
            users.Add(user1);

            var user2 = new User();
            user2.Id = 2;
            user2.Name = "Joe";
            user2.EmailAddress = "joe@example.com";
            users.Add(user2);
            return users;
        }

        [SetUp]
        public void Setup()
        {
            // The Aspose license file should be dropped into the build folder (Ex: bin/Debug/net5.0/). It should never be checked into git.
            SpreadsheetSerializer.AsposeCells.License.SetLicense("Aspose.Total.lic");
        }

        [Test]
        public void CreateExcelFileFromList()
        {
            var users = GetUsers();

            new WorkbookSerializerSingleSheet<User>().Serialize(users, DirectoryPath);
        }

        [Test]
        public void CreateJsonFileFromList()
        {
            var users = GetUsers();

            new JsonFileFromListSerializer<User>().Serialize(users, DirectoryPath);
        }

        [Test]
        public void DeserializeFromJsonFileMatchesList()
        {
            var users = GetUsers();

            string path = Path.Combine(DirectoryPath, nameof(User));
            new WorkbookSerializerSingleSheet<User>().Serialize(users, path);
            List<User> users2 = new WorkbookDeserializerSingleSheet<User>().Deserialize(path);

            users.Should().BeEquivalentTo(users2);
        }

        [Test]
        public void DeserializeEmptyList()
        {
            // ReSharper disable once CollectionNeverUpdated.Local
            var users = new List<User>();
            string path = Path.Combine(DirectoryPath, nameof(User));
            new WorkbookSerializerSingleSheet<User>().Serialize(users, path);
            List<User> users2 = new WorkbookDeserializerSingleSheet<User>().Deserialize(path);
            users.Should().BeEquivalentTo(users2);
        }

        [Test]
        public void DeserializeUsersListWithoutSpecifyingFileName()
        {
            var users = GetUsers();
            new WorkbookSerializerSingleSheet<User>().Serialize(users, DirectoryPath);
            List<User> users2 = new WorkbookDeserializerSingleSheet<User>().Deserialize(DirectoryPath);
            users.Should().BeEquivalentTo(users2);
        }

        [Test]
        public void DeserializeUsersListWithoutSpecifyingFileNameInCurrentWorkingDirectory()
        {
            var users = GetUsers();
            new WorkbookSerializerSingleSheet<User>().Serialize(users);
            List<User> users2 = new WorkbookDeserializerSingleSheet<User>().Deserialize();
            users.Should().BeEquivalentTo(users2);
        }


        [Test]
        public void DeserializeEmptyListWithNamedFileAndSheet()
        {
            // ReSharper disable once CollectionNeverUpdated.Local
            var users = new List<User>();
            string path = Path.Combine(DirectoryPath, "Empty.xlsx");
            new WorkbookSerializerSingleSheet<User>().Serialize(users, path, "Empty");
            List<User> users2 = new WorkbookDeserializerSingleSheet<User>().Deserialize(path, "Empty");
            users.Should().BeEquivalentTo(users2);
        }

        [Test]
        public void DeserializeEmptyListWithNamedFileToCurrentWorkingDirectory()
        {
            // ReSharper disable once CollectionNeverUpdated.Local
            var users = new List<User>();
            string path = "Empty2.xlsx";
            new WorkbookSerializerSingleSheet<User>().Serialize(users, path);
            List<User> users2 = new WorkbookDeserializerSingleSheet<User>().Deserialize(path);
            users.Should().BeEquivalentTo(users2);
        }

        [Test]
        public void DeserializeEmptyListWorksheetNameOnly()
        {
            // ReSharper disable once CollectionNeverUpdated.Local
            var users = new List<User>();
            new WorkbookSerializerSingleSheet<User>().Serialize(users, "", "something");
            List<User> users2 = new WorkbookDeserializerSingleSheet<User>().Deserialize("", "something");
            users.Should().BeEquivalentTo(users2);
        }
    }
}