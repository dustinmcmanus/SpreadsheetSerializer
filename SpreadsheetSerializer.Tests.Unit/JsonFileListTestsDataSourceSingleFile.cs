using FluentAssertions;
using JsonDirectorySerializer;
using NUnit.Framework;
using SpreadsheetSerializer.Examples;
using System.Collections.Generic;
using System.IO;

namespace SpreadsheetSerializer.Tests.Unit
{
    public class JsonFileListTestsDataSourceSingleFile
    {
        private const string DirectoryPath = @"..\..\..\DataSourceSingleFile\";

        private static List<User> GetUsers()
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
            return users;
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

            new JsonFileFromListSerializer<User>().Serialize(users, DirectoryPath);
            List<User> users2 = new JsonFileToListDeserializer<User>().Deserialize(DirectoryPath);

            users.Should().BeEquivalentTo(users2);
        }

        /// <summary>
        /// Outputs an empty JSON array in a User.json file in the current working directory of the running process and reads it in again
        /// </summary>
        [Test]
        public void DeserializeEmptyList()
        {
            // ReSharper disable once CollectionNeverUpdated.Local
            var users = new List<User>();
            new JsonFileFromListSerializer<User>().Serialize(users);
            List<User> users2 = new JsonFileToListDeserializer<User>().Deserialize();
            users.Should().BeEquivalentTo(users2);
        }

        /// <summary>
        /// Tests a fully qualified path.
        /// Outputs an empty JSON array in an Empty.json file in the DataSourceSingleFile directory and reads it in again
        /// </summary>
        [Test]
        public void DeserializeEmptyListWithNamedFile()
        {
            // ReSharper disable once CollectionNeverUpdated.Local
            var users = new List<User>();
            string path = Path.Combine(DirectoryPath, "Empty.json");
            new JsonFileFromListSerializer<User>().Serialize(users, path);
            List<User> users2 = new JsonFileToListDeserializer<User>().Deserialize(path);
            users.Should().BeEquivalentTo(users2);
        }

        /// <summary>
        /// Outputs an empty JSON array in an Empty2.json file in the current working directory of the running process and reads it in again
        /// </summary>
        [Test]
        public void DeserializeEmptyListWithNamedFileToCurrentWorkingDirectory()
        {
            // ReSharper disable once CollectionNeverUpdated.Local
            var users = new List<User>();
            string path = "Empty2.json";
            new JsonFileFromListSerializer<User>().Serialize(users, path);
            List<User> users2 = new JsonFileToListDeserializer<User>().Deserialize(path);
            users.Should().BeEquivalentTo(users2);
        }
    }
}