using System.Data.OleDb;
using System.IO;

namespace Utility.syonoki.DataBase {
    public abstract class FileAsDataBaseDefinition : IDataBaseDefinition {
        
        public OleDbConnection connection(string dataSource)
            =>  new OleDbConnection(connectionString(filePath(dataSource)));
        public abstract string connectionString(string dataSource);

        private string filePath(string dataSource)
        {
            bool hasFullPath = dataSource.Contains("\\");
            string currentDirectory = Directory.GetCurrentDirectory() + "\\";

            return hasFullPath ? dataSource : currentDirectory + "\\" + dataSource;
        }
    }
}