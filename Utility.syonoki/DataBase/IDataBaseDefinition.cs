using System.Data.OleDb;

namespace Utility.syonoki.DataBase {
    public interface IDataBaseDefinition {
        OleDbConnection connection(string dataSource);
        string connectionString(string dataSource);
    }
}