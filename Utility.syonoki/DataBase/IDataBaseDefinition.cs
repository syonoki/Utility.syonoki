using System.Data.OleDb;

namespace Utility.syonoki.DataBase {
    public interface IDataBaseDefinition {
        string connectionString(string dataSource);
    }
}