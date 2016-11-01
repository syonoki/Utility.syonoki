namespace Utility.syonoki.DataBase.FileAsDataBase {
    public class AccessDb:FileAsDataBaseDefinition {
        public override string connectionString(string dataSource) 
            => $@"Provider = Microsoft.ACE.OLEDB.12.0; data source={dataSource}";
    }
}