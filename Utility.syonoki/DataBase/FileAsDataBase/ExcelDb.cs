namespace Utility.syonoki.DataBase.FileAsDataBase {
    class ExcelDb : FileAsDataBaseDefinition {
        public override string connectionString(string dataSource) 
            => $"Provider = Microsoft.ACE.OLEDB.12.0; data source={dataSource}; Extended Properties=Excel 12.0 XML";
    }
}