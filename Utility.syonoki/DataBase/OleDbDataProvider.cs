using System.Data;
using System.Data.Common;
using System.Data.OleDb;

namespace Utility.syonoki.DataBase {
    public abstract class OleDbDataProvider {
        private OleDbConnection conn_;
        public abstract string connectionString { get; }

        public OleDbConnection connection{
            get{
                if (conn_ != null)
                    if (conn_.State == ConnectionState.Open)
                        conn_.Close();

                conn_ = new OleDbConnection(connectionString);
                return conn_;
            }
        }

        public DataSet getDatas(string query, bool withConnectionClose = true){
            if (conn_ == null)
                conn_ = new OleDbConnection(connectionString);

            if (conn_.State != ConnectionState.Open)
                conn_.Open();

            OleDbCommand comm = new OleDbCommand(query, conn_);
            DataAdapter adt = new OleDbDataAdapter(comm);
            DataSet ds = new DataSet();
            adt.Fill(ds);

            if (withConnectionClose)
                conn_.Close();

            return ds;
        }
    }
}