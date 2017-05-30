using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DB
{
    public class Conexao
    {
        public DataSet Conectar()
        {
            string connectionString = @"Data Source=BASONI\SQLEXPRESS;Initial Catalog=TESTE; Integrated Security=true";

            // Provide the query string with a parameter placeholder.
            string queryString = @"SELECT * FROM PESSOA";

            string connetionString = null;
            SqlConnection connection;
            SqlCommand command;
            SqlDataAdapter adapter = new SqlDataAdapter();
            DataSet ds = new DataSet();
            string sql = null;

            ds.Tables.Add(CreateDataTable());

            connetionString = connectionString;// "Data Source=ServerName;Initial Catalog=DatabaseName;User ID=UserName;Password=Password";
            sql = queryString;// "Your SQL Statement Here";

            connection = new SqlConnection(connetionString);

            connection.Open();
            command = new SqlCommand(sql, connection);
            adapter.SelectCommand = command;
            adapter.Fill(ds);
            adapter.Dispose();
            command.Dispose();
            connection.Close();

            return ds;

        }

        public List<Pessoa> ListarPessoa()
        {
            string connectionString = @"Data Source=.\SQLEXPRESS;Initial Catalog=TESTE; Integrated Security=true";

            // Provide the query string with a parameter placeholder.
            string queryString = @"SELECT * FROM PESSOA";

            string connetionString = null;
            SqlConnection connection;
            SqlCommand command;
            SqlDataAdapter adapter = new SqlDataAdapter();
            DataSet ds = new DataSet();
            string sql = null;

            connetionString = connectionString;// "Data Source=ServerName;Initial Catalog=DatabaseName;User ID=UserName;Password=Password";
            sql = queryString;// "Your SQL Statement Here";

            connection = new SqlConnection(connetionString);

            connection.Open();
            command = new SqlCommand(sql, connection);
            adapter.SelectCommand = command;
            adapter.Fill(ds);
            adapter.Dispose();
            command.Dispose();
            connection.Close();

            return ListarPessoa(ds);
        }

        private DataTable CreateDataTable()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Empresa");
            dt.Columns.Add("Mês");
            dt.Columns.Add("Ano");


            DataRow dr = dt.NewRow();
            dr["Empresa"] = "ECORODOVIAS";
            dr["Mês"] = "Janeiro";
            dr["Ano"] = "2016";

            dt.Rows.Add(dr);


            return dt;
        }

        private List<Pessoa> ListarPessoa(DataSet ds)
        {
            var l = new List<Pessoa>();
            foreach (DataRow item in ds.Tables[0].Rows)
            {
                var pessoa = new Pessoa();
                pessoa.CdPessoa = item["CD_PESSOA"] != DBNull.Value ? Convert.ToInt32(item["CD_PESSOA"]) : 0;
                pessoa.Nome = item["NOME"] != DBNull.Value ? item["NOME"].ToString() : "";
                pessoa.Dinheiro = item["DINHEIRO"] != DBNull.Value ? Convert.ToDecimal(item["DINHEIRO"]) : 0;
                pessoa.DhNascimento = item["DH_NASCIMENTO"] != DBNull.Value ? (DateTime?)item["DH_NASCIMENTO"] : null;
                pessoa.BlAtivo = item["BL_ATIVO"] != DBNull.Value ? Convert.ToBoolean(item["BL_ATIVO"]) : false;

                l.Add(pessoa);
            }
            return l;
        }
    }
}
