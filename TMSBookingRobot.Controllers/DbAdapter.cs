using System.Data;
using System.Data.SqlClient;
using System.Linq;

namespace TMSBookingRobot.Controllers
{
    internal class DbAdapter
    {
        static internal DataSet QueryToDS(
            string commandText,
            string sqlConnectionString,
            CommandType commandType = CommandType.Text,
            params SqlParameter[] parameters)
        {
            var dataSet = new DataSet();

            using (SqlConnection conn = new SqlConnection(sqlConnectionString))
            {
                conn.Open();
                var cmd = conn.CreateCommand();
                cmd.CommandText = commandText;
                cmd.CommandType = commandType;

                if (parameters != null && parameters.Count() > 0)
                    cmd.Parameters.AddRange(parameters);

                var adapter = new SqlDataAdapter(cmd);
                adapter.Fill(dataSet);
            }

            return dataSet;
        }

        static internal int Executed(
            string commandText,
            string sqlConnectionString,
            CommandType commandType = CommandType.Text,
            params SqlParameter[] parameters)
        {
            using (SqlConnection conn = new SqlConnection(sqlConnectionString))
            {
                conn.Open();
                var cmd = conn.CreateCommand();
                cmd.CommandText = commandText;
                cmd.CommandType = commandType;

                if (parameters != null && parameters.Count() > 0)
                    cmd.Parameters.AddRange(parameters);

                return cmd.ExecuteNonQuery();
            }
        }
    }
}