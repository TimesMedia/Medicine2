using System;
using System.Data;
using System.Data.SqlClient;


namespace Medicine2.Data
{
	/// <summary>
	/// Summary description for ExceptionData.
	/// </summary>
	public abstract class ExceptionData
	{
		private static SqlConnection Connection = new SqlConnection();

        private ExceptionData()
        {
            // This prevents a default constructor from being created.
        }



        public static string DatabaseConnection
        {
            get
            {
                return @"Data Source = pklwebdb01\mssql2016std; Initial Catalog = Medicine; Integrated Security = True";
            }
        }

		public static void WriteException(int Severity, string Message, string Object, string Method, 
			string Comment) 
		{
            try
            {
                //Remember the stuff in the database

                SqlCommand Command = new SqlCommand();
                SqlDataAdapter Adaptor = new SqlDataAdapter();

                Connection.ConnectionString = DatabaseConnection;

                Connection.Open();
                Command.Connection = Connection;
                Command.CommandType = CommandType.StoredProcedure;
                Command.CommandText = "[ExceptionData.WriteException]";
                SqlCommandBuilder.DeriveParameters(Command);

                Command.Parameters["@Severity"].Value = Severity;
                Command.Parameters["@Message"].Value = Message;
                Command.Parameters["@Object"].Value = Object;
                Command.Parameters["@Method"].Value = Method;
                Command.Parameters["@Comment"].Value = Comment;

                Command.ExecuteScalar();

            }
            finally
            {
                Connection.Close();
            }

		}
	}
}
