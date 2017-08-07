using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Linq;

namespace SalesUtil
{
    public abstract class ExcelDataObject
    {
        public Func<string, string> GetValue;
        public abstract void Init(Func<string, string> Getter);
    }

    public class ExcelContext<TEntity> : IDisposable
        where TEntity : ExcelDataObject, new()
    {
        public string GetValue(string[] columnNames, DataRow RawData, string columnFilter)
        {
            string column = (from c in columnNames
                             where c.StartsWith(columnFilter)
                             select c).FirstOrDefault();
            if (string.IsNullOrEmpty(column))
                throw new Exception(string.Format("Column {0} cannot be found.", columnFilter));

            string cellData = RawData[column].ToString().Trim();

            return cellData;
        }

        public List<TEntity> Factory()
        {
            ColumnNames = table.Columns.Cast<DataColumn>()
                     .Select(x => x.ColumnName)
                     .ToArray();

            var result = new List<TEntity>();
            foreach (DataRow row in table.Rows)
            {
                Func<string, string> getter = (x) => GetValue(ColumnNames, row, x);
                var obj = new TEntity();
                obj.Init(getter);
                result.Add(obj);
            }
            return result;
        }

        string[] ColumnNames { get; set; }
        public DataTable table { get; private set; }

        DataSet ds = new DataSet();

        public string ExcelPath { get; set; }
        public string DataSheet { get; set; }

        OleDbConnection ExcelConnection;

        //public ExcelContext()
        //{
        //    InitConext();
        //}

        public ExcelContext(string ExcelPath, string DataSheet)
        {
            this.ExcelPath = ExcelPath;
            this.DataSheet = DataSheet;            
            InitConext();
        }

        private void InitConnection()
        {
            if (string.IsNullOrEmpty(ExcelPath) || string.IsNullOrEmpty(DataSheet))
                throw new Exception(string.Format("Excel context is not initiated, missing excel path: {0} or data sheet name: {1}", ExcelPath, DataSheet));
            var connectionString = string.Format(ConfigurationManager.ConnectionStrings["ExcelConnectionString"].ConnectionString, ExcelPath);
            ExcelConnection = new OleDbConnection(connectionString);


            ExcelConnection.Open();
        }

        private void InitConext()
        {
            InitConnection();

            var cmd = new OleDbCommand()
            {
                Connection = ExcelConnection,
                CommandText = string.Format("SELECT * FROM `{0}`", DataSheet + '$')
            };

            table = new DataTable() { TableName = DataSheet };
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(table);
            ds.Tables.Add(table);
            Console.WriteLine("Excel context {0} initiated.", ExcelPath.Split('\\').Last());
        }

        public void ListExcelSheets()
        {
            InitConnection();

            DataTable dtSheet = ExcelConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            foreach (DataRow dr in dtSheet.Rows)
            {
                string sheet = dr["TABLE_NAME"].ToString();

                Console.WriteLine(sheet);
            }
        }

        public void Dispose()
        {
            ExcelConnection.Close();
        }
    }
}
