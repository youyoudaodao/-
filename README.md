C# 查询 Excel 数据
```
#region 读取Excel输出
public static DataSet GetDataSet(string sheetName, string filePath)
{
    var ds = new DataSet();
    string connStr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+ filePath + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";
    string fileType= Path.GetExtension(filePath);
    {
        using (var conn = new OleDbConnection(connStr))
        {
            try
            {
                conn.Open();
                var schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                //string sheetName1 = schemaTable.Rows[3]["TABLE_NAME"].ToString();
                string sheetName1 =  sheetName + "$";
                string sheetName2 = "'" + sheetName + "$'";
                foreach (DataRow row in schemaTable.Rows)
                {
                    string current = row["TABLE_NAME"].ToString();
                    if(current== sheetName1|| current == sheetName2)
                    {
                        var adapter = new OleDbDataAdapter($"SELECT * FROM [{current}]", conn);
                        adapter.Fill(ds, "ExcelData");
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                if (ex != null)
                    System.Console.WriteLine(ex.Message);
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                    conn.Dispose();
                }
            }

        }
    }
        
    return ds;
}
#endregion




```
