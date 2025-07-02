C#查询擅长数据
快速读取擅长数据
#地区读取擅长输出
公共静态数据集GetDataSet(字符串sheetName，字符串filePath)
{
var ds = new DataSet()；
string connStr = @ " Provider = Microsoft。ACE . OLEDB.12.0数据源= "+file path+"；扩展属性= \ " Excel 12.0HDR =是；IMEX = 1 \ " "；
字符串文件类型=路径。get extension(file path)；
    {
使用(var conn = new oledb connection(connStr))
        {
尝试
            {
conn . Open()；
var schemaTable = conn . GetOleDbSchemaTable(OleDbSchemaGuid。表，空)；
//string sheetName1 = schemaTable。行[3]["表格名称"]。ToString()；
string sheet name 1 = sheet name+" $ "；
string sheet name 2 = " "+sheet name+" $ ' "；
foreach(schemaTable中的DataRow行。行)
                {
字符串当前=行["表格名称"]。ToString()；
if(当前==工作表名称1||当前==工作表名称2)
                    {
var adapter = new oledb dataadapter($ " SELECT * FROM[{当前}]“，conn)；
适配器。Fill(ds，“excel data”)；
打破；
                    }
                }
            }
catch(例外ex)
            {
如果(ex！=空)
系统。控制台. WriteLine(例如。消息)；
            }
最后
            {
如果(conn！=空)
                {
conn . Close()；
conn . Dispose()；
                }
            }

        }
    }
        
    返回ds；
}
#结束区域



