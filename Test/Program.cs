using ClosedXML.Excel;
using ExcelUtils;
using System.Data;
using System.Drawing;
using System.Reflection;
using System.Text.Json;

public class Program
{
    static void Main(string[] args)
    {
        try
        {
            ExcelTest();
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.ToString());
        }
        Console.ReadLine();
    }

    static void ExcelTest()
    {

        byte[] DocumentBytes;
        var DS = new DataSet();
        String ProgramFolder = Path.GetDirectoryName(Assembly.GetEntryAssembly()?.Location)?.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
        Console.WriteLine($"ExcelTest.");
        Console.WriteLine($"Старт в папке: {ProgramFolder}");
        DS.ReadXml($"{ProgramFolder}ExcelData.xml");
        Console.WriteLine($"DataSet прочитан: {ProgramFolder}ExcelData.xml");

        ExcelTable excel = new ExcelTable(DS.Tables[1], new TableInfo(), 1, 1, true, DS.Tables[0].TableName);
        excel.TablePrepare("ServerTypeName, ServerName", "", "", "ServerTypeName = 'SQL Server' OR ServerTypeName = 'PostgreSQL'");

        TableInfo ti = new TableInfo();
        // создание столбцов на основе данных таблицы
        ti.ColumnsCreate(DS.Tables[1]);
        excel.ti = ti;
        // сортировка
        ti.ColumnSort = "ServerTypeName asc, ServerName, Author";

        // условное форматирование строк таблицы
        var cfs = new TableInfo.ConditionFormattingTableInfo(TableInfo.Conditions.ContainNulls);
        cfs.SourceColumn = "IDAuthor";
        //cfs.TargetColumn = "DBName";
        cfs.Formating.FillColor = Color.Yellow.ToArgb();
        ti.ConditionFormattings.Add(cfs);

        // группировка столбцов
        ti.Groups.Add("ServerTypeName");
        ti.Groups.Add("ServerName");
        ti.Groups.Add("Author");

        // формат столбцов группировки
        ti.FormatingGroups.FontBold = true;
        ti.FormatingGroups.FillColor = Color.LightGray.ToArgb();

        // цикл по опсанию столбцов
        foreach (var ci in ti.Columns)
        {
            if (ci.ColumnName == "ServerTypeName")
            {
                ci.SummatyList.Add(SummaryTypes.Count);
                ci.ColumnCaption = "Тип сервера";
                ci.Width = 50;
                // формат столбца
                ci.Formating.Alignment = ColumnAlignment.Center;
                ci.Formating.FillColor = Color.LightBlue.ToArgb();
                ci.Formating.FontBold = true;
                ci.Formating.FontColor = Color.Red.ToArgb();
                ci.Formating.FontItalic = true;
                ci.Formating.FontName = "Tahoma";
                ci.Formating.FontSize = 14;
                ci.Formating.FontStrikethrough = false;
                ci.Formating.FontUnderline = false;
                ci.Formating.NumberFormat = "";
            }
            if (ci.ColumnName == "ServerType")
            {
                ci.SummatyList.Add(SummaryTypes.Count);
                ci.SummatyList.Add(SummaryTypes.Sum);
                ci.SummatyList.Add(SummaryTypes.Min);
                ci.SummatyList.Add(SummaryTypes.Max);
                ci.SummatyList.Add(SummaryTypes.Average);
            }
            // формат заголовка столбца
            ci.FormatingHeader.FontBold = true;
            ci.FormatingHeader.Alignment = ColumnAlignment.Center;
        }

        // генерация документа Excel
        //excel = new ExcelTable(DS.Tables[1], ti, 1, 1, true, DS.Tables[0].TableName);
        DocumentBytes = excel.ExcelGenegate();

        string jsonString = JsonSerializer.Serialize(ti);
        File.WriteAllText($"{ProgramFolder}Excel.json", jsonString);

        Console.WriteLine("Документ Excel сформирован");
        File.WriteAllBytes($"{ProgramFolder}Excel.xlsx", DocumentBytes);
        Console.WriteLine($"Документ Excel сохранен: {ProgramFolder}Excel.xlsx");
    }
}

