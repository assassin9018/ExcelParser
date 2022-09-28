using System.Globalization;
using CsvHelper;
using CsvHelper.Configuration;
using ExcelParser.Models;

namespace ExcelParser.Writers;

internal class ResultDmWriter
{
    private const string Extention = ".dm";
    private static readonly string EmptyBarCode = string.Concat(Enumerable.Range(0, 13).Select(x=>'0').ToArray());
    private readonly SolutionDocument _settings;
    private readonly CsvConfiguration _configuration;

    public ResultDmWriter(Settings settings)
    {
        _settings = settings.SolutionDocument;
        _configuration = new CsvConfiguration(CultureInfo.CurrentCulture)
        {
            Delimiter = ";"
        };
    }

    public void Write(ICollection<ResultTableRow> rows, string sourceFileName)
    {
        string fileName = GetFileName(sourceFileName);
        DateTime generationTime = DateTime.Now;
        
        using FileStream fs = File.Create(fileName);
        using TextWriter textWriter = new StreamWriter(fs);
        using var writer = new CsvWriter(textWriter, _configuration);
        WriteCount(writer, rows.Count + 2);
        WriteLoadMode(writer, generationTime);
        WriteTeplate(writer,rows.Count + 2);
        WriteProductRows(writer, generationTime, fileName, rows);
    }

    private static void WriteCount(CsvWriter writer, int count)
    {
        writer.WriteField(count);
        writer.NextRecord();
    }

    private void WriteLoadMode(CsvWriter writer, DateTime generationTime)
    {
        //+;dk8#b038024c-9e62-11ea-979a-341a4c115056;БПРЦ-000167;2020-05-25 02:10:10.000;БПРЦ-000167;;;;
        //РежимЗагрузкиФайла
        writer.WriteField("+");
        //ИдентификаторДокумента
        writer.WriteField(Guid.NewGuid());
        //НомерДокумента
        writer.WriteField($"БПРЦ-{_settings.Iterator.ToString().PadLeft(6, '0')}");
        //ДатаДокумента
        writer.WriteField(generationTime);
        //ШтрихКодДокумента
        writer.WriteField(EmptyBarCode);
        //КомментарийДокумента
        writer.WriteField("");
        //Контрагент
        writer.WriteField("");
        //Склад
        writer.WriteField("");
        // //ИНН
        // writer.WriteField("");
        // //РодительскийДокумент
        // writer.WriteField("");
        // //IDРодительского
        // writer.WriteField("");
        // //ТипЛогикиГрупповогоДокумента
        // writer.WriteField("");
        // //Приоритет
        // writer.WriteField("");
        // //Склад2
        // writer.WriteField("");
        writer.NextRecord();
    }

    private void WriteTeplate(CsvWriter writer, int unknownValue)
    {
        //7;Перемещение;ПеремещениеТоваров;
        //1 - 1;
        //4 - 0;0;0;0;0;
        //2 - 1;1;
        //4 - 0;0;0;0;0;
        //1 - 1;
        //2 - 0;0;
        //2 - 1;1;
        //20 - 0;0;0;0; 0;0;0;0; 0;0;0;0; 0;0;0;0; 0;0;0;0;
        //1 - 2;
        //24 - 0;0;0;0; 0;0;0;0; 0;0;0;0; 0;0;0;0; 0;0;0;0; 0;0;0;0; 0;0;0;0;
        //2 - ;;
        //7 - 0;0;0;0;0;0;0;
        writer.WriteField(unknownValue);
        writer.WriteField("Перемещение");
        writer.WriteField("ПеремещениеТоваров");
        AddIdenticalValues(writer, 1, 1);
        AddIdenticalValues(writer, 4, 0);
        AddIdenticalValues(writer, 2, 1);
        AddIdenticalValues(writer, 4, 0);
        AddIdenticalValues(writer, 1, 1);
        AddIdenticalValues(writer, 2, 0);
        AddIdenticalValues(writer, 2, 1);
        AddIdenticalValues(writer, 20, 0);
        AddIdenticalValues(writer, 1, 2);
        AddIdenticalValues(writer, 28, 0);
        AddIdenticalValues(writer, 2, "");
        AddIdenticalValues(writer, 7, 0);
        writer.NextRecord();
    }

    private static void WriteProductRows(CsvWriter writer, DateTime generationTime, string fileName,
        ICollection<ResultTableRow> rows)
    {
        //1;S;8U-b0380245-9e62-11ea-979a-341a4c11505600000000-0000-0000-0000-000000000000;5201409809378;;;4;4;
        
        int i = 1;
        foreach (var row in rows)
        {
            //НомерСтроки - 1
            writer.WriteField(i++);
            //S|I - S
            writer.WriteField("S");
            //ИдентификаторТовара - 8U-b0380245-9e62-11ea-979a-341a4c11505600000000-0000-0000-0000-000000000000
            writer.WriteField(row.VendorCode2);
            //ШтрихкодТовара - 5201409809378
            writer.WriteField(row.Barcode);
            //ШтрихКодЯчейки - 
            writer.WriteField("");
            //СерийныйНомер - 
            writer.WriteField("");
            //Количество
            writer.WriteField(row.Count);
            //Лимит
            writer.WriteField(row.Count);
            writer.NextRecord();
        }
    }

    private string GetFileName(string sourceFileName)
    {
        string dir = Path.GetDirectoryName(_settings.DmFolder) ?? string.Empty;
        string fileName = $"19102020184919_v83_Doc_БПРЦ-{_settings.Iterator.ToString().PadLeft(6, '0')}";
        string fullName = Path.Combine(dir, fileName + Extention); 
        
        if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
            Directory.CreateDirectory(dir);
        if (File.Exists(fullName))
            File.Delete(fullName);

        return fullName;
    }

    private static void AddIdenticalValues<T>(CsvWriter writer, int valuesCount, T value)
    {
        for (int i = 0; i < valuesCount; i++)
            writer.WriteField(value);
    }
}