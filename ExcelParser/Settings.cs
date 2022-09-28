using System.Diagnostics.CodeAnalysis;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Unicode;

namespace ExcelParser
{
    internal class Settings
    {
        private static readonly JsonSerializerOptions SerializerOptions = new(JsonSerializerDefaults.General)
        {
            DefaultIgnoreCondition = JsonIgnoreCondition.Never,
            MaxDepth = 10,
            WriteIndented = true,
            Encoder = JavaScriptEncoder.Create(UnicodeRanges.BasicLatin, UnicodeRanges.Cyrillic),
        };

        private const string SettingsFileName = "settings.json";

        public FirstDocument FirstDocument { get; set; } = new();
        public SecondDocument SecondDocument { get; set; } = new();
        public ThirdDocument ThirdDocument { get; set; } = new();
        public SolutionDocument SolutionDocument { get; set; } = new();

        public Settings()
        {
        }

        [UnconditionalSuppressMessage("Trimming", "IL2026:Members annotated with 'RequiresUnreferencedCodeAttribute' require dynamic access otherwise can break functionality when trimming application code", Justification = "<Pending>")]
        public static Settings Load()
        {
            Settings settings;
            try
            {
                settings = File.Exists(SettingsFileName)
                    //? (JsonSerializer.Deserialize(File.ReadAllText(SettingsFileName), typeof(Settings), ExcelParserContext.Default) as Settings) ?? new()
                    ? JsonSerializer.Deserialize<Settings>(File.ReadAllText(SettingsFileName), SerializerOptions) ?? new()
                    : new Settings();
            }
            catch
            {
                settings = new();
            }

            settings.Save();

            return settings;
        }

        [UnconditionalSuppressMessage("Trimming", "IL2026:Members annotated with 'RequiresUnreferencedCodeAttribute' require dynamic access otherwise can break functionality when trimming application code", Justification = "<Pending>")]
        public void Save()
        {
            try
            {
                string json = JsonSerializer.Serialize(this, SerializerOptions);
                File.WriteAllText(SettingsFileName, json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class FirstDocument
    {
        public string FodlerName { get; set; }
        public int VendorCodeColumnNumber { get; set; }
        public string? WorksheetName { get; set; }
        public List<string> WordsFilter { get; set; }

        public FirstDocument()
        {
            FodlerName = "Orders\\";
            VendorCodeColumnNumber = 3;
            WorksheetName = "ДСП";
            WordsFilter = new()
            {
                "Категория в базе не найдена Материал не найден",
                "Артикул"
            };
        }
    }

    public class SecondDocument
    {
        public string FileName { get; set; } = "Examples\\2.xlsx";
        public string? WorksheetName { get; set; } = null;


        public int VendorCode1ColumnNumber { get; set; } = 1;
        public int VendorCode2ColumnNumber { get; set; } = 2;
        public int NameColumnNumber { get; set; } = 3;
        public int CountColumnNumber { get; set; } = 4;
    }

    public class ThirdDocument
    {
        public string FileName { get; set; } = "Examples\\3.xlsx";
        public string? WorksheetName { get; set; } = null;

        public int VendorCode2ColumnNumber { get; set; } = 2;
        public int BarcodeColumnNumber { get; set; } = 4;
    }

    public class SolutionDocument
    {
        public bool OutExcel { get; set; } = true;
        public bool OutDm { get; set; } = true;
        public int BarcodeLength { get; set; } = 12;
        public string WorksheetName { get; set; } = "Todo add name";
        public string VendorCode2Header { get; set; } = "Артикул 2";
        public string NameHeader { get; set; } = "Наименование";
        public string CountHeader { get; set; } = "Количество";
        public string BarcodeHeader { get; set; } = "Штрих-код";
        public string ExcelFolder { get; set; } = "ExcelSolutions\\";
        public string DmFolder { get; set; } = "DmSolutions\\";
        public string SolutionFileNamePrefix { get; set; } = "S.";
        public string SolutionFileNameSuffix { get; set; } = "_";
        public int Iterator { get; set; }
    }
}
