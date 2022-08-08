using System.Diagnostics.CodeAnalysis;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Unicode;

namespace ExcelParser
{
    internal partial class Settings
    {
        private static readonly JsonSerializerOptions _serializerOptions = new(JsonSerializerDefaults.General)
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
                    ? JsonSerializer.Deserialize<Settings>(File.ReadAllText(SettingsFileName), _serializerOptions) ?? new()
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
                string json = JsonSerializer.Serialize(this, _serializerOptions);
                File.WriteAllText(SettingsFileName, json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static Settings Change()
        {
            int command = -1;
            ShowCommands();
            var settings = Load();
            do
            {
                command = int.TryParse(Console.ReadLine(), out int value) ? value : -1;
                switch (command)
                {
                    case 1:
                        settings.Save();
                        break;
                }
            } while (command != 0 && command != 1);

            return Load();
        }

        private static void ShowCommands()
        {
            Console.WriteLine("0 - отменить");
            Console.WriteLine("1 - сохранить");
            Console.WriteLine("2 - столбец первого файла");
            Console.WriteLine("3 - столбец второго файла");
            Console.WriteLine("4 - столбец третьего файла");
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
        public string DefaultBarcode { get; set; } = "00000000000";
        public string WorksheetName { get; set; } = "Todo add name";
        public string VendorCode2Header { get; set; } = "Артикул 2";
        public string NameHeader { get; set; } = "Наименование";
        public string CountHeader { get; set; } = "Количество";
        public string BarcodeHeader { get; set; } = "Штрих-код";
        public string SolutionFolder { get; set; } = "Solutions\\";
        public string SolutionFileNamePrefix { get; set; } = "S.";
        public string SolutionFileNameSuffix { get; set; } = "_";
    }
}
