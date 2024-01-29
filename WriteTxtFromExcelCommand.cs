using Microsoft.Extensions.Logging;
using System.CommandLine;
using System.Text.Json;

namespace AudioCuesUtil
{
    public class WriteTxtFromExcelCommand : BaseCommand
    {
        private readonly ExcelReader _excelReader;
        private readonly JsonSerializerOptions _serializerOptions;
        private readonly ILogger<WriteTxtFromExcelCommand> _logger;

        public WriteTxtFromExcelCommand(ExcelReader excelReader, JsonSerializerOptions serializerOptions, ILogger<WriteTxtFromExcelCommand> logger) : base(typeof(WriteTxtFromExcelCommand), "Writes the JSON file.")
        {
            _excelReader = excelReader;
            _serializerOptions = serializerOptions;
            _logger = logger;
            var excelFileOption = new Option<FileInfo>("--excelFile")
            {
                Name = "excelFile",
                Description = "The path of the Excel file that should be read",
                IsRequired = true
            };
            AddOption(excelFileOption);

            var txtFileOption = new Option<FileInfo>(aliases: ["--txtFile", "--tf", "--txtfile", "--textFile", "--textfile"], getDefaultValue: () => new FileInfo("AudioCuesBackup.txt"))
            {
                Name = "txtFile",
                Description = "The path of the JSON file (show design) from Audio Cues, that should be writter",
                IsRequired = true
            };
            AddOption(txtFileOption);

            var timestampOption = new Option<bool>(aliases: ["--addTimestamp", "--ts", "--addtimestamp", "--addTimeStamp"], getDefaultValue: () => false)
            {
                Name = "addTimestamp",
                Description = "Adds a Timestamp to the output file name",
                IsRequired = false
            };
            AddOption(timestampOption);

            this.SetHandler(HandleCommandAsync, excelFileOption, txtFileOption, timestampOption);

        }

        private async Task HandleCommandAsync(FileInfo excelInputFile, FileInfo textOutputFile, bool addTimestamp)
        {
            CancellationToken token = default;
            var show = _excelReader.ReadShow(excelInputFile);
            string json = JsonSerializer.Serialize(show, _serializerOptions);

            var filenName = GetFileName(textOutputFile, addTimestamp);
            
            await File.WriteAllTextAsync(filenName, json, token);
            _logger.LogInformation("File saved to '{jsonPath}'", filenName);
        }

        private string GetFileName(FileInfo textOutputFile, bool addTimestamp)
        {
            if(!addTimestamp)return textOutputFile.FullName;
            return $"{textOutputFile.Directory.FullName}{Path.DirectorySeparatorChar}{Path.GetFileNameWithoutExtension(textOutputFile.FullName)}.{DateTime.Now:yyyyMMddHHmmss}{textOutputFile.Extension}";
        }
    }
}
