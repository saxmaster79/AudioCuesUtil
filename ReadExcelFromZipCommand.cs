using System;
using System.Collections.Generic;
using System.CommandLine.Invocation;
using System.CommandLine;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AudioCuesUtil
{
    public class ReadExcelFromZipCommand : Command
    {
        private readonly IBackupFileReader _backupFileReader;

        public ReadExcelFromZipCommand(IBackupFileReader backupFileReader) : base(nameof(ReadExcelFromZipCommand).Replace("Command",""), "Extracts an Excel file from the AudioCuesBackup.txt in the ZIP file.")
        {
            _backupFileReader = backupFileReader;
            var zipInputFile = new Option<FileInfo>(aliases: ["--zipfile", "--zf", "--ZIPFile", "--ZipFile", "--zipFile"])
            {
                Name = "zipfile",
                Description = "The path of the Audio Cues backup Zip.",
                IsRequired = true
            };

            AddOption(zipInputFile);
            var excelOutputFile = new Option<FileInfo>("--excelFile", getDefaultValue: () => new FileInfo("Export.xlsx"))
            {
                Name = "excelFile",
                Description = "The path of the Excel file that should be written",
                IsRequired = false
            };
            AddOption(excelOutputFile);


            this.SetHandler(HandleCommandAsync, zipInputFile, excelOutputFile);

        }

        private async Task HandleCommandAsync(FileInfo zipInputFile, FileInfo excelOutputFile)
        {
            await _backupFileReader.ReadBackupFileAsync(zipInputFile,excelOutputFile, default);
           
        }
    }
}
