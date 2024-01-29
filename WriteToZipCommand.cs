using System;
using System.Collections.Generic;
using System.CommandLine.Invocation;
using System.CommandLine;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AudioCuesUtil
{
    public class WriteToZipCommand : Command
    {
        private readonly IBackupFileWriter _backupFileWriter;

        public WriteToZipCommand(IBackupFileWriter backupFileWriter) : base("WriteToZip", "Updates the AudioCuesBackup.txt in the ZIP file and copies the existing AudioCuesBackup.txt to a .bak file.")
        {
            _backupFileWriter = backupFileWriter;
            var zipFileToUpdate = new Option<FileInfo>(aliases: ["--zipfile", "--zf", "--ZIPFile", "--ZipFile", "--zipFile"])
            {
                Name = "zipfile",
                Description = "The path of the Audio Cues backup Zip.",
                IsRequired = true
            };
            AddOption(zipFileToUpdate);
            var excelOutputFile = new Option<FileInfo>("--excelFile", getDefaultValue: () => new FileInfo("Export.xlsx"))
            {
                Name = "excelFile",
                Description = "The path of the Excel file that should be written",
                IsRequired = false
            };
            AddOption(excelOutputFile);
            this.SetHandler(HandleCommandAsync, excelOutputFile, zipFileToUpdate);

        }

        private async Task HandleCommandAsync(FileInfo excelInputFile, FileInfo zipFileToUpdate)
        {

                var writer = _backupFileWriter;
                await writer.WriteBackupFileAsync(excelInputFile, zipFileToUpdate, default);
           
        }
    }
}
