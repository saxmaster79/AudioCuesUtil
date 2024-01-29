using System;
using System.Collections.Generic;
using System.CommandLine.Invocation;
using System.CommandLine;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AudioCuesUtil
{
    public class ReadExcelFromTxtCommand : Command
    {
        private readonly ExcelWriter _excelWriter;

        public ReadExcelFromTxtCommand(ExcelWriter excelWriter) : base(nameof(ReadExcelFromTxtCommand).Replace("Command",""), "Extracts an Excel file from JSON in a TXT-File")
        {
            _excelWriter = excelWriter;
            var txtFileOption = new Option<string>(aliases:["--txtFile","--tf", "--txtfile", "--textFile", "--textfile"])
            {
                Name = "txtFile",
                Description = "The path of the JSON file (show design) from Audio Cues",
                IsRequired = true
            };
            
            AddOption(txtFileOption);
                        var excelFileOption = new Option<FileInfo>("--excelFile", getDefaultValue:()=>new FileInfo("Export.xlsx"))
            {
                Name = "excelFile",
                Description = "The path of the Excel file that should be written",
                IsRequired = false
            };
            AddOption(excelFileOption);


            this.SetHandler(HandleItAsync, txtFileOption, excelFileOption);
        }

        private async Task HandleItAsync(string txtFile, FileInfo excelOutputFileName)
        {
            var stream = new FileStream(txtFile, FileMode.Open);
            await _excelWriter.WriteExcelAsync(stream, excelOutputFileName, default);
          
        }
    }
}
