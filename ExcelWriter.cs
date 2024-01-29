using Microsoft.Extensions.Logging;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace AudioCuesUtil;

public class ExcelWriter(JsonSerializerOptions serializerOptions, ILogger<ExcelWriter> logger)
{
    public async Task WriteExcelAsync(Stream stream, FileInfo excelOutputFileName, CancellationToken cancellationToken)
    {
        var show = await JsonSerializer.DeserializeAsync<Show>(stream, options: serializerOptions, cancellationToken);
        if (show is null) return;
        using var sl = new SLDocument();
        sl.RenameWorksheet("Sheet1", "Show");
        WriteSheetContent(sl, new List<Show>() { show }, nameof(Show.CueList));
        sl.AddWorksheet("Cues");
        WriteSheetContent(sl, show.CueList);

        sl.SaveAs(excelOutputFileName.FullName);
        logger.LogInformation("File saved to '{filePath}'", excelOutputFileName.FullName);
    }
    private static void WriteSheetContent<T>(SLDocument sl, IList<T> rows, params string[] propsToIgnore)
    {
        var properties = rows[0]!.GetType().GetProperties().Where(p => !propsToIgnore.Contains(p.Name));
        int columnIndex = 1;
        int rowindex = 2;
        foreach (var row in rows)
        {
            columnIndex = 1;

            foreach (var prop in properties)
            {
                string? propertyValue = prop.GetValue(row)?.ToString() ?? null;
                var val = sl.SetCellValue(rowindex, columnIndex, propertyValue);
                columnIndex++;
            }

            rowindex++;
        }
        columnIndex = 1;
        foreach (var prop in properties)
        {
            var val = sl.SetCellValue(1, columnIndex, prop.Name);
            sl.AutoFitColumn(columnIndex);
            columnIndex++;
        }
    }

}

public record Show(List<Cue> CueList, string ShowTitle, string ShowId);


public class Cue
{
    public string CueId { get; set; }
    public int CueIndex { get; set; }
    public string? CueNote { get; set; }
    public string? CueNumber { get; set; }
    public float Delay { get; set; }
    public float Duration { get; set; }
    public float FadeInDuration { get; set; }
    public float FadeOutDuration { get; set; }
    public string? FlicAddress { get; set; }
    public string? FlicEvent { get; set; }
    public int LeftVolume { get; set; }
    public int LoopTimes { get; set; }
    public int MasterVolume { get; set; }
    public int PlayNow { get; set; }
    public int RightVolume { get; set; }
    public int Shortcut { get; set; }
    public string ShowId { get; set; }
    public float StartTime { get; set; }
    public int StopTarget { get; set; }
    public float StopTime { get; set; }
    public string Target { get; set; }
    public string Title { get; set; }
    public string TriggerType { get; set; }
    public string Type { get; set; }
}
