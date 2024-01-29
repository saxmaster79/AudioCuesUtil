using System.IO.Compression;

namespace AudioCuesUtil;
internal class BackupFileReader(ExcelWriter excelWriter) : IBackupFileReader
{
    public async Task ReadBackupFileAsync(FileInfo zipInputFile, FileInfo excelOutputFile, CancellationToken cancellationToken)
    {
        using var zipfile = ZipFile.OpenRead(zipInputFile.FullName);
        using var stream = zipfile.GetEntry("AudioCuesBackup.txt")?.Open();
        if (stream is null) return;
        await excelWriter.WriteExcelAsync(stream, excelOutputFile, cancellationToken);
    }
   

}