using System.IO.Compression;
using System.Text.Json;

namespace AudioCuesUtil;

internal class BackupFileWriter(JsonSerializerOptions serializerOptions, ExcelReader excelReader) : IBackupFileWriter
{


    public async Task WriteBackupFileAsync(FileInfo excelInputFile, FileInfo zipFileToUpdate, CancellationToken token)
    {
        var show = excelReader.ReadShow(excelInputFile);

        using (var archive = new ZipArchive(File.Open(zipFileToUpdate.FullName, FileMode.Open, FileAccess.ReadWrite), ZipArchiveMode.Update))
        {
            foreach (var entry in archive.Entries.ToArray().Where(e => e.Name == "AudioCuesBackup.txt"))
            {
                var copyEntry = archive.CreateEntry(CreateBackupFileName(entry.Name));
                using (var a = entry.Open())
                using (var b = copyEntry.Open())
                    a.CopyTo(b);
                entry.Delete();
            }
            var newEntry= archive.CreateEntry("AudioCuesBackup.txt");
            using var stream = newEntry.Open();
            await JsonSerializer.SerializeAsync(stream, show, options: serializerOptions, cancellationToken: token);
        }

    }

    private string CreateBackupFileName(string name)
    {
        return $"{name}.{DateTime.Now:yyyyMMddHHmmss}.bak";
    }

}

