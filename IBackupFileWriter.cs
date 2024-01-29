namespace AudioCuesUtil;
public interface IBackupFileWriter
{
    Task WriteBackupFileAsync(FileInfo excelInputFile, FileInfo zipFileToUpdate, CancellationToken token);
}