namespace AudioCuesUtil;

    public interface IBackupFileReader
    {
        Task ReadBackupFileAsync(FileInfo zipInputFile, FileInfo excelOutputFile, CancellationToken cancellationToken);
    }
