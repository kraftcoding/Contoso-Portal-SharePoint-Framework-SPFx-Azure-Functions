namespace Contoso.Portal.Data.DAL.Helpers
{
    public class TemporalFile : IDisposable
    {
        public string FilePath { get; private set; }

        private TemporalFile(string filePath)
        {
            FilePath = filePath;
        }

        public static TemporalFile Create()
        {
            var outputFilePath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName(), Path.GetRandomFileName());
            Directory.CreateDirectory(Path.GetDirectoryName(outputFilePath)!);
            return new TemporalFile(outputFilePath);
        }

        public static TemporalFile Create(string fileName)
        {
            var outputFilePath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName(), fileName);
            Directory.CreateDirectory(Path.GetDirectoryName(outputFilePath)!);
            File.Create(outputFilePath).Dispose();
            return new TemporalFile(outputFilePath);
        }

        public static TemporalFile Create(string fileName, Stream content)
        {
            var outputFilePath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName(), fileName);
            Directory.CreateDirectory(Path.GetDirectoryName(outputFilePath)!);
            using var fileStream = File.Create(outputFilePath);
            content.Seek(0, SeekOrigin.Begin);
            content.CopyTo(fileStream);
            return new TemporalFile(outputFilePath);
        }

        public static TemporalFile Create(string fileName, byte[] content)
        {
            var outputFilePath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName(), fileName);
            Directory.CreateDirectory(Path.GetDirectoryName(outputFilePath)!);
            using var fileStream = File.Create(outputFilePath);
            fileStream.Write(content, 0, content.Length);
            return new TemporalFile(outputFilePath);
        }

        public void Dispose() => Directory.Delete(Path.GetDirectoryName(FilePath)!, true);
    }
}