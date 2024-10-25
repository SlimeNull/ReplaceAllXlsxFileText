using System.Diagnostics;
using System.IO.Compression;
using System.Text;

namespace ReplaceAllXlsxFileText
{
    internal class Program
    {
        static void Pause()
        {
            Console.WriteLine("按任意键继续");
            Console.ReadKey();
        }

        static string EscapeForXml(string toEscape)
        {
            return new StringBuilder(toEscape)
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;")
                .ToString();
        }

        static void SafeDoForAllFilesUnderFolder(string folderPath, Action<string> action)
        {
            ArgumentNullException.ThrowIfNull(folderPath);
            ArgumentNullException.ThrowIfNull(action);

            try
            {
                if (!Directory.Exists(folderPath))
                {
                    return;
                }

                // 执行操作
                foreach (var file in Directory.EnumerateFiles(folderPath, "*", SearchOption.TopDirectoryOnly))
                {
                    action.Invoke(file);
                }

                // 手动递归
                // 之所以手动递归, 是因为直接调用递归搜索, 遇到异常, 递归会停止
                foreach (var subFolder in Directory.EnumerateDirectories(folderPath, "*", SearchOption.TopDirectoryOnly))
                {
                    SafeDoForAllFilesUnderFolder(subFolder, action);
                }
            }
            catch
            {
                // pass
            }
        }

        static void Main(string[] args)
        {
            Console.WriteLine("输入目录路径: ");
            var folderPath = Console.ReadLine();

            if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
            {
                Console.WriteLine("目录不存在");
                Pause();
                return;
            }

            Console.WriteLine("输入要替换的字符串: ");
            var toReplace = Console.ReadLine();
            if (string.IsNullOrWhiteSpace(toReplace))
            {
                Console.WriteLine("输入为空");
                Pause();
                return;
            }

            Console.WriteLine("输入替换后的字符串: ");
            var newString = Console.ReadLine();
            if (string.IsNullOrWhiteSpace(newString))
            {
                Console.WriteLine("输入为空");
                Pause();
                return;
            }

            var escapedToReplace = EscapeForXml(toReplace);
            var escapedNewString = EscapeForXml(newString);

            int xlsxCount = 0;
            int replacedFileCount = 0;
            int exceptionCount = 0;
            Stopwatch stopwatch = Stopwatch.StartNew();

            SafeDoForAllFilesUnderFolder(folderPath, filePath =>
            {
                if (!filePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                try
                {
                    using (ZipArchive archive = ZipFile.Open(filePath, ZipArchiveMode.Update))
                    {
                        var entry = archive.GetEntry("xl/sharedStrings.xml");

                        if (entry is null)
                        {
                            return;
                        }

                        using var entryStream = entry.Open();
                        using var entryStreamReader = new StreamReader(entryStream);

                        var entryContent = entryStreamReader.ReadToEnd();
                        var newSharedStringsContent = entryContent.Replace(escapedToReplace, escapedNewString, StringComparison.Ordinal);

                        if (entryContent == newSharedStringsContent)
                        {
                            return;
                        }

                        entryStream.Seek(0, SeekOrigin.Begin);
                        entryStream.SetLength(0);
                        using var entryStreamWriter = new StreamWriter(entryStream);
                        entryStreamWriter.Write(newSharedStringsContent);
                        replacedFileCount++;
                    }
                }
                catch
                {
                    exceptionCount++;
                    // ignore
                }
            });

            stopwatch.Stop();
            Console.WriteLine($"已完成, 共扫描到 {xlsxCount} 个 XLSX 文件, 对 {replacedFileCount} 执行了替换操作, 失败 {exceptionCount} 次. 共用时 {stopwatch.ElapsedMilliseconds}ms");
            Pause();
        }
    }
}
