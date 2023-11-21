using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security.Principal;
using System.Text.Json;
using Microsoft.Win32.SafeHandles;

namespace BandizipAutoExtractor;

public class Config
{
    /// <summary>
    ///     配置文件的版本
    /// </summary>
    public string Version { get; set; } = "1.0";

    /// <summary>
    ///     Bandizip路径
    /// </summary>
    public string BandizipPath { get; set; }

    /// <summary>
    ///     解压输出路径
    /// </summary>
    public string OutputPath { get; set; }

    /// <summary>
    ///     是否每次解压完成后打开输出目录
    /// </summary>
    public bool OpenOutputPath { get; set; }

    /// <summary>
    ///   是否使用Bandizip命令行程序解压
    /// </summary>
    public bool UseBandizipConsoleTool { get; set; }
}

internal class Program
{
    private static void Main(string[] args)
    {
        if (args.Length < 1)
        {
            Console.WriteLine("用法：BandizipAutoExtractor.exe <压缩文件路径>");
            Console.WriteLine(
                "示例：BandizipAutoExtractor.exe \"C:\\Users\\happy\\Downloads\\Example.zip\"");
            Console.WriteLine("点按任意按键退出...");
            Console.ReadKey();
            return;
        }

        var configPath =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                         ".BandizipAutoExtractor", "config.json");
        string bandizipPath;
        string outputPath;
        bool openOutputPath;
        bool useBandizipConsoleTool;

        if (File.Exists(configPath))
        {
            try
            {
                var config = JsonSerializer.Deserialize<Config>(File.ReadAllText(configPath));
                if (config == null)
                {
                    Console.WriteLine("配置文件解析失败，请考虑删除配置文件后重启软件，程序退出...");
                    return;
                }

                // 验证配置对象的属性
                if (string.IsNullOrEmpty(config.BandizipPath) ||
                    string.IsNullOrEmpty(config.OutputPath))
                {
                    Console.WriteLine("配置文件中的Bandizip路径或输出路径无效，请考虑删除配置文件后重启软件，程序退出...");
                    return;
                }

                bandizipPath = config.BandizipPath;
                outputPath = config.OutputPath;
                openOutputPath = config.OpenOutputPath;
                useBandizipConsoleTool = config.UseBandizipConsoleTool;
            }
            catch (JsonException ex)
            {
                Console.WriteLine("无法正确解析配置文件 JSON 格式：" + ex.Message);
                Console.WriteLine("请考虑删除配置文件后重启软件，程序退出...");
                return;
            }
            catch (IOException ex)
            {
                Console.WriteLine("无法读取配置文件：" + ex.Message);
                Console.WriteLine("请考虑删除配置文件后重启软件，程序退出...");
                return;
            }
            catch (Exception ex)
            {
                Console.WriteLine("读取配置文件时发生未知错误：" + ex.Message);
                Console.WriteLine("请考虑删除配置文件后重启软件，程序退出...");
                return;
            }
        }
        else
        {
            Console.WriteLine("首次运行BandizipAutoExtractor，开始配置...");
            // 查找Bandizip路径
            Console.WriteLine("正在查找Bandizip安装目录...");
            var searchResult = SafeFileSearch(@"C:\", "Bandizip.exe");
            // 如果找到多个Bandizip.exe，让用户选择
            var enumerable = searchResult.ToList();
            if (enumerable.Count > 1)
            {
                Console.WriteLine("找到多个Bandizip安装目录，请选择：");
                var i = 1;
                foreach (var result in enumerable)
                {
                    Console.WriteLine($"{i}. {result}");
                    i++;
                }

                Console.WriteLine("请输入序号：");
                var choice = Console.ReadLine();
                if (int.TryParse(choice, out var index) && index > 0 &&
                    index <= enumerable.Count)
                {
                    bandizipPath = enumerable.ElementAt(index - 1);
                }
                else
                {
                    Console.WriteLine("输入错误，程序退出...");
                    return;
                }
            }
            else
            {
                bandizipPath = enumerable.FirstOrDefault();
            }

            if (string.IsNullOrEmpty(bandizipPath))
            {
                Console.WriteLine("未找到Bandizip安装目录，程序退出");
                return;
            }

            Console.WriteLine($"找到Bandizip安装目录: {bandizipPath}，是否确认？(y/n)");
            if (Console.ReadLine()?.ToLower() != "y")
            {
                return;
            }

            Console.WriteLine(
                "BandizipAutoExtractor提供了普通Bandizip.exe解压（调用常用UI界面）和命令行程序bz.exe，请选择：");
            Console.WriteLine("1. 普通Bandizip.exe解压（推荐）");
            Console.WriteLine("2. 命令行程序bz.exe");
            Console.WriteLine("请输入序号：");
            var userChoice1 = Console.ReadLine();
            switch (userChoice1)
            {
                case "1":
                    bandizipPath = Path.Combine(Path.GetDirectoryName(bandizipPath) ?? string.Empty,
                                                "Bandizip.exe");
                    useBandizipConsoleTool = false;
                    break;
                case "2":
                    bandizipPath = Path.Combine(Path.GetDirectoryName(bandizipPath) ?? string.Empty,
                                                "bz.exe");
                    useBandizipConsoleTool = true;
                    break;
                default:
                    Console.WriteLine("输入错误，程序退出...");
                    return;
            }

            // 获取输出路径
            outputPath =
                Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "Programs");
            Console.WriteLine($"默认解压输出路径为: {outputPath}，你也可以自定义输出路径，请选择：");
            Console.WriteLine("1. 默认解压输出路径");
            Console.WriteLine("2. 自定义输出路径");
            Console.WriteLine("请输入序号：");
            var userChoice2 = Console.ReadLine();
            if (userChoice2 == "2")
            {
                Console.WriteLine("请输入自定义输出路径：");
                var customPath = Console.ReadLine();

                // 验证自定义路径是否有效
                if (!string.IsNullOrWhiteSpace(customPath) &&
                    Directory.Exists(Path.GetDirectoryName(customPath)))
                {
                    outputPath = customPath;
                }
                else
                {
                    Console.WriteLine("无效的自定义路径，将使用默认路径。");
                }
            }
            else if (userChoice2 != "1")
            {
                Console.WriteLine("输入错误，程序退出...");
                return;
            }

            // 是否每次解压完成后打开输出目录
            Console.WriteLine("是否每次解压完成后打开输出目录？(y/n)");
            var userChoice3 = Console.ReadLine();
            switch (userChoice3?.ToLower())
            {
                case "y":
                    openOutputPath = true;
                    break;
                case "n":
                    openOutputPath = false;
                    break;
                default:
                    Console.WriteLine("输入错误，程序退出...");
                    return;
            }

            // 保存Bandizip路径和输出路径到配置文件
            var config = new Config
            {
                BandizipPath = bandizipPath,
                OutputPath = outputPath,
                OpenOutputPath = openOutputPath
            };
            Directory.CreateDirectory(Path.GetDirectoryName(configPath) ?? string.Empty);
            File.WriteAllText(configPath, JsonSerializer.Serialize(config));
            Console.WriteLine($"配置已保存到: {configPath}，你可以自行编辑");
        }

        // 拼接Bandizip命令行参数
        var command = $"x -o:{outputPath} -target:auto \"{args[0]}\"";

        // 检查当前进程是否以管理员权限运行
        var identity = WindowsIdentity.GetCurrent();
        var principal = new WindowsPrincipal(identity);
        if (principal.IsInRole(WindowsBuiltInRole.Administrator))
        {
            Console.WriteLine("当前进程以管理员权限运行");
        }

        // 启动Bandizip进程
        var processStartInfo = new ProcessStartInfo
        {
            FileName = bandizipPath,
            Arguments = command,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            UseShellExecute = false
        };

        using var process = new Process();
        process.StartInfo = processStartInfo;

        process.Start();

        Console.WriteLine("开始解压...");

        process.WaitForExit();

        // 输出Bandizip的标准输出和标准错误
        var output = process.StandardOutput.ReadToEnd();
        var error = process.StandardError.ReadToEnd();
        
        if (useBandizipConsoleTool)
        {
            Console.WriteLine("Bandizip Console Tool 输出信息:");
            Console.WriteLine(output);
        
            if (!string.IsNullOrWhiteSpace(error))
            {
                Console.WriteLine("Bandizip出现错误！错误信息:");
                Console.WriteLine(error);
            }
        }

        // 打开输出目录

        if (outputPath != null && openOutputPath)
        {
            // Process.Start("explorer.exe", outputPath);
            Process.Start("explorer.exe", "/n, " + outputPath);
        }
    }

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    private static extern SafeFindHandle FindFirstFile(string lpFileName,
                                                       out WIN32_FIND_DATA lpFindFileData);

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    private static extern bool FindNextFile(SafeFindHandle hFindFile,
                                            out WIN32_FIND_DATA lpFindFileData);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool FindClose(IntPtr hFindFile);

    private static IEnumerable<string> SafeFileSearch(string root, string searchPattern)
    {
        var findHandle = FindFirstFile(Path.Combine(root, searchPattern), out var findData);

        if (!findHandle.IsInvalid)
        {
            do
            {
                yield return Path.Combine(root, findData.cFileName);
            } while (FindNextFile(findHandle, out findData));

            FindClose(findHandle.DangerousGetHandle());
        }

        string[] subDirs;
        try
        {
            subDirs = Directory.GetDirectories(root);
        }
        catch (UnauthorizedAccessException)
        {
            yield break;
        }

        foreach (var subDir in subDirs)
        {
            foreach (var file in SafeFileSearch(subDir, searchPattern))
            {
                yield return file;
            }
        }
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    private struct WIN32_FIND_DATA
    {
        public uint dwFileAttributes;
        public FILETIME ftCreationTime;
        public FILETIME ftLastAccessTime;
        public FILETIME ftLastWriteTime;
        public uint nFileSizeHigh;
        public uint nFileSizeLow;
        public uint dwReserved0;
        public uint dwReserved1;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string cFileName;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 14)]
        public string cAlternateFileName;
    }

    private sealed class SafeFindHandle : SafeHandleZeroOrMinusOneIsInvalid
    {
        private SafeFindHandle() : base(true)
        {
        }

        protected override bool ReleaseHandle()
        {
            return FindClose(handle);
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool FindClose(IntPtr handle);
    }
}
