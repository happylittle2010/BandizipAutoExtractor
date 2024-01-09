using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security.Principal;
using System.Text;
using System.Text.Json;
using Microsoft.Win32.SafeHandles;

namespace BandizipAutoExtractor;

public class Config
{
    /// <summary>
    ///     配置文件的版本
    /// </summary>
    public string Version { get; init; } = "1.0";

    /// <summary>
    ///     Bandizip路径
    /// </summary>
    public string? BandizipPath { get; init; }

    /// <summary>
    ///     解压输出路径
    /// </summary>
    public string? OutputPath { get; init; }

    /// <summary>
    ///     是否每次解压完成后打开输出目录
    /// </summary>
    public bool OpenOutputPath { get; init; }

    /// <summary>
    ///     是否使用Bandizip命令行程序解压
    /// </summary>
    public bool UseBandizipConsoleTool { get; init; }
}

internal class Program
{
    private static void Main(string[] args)
    {
        if (args.Length < 1)
        {
            // Console.WriteLine("欢迎使用Bandizip Auto Extractor！");
            // Console.WriteLine("Bandizip Auto Extractor可以帮助你快速解压文件。");
            // Console.WriteLine();
            // Console.WriteLine(
            //     "Bandizip Auto Extractor在运行时会请求获取管理员权限，你可以为软件配置默认的管理员权限，以避免每次运行时都需要授权。");
            // Console.WriteLine();
            // Console.WriteLine("用法：BandizipAutoExtractor.exe <压缩文件路径>");
            // var downloadsPath =
            //     Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            //                  "Downloads");
            // Console.WriteLine(
            //     $"示例：BandizipAutoExtractor.exe \"{Path.Combine(downloadsPath, "Example.zip")}\"");
            // Console.WriteLine("你也可以通过右键选中压缩文件，并选择\"BandizipAutoExtractor\"来快速解压。");
            // Console.WriteLine("点按任意按键退出...");
            // Console.ReadKey();
            // return;

            args = new[] { @"C:\Users\happy\Downloads\sonarlint-intellij-10.0.0.76954.zip" };
        }

        var configPath =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                         ".BandizipAutoExtractor", "config.json");
        string? bandizipPath = null;
        string? outputPath;
        bool    openOutputPath;
        bool    useBandizipConsoleTool;

        if (File.Exists(configPath))
        {
            try
            {
                var config = JsonSerializer.Deserialize<Config>(File.ReadAllText(configPath));
                if (config == null)
                {
                    Console.WriteLine("配置文件解析失败，请考虑删除配置文件后重启软件。");
                    Console.WriteLine($"配置文件位于: {configPath}。");
                    Console.WriteLine("点按任意按键退出...");
                    Console.ReadKey();
                    return;
                }

                // 验证配置对象的属性
                if (string.IsNullOrEmpty(config.BandizipPath) ||
                    string.IsNullOrEmpty(config.OutputPath))
                {
                    Console.WriteLine("配置文件中的Bandizip路径或输出路径无效，请考虑删除配置文件后重启软件。");
                    Console.WriteLine($"配置文件位于: {configPath}。");
                    Console.WriteLine("点按任意按键退出...");
                    Console.ReadKey();
                    return;
                }

                // 验证Bandizip路径中bandizip.exe和bz.exe是否存在
                var isBandizipDirectoryExisting = Directory.Exists(
                    Path.GetDirectoryName(config.BandizipPath));
                var isBandizipExeExisting =
                    File.Exists(Path.Combine(
                                    Path.GetDirectoryName(config.BandizipPath) ??
                                    string.Empty,
                                    "Bandizip.exe"));
                var isBzExeExisting =
                    File.Exists(Path.Combine(
                                    Path.GetDirectoryName(config.BandizipPath) ??
                                    string.Empty,
                                    "bz.exe"));
                if (!isBandizipDirectoryExisting || !isBandizipExeExisting || !isBzExeExisting)
                {
                    Console.WriteLine("配置文件中的Bandizip路径没有找到Bandizip.exe或bz.exe，请考虑删除配置文件后重启软件。");
                    Console.WriteLine($"配置文件位于: {configPath}。");
                    Console.WriteLine("点按任意按键退出...");
                    Console.ReadKey();
                    return;
                }

                bandizipPath           = config.BandizipPath;
                outputPath             = config.OutputPath;
                openOutputPath         = config.OpenOutputPath;
                useBandizipConsoleTool = config.UseBandizipConsoleTool;
            }
            catch (JsonException ex)
            {
                Console.WriteLine("无法正确解析配置文件 JSON 格式：" + ex.Message);
                Console.WriteLine("请考虑删除配置文件后重启软件。");
                Console.WriteLine($"配置文件位于: {configPath}。");
                Console.WriteLine("点按任意按键退出...");
                Console.ReadKey();
                return;
            }
            catch (IOException ex)
            {
                Console.WriteLine("无法读取配置文件：" + ex.Message);
                Console.WriteLine("请考虑删除配置文件后重启软件。");
                Console.WriteLine($"配置文件位于: {configPath}。");
                Console.WriteLine("点按任意按键退出...");
                Console.ReadKey();
                return;
            }
            catch (Exception ex)
            {
                Console.WriteLine("读取配置文件时发生未知错误：" + ex.Message);
                Console.WriteLine("请考虑删除配置文件后重启软件。");
                Console.WriteLine($"配置文件位于: {configPath}。");
                Console.WriteLine("点按任意按键退出...");
                Console.ReadKey();
                return;
            }
        }
        else
        {
            Console.WriteLine("欢迎使用Bandizip Auto Extractor！");
            Console.WriteLine("Bandizip Auto Extractor可以帮助你快速解压文件。");
            Console.WriteLine();
            Console.WriteLine(
                "Bandizip Auto Extractor在运行时会请求获取管理员权限，你可以为软件配置默认的管理员权限，以避免每次运行时都需要授权。");
            Console.WriteLine();
            Console.WriteLine("首次运行Bandizip Auto Extractor，开始配置...");
            Console.WriteLine("Bandizip Auto Extractor调用Bandizip解压文件，因此需要你的电脑安装有Bandizip。");
            const string bandizipDefaultPath = @"C:\Program Files\Bandizip\Bandizip.exe";
            if (File.Exists(bandizipDefaultPath))
            {
                bandizipPath = bandizipDefaultPath;
            }
            else
            {
                string[] searchPaths = { @"C:\Program Files", @"C:\Program Files (x86)" };
                Console.WriteLine(
                    @"未找到Bandizip默认安装目录，正在从C:\Program Files和C:\Program Files (x86)中搜索...");
                foreach (var searchPath in searchPaths)
                {
                    var searchResult = SafeFileSearch(searchPath, "Bandizip.exe").FirstOrDefault();
                    if (searchResult != null)
                    {
                        bandizipPath = searchResult;
                        break;
                    }
                }
            }

            if (string.IsNullOrEmpty(bandizipPath))
            {
                Console.WriteLine("未找到Bandizip安装目录，正在从所有固定硬盘中搜索...");
                var searchResult = new List<string?>();
                foreach (var drive in DriveInfo.GetDrives())
                {
                    if (drive.DriveType == DriveType.Fixed) // 只搜索固定硬盘
                    {
                        searchResult.AddRange(SafeFileSearch(drive.Name, "Bandizip.exe"));
                    }
                }

                // 如果找到多个Bandizip.exe，让用户选择
                if (searchResult.Count > 1)
                {
                    Console.WriteLine("找到多个Bandizip安装目录，请选择：");
                    var i = 1;
                    foreach (var result in searchResult)
                    {
                        Console.WriteLine($"{i}. {result}");
                        i++;
                    }

                    Console.WriteLine("请输入序号：");
                    var choice = Console.ReadLine();
                    if (int.TryParse(choice, out var index) && index > 0 &&
                        index                                        <= searchResult.Count)
                    {
                        bandizipPath = searchResult.ElementAt(index - 1);
                    }
                    else
                    {
                        Console.WriteLine("输入错误！");
                        Console.WriteLine("点按任意按键退出...");
                        Console.ReadKey();
                        return;
                    }
                }
                else
                {
                    bandizipPath = searchResult.FirstOrDefault();
                }
            }

            if (string.IsNullOrEmpty(bandizipPath))
            {
                Console.WriteLine("未找到Bandizip安装目录！");
                Console.WriteLine("点按任意按键退出...");
                Console.ReadKey();
                return;
            }

            Console.WriteLine($"找到Bandizip安装目录: {bandizipPath}，是否确认？(y/n)");
            if (Console.ReadLine()?.ToLower() != "y")
            {
                Console.WriteLine("未确认Bandizip安装目录！");
                Console.WriteLine("点按任意按键退出...");
                Console.ReadKey();
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
                    Console.WriteLine("输入错误！");
                    Console.WriteLine("点按任意按键退出...");
                    Console.ReadKey();
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
                Console.WriteLine("输入错误！");
                Console.WriteLine("点按任意按键退出...");
                Console.ReadKey();
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
                    Console.WriteLine("输入错误！");
                    Console.WriteLine("点按任意按键退出...");
                    Console.ReadKey();
                    return;
            }

            // 保存Bandizip路径和输出路径到配置文件
            var config = new Config
            {
                BandizipPath   = bandizipPath,
                OutputPath     = outputPath,
                OpenOutputPath = openOutputPath
            };
            Directory.CreateDirectory(Path.GetDirectoryName(configPath) ?? string.Empty);
            File.WriteAllText(configPath, JsonSerializer.Serialize(config));
            Console.WriteLine($"配置已保存到: {configPath}，你可以自行编辑");
        }

        // 拼接Bandizip命令行参数
        var command = $"x -o:{outputPath} -target:auto \"{args[0]}\"";

        // 检查当前进程是否以管理员权限运行
        var identity  = WindowsIdentity.GetCurrent();
        var principal = new WindowsPrincipal(identity);
        if (principal.IsInRole(WindowsBuiltInRole.Administrator))
        {
            Console.WriteLine("当前软件以管理员权限运行！");
        }

        try
        {
            var processStartInfo = new ProcessStartInfo
            {
                FileName               = bandizipPath,
                Arguments              = command,
                RedirectStandardOutput = true,
                RedirectStandardError  = true,
                RedirectStandardInput  = true, // 重定向标准输入
                UseShellExecute        = false,
                CreateNoWindow         = true
            };
            using (var process = Process.Start(processStartInfo))
            {
                if (process != null)
                {
                    Console.WriteLine("开始解压...");

                    // 启动新线程来读取输出
                    void ThreadStart()
                    {
                        while (!process.StandardOutput.EndOfStream)
                        {
                            var line = process.StandardOutput.ReadLine();
                            Console.WriteLine(line);
                        }
                    }

                    var outputThread = new Thread(ThreadStart);
                    outputThread.Start();

                    // // 启动新线程来读取错误
                    // void Start()
                    // {
                    //     while (!process.StandardError.EndOfStream)
                    //     {
                    //         var line = process.StandardError.ReadLine();
                    //         Console.Error.WriteLine(line);
                    //     }
                    // }
                    //
                    // var errorThread = new Thread(Start);
                    // errorThread.Start();

                    // 启动新线程处理标准输入
                    void inputThreadDelegate()
                    {
                        var inputBuffer = new StringBuilder();
                        while (!process.HasExited)
                        {
                            if (Console.KeyAvailable)
                            {
                                var key = Console.ReadKey(true);
                                if (key.Key == ConsoleKey.Enter)
                                {
                                    process.StandardInput.WriteLine(inputBuffer.ToString());
                                    inputBuffer.Clear();
                                }
                                else if (key.Key == ConsoleKey.Backspace && inputBuffer.Length > 0)
                                {
                                    inputBuffer.Length--;
                                }
                                else if (key.KeyChar != '\u0000')
                                {
                                    inputBuffer.Append(key.KeyChar);
                                }
                            }

                            Thread.Sleep(50); // To reduce CPU usage
                        }
                    }

                    var inputThread = new Thread(inputThreadDelegate);
                    inputThread.Start();

                    // 等待进程退出
                    process.WaitForExit();

                    // 确保所有的输出都已经被打印到控制台
                    outputThread.Join();
                    // errorThread.Join();
                    inputThread.Join();
                }
            }
        }
        catch (Win32Exception ex)
        {
            Console.WriteLine($"启动Bandizip进程时出错: {ex.Message}");
            Console.WriteLine("点按任意按键退出...");
            Console.ReadKey();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"启动Bandizip进程时发生未知错误: {ex.Message}");
            Console.WriteLine("点按任意按键退出...");
            Console.ReadKey();
        }

        // 打开输出目录

        if (outputPath != null && openOutputPath)
        {
            // Process.Start("explorer.exe", outputPath);
            Process.Start("explorer.exe", "/n, " + outputPath);
        }
    }

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    private static extern SafeFindHandle FindFirstFile(string              lpFileName,
                                                       out WIN32_FIND_DATA lpFindFileData);

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    private static extern bool FindNextFile(SafeFindHandle      hFindFile,
                                            out WIN32_FIND_DATA lpFindFileData);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool FindClose(IntPtr hFindFile);

    private static IEnumerable<string?> SafeFileSearch(string root, string searchPattern)
    {
        // 检查参数合法性
        if (string.IsNullOrEmpty(root) || string.IsNullOrEmpty(searchPattern))
        {
            Console.WriteLine("参数错误...");
            yield return null;
        }

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
        catch (UnauthorizedAccessException ex)
        {
            Console.WriteLine($"搜索时出现未授权访问异常：{ex.Message}，继续搜索其他目录...");
            yield break;
        }
        catch (DirectoryNotFoundException ex)
        {
            Console.WriteLine($"搜索时出现目录不存在异常：{ex.Message}，继续搜索其他目录...");
            yield break;
        }
        catch (PathTooLongException ex)
        {
            Console.WriteLine($"搜索时出现路径过长异常：{ex.Message}，继续搜索其他目录...");
            yield break;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"搜索时出现预期之外的异常：{ex.Message}，继续搜索其他目录...");
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
        public uint     dwFileAttributes;
        public FILETIME ftCreationTime;
        public FILETIME ftLastAccessTime;
        public FILETIME ftLastWriteTime;
        public uint     nFileSizeHigh;
        public uint     nFileSizeLow;
        public uint     dwReserved0;
        public uint     dwReserved1;

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
