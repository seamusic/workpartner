using System;
using System.IO;

namespace WorkPartner.Utils
{
    /// <summary>
    /// 命令行参数模型
    /// </summary>
    public class CommandLineArguments
    {
        public string InputPath { get; set; } = string.Empty;
        public string OutputPath { get; set; } = string.Empty;
        public bool Verbose { get; set; } = false;
        public bool CompareMode { get; set; } = false;
        public string CompareOriginalPath { get; set; } = string.Empty;
        public string CompareProcessedPath { get; set; } = string.Empty;
        public bool ShowDetailedDifferences { get; set; } = false;
        public double Tolerance { get; set; } = 0.001;
        public int MaxDifferencesToShow { get; set; } = 10;
        public bool CheckLargeValues { get; set; } = false;
        public double LargeValueThreshold { get; set; } = 4.0;
        public bool DataCorrectionMode { get; set; } = false;
        public string CorrectionOriginalPath { get; set; } = string.Empty;
        public string CorrectionProcessedPath { get; set; } = string.Empty;
        public bool ValidateProcessedMode { get; set; } = false;
        public string ValidateProcessedDirectory { get; set; } = string.Empty;
    }

    /// <summary>
    /// 命令行参数解析器
    /// </summary>
    public static class CommandLineParser
    {
        /// <summary>
        /// 解析命令行参数
        /// </summary>
        /// <param name="args">命令行参数数组</param>
        /// <returns>解析后的参数对象，如果解析失败返回null</returns>
        public static CommandLineArguments? ParseCommandLineArguments(string[] args)
        {
            if (args.Length == 0)
            {
                return null;
            }

            var arguments = new CommandLineArguments();

            for (int i = 0; i < args.Length; i++)
            {
                switch (args[i].ToLower())
                {
                    case "-i":
                    case "--input":
                        if (i + 1 < args.Length)
                        {
                            arguments.InputPath = args[++i];
                        }
                        break;
                    case "-o":
                    case "--output":
                        if (i + 1 < args.Length)
                        {
                            arguments.OutputPath = args[++i];
                        }
                        break;
                    case "-v":
                    case "--verbose":
                        arguments.Verbose = true;
                        break;
                    case "-c":
                    case "--compare":
                        arguments.CompareMode = true;
                        // 比较模式需要两个路径参数
                        if (i + 2 < args.Length)
                        {
                            arguments.CompareOriginalPath = args[++i];
                            arguments.CompareProcessedPath = args[++i];
                        }
                        else if (i + 1 < args.Length)
                        {
                            // 如果只有一个路径，假设是原始路径，输出路径使用默认值
                            arguments.CompareOriginalPath = args[++i];
                            arguments.CompareProcessedPath = Path.Combine(arguments.CompareOriginalPath, "processed");
                        }
                        break;
                    case "--detailed":
                        arguments.ShowDetailedDifferences = true;
                        break;
                    case "--tolerance":
                        if (i + 1 < args.Length && double.TryParse(args[++i], out var tolerance))
                        {
                            arguments.Tolerance = tolerance;
                        }
                        break;
                    case "--max-differences":
                        if (i + 1 < args.Length && int.TryParse(args[++i], out var maxDiff))
                        {
                            arguments.MaxDifferencesToShow = maxDiff;
                        }
                        break;
                    case "--check-large-values":
                        arguments.CheckLargeValues = true;
                        break;
                    case "--large-value-threshold":
                        if (i + 1 < args.Length && double.TryParse(args[++i], out var largeValueThreshold))
                        {
                            arguments.LargeValueThreshold = largeValueThreshold;
                        }
                        break;
                    case "--data-correction":
                        arguments.DataCorrectionMode = true;
                        // 数据修正模式需要两个路径参数
                        if (i + 2 < args.Length)
                        {
                            arguments.CorrectionOriginalPath = args[++i];
                            arguments.CorrectionProcessedPath = args[++i];
                        }
                        else if (i + 1 < args.Length)
                        {
                            // 如果只有一个路径，假设是原始路径，输出路径使用默认值
                            arguments.CorrectionOriginalPath = args[++i];
                            arguments.CorrectionProcessedPath = Path.Combine(arguments.CorrectionOriginalPath, "processed");
                        }
                        break;
                    case "--validate-processed":
                        arguments.ValidateProcessedMode = true;
                        if (i + 1 < args.Length)
                        {
                            arguments.ValidateProcessedDirectory = args[++i];
                        }
                        break;
                    case "-h":
                    case "--help":
                        return null;
                    default:
                        // 检查是否是比较模式的简化语法：-v 原始目录 对比目录
                        if (args[i] == "-v" && i + 2 < args.Length)
                        {
                            arguments.CompareMode = true;
                            arguments.Verbose = true;
                            arguments.ShowDetailedDifferences = true;
                            arguments.CompareOriginalPath = args[++i];
                            arguments.CompareProcessedPath = args[++i];
                        }
                        // 检查是否是大值检查模式：--check-large-values 目录路径
                        else if (args[i] == "--check-large-values" && i + 1 < args.Length)
                        {
                            arguments.CheckLargeValues = true;
                            arguments.InputPath = args[++i]; // 将下一个参数作为要检查的目录路径
                        }
                        // 检查是否是数据修正模式：--data-correction 原目录 处理后目录
                        else if (args[i] == "--data-correction" && i + 2 < args.Length)
                        {
                            arguments.DataCorrectionMode = true;
                            arguments.CorrectionOriginalPath = args[++i];
                            arguments.CorrectionProcessedPath = args[++i];
                        }
                        // 校验已处理目录：--validate-processed 目录路径
                        else if (args[i] == "--validate-processed" && i + 1 < args.Length)
                        {
                            arguments.ValidateProcessedMode = true;
                            arguments.ValidateProcessedDirectory = args[++i];
                        }
                        // 如果没有指定参数，第一个参数作为输入路径
                        else if (string.IsNullOrEmpty(arguments.InputPath))
                        {
                            arguments.InputPath = args[i];
                        }
                        break;
                }
            }

            // 如果没有指定输出路径，使用默认路径
            if (string.IsNullOrEmpty(arguments.OutputPath))
            {
                arguments.OutputPath = Path.Combine(arguments.InputPath, "processed");
            }

            return arguments;
        }

        /// <summary>
        /// 显示使用说明
        /// </summary>
        public static void ShowUsage()
        {
            Console.WriteLine("使用方法:");
            Console.WriteLine("  WorkPartner.exe <输入目录> [选项]");
            Console.WriteLine("  WorkPartner.exe -c <原始目录> <对比目录> [选项]");
            Console.WriteLine("  WorkPartner.exe -v <原始目录> <对比目录>");
            Console.WriteLine("  WorkPartner.exe --check-large-values <目录路径> [选项]");
            Console.WriteLine("");
            Console.WriteLine("参数:");
            Console.WriteLine("  <输入目录>              包含Excel文件的目录路径");
            Console.WriteLine("  <原始目录>              原始Excel文件目录");
            Console.WriteLine("  <对比目录>              已处理的Excel文件目录");
            Console.WriteLine("  <目录路径>              要检查的Excel文件目录");
            Console.WriteLine("");
            Console.WriteLine("选项:");
            Console.WriteLine("  -o, --output <目录>     输出目录路径 (默认: <输入目录>/processed)");
            Console.WriteLine("  -v, --verbose           详细输出模式");
            Console.WriteLine("  -c, --compare           文件比较模式");
            Console.WriteLine("  --detailed              显示详细差异信息");
            Console.WriteLine("  --tolerance <数值>      设置比较容差 (默认: 0.001)");
            Console.WriteLine("  --max-differences <数量> 限制显示差异数量 (默认: 10)");
            Console.WriteLine("  --check-large-values    大值数据检查模式");
            Console.WriteLine("  --large-value-threshold <数值> 设置大值检查阈值 (默认: 4.0)");
            Console.WriteLine("  --data-correction       数据修正模式");
            Console.WriteLine("  --validate-processed <目录> 校验已处理目录的累计逻辑(仅校验, 不修正)");
            Console.WriteLine("  -h, --help              显示此帮助信息");
            Console.WriteLine("");
            Console.WriteLine("支持的文件格式:");
            Console.WriteLine("  ✅ .xlsx (Excel 2007+)");
            Console.WriteLine("  ✅ .xls (Excel 97-2003)");
            Console.WriteLine("");
            Console.WriteLine("示例:");
            Console.WriteLine("  数据处理模式:");
            Console.WriteLine("    WorkPartner.exe C:\\excel\\");
            Console.WriteLine("    WorkPartner.exe ..\\excel\\");
            Console.WriteLine("    WorkPartner.exe C:\\excel\\ -o C:\\output\\ -v");
            Console.WriteLine("");
            Console.WriteLine("  文件比较模式:");
            Console.WriteLine("    WorkPartner.exe -c C:\\original C:\\processed");
            Console.WriteLine("    WorkPartner.exe -v C:\\original C:\\processed");
            Console.WriteLine("    WorkPartner.exe -c C:\\original C:\\processed --detailed --tolerance 0.01");
            Console.WriteLine("");
            Console.WriteLine("  大值检查模式:");
            Console.WriteLine("    WorkPartner.exe --check-large-values C:\\output");
            Console.WriteLine("    WorkPartner.exe --check-large-values C:\\output --large-value-threshold 5.0");
            Console.WriteLine("    WorkPartner.exe --check-large-values C:\\output --large-value-threshold 3.0 -v");
            Console.WriteLine("  数据修正模式:");
            Console.WriteLine("    WorkPartner.exe --data-correction C:\\original C:\\processed");
            Console.WriteLine("");
            Console.WriteLine("  校验已处理目录累计逻辑:");
            Console.WriteLine("    WorkPartner.exe --validate-processed C:\\processed -v --tolerance 0.001");
        }
    }
}
