using System;

namespace WorkPartner.Utils
{
    /// <summary>
    /// 处理流程的可配置选项（从命令行或配置绑定）
    /// </summary>
    public class ProcessingOptions
    {
        public string InputPath { get; set; } = string.Empty;
        public string OutputPath { get; set; } = string.Empty;

        // 文件比较选项
        public bool ShowDetailedDifferences { get; set; } = true;
        public double Tolerance { get; set; } = 0.001;
        public int MaxDifferencesToShow { get; set; } = 10;
    }
}


