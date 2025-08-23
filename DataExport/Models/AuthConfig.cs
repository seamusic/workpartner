namespace DataExport.Models
{
    /// <summary>
    /// 认证配置类
    /// </summary>
    public class AuthConfig
    {
        /// <summary>
        /// 基础URL
        /// </summary>
        public string BaseUrl { get; set; } = "http://localhost:20472";

        /// <summary>
        /// Cookie字符串
        /// </summary>
        public string CookieString { get; set; } = string.Empty;

        /// <summary>
        /// 用户代理
        /// </summary>
        public string UserAgent { get; set; } = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36";

        /// <summary>
        /// 接受语言
        /// </summary>
        public string AcceptLanguage { get; set; } = "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6";

        /// <summary>
        /// 接受类型
        /// </summary>
        public string Accept { get; set; } = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,application/ms-excel,application/octet-stream";

        /// <summary>
        /// 连接类型
        /// </summary>
        public string Connection { get; set; } = "keep-alive";

        /// <summary>
        /// 升级不安全请求
        /// </summary>
        public string UpgradeInsecureRequests { get; set; } = "1";

        /// <summary>
        /// 安全获取目标
        /// </summary>
        public string SecFetchDest { get; set; } = "download";

        /// <summary>
        /// 安全获取模式
        /// </summary>
        public string SecFetchMode { get; set; } = "navigate";

        /// <summary>
        /// 安全获取站点
        /// </summary>
        public string SecFetchSite { get; set; } = "same-origin";

        /// <summary>
        /// 安全获取用户
        /// </summary>
        public string SecFetchUser { get; set; } = "?1";

        /// <summary>
        /// 检查Cookie是否有效
        /// </summary>
        /// <returns>是否有效</returns>
        public bool IsCookieValid()
        {
            return !string.IsNullOrWhiteSpace(CookieString) && 
                   CookieString.Contains("Qianchen_ADMS_V7_Token=");
        }

        /// <summary>
        /// 获取完整的导出URL
        /// </summary>
        /// <returns>完整URL</returns>
        public string GetExportUrl()
        {
            return $"{BaseUrl}/QC_FoundationPit/ResultsQuery/ExportDataList";
        }

        /// <summary>
        /// 获取引用页面URL
        /// </summary>
        /// <param name="projectId">项目ID</param>
        /// <param name="projectCode">项目代码</param>
        /// <param name="projectName">项目名称</param>
        /// <returns>引用页面URL</returns>
        public string GetRefererUrl(string projectId, string projectCode, string projectName)
        {
            var encodedProjectName = Uri.EscapeDataString(projectName);
            return $"{BaseUrl}/QC_FoundationPit/ResultsQuery/Index?projectidfordetail={projectId}&projectId={projectId}&projectcodefordetail={projectCode}&parentID=&subwayID=c05000e5-bd0b-11eb-ab60-b025aa3442a2&type=3&projectnamefordetail={encodedProjectName}&sidetype=&intype=true&openarg=undefined&dataCode=";
        }
    }
}
