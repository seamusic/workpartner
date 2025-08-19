using DataFixter.Services;

namespace DataFixter.Models;

/// <summary>
/// 数据完整性问题
/// </summary>
public class DataIntegrityIssue
{
    /// <summary>
    /// 点名
    /// </summary>
    public string PointName { get; set; } = string.Empty;

    /// <summary>
    /// 问题类型
    /// </summary>
    public DataIntegrityIssueType IssueType { get; set; }

    /// <summary>
    /// 问题描述
    /// </summary>
    public string Description { get; set; } = string.Empty;

    /// <summary>
    /// 严重程度
    /// </summary>
    public DataIntegrityIssueSeverity Severity { get; set; }
}