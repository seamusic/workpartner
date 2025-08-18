namespace DataFixter.Models
{
    /// <summary>
    /// 数据方向枚举
    /// </summary>
    public enum DataDirection
    {
        /// <summary>
        /// X方向
        /// </summary>
        X,
        
        /// <summary>
        /// Y方向
        /// </summary>
        Y,
        
        /// <summary>
        /// Z方向
        /// </summary>
        Z
    }

    /// <summary>
    /// 验证状态枚举
    /// </summary>
    public enum ValidationStatus
    {
        /// <summary>
        /// 未验证
        /// </summary>
        NotValidated,
        
        /// <summary>
        /// 验证通过
        /// </summary>
        Valid,
        
        /// <summary>
        /// 验证失败
        /// </summary>
        Invalid,
        
        /// <summary>
        /// 需要修正
        /// </summary>
        NeedsAdjustment,

        /// <summary>
        /// 可以修正
        /// </summary>
        CanAdjustment,
    }

    /// <summary>
    /// 调整类型枚举
    /// </summary>
    public enum AdjustmentType
    {
        /// <summary>
        /// 未调整
        /// </summary>
        None,
        
        /// <summary>
        /// 调整本期变化量
        /// </summary>
        CurrentPeriod,
        
        /// <summary>
        /// 调整累计变化量
        /// </summary>
        Cumulative,
        
        /// <summary>
        /// 调整日变化量
        /// </summary>
        Daily
    }

    /// <summary>
    /// 统计维度枚举
    /// </summary>
    public enum StatisticsDimension
    {
        /// <summary>
        /// 按点名统计
        /// </summary>
        ByPointName,
        
        /// <summary>
        /// 按文件统计
        /// </summary>
        ByFile,
        
        /// <summary>
        /// 按调整类型统计
        /// </summary>
        ByAdjustmentType
    }
}
