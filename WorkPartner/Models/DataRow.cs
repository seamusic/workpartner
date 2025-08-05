namespace WorkPartner.Models
{
    /// <summary>
    /// 数据行模型
    /// </summary>
    public class DataRow
    {
        /// <summary>
        /// 数据名称（B列）
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// 数据值列表（D-I列）
        /// </summary>
        public List<double?> Values { get; set; } = new List<double?>();

        /// <summary>
        /// 在Excel中的行索引（从1开始）
        /// </summary>
        public int RowIndex { get; set; }

        /// <summary>
        /// 是否包含缺失数据
        /// </summary>
        public bool HasMissingData => Values.Any(v => !v.HasValue);

        /// <summary>
        /// 缺失数据数量
        /// </summary>
        public int MissingDataCount => Values.Count(v => !v.HasValue);

        /// <summary>
        /// 有效数据数量
        /// </summary>
        public int ValidDataCount => Values.Count(v => v.HasValue);

        /// <summary>
        /// 数据完整性百分比
        /// </summary>
        public double CompletenessPercentage => Values.Count > 0 ? (double)ValidDataCount / Values.Count * 100 : 0;

        /// <summary>
        /// 平均值（仅计算有效数据）
        /// </summary>
        public double? AverageValue
        {
            get
            {
                var validValues = Values.Where(v => v.HasValue).ToList();
                return validValues.Any() ? validValues.Average() : null;
            }
        }

        /// <summary>
        /// 最小值（仅计算有效数据）
        /// </summary>
        public double? MinValue
        {
            get
            {
                var validValues = Values.Where(v => v.HasValue).ToList();
                return validValues.Any() ? validValues.Min() : null;
            }
        }

        /// <summary>
        /// 最大值（仅计算有效数据）
        /// </summary>
        public double? MaxValue
        {
            get
            {
                var validValues = Values.Where(v => v.HasValue).ToList();
                return validValues.Any() ? validValues.Max() : null;
            }
        }

        /// <summary>
        /// 数据范围（最大值 - 最小值）
        /// </summary>
        public double? DataRange
        {
            get
            {
                if (MinValue.HasValue && MaxValue.HasValue)
                {
                    return MaxValue.Value - MinValue.Value;
                }
                return null;
            }
        }

        /// <summary>
        /// 是否所有数据都缺失
        /// </summary>
        public bool IsAllMissing => Values.All(v => !v.HasValue);

        /// <summary>
        /// 是否所有数据都有效
        /// </summary>
        public bool IsAllValid => Values.All(v => v.HasValue);

        /// <summary>
        /// 获取指定索引的数据值
        /// </summary>
        /// <param name="index">数据索引</param>
        /// <returns>数据值</returns>
        public double? GetValue(int index)
        {
            return index >= 0 && index < Values.Count ? Values[index] : null;
        }

        /// <summary>
        /// 设置指定索引的数据值
        /// </summary>
        /// <param name="index">数据索引</param>
        /// <param name="value">数据值</param>
        public void SetValue(int index, double? value)
        {
            if (index >= 0 && index < Values.Count)
            {
                Values[index] = value;
            }
        }

        /// <summary>
        /// 添加数据值
        /// </summary>
        /// <param name="value">数据值</param>
        public void AddValue(double? value)
        {
            Values.Add(value);
        }

        /// <summary>
        /// 清空所有数据值
        /// </summary>
        public void ClearValues()
        {
            Values.Clear();
        }

        public override string ToString()
        {
            return $"Row {RowIndex}: {Name} ({ValidDataCount}/{Values.Count} valid)";
        }
    }
} 