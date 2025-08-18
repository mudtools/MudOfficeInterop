//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


using System.Drawing;

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel Borders 集合对象的二次封装实现类
/// 负责对 Microsoft.Office.Interop.Excel.Borders 对象的安全访问和资源管理
/// </summary>
internal class ExcelBorders : IExcelBorders
{
    /// <summary>
    /// 底层的 COM Borders 集合对象
    /// </summary>
    private MsExcel.Borders _borders;

    /// <summary>
    /// 标记对象是否已被释放
    /// </summary>
    private bool _disposedValue;

    #region 构造函数和释放

    /// <summary>
    /// 初始化 ExcelBorders 实例
    /// </summary>
    /// <param name="borders">底层的 COM Borders 集合对象</param>
    internal ExcelBorders(MsExcel.Borders borders)
    {
        _borders = borders;
        _disposedValue = false;
    }

    /// <summary>
    /// 释放资源的核心方法
    /// </summary>
    /// <param name="disposing">是否为显式释放</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            try
            {
                // 释放所有子边框对象
                foreach (object? item in _borders)
                {
                    ExcelBorder? border = item as ExcelBorder;
                    border?.Dispose();
                }
                // 释放底层COM对象
                if (_borders != null)
                    Marshal.ReleaseComObject(_borders);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _borders = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 实现 IDisposable 接口的释放方法
    /// </summary>
    public void Dispose() => Dispose(true);

    #endregion

    #region 基础属性
    /// <summary>
    /// 应用到全局。
    /// </summary>
    public bool ApplyToAll { get; set; }

    public Dictionary<XlBordersIndex, IExcelCellFormat> CustomBorders { get; set; } = [];

    /// <summary>
    /// 获取边框集合中的边框数量
    /// </summary>
    public int Count => _borders?.Count ?? 0;

    /// <summary>
    /// 获取指定类型的边框对象
    /// </summary>
    /// <param name="borderType">边框类型</param>
    /// <returns>边框对象</returns>
    public IExcelBorder this[XlBordersIndex borderType]
    {
        get
        {
            if (_borders == null)
                return null;

            try
            {
                var bt = (MsExcel.XlBordersIndex)borderType;
                var border = _borders[bt];
                return border != null ? new ExcelBorder(border) : null;
            }
            catch
            {
                return null;
            }
        }
    }

    /// <summary>
    /// 获取边框集合所在的父对象
    /// </summary>
    public object Parent => _borders?.Parent;

    /// <summary>
    /// 获取边框集合所在的Application对象
    /// </summary>
    public IExcelApplication Application
    {
        get
        {
            var application = _borders?.Application as MsExcel.Application;
            return application != null ? new ExcelApplication(application) : null;
        }
    }

    public XlLineStyle LineStyle
    {
        get => (XlLineStyle)_borders?.LineStyle;
        set
        {
            _borders.LineStyle = (MsExcel.XlLineStyle)value;
            SetLineStyle(LineStyle);
        }
    }

    public XlBorderWeight Weight
    {
        get => (XlBorderWeight)_borders?.Weight;
        set
        {
            _borders.Weight = (MsExcel.XlBorderWeight)value;
            SetWeight(Convert.ToInt32(value));
        }
    }

    /// <summary>
    /// 获取或设置字体颜色（RGB值）
    /// </summary>
    public Color Color { get; set; }
    #endregion   

    #region 查找和筛选

    /// <summary>
    /// 根据线条样式查找边框
    /// </summary>
    /// <param name="lineStyle">线条样式</param>
    /// <returns>匹配的边框数组</returns>
    public IExcelBorder[] FindByLineStyle(int lineStyle)
    {
        if (_borders == null || Count == 0)
            return [];

        List<IExcelBorder> result = [];

        foreach (object? item in _borders)
        {
            if (item is MsExcel.Border border && (int)border.LineStyle == lineStyle)
            {
                try
                {
                    IExcelBorder excelBorder = new ExcelBorder(border);
                    result.Add(excelBorder);
                }
                catch
                {
                    // 忽略单个边框访问异常
                }
            }
        }
        return result.ToArray();
    }

    /// <summary>
    /// 根据颜色查找边框
    /// </summary>
    /// <param name="color">边框颜色</param>
    /// <returns>匹配的边框数组</returns>
    public IExcelBorder[] FindByColor(int color)
    {
        if (_borders == null || Count == 0)
            return [];

        List<IExcelBorder> result = [];

        foreach (object? item in _borders)
        {
            if (item is MsExcel.Border border && (int)border.Color == color)
            {
                try
                {
                    IExcelBorder excelBorder = new ExcelBorder(border);
                    result.Add(excelBorder);
                }
                catch
                {
                    // 忽略单个边框访问异常
                }
            }
        }
        return result.ToArray();
    }

    /// <summary>
    /// 根据粗细查找边框
    /// </summary>
    /// <param name="weight">边框粗细</param>
    /// <returns>匹配的边框数组</returns>
    public IExcelBorder[] FindByWeight(int weight)
    {
        if (_borders == null || Count == 0)
            return [];

        List<IExcelBorder> result = [];

        foreach (object? item in _borders)
        {
            if (item is MsExcel.Border border && (int)border.Weight == weight)
            {
                IExcelBorder excelBorder = new ExcelBorder(border);
                result.Add(excelBorder);
            }
        }
        return result.ToArray();
    }

    #endregion

    #region 格式设置

    /// <summary>
    /// 设置所有边框的线条样式
    /// </summary>
    /// <param name="lineStyle">线条样式</param>
    public void SetLineStyle(XlLineStyle lineStyle, int weight = 1)
    {
        if (_borders == null || Count == 0)
            return;

        try
        {
            _borders.LineStyle = (MsExcel.XlLineStyle)lineStyle;
            _borders.Weight = weight;
            _borders[MsExcel.XlBordersIndex.xlEdgeLeft].LineStyle = (MsExcel.XlLineStyle)lineStyle;
            _borders[MsExcel.XlBordersIndex.xlEdgeRight].LineStyle = (MsExcel.XlLineStyle)lineStyle;
            _borders[MsExcel.XlBordersIndex.xlEdgeTop].LineStyle = (MsExcel.XlLineStyle)lineStyle;
            _borders[MsExcel.XlBordersIndex.xlEdgeBottom].LineStyle = (MsExcel.XlLineStyle)lineStyle;
        }
        catch
        {
            // 忽略设置过程中的异常
        }
    }

    /// <summary>
    /// 设置所有边框的颜色
    /// </summary>
    /// <param name="color">边框颜色</param>
    public void SetColor(Color color)
    {
        if (_borders == null || Count == 0)
            return;

        try
        {
            foreach (object? item in _borders)
            {
                if (item is MsExcel.Border border)
                {
                    border.Color = (int)((color.B << 16) | (color.G << 8) | color.R);
                }
            }
        }
        catch
        {
            // 忽略设置过程中的异常
        }
    }

    /// <summary>
    /// 设置所有边框的粗细
    /// </summary>
    /// <param name="weight">边框粗细</param>
    public void SetWeight(int weight)
    {
        if (_borders == null || Count == 0)
            return;

        try
        {
            foreach (object? item in _borders)
            {
                if (item is MsExcel.Border border)
                {
                    border.Weight = weight;
                }
            }
        }
        catch
        {
            // 忽略设置过程中的异常
        }
    }

    /// <summary>
    /// 统一所有边框的格式
    /// </summary>
    /// <param name="lineStyle">线条样式</param>
    /// <param name="color">边框颜色</param>
    /// <param name="weight">边框粗细</param>
    public void UniformFormat(Color color, XlLineStyle lineStyle = XlLineStyle.xlLineStyleNone, int weight = 2)
    {
        if (_borders == null || Count == 0)
            return;

        try
        {
            foreach (object? item in _borders)
            {
                if (item is MsExcel.Border border)
                {
                    border.LineStyle = lineStyle;
                    border.Color = color;
                    border.Weight = weight;
                }
            }
        }
        catch
        {
            // 忽略统一格式过程中的异常
        }
    }

    /// <summary>
    /// 复制边框格式
    /// </summary>
    /// <param name="sourceBorder">源边框</param>
    /// <param name="applyToAll">是否应用到所有边框</param>
    public void CopyFormat(IExcelBorder sourceBorder, bool applyToAll = false)
    {
        if (_borders == null || sourceBorder == null)
            return;

        try
        {
            if (applyToAll)
            {
                UniformFormat(
                    sourceBorder.Color,
                    sourceBorder.LineStyle,
                    sourceBorder.Weight
                );
            }
            else
            {
                // 应用到第一个边框
                if (Count > 0)
                {
                    foreach (object? item in _borders)
                    {
                        if (item is MsExcel.Border firstBorder)
                        {
                            firstBorder.LineStyle = sourceBorder.LineStyle;
                            firstBorder.Color = sourceBorder.Color;
                            firstBorder.Weight = sourceBorder.Weight;
                            break;
                        }
                    }
                }
            }
        }
        catch
        {
            // 忽略复制格式过程中的异常
        }
    }

    /// <summary>
    /// 应用预设边框样式
    /// </summary>
    /// <param name="presetStyle">预设样式类型</param>
    public void ApplyPresetStyle(int presetStyle)
    {
        if (_borders == null || Count == 0)
            return;

        try
        {
            switch (presetStyle)
            {
                case 1: // 实线边框
                    UniformFormat(Color.Black, XlLineStyle.xlContinuous, 2); // xlContinuous, black, xlThin
                    break;
                case 2: // 虚线边框
                    UniformFormat(Color.Black, XlLineStyle.xlDash, 2); // xlDash, black, xlThin
                    break;
                case 3: // 点线边框
                    UniformFormat(Color.Black, XlLineStyle.xlContinuous, 2); // xlDot, black, xlThin
                    break;
                case 4: // 双线边框
                    UniformFormat(Color.Black, XlLineStyle.xlDouble, 3); // xlDouble, black, xlMedium
                    break;
                case 5: // 粗边框
                    UniformFormat(Color.Black, XlLineStyle.xlContinuous, 4); // xlContinuous, black, xlThick
                    break;
                default:
                    // 默认样式
                    UniformFormat(Color.Black, XlLineStyle.xlContinuous, 2);
                    break;
            }
        }
        catch
        {
            // 忽略应用预设样式过程中的异常
        }
    }

    #endregion    

    #region 导出和导入

    /// <summary>
    /// 导出所有边框到文件
    /// </summary>
    /// <param name="filename">导出文件路径</param>
    /// <returns>是否导出成功</returns>
    public bool ExportToFile(string filename)
    {
        if (_borders == null || Count == 0 || string.IsNullOrEmpty(filename))
            return false;

        try
        {
            using StreamWriter writer = new(filename, false, System.Text.Encoding.UTF8);
            writer.WriteLine("Excel Borders Export");
            writer.WriteLine("====================");
            writer.WriteLine($"Export Date: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            writer.WriteLine($"Total Borders: {Count}");
            writer.WriteLine();
            int i = 0;
            foreach (object? item in _borders)
            {
                if (item is MsExcel.Border border)
                {
                    try
                    {
                        if (border != null)
                        {
                            writer.WriteLine($"Border #{i}");
                            writer.WriteLine($"LineStyle: {border.LineStyle}");
                            writer.WriteLine($"Weight: {border.Weight}");
                            writer.WriteLine($"Color: {border.Color}");
                            writer.WriteLine($"ThemeColor: {border.ThemeColor}");
                            writer.WriteLine($"TintAndShade: {border.TintAndShade}");
                            writer.WriteLine(new string('-', 40));
                            writer.WriteLine();
                        }
                    }
                    catch
                    {
                        // 忽略单个边框导出异常
                    }
                }
                i++;
            }
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// 从文件导入边框
    /// </summary>
    /// <param name="filename">导入文件路径</param>
    /// <returns>成功导入的边框数量</returns>
    public int ImportFromFile(string filename)
    {
        if (_borders == null || string.IsNullOrEmpty(filename))
            return 0;

        // 注意：Excel Borders不支持直接导入
        // 这里提供一个示例实现框架
        return 0;
    }

    /// <summary>
    /// 获取所有边框的信息
    /// </summary>
    /// <returns>边框信息数组</returns>
    public BorderInfo[] GetAllBorderInfo()
    {
        if (_borders == null || Count == 0)
            return new BorderInfo[0];

        var result = new System.Collections.Generic.List<BorderInfo>();
        int i = 0;
        foreach (object? item in _borders)
        {
            if (item is MsExcel.Border border)
            {
                try
                {
                    var info = new BorderInfo
                    {
                        Index = i,
                        LineStyle = (int)border.LineStyle,
                        Weight = (int)border.Weight,
                        Color = (int)border.Color,
                        ThemeColor = (int)border.ThemeColor,
                        TintAndShade = (int)border.TintAndShade,
                        Created = DateTime.Now,
                        Modified = DateTime.Now,
                        Locked = false, // 需要从父对象获取
                        Worksheet = "", // 需要从父对象获取
                        RangeAddress = "" // 需要从父对象获取
                    };
                    result.Add(info);
                }
                catch
                {
                    // 忽略单个边框访问异常
                }
            }
            i++;
        }

        return result.ToArray();
    }
    #endregion

    #region 统计和分析    

    /// <summary>
    /// 获取线条样式统计
    /// </summary>
    /// <returns>线条样式统计信息数组</returns>
    public LineStyleStatistics[] GetLineStyleStatistics()
    {
        if (_borders == null || Count == 0)
            return [];

        var styleCount = new System.Collections.Generic.Dictionary<int, int>();
        int i = 0;
        foreach (object? item in _borders)
        {
            if (item is MsExcel.Border border)
            {
                try
                {
                    int style = (int)border.LineStyle;

                    if (styleCount.ContainsKey(style))
                        styleCount[style]++;
                    else
                        styleCount[style] = 1;
                }
                catch
                {
                    // 忽略单个边框访问异常
                }
            }
            i++;
        }
        List<LineStyleStatistics> result = [];
        int totalCount = Count;

        foreach (KeyValuePair<int, int> kvp in styleCount)
        {
            result.Add(new LineStyleStatistics
            {
                LineStyle = kvp.Key,
                Count = kvp.Value,
                Percentage = totalCount > 0 ? (double)kvp.Value / totalCount * 100 : 0,
                StyleName = GetLineStyleName(kvp.Key),
                IsSolid = kvp.Key == 1,      // xlContinuous
                IsDash = kvp.Key == -4115,   // xlDash
                IsDot = kvp.Key == -4118,    // xlDot
                IsDouble = kvp.Key == -4119  // xlDouble
            });
        }

        return result.ToArray();
    }

    /// <summary>
    /// 获取颜色统计
    /// </summary>
    /// <returns>颜色统计信息</returns>
    public BorderColorStatistics GetColorStatistics()
    {
        var stats = new BorderColorStatistics
        {
            Color = 0,
            ColorName = "Unknown",
            Count = 0,
            Percentage = 0,
            IsPrimary = false,
            IsCustom = false,
            Brightness = 0,
            Saturation = 0,
            Hue = 0
        };

        if (_borders == null || Count == 0)
            return stats;

        Dictionary<int, int> colorCount = [];

        int i = 0;
        foreach (object? item in _borders)
        {
            if (item is MsExcel.Border border)
            {
                try
                {
                    int color = (int)border.Color;

                    if (colorCount.ContainsKey(color))
                        colorCount[color]++;
                    else
                        colorCount[color] = 1;
                }
                catch
                {
                    // 忽略单个边框访问异常
                }
            }
            i++;
        }

        // 返回最常见的颜色统计
        int maxCount = 0;
        int mostCommonColor = 0;

        foreach (var kvp in colorCount)
        {
            if (kvp.Value > maxCount)
            {
                maxCount = kvp.Value;
                mostCommonColor = kvp.Key;
            }
        }

        stats.Color = mostCommonColor;
        stats.ColorName = GetColorName(mostCommonColor);
        stats.Count = maxCount;
        stats.Percentage = Count > 0 ? (double)maxCount / Count * 100 : 0;
        stats.IsPrimary = IsPrimaryColor(mostCommonColor);
        stats.IsCustom = IsCustomColor(mostCommonColor);
        stats.Brightness = GetBrightness(mostCommonColor);
        stats.Saturation = GetSaturation(mostCommonColor);
        stats.Hue = GetHue(mostCommonColor);

        return stats;
    }

    /// <summary>
    /// 获取粗细统计
    /// </summary>
    /// <returns>粗细统计信息</returns>
    public WeightStatistics GetWeightStatistics()
    {
        WeightStatistics stats = new WeightStatistics
        {
            Weight = 0,
            Count = 0,
            Percentage = 0,
            WeightName = "Unknown",
            IsThin = false,
            IsMedium = false,
            IsThick = false,
            IsExtraThick = false
        };

        if (_borders == null || Count == 0)
            return stats;

        Dictionary<int, int> weightCount = [];

        int i = 0;
        foreach (object? item in _borders)
        {
            if (item is MsExcel.Border border)
            {
                try
                {
                    int weight = (int)border.Weight;

                    if (weightCount.ContainsKey(weight))
                        weightCount[weight]++;
                    else
                        weightCount[weight] = 1;
                }
                catch
                {
                    // 忽略单个边框访问异常
                }
            }
            i++;
        }

        // 返回最常见的粗细统计
        int maxCount = 0;
        int mostCommonWeight = 0;

        foreach (KeyValuePair<int, int> kvp in weightCount)
        {
            if (kvp.Value > maxCount)
            {
                maxCount = kvp.Value;
                mostCommonWeight = kvp.Key;
            }
        }

        stats.Weight = mostCommonWeight;
        stats.Count = maxCount;
        stats.Percentage = Count > 0 ? (double)maxCount / Count * 100 : 0;
        stats.WeightName = GetWeightName(mostCommonWeight);
        stats.IsThin = mostCommonWeight == 1;      // xlThin
        stats.IsMedium = mostCommonWeight == 2;    // xlMedium
        stats.IsThick = mostCommonWeight == 3;     // xlThick
        stats.IsExtraThick = mostCommonWeight == 4; // xlThick
        return stats;
    }
    #endregion

    #region 高级功能
    /// <summary>
    /// 重置边框为默认值
    /// </summary>
    public void Reset()
    {
        if (_borders == null) return;

        try
        {
            UniformFormat(Color.Black, XlLineStyle.xlContinuous, 2); // xlContinuous, black, xlThin, visible
        }
        catch
        {
            // 忽略重置过程中的异常
        }
    }

    /// <summary>
    /// 验证边框设置
    /// </summary>
    /// <returns>验证结果</returns>
    public BorderValidationResult Validate()
    {
        BorderValidationResult result = new()
        {
            IsValid = true,
            ErrorMessage = "",
            SuggestedFix = "",
            ValidBorderType = true,
            ValidLineStyle = true,
            ValidColor = true,
            ValidWeight = true,
            ValidPosition = true,
            ValidSize = true,
            OutOfBounds = false,
            Overlapping = false,
            NegativeDimensions = false
        };

        if (_borders == null)
        {
            result.IsValid = false;
            result.ErrorMessage = "边框集合为空";
            return result;
        }

        int i = 0;
        foreach (object? item in _borders)
        {
            if (item is MsExcel.Border border)
            {
                try
                {
                    IExcelBorder excelBorder = new ExcelBorder(border);
                    var borderResult = excelBorder.Validate();
                    if (!borderResult.IsValid)
                    {
                        result.IsValid = false;
                        result.ErrorMessage += $"边框{i}: {borderResult.ErrorMessage}; ";
                    }
                }
                catch
                {
                    // 忽略单个边框访问异常
                }
            }
            i++;
        }
        return result;
    }
    #endregion

    #region 私有辅助方法

    /// <summary>
    /// 获取线条样式名称
    /// </summary>
    /// <param name="lineStyle">线条样式</param>
    /// <returns>线条样式名称</returns>
    private string GetLineStyleName(int lineStyle)
    {
        return lineStyle switch
        {
            1 => "实线",
            -4115 => "虚线",
            -4118 => "点线",
            -4119 => "双线",
            -4107 => "点划线",
            -4114 => "双点划线",
            -4117 => "斜点划线",
            -4121 => "细点划线",
            -4142 => "自动",
            -4166 => "无",
            _ => "未知",
        };
    }

    /// <summary>
    /// 获取颜色名称
    /// </summary>
    /// <param name="color">颜色值</param>
    /// <returns>颜色名称</returns>
    private string GetColorName(int color)
    {
        return color switch
        {
            0 => "黑色",
            255 => "红色",
            65280 => "绿色",
            16711680 => "蓝色",
            16777215 => "白色",
            16776960 => "黄色",
            16711935 => "洋红色",
            65535 => "青色",
            _ => $"RGB({color & 0xFF}, {(color >> 8) & 0xFF}, {(color >> 16) & 0xFF})",
        };
    }

    /// <summary>
    /// 判断是否为主要颜色
    /// </summary>
    /// <param name="color">颜色值</param>
    /// <returns>是否为主要颜色</returns>
    private bool IsPrimaryColor(int color)
    {
        return color == 0 || color == 255 || color == 65280 || color == 16711680 ||
               color == 16777215 || color == 16776960 || color == 16711935 || color == 65535;
    }

    /// <summary>
    /// 判断是否为自定义颜色
    /// </summary>
    /// <param name="color">颜色值</param>
    /// <returns>是否为自定义颜色</returns>
    private bool IsCustomColor(int color)
    {
        return !IsPrimaryColor(color);
    }

    /// <summary>
    /// 获取亮度值
    /// </summary>
    /// <param name="color">颜色值</param>
    /// <returns>亮度值</returns>
    private double GetBrightness(int color)
    {
        int r = color & 0xFF;
        int g = (color >> 8) & 0xFF;
        int b = (color >> 16) & 0xFF;
        return (r * 0.299 + g * 0.587 + b * 0.114) / 255.0;
    }

    /// <summary>
    /// 获取饱和度值
    /// </summary>
    /// <param name="color">颜色值</param>
    /// <returns>饱和度值</returns>
    private double GetSaturation(int color)
    {
        int r = color & 0xFF;
        int g = (color >> 8) & 0xFF;
        int b = (color >> 16) & 0xFF;

        int max = Math.Max(Math.Max(r, g), b);
        int min = Math.Min(Math.Min(r, g), b);

        if (max == min) return 0;

        int diff = max - min;
        int sum = max + min;
        return sum <= 255 ? (double)diff / sum : (double)diff / (510 - sum);
    }

    /// <summary>
    /// 获取色调值
    /// </summary>
    /// <param name="color">颜色值</param>
    /// <returns>色调值</returns>
    private double GetHue(int color)
    {
        int r = color & 0xFF;
        int g = (color >> 8) & 0xFF;
        int b = (color >> 16) & 0xFF;

        int max = Math.Max(Math.Max(r, g), b);
        int min = Math.Min(Math.Min(r, g), b);

        if (max == min) return 0;

        double hue;
        if (max == r)
            hue = (g - b) / (double)(max - min);
        else if (max == g)
            hue = 2 + (b - r) / (double)(max - min);
        else
            hue = 4 + (r - g) / (double)(max - min);

        hue *= 60;
        if (hue < 0) hue += 360;

        return hue;
    }

    /// <summary>
    /// 获取粗细名称
    /// </summary>
    /// <param name="weight">粗细值</param>
    /// <returns>粗细名称</returns>
    private string GetWeightName(int weight)
    {
        return weight switch
        {
            1 => "细",
            2 => "中",
            3 => "粗",
            4 => "超粗",
            _ => "未知",
        };
    }

    public IEnumerator<IExcelBorder> GetEnumerator()
    {
        foreach (object? item in _borders)
        {
            if (item is MsExcel.Border border)
                yield return new ExcelBorder(border);
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion
}
