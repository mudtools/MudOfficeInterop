//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.Excel.Imps;


/// <summary>
/// Excel Border 对象的二次封装实现类
/// 负责对 Microsoft.Office.Interop.Excel.Border 对象的安全访问和资源管理
/// </summary>
internal class ExcelBorder : IExcelBorder
{
    /// <summary>
    /// 底层的 COM Border 对象
    /// </summary>
    internal MsExcel.Border _border;

    /// <summary>
    /// 标记对象是否已被释放
    /// </summary>
    private bool _disposedValue;

    #region 构造函数和释放

    /// <summary>
    /// 初始化 ExcelBorder 实例
    /// </summary>
    /// <param name="border">底层的 COM Border 对象</param>
    internal ExcelBorder(MsExcel.Border border)
    {
        _border = border ?? throw new ArgumentNullException(nameof(border));
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
                // 释放底层COM对象
                if (_border != null)
                    Marshal.ReleaseComObject(_border);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _border = null;
        }

        _disposedValue = true;
    }

    ~ExcelBorder()
    {
        Dispose(disposing: false);
    }

    /// <summary>
    /// 实现 IDisposable 接口的释放方法
    /// </summary>
    public void Dispose() => Dispose(true);

    #endregion

    #region 基础属性

    /// <summary>
    /// 获取或设置边框的线条样式
    /// </summary>
    public XlLineStyle LineStyle
    {
        get => _border != null ? (XlLineStyle)Enum.ToObject(typeof(XlLineStyle), _border.LineStyle) : XlLineStyle.xlLineStyleNone;
        set
        {
            if (_border != null)
                _border.LineStyle = (MsExcel.XlLineStyle)Enum.ToObject(typeof(MsExcel.XlLineStyle), (int)value);
        }
    }

    /// <summary>
    /// 获取或设置边框的粗细
    /// </summary>
    public int Weight
    {
        get => _border != null ? Convert.ToInt32(_border.Weight) : 0;
        set
        {
            if (_border != null)
                _border.Weight = value;
        }
    }

    /// <summary>
    /// 获取或设置边框的颜色
    /// </summary>
    public Color Color
    {
        get
        {
            if (_border != null)
            {
                var color = Convert.ToInt32(_border.Color);
                return Color.FromArgb((int)(color & 0xFF), (int)((color >> 8) & 0xFF), (int)((color >> 16) & 0xFF));
            }
            return Color.White;
        }
        set
        {
            if (_border != null)
                _border.Color = (int)((value.B << 16) | (value.G << 8) | value.R);
        }
    }

    /// <summary>
    /// 获取或设置边框的主题颜色
    /// </summary>
    public Color ThemeColor
    {
        get
        {
            if (_border != null)
            {
                var color = Convert.ToInt32(_border.ThemeColor);
                return Color.FromArgb((int)(color & 0xFF), (int)((color >> 8) & 0xFF), (int)((color >> 16) & 0xFF));
            }
            return Color.White;
        }
        set
        {
            if (_border != null)
                _border.ThemeColor = (int)((value.B << 16) | (value.G << 8) | value.R);
        }
    }

    /// <summary>
    /// 获取或设置边框的着色和阴影
    /// </summary>
    public double TintAndShade
    {
        get => _border != null ? Convert.ToDouble(_border.TintAndShade) : 0;
        set
        {
            if (_border != null)
                _border.TintAndShade = value;
        }
    }

    /// <summary>
    /// 获取边框所在的父对象
    /// </summary>
    public object? Parent => _border?.Parent;

    /// <summary>
    /// 获取边框所在的Application对象
    /// </summary>
    public IExcelApplication? Application
    {
        get
        {
            var application = _border?.Application as MsExcel.Application;
            return application != null ? new ExcelApplication(application) : null;
        }
    }
    #endregion


    #region 格式设置    

    /// <summary>
    /// 重置边框为默认值
    /// </summary>
    public void Reset()
    {
        if (_border == null) return;

        try
        {
            LineStyle = XlLineStyle.xlContinuous;
            Color = Color.Black;
            Weight = 2;     // xlThin
        }
        catch
        {
            // 忽略重置过程中的异常
        }
    }

    /// <summary>
    /// 复制边框格式
    /// </summary>
    /// <param name="sourceBorder">源边框</param>
    public void CopyFormat(IExcelBorder sourceBorder)
    {
        if (_border == null || sourceBorder == null) return;

        try
        {
            LineStyle = sourceBorder.LineStyle;
            Color = sourceBorder.Color;
            Weight = sourceBorder.Weight;
        }
        catch
        {
            // 忽略复制格式过程中的异常
        }
    }

    /// <summary>
    /// 应用预设样式
    /// </summary>
    /// <param name="presetStyle">预设样式类型</param>
    public void ApplyPresetStyle(int presetStyle)
    {
        if (_border == null) return;

        try
        {
            switch (presetStyle)
            {
                case 1: // 实线边框
                    LineStyle = XlLineStyle.xlContinuous;
                    Color = Color.Black;
                    Weight = 2;     // xlThin
                    break;
                case 2: // 虚线边框
                    LineStyle = XlLineStyle.xlDash;
                    Color = Color.Black;
                    Weight = 2;        // xlThin
                    break;
                case 3: // 点线边框
                    LineStyle = XlLineStyle.xlDot;
                    Color = Color.Black;
                    Weight = 2;        // xlThin
                    break;
                case 4: // 双线边框
                    LineStyle = XlLineStyle.xlDouble;
                    Color = Color.Black;
                    Weight = 3;        // xlMedium
                    break;
                case 5: // 粗边框
                    LineStyle = XlLineStyle.xlContinuous;
                    Color = Color.Black;
                    Weight = 4;     // xlThick
                    break;
                default:
                    // 默认样式
                    LineStyle = XlLineStyle.xlContinuous;
                    Color = Color.Black;
                    Weight = 2;     // xlThin
                    break;
            }
        }
        catch
        {
            // 忽略应用预设样式过程中的异常
        }
    }

    #endregion

    #region 导出和转换

    /// <summary>
    /// 导出边框到文件
    /// </summary>
    /// <param name="filename">导出文件路径</param>
    /// <param name="overwrite">是否覆盖已存在文件</param>
    /// <returns>是否导出成功</returns>
    public bool Export(string filename, bool overwrite = true)
    {
        if (_border == null || string.IsNullOrEmpty(filename))
            return false;

        try
        {
            // 验证文件扩展名
            string extension = System.IO.Path.GetExtension(filename)?.ToLower();
            if (string.IsNullOrEmpty(extension))
            {
                filename += ".txt";
            }

            // 检查是否覆盖
            if (System.IO.File.Exists(filename) && !overwrite)
                return false;

            using (var writer = new System.IO.StreamWriter(filename, false, System.Text.Encoding.UTF8))
            {
                writer.WriteLine("Excel Border Export");
                writer.WriteLine("==================");
                writer.WriteLine($"Export Date: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                writer.WriteLine($"Line Style: {LineStyle}");
                writer.WriteLine($"Weight: {Weight}");
                writer.WriteLine($"Color: {Color}");
                writer.WriteLine($"Theme Color: {ThemeColor}");
                writer.WriteLine($"Tint And Shade: {TintAndShade}");
            }
            return true;
        }
        catch
        {
            return false;
        }
    }
    #endregion

    #region 私有辅助方法

    /// <summary>
    /// 获取线条样式名称
    /// </summary>
    /// <param name="lineStyle">线条样式</param>
    /// <returns>线条样式名称</returns>
    private string GetLineStyleName(XlLineStyle lineStyle)
    {
        return lineStyle switch
        {
            XlLineStyle.xlContinuous => "实线",
            XlLineStyle.xlDash => "虚线",
            XlLineStyle.xlDot => "点线",
            XlLineStyle.xlDouble => "双线",
            XlLineStyle.xlDashDot => "点划线",
            XlLineStyle.xlDashDotDot => "双点划线",
            XlLineStyle.xlSlantDashDot => "斜点划线",
            XlLineStyle.xlLineStyleNone => "无",
            _ => "未知",
        };
    }

    /// <summary>
    /// 获取颜色名称
    /// </summary>
    /// <param name="color">颜色值</param>
    /// <returns>颜色名称</returns>
    private string GetColorName(Color color)
    {
        // 将Color转换为RGB值进行比较
        int colorValue = (color.B << 16) | (color.G << 8) | color.R;

        return colorValue switch
        {
            0 => "黑色",
            255 => "红色",
            65280 => "绿色",           // 255 << 8
            16711680 => "蓝色",        // 255 << 16
            16777215 => "白色",        // 255 | (255 << 8) | (255 << 16)
            16776960 => "黄色",        // 255 | (255 << 8)
            16711935 => "洋红色",      // 255 | (255 << 16)
            65535 => "青色",           // 255 | (255 << 8)
            _ => $"RGB({color.R}, {color.G}, {color.B})",
        };
    }

    /// <summary>
    /// 判断是否为主要颜色
    /// </summary>
    /// <param name="color">颜色值</param>
    /// <returns>是否为主要颜色</returns>
    private bool IsPrimaryColor(Color color)
    {
        int colorValue = (color.B << 16) | (color.G << 8) | color.R;
        return colorValue == 0 || colorValue == 255 || colorValue == 65280 || colorValue == 16711680 ||
               colorValue == 16777215 || colorValue == 16776960 || colorValue == 16711935 || colorValue == 65535;
    }

    /// <summary>
    /// 判断是否为自定义颜色
    /// </summary>
    /// <param name="color">颜色值</param>
    /// <returns>是否为自定义颜色</returns>
    private bool IsCustomColor(Color color)
    {
        return !IsPrimaryColor(color);
    }

    /// <summary>
    /// 获取亮度值
    /// </summary>
    /// <param name="color">颜色值</param>
    /// <returns>亮度值</returns>
    private double GetBrightness(Color color)
    {
        return (color.R * 0.299 + color.G * 0.587 + color.B * 0.114) / 255.0;
    }

    /// <summary>
    /// 获取饱和度值
    /// </summary>
    /// <param name="color">颜色值</param>
    /// <returns>饱和度值</returns>
    private double GetSaturation(Color color)
    {
        int max = Math.Max(Math.Max(color.R, color.G), color.B);
        int min = Math.Min(Math.Min(color.R, color.G), color.B);

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
    private double GetHue(Color color)
    {
        int max = Math.Max(Math.Max(color.R, color.G), color.B);
        int min = Math.Min(Math.Min(color.R, color.G), color.B);

        if (max == min) return 0;

        double hue;
        if (max == color.R)
            hue = (color.G - color.B) / (double)(max - min);
        else if (max == color.G)
            hue = 2 + (color.B - color.R) / (double)(max - min);
        else
            hue = 4 + (color.R - color.G) / (double)(max - min);

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

    #endregion

    /// <summary>
    /// 获取边框的详细信息
    /// </summary>
    /// <returns>边框详细信息对象</returns>
    public BorderDetails GetDetails()
    {
        var details = new BorderDetails
        {
            LineStyle = (int)LineStyle,
            LineStyleName = GetLineStyleName(LineStyle),
            Weight = Weight,
            WeightName = GetWeightName(Weight),
            Color = Color.ToArgb(),
            ColorName = GetColorName(Color),
            ThemeColor = ThemeColor.ToArgb(),
            TintAndShade = TintAndShade,
            Locked = false, // 需要从父对象获取
            IsSolid = LineStyle == XlLineStyle.xlContinuous,
            IsDash = LineStyle == XlLineStyle.xlDash,
            IsDot = LineStyle == XlLineStyle.xlDot,
            IsDouble = LineStyle == XlLineStyle.xlDouble,
            IsPrimaryColor = IsPrimaryColor(Color),
            IsCustomColor = IsCustomColor(Color),
            IsThin = Weight == 1,
            IsMedium = Weight == 2,
            IsThick = Weight == 3,
            Created = DateTime.Now,
            Modified = DateTime.Now,
            Worksheet = "", // 需要从父对象获取
            RangeAddress = "", // 需要从父对象获取
            Category = "Border",
            Description = "Excel Border Object",
            Tags = [],
            Priority = 0,
            IsEnabled = true,
            Transparency = 0,
            GradientType = 0,
            GradientAngle = 0,
            GradientColor1 = Color.ToArgb(),
            GradientColor2 = Color.ToArgb(),
            GradientStop1 = 0,
            GradientStop2 = 1
        };

        return details;
    }

    /// <summary>
    /// 验证边框设置
    /// </summary>
    /// <returns>验证结果</returns>
    public BorderValidationResult Validate()
    {
        var result = new BorderValidationResult
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

        if (_border == null)
        {
            result.IsValid = false;
            result.ErrorMessage = "边框对象为空";
            return result;
        }

        // 验证线条样式
        if (!Enum.IsDefined(typeof(XlLineStyle), LineStyle))
        {
            result.IsValid = false;
            result.ValidLineStyle = false;
            result.ErrorMessage += "无效的线条样式; ";
        }

        // 验证粗细
        if (Weight < 1 || Weight > 4)
        {
            result.IsValid = false;
            result.ValidWeight = false;
            result.ErrorMessage += "无效的边框粗细; ";
        }

        return result;
    }

    /// <summary>
    /// 比较两个边框
    /// </summary>
    /// <param name="otherBorder">要比较的边框</param>
    /// <returns>比较结果</returns>
    public BorderComparisonResult Compare(IExcelBorder otherBorder)
    {
        var result = new BorderComparisonResult
        {
            AreEqual = false,
            Similarity = 0,
            Differences = [],
            Similarities = [],
            ComparisonTime = DateTime.Now,
            ComparisonType = "FullComparison",
            Description = "边框比较结果",
            RecommendedAction = "Review differences",
            CanBeMerged = true,
            CanBeReplaced = true,
            CanBeInherited = true,
            CanBeOverridden = true,
            NeedsUpdate = false,
            NeedsSynchronization = false,
            NeedsValidation = false,
            NeedsOptimization = false
        };

        if (otherBorder == null)
        {
            result.Differences = new[] { "对比边框为空" };
            return result;
        }

        var differences = new System.Collections.Generic.List<string>();
        var similarities = new System.Collections.Generic.List<string>();

        try
        {
            // 比较各种属性
            if (LineStyle != otherBorder.LineStyle)
                differences.Add($"线条样式不同: {LineStyle} vs {otherBorder.LineStyle}");
            else
                similarities.Add("线条样式相同");

            if (Weight != otherBorder.Weight)
                differences.Add($"边框粗细不同: {Weight} vs {otherBorder.Weight}");
            else
                similarities.Add("边框粗细相同");

            if (Color != otherBorder.Color)
                differences.Add($"边框颜色不同: {Color} vs {otherBorder.Color}");
            else
                similarities.Add("边框颜色相同");

            result.AreEqual = differences.Count == 0;
            result.Similarity = similarities.Count / (double)(similarities.Count + differences.Count);
            result.Differences = differences.ToArray();
            result.Similarities = similarities.ToArray();
        }
        catch
        {
            result.Differences = ["比较过程中发生异常"];
        }

        return result;
    }

    /// <summary>
    /// 克隆边框
    /// </summary>
    /// <returns>克隆的边框对象</returns>
    public IExcelBorder Clone()
    {
        if (_border?.Parent == null)
            return null;

        try
        {
            // 注意：Excel Border对象通常不能直接克隆
            // 这里提供一个示例实现框架
            return null;
        }
        catch
        {
            return null;
        }
    }
}