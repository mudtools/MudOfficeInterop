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

}