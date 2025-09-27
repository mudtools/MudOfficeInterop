//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
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
    private static readonly ILog log = LogManager.GetLogger(typeof(ExcelBorders));
    /// <summary>
    /// 底层的 COM Borders 集合对象
    /// </summary>
    private MsExcel.Borders _borders;

    private DisposableList _disposables = [];
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
            _disposables?.Dispose();
            if (_borders != null)
                Marshal.ReleaseComObject(_borders);
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
    public IExcelBorder? this[XlBordersIndex borderType]
    {
        get
        {
            if (_borders == null)
                return null;

            try
            {
                var bt = (MsExcel.XlBordersIndex)borderType;
                var border = _borders[bt];
                var b = border != null ? new ExcelBorder(border) : null;
                if (b != null)
                    _disposables.Add(b);
                return b;
            }
            catch (COMException e)
            {
                log.Error("获取指定类型的边框对象失败：" + e.Message, e);
                return null;
            }
            catch (Exception e)
            {
                log.Error("获取指定类型的边框对象失败：" + e.Message, e);
                return null;
            }
        }
    }

    /// <summary>
    /// 获取边框集合所在的父对象
    /// </summary>
    public object? Parent => _borders?.Parent;

    /// <summary>
    /// 获取边框集合所在的Application对象
    /// </summary>
    public IExcelApplication? Application
    {
        get
        {
            var application = _borders?.Application;
            return application != null ? new ExcelApplication(application) : null;
        }
    }

    public XlLineStyle LineStyle
    {
        get => _borders != null ? _borders.ColorIndex.ObjectConvertEnum(XlLineStyle.xlContinuous) : XlLineStyle.xlContinuous;
        set
        {
            if (_borders != null)
                _borders.ColorIndex = value.EnumConvert(MsExcel.XlLineStyle.xlContinuous);
        }
    }

    public XlBorderWeight Weight
    {
        get => _borders != null ? _borders.Weight.ObjectConvertEnum(XlBorderWeight.xlMedium) : XlBorderWeight.xlMedium;
        set
        {
            if (_borders != null)
                _borders.Weight = value.EnumConvert(MsExcel.XlBorderWeight.xlMedium);
        }
    }

    /// <summary>
    /// 获取或设置边框的颜色
    /// </summary>
    public Color Color
    {
        get
        {
            if (_borders != null)
            {
                var color = Convert.ToInt32(_borders.Color);
                return Color.FromArgb((int)(color & 0xFF), (int)((color >> 8) & 0xFF), (int)((color >> 16) & 0xFF));
            }
            return Color.White;
        }
        set
        {
            if (_borders != null)
                _borders.Color = (int)((value.B << 16) | (value.G << 8) | value.R);
        }
    }

    public XlColorIndex ColorIndex
    {
        get
        {
            if (_borders != null)
                return _borders.ColorIndex.ObjectConvertEnum(XlColorIndex.xlColorIndexAutomatic);
            return XlColorIndex.xlColorIndexAutomatic;
        }
        set
        {
            if (_borders != null)
                _borders.ColorIndex = value.EnumConvert(MsExcel.XlColorIndex.xlColorIndexAutomatic);
        }
    }
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

        try
        {
            foreach (object? item in _borders)
            {
                if (item is MsExcel.Border border && (int)border.LineStyle == lineStyle)
                {
                    IExcelBorder excelBorder = new ExcelBorder(border);
                    result.Add(excelBorder);
                    _disposables.Add(excelBorder);
                }
            }
            return result.ToArray();
        }
        catch (COMException x)
        {
            log.Error($"根据线条样式查找边框时，访问索引的边框发生异常:" + x.Message, x);
            return [];
        }
        catch (Exception x)
        {
            log.Error($"根据线条样式查找边框时，访问索引的边框发生异常:" + x.Message, x);
            return [];
        }
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
        try
        {
            foreach (object? item in _borders)
            {
                if (item is MsExcel.Border border && (int)border.Color == color)
                {
                    IExcelBorder excelBorder = new ExcelBorder(border);
                    result.Add(excelBorder);
                    _disposables.Add(excelBorder);
                }
            }
            return result.ToArray();
        }
        catch (COMException x)
        {
            log.Error($"根据颜色查找边框时，访问索引的边框发生异常:" + x.Message, x);
            return [];
        }
        catch (Exception x)
        {
            log.Error($"根据颜色查找边框时，访问索引的边框发生异常:" + x.Message, x);
            return [];
        }
    }

    /// <summary>
    /// 根据粗细查找边框
    /// </summary>
    /// <param name="weight">边框粗细</param>
    /// <returns>匹配的边框数组</returns>
    public IExcelBorder[] FindByWeight(XlBorderWeight weight)
    {
        if (_borders == null || Count == 0)
            return [];

        List<IExcelBorder> result = [];

        try
        {
            foreach (object? item in _borders)
            {
                if (item is MsExcel.Border border && border.Weight.ConvertToInt() == (int)weight)
                {
                    IExcelBorder excelBorder = new ExcelBorder(border);
                    result.Add(excelBorder);
                    _disposables.Add(excelBorder);
                }
            }
            return result.ToArray();
        }
        catch (COMException x)
        {
            log.Error($"根据粗细查找边框时，访问索引的边框发生异常:" + x.Message, x);
            return [];
        }
        catch (Exception x)
        {
            log.Error($"根据粗细查找边框时，访问索引的边框发生异常:" + x.Message, x);
            return [];
        }

    }

    #endregion

    #region 格式设置

    /// <summary>
    /// 设置所有边框的线条样式
    /// </summary>
    /// <param name="lineStyle">线条样式</param>
    /// <param name="weight">边框粗细</param>
    public void SetLineStyle(XlLineStyle lineStyle, XlBorderWeight weight = XlBorderWeight.xlHairline)
    {
        if (_borders == null || Count == 0)
            return;

        try
        {
            _borders.LineStyle = (MsExcel.XlLineStyle)lineStyle;
            _borders.Weight = weight.EnumConvert(MsExcel.XlBorderWeight.xlHairline);
            _borders[MsExcel.XlBordersIndex.xlEdgeLeft].LineStyle = (MsExcel.XlLineStyle)lineStyle;
            _borders[MsExcel.XlBordersIndex.xlEdgeRight].LineStyle = (MsExcel.XlLineStyle)lineStyle;
            _borders[MsExcel.XlBordersIndex.xlEdgeTop].LineStyle = (MsExcel.XlLineStyle)lineStyle;
            _borders[MsExcel.XlBordersIndex.xlEdgeBottom].LineStyle = (MsExcel.XlLineStyle)lineStyle;
        }
        catch (COMException x)
        {
            log.Error($"设置所有边框的线条样式时，访问索引的边框发生异常:" + x.Message, x);
        }
        catch (Exception x)
        {
            log.Error($"设置所有边框的线条样式时，访问索引的边框发生异常:" + x.Message, x);
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
        catch (COMException x)
        {
            log.Error($"设置所有边框的颜色时，访问索引的边框发生异常:" + x.Message, x);
        }
        catch (Exception x)
        {
            log.Error($"设置所有边框的颜色时，访问索引的边框发生异常:" + x.Message, x);
        }
    }

    /// <summary>
    /// 设置所有边框的粗细
    /// </summary>
    /// <param name="weight">边框粗细</param>
    public void SetWeight(XlBorderWeight weight)
    {
        if (_borders == null || Count == 0)
            return;

        try
        {
            foreach (object? item in _borders)
            {
                if (item is MsExcel.Border border)
                {
                    border.Weight = weight.EnumConvert(MsExcel.XlBorderWeight.xlHairline);
                }
            }
        }
        catch (COMException x)
        {
            log.Error($"设置所有边框的粗细时，访问索引的边框发生异常:" + x.Message, x);
        }
        catch (Exception x)
        {
            log.Error($"设置所有边框的粗细时，访问索引的边框发生异常:" + x.Message, x);
        }
    }

    /// <summary>
    /// 统一所有边框的格式
    /// </summary>
    /// <param name="lineStyle">线条样式</param>
    /// <param name="color">边框颜色</param>
    /// <param name="weight">边框粗细</param>
    public void UniformFormat(Color color,
                            XlLineStyle lineStyle = XlLineStyle.xlLineStyleNone,
                            XlBorderWeight weight = XlBorderWeight.xlThin)
    {
        if (_borders == null || Count == 0)
            return;

        try
        {
            foreach (object? item in _borders)
            {
                if (item is MsExcel.Border border)
                {
                    border.LineStyle = lineStyle.EnumConvert(MsExcel.XlLineStyle.xlLineStyleNone);
                    border.Color = (int)((color.B << 16) | (color.G << 8) | color.R);
                    border.Weight = weight.EnumConvert(MsExcel.XlBorderWeight.xlHairline);
                }
            }
        }
        catch (COMException x)
        {
            log.Error($"统一所有边框的格式时，访问索引的边框发生异常:" + x.Message, x);
        }
        catch (Exception x)
        {
            log.Error($"统一所有边框的格式时，访问索引的边框发生异常:" + x.Message, x);
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
        catch (COMException x)
        {
            log.Error($"复制边框格式时，访问索引的边框发生异常", x);
        }
        catch (Exception x)
        {
            log.Error($"复制边框格式时，访问索引的边框发生异常", x);
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
                    UniformFormat(Color.Black, XlLineStyle.xlContinuous, XlBorderWeight.xlThin); // xlContinuous, black, xlThin
                    break;
                case 2: // 虚线边框
                    UniformFormat(Color.Black, XlLineStyle.xlDash, XlBorderWeight.xlThin); // xlDash, black, xlThin
                    break;
                case 3: // 点线边框
                    UniformFormat(Color.Black, XlLineStyle.xlContinuous, XlBorderWeight.xlThin); // xlDot, black, xlThin
                    break;
                case 4: // 双线边框
                    UniformFormat(Color.Black, XlLineStyle.xlDouble, XlBorderWeight.xlMedium); // xlDouble, black, xlMedium
                    break;
                case 5: // 粗边框
                    UniformFormat(Color.Black, XlLineStyle.xlContinuous, XlBorderWeight.xlThick); // xlContinuous, black, xlThick
                    break;
                default:
                    // 默认样式
                    UniformFormat(Color.Black, XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
                    break;
            }
        }
        catch (COMException x)
        {
            log.Error($"应用预设边框样式时，访问索引的边框发生异常:" + x.Message, x);
        }
        catch (Exception x)
        {
            log.Error($"应用预设边框样式时，访问索引的边框发生异常:" + x.Message, x);
        }
    }

    #endregion

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
}
