//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// CalloutFormat COM 对象的封装实现类。
/// 负责管理 COM 对象生命周期，提供安全的属性访问和资源释放。
/// </summary>
internal class ExcelCalloutFormat : IExcelCalloutFormat
{
    /// <summary>
    /// 内部持有的原始 COM 对象。
    /// </summary>
    internal MsExcel.CalloutFormat _calloutFormat;

    /// <summary>
    /// 标记对象是否已被释放。
    /// </summary>
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装类。
    /// </summary>
    /// <param name="calloutFormat">原始的 CalloutFormat COM 对象，不可为 null。</param>
    /// <exception cref="ArgumentNullException">当传入的 calloutFormat 为 null 时抛出。</exception>
    internal ExcelCalloutFormat(MsExcel.CalloutFormat calloutFormat)
    {
        _calloutFormat = calloutFormat ?? throw new ArgumentNullException(nameof(calloutFormat));
        _disposedValue = false;
    }

    /// <summary>
    /// 释放资源的受保护虚方法，支持派生类重写。
    /// </summary>
    /// <param name="disposing">是否由用户代码显式调用释放。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            // 释放托管资源：释放 COM 对象
            if (_calloutFormat != null)
            {
                Marshal.ReleaseComObject(_calloutFormat);
                _calloutFormat = null;
            }
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 公开的 Dispose 方法，用于显式释放资源。
    /// 调用后对象不应再被使用。
    /// </summary>
    public void Dispose() => Dispose(true);

    /// <summary>
    /// 获取此对象的父对象（通常是 Shape）。
    /// </summary>
    public object Parent => _calloutFormat?.Parent;

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// 返回封装后的 <see cref="IExcelApplication"/> 接口实例。
    /// </summary>
    public IExcelApplication Application =>
        _calloutFormat?.Application != null
            ? new ExcelApplication(_calloutFormat.Application as MsExcel.Application)
            : null;

    /// <summary>
    /// 获取或设置标注的类型（如无引线、单引线、角度引线等）。
    /// 默认值：msoCalloutTypeMixed
    /// </summary>
    public MsoCalloutType Type
    {
        get => _calloutFormat != null
            ? _calloutFormat.Type.EnumConvert(MsoCalloutType.msoCalloutMixed)
            : MsoCalloutType.msoCalloutMixed;

        set
        {
            if (_calloutFormat != null)
                _calloutFormat.Type = value.EnumConvert(MsCore.MsoCalloutType.msoCalloutMixed);
        }
    }

    /// <summary>
    /// 获取或设置标注引线的角度（仅对角度引线有效）。
    /// 默认值：msoCalloutAngleMixed
    /// </summary>
    public MsoCalloutAngleType Angle
    {
        get => _calloutFormat != null
            ? _calloutFormat.Angle.EnumConvert(MsoCalloutAngleType.msoCalloutAngleMixed)
            : MsoCalloutAngleType.msoCalloutAngleMixed;

        set
        {
            if (_calloutFormat != null)
                _calloutFormat.Angle = value.EnumConvert(MsCore.MsoCalloutAngleType.msoCalloutAngleMixed);
        }
    }

    public float Gap
    {
        get => _calloutFormat != null ? _calloutFormat.Gap : 0f;
        set
        {
            if (_calloutFormat != null)
                _calloutFormat.Gap = value;
        }
    }

    /// <summary>
    /// 获取或设置标注引线的起点相对于文本框的垂直偏移量（单位：磅）。
    /// 仅对部分引线类型有效。
    /// </summary>
    public float Drop
    {
        get => _calloutFormat != null ? _calloutFormat.Drop : 0f;
    }

    /// <summary>
    /// 获取标注引线起点类型（自动/手动）。
    /// 默认值：msoCalloutDropMixed
    /// </summary>
    public MsoCalloutDropType DropType =>
        _calloutFormat != null
            ? _calloutFormat.DropType.EnumConvert(MsoCalloutDropType.msoCalloutDropMixed)
            : MsoCalloutDropType.msoCalloutDropMixed;

    /// <summary>
    /// 获取或设置是否在标注中显示强调引线（如加粗或特殊样式）。
    /// </summary>
    public bool Accent
    {
        get => _calloutFormat != null && _calloutFormat.Accent.ConvertToBool();
        set
        {
            if (_calloutFormat != null)
                _calloutFormat.Accent = value.ConvertTriState();
        }
    }

    public bool Border
    {
        get => _calloutFormat != null && _calloutFormat.Border.ConvertToBool();
        set
        {
            if (_calloutFormat != null)
                _calloutFormat.Border = value.ConvertTriState();
        }

    }

    /// <summary>
    /// 获取或设置标注引线是否自动调整以避免遮挡文本。
    /// </summary>
    public bool AutoAttach
    {
        get => _calloutFormat != null && _calloutFormat.AutoAttach.ConvertToBool();
        set
        {
            if (_calloutFormat != null)
                _calloutFormat.AutoAttach = value.ConvertTriState();
        }
    }

    /// <summary>
    /// 获取或设置标注引线的长度（仅对部分类型有效，单位：磅）。
    /// </summary>
    public float Length
    {
        get => _calloutFormat != null ? _calloutFormat.Length : 0f;
    }

    public void AutomaticLength()
    {
        _calloutFormat?.AutomaticLength();
    }

    public void CustomDrop(float Drop)
    {
        _calloutFormat?.CustomDrop(Drop);
    }

    public void CustomLength(float Length)
    {
        _calloutFormat?.CustomLength(Length);
    }

    public void PresetDrop(MsoCalloutDropType DropType)
    {
        _calloutFormat?.PresetDrop(DropType.EnumConvert(MsCore.MsoCalloutDropType.msoCalloutDropMixed));
    }
}