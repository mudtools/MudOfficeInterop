//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// ListDataFormat COM 对象的封装实现类。
/// 负责管理 COM 对象生命周期，提供安全的属性访问和资源释放。
/// 注意：ListDataFormat 是只读对象，所有属性均为 getter。
/// </summary>
internal class ExcelListDataFormat : IExcelListDataFormat
{
    /// <summary>
    /// 内部持有的原始 COM 对象。
    /// </summary>
    internal MsExcel.ListDataFormat _listDataFormat;

    /// <summary>
    /// 标记对象是否已被释放。
    /// </summary>
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装类。
    /// </summary>
    /// <param name="listDataFormat">原始的 ListDataFormat COM 对象，不可为 null。</param>
    /// <exception cref="ArgumentNullException">当传入的 listDataFormat 为 null 时抛出。</exception>
    internal ExcelListDataFormat(MsExcel.ListDataFormat listDataFormat)
    {
        _listDataFormat = listDataFormat ?? throw new ArgumentNullException(nameof(listDataFormat));
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
            if (_listDataFormat != null)
            {
                Marshal.ReleaseComObject(_listDataFormat);
                _listDataFormat = null;
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
    /// 获取此对象的父对象（通常是 ListColumn）。
    /// </summary>
    public object Parent => _listDataFormat?.Parent;

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// 返回封装后的 <see cref="IExcelApplication"/> 接口实例。
    /// </summary>
    public IExcelApplication Application =>
        _listDataFormat?.Application != null
            ? new ExcelApplication(_listDataFormat.Application as MsExcel.Application)
            : null;

    /// <summary>
    /// 获取该列数据的默认值（如果未设置则返回 null）。
    /// </summary>
    public object DefaultValue => _listDataFormat?.DefaultValue;

    /// <summary>
    /// 获取该列是否允许“自由填写”（即允许用户输入列表外的值）。
    /// 注意：此属性名称在 COM 中为 AllowFillIn，常用于下拉列表列。
    /// </summary>
    public bool AllowFillIn => _listDataFormat != null && _listDataFormat.AllowFillIn.ConvertToBool();

    /// <summary>
    /// 获取该列是否为“必需”字段（即不允许空值）。
    /// </summary>
    public bool Required => _listDataFormat != null && _listDataFormat.Required.ConvertToBool();

    /// <summary>
    /// 获取该列数据类型（如文本、数字、日期等）。
    /// 使用 Excel.XlListDataType 枚举。
    /// 默认值：xlListDataTypeNone。
    /// </summary>
    public XlListDataType Type =>
        _listDataFormat != null
            ? _listDataFormat.Type.EnumConvert(XlListDataType.xlListDataTypeNone)
            : XlListDataType.xlListDataTypeNone;

    public int MaxCharacters => _listDataFormat?.MaxCharacters ?? 0;

    /// <summary>
    /// 获取该列数据校验的最小值（仅对数字/日期类型有效）。
    /// 如果未设置或类型不支持，返回 null。
    /// </summary>
    public object MinNumber => _listDataFormat?.MinNumber;

    /// <summary>
    /// 获取该列数据校验的最大值（仅对数字/日期类型有效）。
    /// 如果未设置或类型不支持，返回 null。
    /// </summary>
    public object MaxNumber => _listDataFormat?.MaxNumber;

    /// <summary>
    /// 获取该列是否启用“只读”模式（用户不能编辑）。
    /// </summary>
    public bool ReadOnly => _listDataFormat != null && _listDataFormat.ReadOnly.ConvertToBool();


    /// <summary>
    /// 获取该列自定义错误提示信息（数据校验失败时显示）。
    /// </summary>
    public int DecimalPlaces => _listDataFormat?.DecimalPlaces ?? 0;
}