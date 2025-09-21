//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// LinkFormat COM 对象的封装实现类。
/// 负责管理 COM 对象生命周期，提供安全的属性访问和资源释放。
/// </summary>
internal class ExcelLinkFormat : IExcelLinkFormat
{
    /// <summary>
    /// 内部持有的原始 COM 对象。
    /// </summary>
    internal MsExcel.LinkFormat _linkFormat;

    /// <summary>
    /// 标记对象是否已被释放。
    /// </summary>
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装类。
    /// </summary>
    /// <param name="linkFormat">原始的 LinkFormat COM 对象，不可为 null。</param>
    /// <exception cref="ArgumentNullException">当传入的 linkFormat 为 null 时抛出。</exception>
    internal ExcelLinkFormat(MsExcel.LinkFormat linkFormat)
    {
        _linkFormat = linkFormat ?? throw new ArgumentNullException(nameof(linkFormat));
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
            if (_linkFormat != null)
            {
                Marshal.ReleaseComObject(_linkFormat);
                _linkFormat = null;
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
    /// 获取此对象的父对象（通常是 Shape 或 OLEObject）。
    /// </summary>
    public object Parent => _linkFormat?.Parent;

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// 返回封装后的 <see cref="IExcelApplication"/> 接口实例。
    /// </summary>
    public IExcelApplication Application =>
        _linkFormat?.Application != null
            ? new ExcelApplication(_linkFormat.Application as MsExcel.Application)
            : null;

    /// <summary>
    /// 获取或设置链接的自动更新方式。
    /// 默认值：xlLinkTypeExcelLinks
    /// </summary>
    public bool AutoUpdate
    {
        get => _linkFormat != null
            ? _linkFormat.AutoUpdate
            : false;

        set
        {
            if (_linkFormat != null)
                _linkFormat.AutoUpdate = value;
        }
    }

    public bool Locked
    {
        get => _linkFormat != null
            ? _linkFormat.Locked
            : false;

        set
        {
            if (_linkFormat != null)
                _linkFormat.Locked = value;
        }
    }

    /// <summary>
    /// 立即从源更新链接内容。
    /// </summary>
    public void Update()
    {
        _linkFormat?.Update();
    }
}