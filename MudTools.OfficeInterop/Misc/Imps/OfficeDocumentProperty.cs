//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;

namespace MudTools.OfficeInterop.Imps;

/// <summary>
/// IOfficeDocumentProperty 接口的内部实现类。
/// 封装了 Microsoft.Office.Core.DocumentProperty COM 对象，并负责其资源释放。
/// </summary>
internal class OfficeDocumentProperty : IOfficeDocumentProperty
{
    private static readonly ILog log = LogManager.GetLogger(typeof(OfficeDocumentProperty));
    internal MsCore.DocumentProperty? _property;
    private bool _disposedValue;

    /// <summary>
    /// 使用指定的 DocumentProperty COM 对象初始化 OfficeDocumentProperty 的新实例。
    /// </summary>
    /// <param name="property">要封装的 DocumentProperty COM 对象。</param>
    /// <exception cref="ArgumentNullException">当 <paramref name="property"/> 为 null 时抛出。</exception>
    internal OfficeDocumentProperty(MsCore.DocumentProperty property)
    {
        _property = property ?? throw new ArgumentNullException(nameof(property));
        _disposedValue = false;
    }

    #region IOfficeDocumentProperty 属性实现

    public object Application => _property?.Application ?? new object();
    public int Creator => _property?.Creator ?? 0;

    public string Name
    {
        get => _property?.Name ?? string.Empty;
        set
        {
            if (_property != null)
                _property.Name = value;
        }
    }

    public MsoDocProperties Type
    {
        get => _property?.Type.EnumConvert(MsoDocProperties.msoPropertyTypeString) ?? MsoDocProperties.msoPropertyTypeString;
        set
        {
            if (_property != null)
                _property.Type = value.EnumConvert(MsCore.MsoDocProperties.msoPropertyTypeString);
        }
    }

    public object Value
    {
        get => _property?.Value ?? new object();
        set
        {
            if (_property != null)
                _property.Value = value;
        }
    }

    public bool IsBuiltIn => _property != null ? _property.LinkToContent == false && _property.Type.EnumConvert(MsoDocProperties.msoPropertyTypeString) != MsoDocProperties.msoPropertyTypeString : false;


    public bool LinkToContent
    {
        get => _property != null && _property.LinkToContent;
        set
        {
            if (_property != null)
                _property.LinkToContent = value;
        }
    }

    #endregion

    #region IOfficeDocumentProperty 方法实现

    public void Delete()
    {
        if (_property == null)
            return;

        try
        {
            _property.Delete();
        }
        catch (Exception ex)
        {
            log.Error($"删除文档属性 '{Name}' 失败: {ex.Message}");
        }
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _property != null)
        {
            Marshal.ReleaseComObject(_property);
            _property = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}