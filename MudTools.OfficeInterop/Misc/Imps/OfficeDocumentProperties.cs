//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;

namespace MudTools.OfficeInterop.Imps;

/// <summary>
/// IOfficeDocumentProperties 接口的内部实现类。
/// 封装了 Microsoft.Office.Core.DocumentProperties COM 对象，并负责其资源释放。
/// </summary>
internal class OfficeDocumentProperties : IOfficeDocumentProperties
{
    private static readonly ILog log = LogManager.GetLogger(typeof(OfficeDocumentProperties));

    private dynamic _dynamicProperties;
    private bool _disposedValue;

    /// <summary>
    /// 使用指定的 DocumentProperties COM 对象初始化 OfficeDocumentProperties 的新实例。
    /// </summary>
    /// <param name="properties">要封装的 DocumentProperties COM 对象。</param>
    /// <exception cref="ArgumentNullException">当 <paramref name="properties"/> 为 null 时抛出。</exception>
    internal OfficeDocumentProperties(object properties)
    {
        _dynamicProperties = properties ?? throw new ArgumentNullException(nameof(properties));
        _disposedValue = false;
    }

    #region IOfficeDocumentProperties 属性实现

    public int Count => _dynamicProperties?.Count ?? 0;

    public IOfficeDocumentProperty? this[int index]
    {
        get
        {
            if (_dynamicProperties == null || index < 1 || index > _dynamicProperties.Count)
                return null;

            try
            {
                var comProperty = _dynamicProperties[index];
                return comProperty != null ? new OfficeDocumentProperty(comProperty) : null;
            }
            catch (Exception ex)
            {
                log.Error($"获取索引为 {index} 的文档属性失败: {ex.Message}");
                return null;
            }
        }
    }

    public IOfficeDocumentProperty? this[string name]
    {
        get
        {
            if (_dynamicProperties == null)
                return null;

            try
            {
                var comProperty = _dynamicProperties[name];
                return comProperty != null ? new OfficeDocumentProperty(comProperty) : null;
            }
            catch (Exception ex)
            {
                log.Warn($"未找到名为 '{name}' 的文档属性: {ex.Message}");
                return null;
            }
        }
    }

    #endregion

    #region IOfficeDocumentProperties 方法实现

    public IOfficeDocumentProperty? Add(string name, bool linkToContent, MsoDocProperties type, object value, object? linkSource = null)
    {
        if (_dynamicProperties == null)
            return null;

        try
        {
            // 根据是否链接到内容，调用不同的Add方法重载
            MsCore.DocumentProperty? comProperty;
            if (linkToContent)
            {
                // 当链接到内容时，value 参数通常代表链接源
                comProperty = _dynamicProperties.Add(name, linkToContent, type, value, linkSource);
            }
            else
            {
                comProperty = _dynamicProperties.Add(name, linkToContent, type, value);
            }

            return comProperty != null ? new OfficeDocumentProperty(comProperty) : null;
        }
        catch (Exception ex)
        {
            log.Error($"添加文档属性 '{name}' 失败: {ex.Message}");
            return null;
        }
    }

    #endregion

    #region IEnumerable<IOfficeDocumentProperty> 实现

    public IEnumerator<IOfficeDocumentProperty> GetEnumerator()
    {
        if (_dynamicProperties == null)
            yield break;

        for (int i = 1; i <= _dynamicProperties.Count; i++)
        {
            var property = this[i];
            if (property != null)
                yield return property;
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _dynamicProperties != null)
        {
            Marshal.ReleaseComObject(_dynamicProperties);
            _dynamicProperties = null;
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
