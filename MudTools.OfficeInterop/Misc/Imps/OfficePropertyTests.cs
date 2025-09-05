//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;

namespace MudTools.OfficeInterop.Imps;
/// <summary>
/// 对 Microsoft.Office.Core.PropertyTests 的二次封装实现类。
/// 提供安全访问属性测试条件集合的方式，并管理 COM 对象生命周期。
/// </summary>
internal class OfficePropertyTests : IOfficePropertyTests
{
    private static readonly ILog log = LogManager.GetLogger(typeof(OfficePropertyTests));

    private MsCore.PropertyTests _propertyTests;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装的 PropertyTests 对象。
    /// </summary>
    /// <param name="propertyTests">原始的 COM PropertyTests 对象。</param>
    internal OfficePropertyTests(MsCore.PropertyTests propertyTests)
    {
        _propertyTests = propertyTests ?? throw new ArgumentNullException(nameof(propertyTests));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public int Count => _propertyTests?.Count ?? 0;

    /// <inheritdoc/>
    public IOfficePropertyTest this[int index]
    {
        get
        {
            if (_propertyTests == null || index < 1 || index > Count)
                return null;

            try
            {
                var propertyTest = _propertyTests[index];
                return propertyTest != null ? new OfficePropertyTest(propertyTest) : null;
            }
            catch (COMException ce)
            {
                log.Error($"Failed to retrieve object based on index: {ce.Message}", ce);
                return null;
            }
        }
    }
    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public IOfficePropertyTest? Add(string name,
        MsoCondition condition,
        object value, object? secondValue = null,
        MsoConnector connector = MsoConnector.msoConnectorAnd)
    {
        if (_propertyTests == null || string.IsNullOrEmpty(name))
            return null;

        _propertyTests.Add(name, (MsCore.MsoCondition)(int)condition, value, (MsCore.MsoConnector)(int)secondValue);
        var propertyTest = _propertyTests[Count];
        return propertyTest != null ? new OfficePropertyTest(propertyTest) : null;
    }

    /// <inheritdoc/>
    public void Remove(int index)
    {
        if (_propertyTests == null || index < 1 || index > Count)
            return;

        _propertyTests.Remove(index);
    }

    /// <inheritdoc/>
    public void Clear()
    {
        while (Count > 0)
        {
            _propertyTests.Remove(1);
        }
    }

    #endregion

    #region IEnumerable<IOfficePropertyTest> 实现

    /// <inheritdoc/>
    public IEnumerator<IOfficePropertyTest> GetEnumerator()
    {
        if (_propertyTests == null)
            yield break;

        for (int i = 1; i <= Count; i++)
        {
            var propertyTest = _propertyTests[i];
            if (propertyTest != null)
                yield return new OfficePropertyTest(propertyTest);
        }
    }

    /// <inheritdoc/>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放资源的核心方法。
    /// </summary>
    /// <param name="disposing">是否由 Dispose() 调用。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _propertyTests != null)
        {
            Marshal.ReleaseComObject(_propertyTests);
            _propertyTests = null;
        }

        _disposedValue = true;
    }

    /// <inheritdoc/>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}