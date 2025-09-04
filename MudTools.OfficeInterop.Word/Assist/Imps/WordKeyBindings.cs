//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;
using MudTools.OfficeInterop.Word.Imps;

namespace MudTools.OfficeInterop.Word.Assist.Imps;
/// <summary>
/// 表示当前上下文中的自定义键分配集合的封装实现类。
/// </summary>
internal class WordKeyBindings : IWordKeyBindings
{
    private static readonly ILog log = LogManager.GetLogger(typeof(WordKeyBindings));
    private MsWord.KeyBindings _keyBindings;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 <see cref="WordKeyBindings"/> 类的新实例。
    /// </summary>
    /// <param name="keyBindings">要封装的原始 COM KeyBindings 对象。</param>
    internal WordKeyBindings(MsWord.KeyBindings keyBindings)
    {
        _keyBindings = keyBindings ?? throw new ArgumentNullException(nameof(keyBindings));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _keyBindings != null ? new WordApplication(_keyBindings.Application) : null;

    /// <inheritdoc/>
    public object Parent => _keyBindings?.Parent;

    /// <inheritdoc/>
    public int Count => _keyBindings?.Count ?? 0;

    /// <inheritdoc/>
    public IWordKeyBinding this[int index]
    {
        get
        {
            if (_keyBindings == null || index < 1 || index > Count) return null;
            try
            {
                var comKeyBinding = _keyBindings[index];
                return comKeyBinding != null ? new WordKeyBinding(comKeyBinding) : null;
            }
            catch (COMException ce)
            {
                log.Error($"Failed to retrieve object based on index: {ce.Message}", ce);
                return null;
            }
        }
    }

    /// <inheritdoc/>
    public object Context => _keyBindings?.Context;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public IWordKeyBinding Add(WdKeyCategory keyCategory, string command, int keyCode,
                              object keyCode2, object commandParameter)
    {
        if (_keyBindings == null) return null;
        try
        {
            var newKeyBinding = _keyBindings.Add((MsWord.WdKeyCategory)(int)keyCategory, command, keyCode, ref keyCode2, ref commandParameter);
            return newKeyBinding != null ? new WordKeyBinding(newKeyBinding) : null;
        }
        catch (COMException ex)
        {
            log.Error($"Failed to add key binding: {ex.Message}");
            return null;
        }
    }

    /// <inheritdoc/>
    public void ClearAll()
    {
        _keyBindings?.ClearAll();
    }

    /// <inheritdoc/>
    public IWordKeyBinding Key(int keyCode, object keyCode2)
    {
        if (_keyBindings == null) return null;
        try
        {
            var keyBinding = _keyBindings.Key(keyCode, ref keyCode2);
            return keyBinding != null ? new WordKeyBinding(keyBinding) : null;
        }
        catch (COMException)
        {
            // 如果按键组合不存在，则返回 null。
            return null;
        }
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放由 <see cref="WordKeyBindings"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _keyBindings != null)
        {
            Marshal.ReleaseComObject(_keyBindings);
            _keyBindings = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 释放由 <see cref="WordKeyBindings"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion

    #region IEnumerable<IWordKeyBinding> 实现

    /// <inheritdoc/>
    public IEnumerator<IWordKeyBinding> GetEnumerator()
    {
        for (int i = 1; i <= Count; i++)
        {
            yield return this[i];
        }
    }

    /// <inheritdoc/>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion
}