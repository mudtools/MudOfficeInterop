//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 表示一个自定义组合键的封装实现类。
/// </summary>
internal class WordKeyBinding : IWordKeyBinding
{
    private MsWord.KeyBinding _keyBinding;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 <see cref="WordKeyBinding"/> 类的新实例。
    /// </summary>
    /// <param name="keyBinding">要封装的原始 COM KeyBinding 对象。</param>
    internal WordKeyBinding(MsWord.KeyBinding keyBinding)
    {
        _keyBinding = keyBinding ?? throw new ArgumentNullException(nameof(keyBinding));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _keyBinding != null ? new WordApplication(_keyBinding.Application) : null;

    /// <inheritdoc/>
    public object Parent => _keyBinding?.Parent;

    /// <inheritdoc/>
    public int Creator => _keyBinding?.Creator ?? 0;

    /// <inheritdoc/>
    public bool Protected
    {
        get => _keyBinding?.Protected ?? false;
    }

    /// <inheritdoc/>
    public WdKey KeyCode
        => _keyBinding?.KeyCode != null ? (WdKey)(int)_keyBinding?.KeyCode : WdKey.wdNoKey;

    /// <inheritdoc/>
    public WdKey KeyCode2
        => _keyBinding?.KeyCode2 != null ? (WdKey)(int)_keyBinding?.KeyCode2 : WdKey.wdNoKey;

    /// <inheritdoc/>
    public WdKeyCategory KeyCategory
         => _keyBinding?.KeyCategory != null ? (WdKeyCategory)(int)_keyBinding?.KeyCategory : WdKeyCategory.wdKeyCategoryNil;

    /// <inheritdoc/>
    public string Command => _keyBinding?.Command ?? string.Empty;

    /// <inheritdoc/>
    public string CommandParameter => _keyBinding?.CommandParameter ?? string.Empty;
    /// <inheritdoc/>
    public object Context => _keyBinding?.Context;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Clear()
    {
        _keyBinding?.Clear();
    }

    /// <inheritdoc/>
    public void Execute()
    {
        _keyBinding?.Execute();
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放由 <see cref="WordKeyBinding"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _keyBinding != null)
        {
            Marshal.ReleaseComObject(_keyBinding);
            _keyBinding = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 释放由 <see cref="WordKeyBinding"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}