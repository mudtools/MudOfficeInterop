//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// Word 文档样式实现类
/// </summary>
internal class WordStyle : IWordStyle
{
    private readonly MsWord.Style _style;
    private bool _disposedValue;

    /// <summary>
    /// 获取样式名称
    /// </summary>
    public string Name => _style.NameLocal;

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object Parent => _style.Parent;

    /// <summary>
    /// 获取或设置是否基于其他样式
    /// </summary>
    public bool BasedOn
    {
        get => _style.get_BaseStyle != null;
        set
        {
            // 注意：设置 BasedOn 需要更复杂的实现
            throw new NotImplementedException("Setting BasedOn property is not implemented.");
        }
    }

    /// <summary>
    /// 获取或设置字体名称
    /// </summary>
    public string FontName
    {
        get => _style.Font.Name;
        set => _style.Font.Name = value;
    }

    /// <summary>
    /// 获取或设置字体大小
    /// </summary>
    public float FontSize
    {
        get => _style.Font.Size;
        set => _style.Font.Size = value;
    }

    /// <summary>
    /// 获取或设置是否加粗
    /// </summary>
    public bool Bold
    {
        get => _style.Font.Bold == 1;
        set => _style.Font.Bold = value ? 1 : 0;
    }

    /// <summary>
    /// 获取或设置是否斜体
    /// </summary>
    public bool Italic
    {
        get => _style.Font.Italic == 1;
        set => _style.Font.Italic = value ? 1 : 0;
    }

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="style">COM Style 对象</param>
    internal WordStyle(MsWord.Style style)
    {
        _style = style ?? throw new ArgumentNullException(nameof(style));
        _disposedValue = false;
    }

    /// <summary>
    /// 删除样式
    /// </summary>
    public void Delete()
    {
        try
        {
            _style.Delete();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to delete style.", ex);
        }
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    /// <param name="disposing">是否正在 disposing</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;
        _disposedValue = true;
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}
