//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word 文档样式集合实现类
/// </summary>
internal class WordStyles : IWordStyles
{
    private readonly MsWord.Styles _styles;
    private readonly IWordDocument _document;
    private bool _disposedValue;

    /// <summary>
    /// 获取样式数量
    /// </summary>
    public int Count => _styles.Count;


    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="styles">COM Styles 对象</param>
    /// <param name="document">关联的文档对象</param>
    internal WordStyles(MsWord.Styles styles, IWordDocument document)
    {
        _styles = styles ?? throw new ArgumentNullException(nameof(styles));
        _document = document ?? throw new ArgumentNullException(nameof(document));
        _disposedValue = false;
    }

    /// <summary>
    /// 根据索引获取样式
    /// </summary>
    /// <param name="index">样式索引</param>
    /// <returns>样式对象</returns>
    public IWordStyle Item(int index)
    {
        if (index < 1 || index > Count)
            throw new ArgumentOutOfRangeException(nameof(index), $"Index must be between 1 and {Count}.");

        try
        {
            var style = _styles[index];
            return new WordStyle(style);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to get style at index {index}.", ex);
        }
    }

    /// <summary>
    /// 根据名称获取样式
    /// </summary>
    /// <param name="name">样式名称</param>
    /// <returns>样式对象</returns>
    public IWordStyle Item(string name)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Style name cannot be null or empty.", nameof(name));

        try
        {
            var style = _styles[name];
            return new WordStyle(style);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to get style with name '{name}'.", ex);
        }
    }

    /// <summary>
    /// 添加样式
    /// </summary>
    /// <param name="name">样式名称</param>
    /// <param name="type">样式类型</param>
    /// <returns>样式对象</returns>
    public IWordStyle Add(string name, int type = 1)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Style name cannot be null or empty.", nameof(name));

        try
        {
            var style = _styles.Add(name, (MsWord.WdStyleType)type);
            return new WordStyle(style);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to add style '{name}'.", ex);
        }
    }

    /// <summary>
    /// 获取枚举器
    /// </summary>
    /// <returns>样式枚举器</returns>
    public IEnumerator<IWordStyle> GetEnumerator()
    {
        try
        {
            var styles = new List<IWordStyle>();
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    styles.Add(Item(i));
                }
                catch
                {
                    // 忽略获取失败的样式
                    continue;
                }
            }
            return styles.GetEnumerator();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to enumerate styles.", ex);
        }
    }

    /// <summary>
    /// 获取枚举器
    /// </summary>
    /// <returns>枚举器</returns>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
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
