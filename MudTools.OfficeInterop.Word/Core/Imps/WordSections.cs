//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word 文档节集合实现类
/// </summary>
internal class WordSections : IWordSections
{
    private readonly MsWord.Sections _sections;
    private readonly IWordDocument _document;
    private bool _disposedValue;

    /// <summary>
    /// 获取节数量
    /// </summary>
    public int Count => _sections.Count;

    public IWordApplication? Application => _sections != null ? new WordApplication(_sections.Application) : null;


    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="sections">COM Sections 对象</param>
    /// <param name="document">关联的文档对象</param>
    internal WordSections(MsWord.Sections sections, IWordDocument document)
    {
        _sections = sections ?? throw new ArgumentNullException(nameof(sections));
        _document = document ?? throw new ArgumentNullException(nameof(document));
        _disposedValue = false;
    }

    /// <summary>
    /// 根据索引获取节
    /// </summary>
    /// <param name="index">节索引</param>
    /// <returns>节对象</returns>
    public IWordSection Item(int index)
    {
        if (index < 1 || index > Count)
            throw new ArgumentOutOfRangeException(nameof(index), $"Index must be between 1 and {Count}.");

        try
        {
            var section = _sections[index];
            return new WordSection(section);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to get section at index {index}.", ex);
        }
    }

    /// <summary>
    /// 添加节
    /// </summary>
    /// <param name="range">插入范围</param>
    /// <param name="type">分节符类型</param>
    /// <returns>节对象</returns>
    public IWordSection Add(IWordRange range, int type = 2)
    {
        try
        {
            // 注意：这里需要将 IWordRange 转换为 COM Range 对象
            // 由于缺少具体实现，这里使用占位符
            var comRange = GetComRange(range);
            var section = _sections.Add(comRange, (MsWord.WdBreakType)type);
            return new WordSection(section);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to add section.", ex);
        }
    }

    /// <summary>
    /// 获取枚举器
    /// </summary>
    /// <returns>节枚举器</returns>
    public IEnumerator<IWordSection> GetEnumerator()
    {
        try
        {
            var sections = new List<IWordSection>();
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    sections.Add(Item(i));
                }
                catch
                {
                    // 忽略获取失败的节
                    continue;
                }
            }
            return sections.GetEnumerator();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to enumerate sections.", ex);
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
    /// 将 IWordRange 转换为 COM Range 对象
    /// </summary>
    /// <param name="range">IWordRange 对象</param>
    /// <returns>COM Range 对象</returns>
    private MsWord.Range GetComRange(IWordRange range)
    {
        // 这里需要具体的实现来获取 COM Range 对象
        // 由于缺少具体实现，返回 null 作为占位符
        return null;
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

