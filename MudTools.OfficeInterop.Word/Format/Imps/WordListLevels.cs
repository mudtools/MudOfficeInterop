//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// Word 文档列表级别集合实现类
/// </summary>
internal class WordListLevels : IWordListLevels
{
    private readonly MsWord.ListLevels _listLevels;
    private bool _disposedValue;


    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="listLevels">COM ListLevels 对象</param>
    internal WordListLevels(MsWord.ListLevels listLevels)
    {
        _listLevels = listLevels ?? throw new ArgumentNullException(nameof(listLevels));
        _disposedValue = false;
    }

    /// <summary>
    /// 根据索引获取列表级别
    /// </summary>
    /// <param name="index">列表级别索引</param>
    /// <returns>列表级别对象</returns>
    public IWordListLevel Item(int index)
    {
        if (index < 1 || index > 9) // Word 最多支持 9 个列表级别
            throw new ArgumentOutOfRangeException(nameof(index), "Index must be between 1 and 9.");

        try
        {
            var listLevel = _listLevels[index];
            return new WordListLevel(listLevel);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to get list level at index {index}.", ex);
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
