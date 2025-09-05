//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

internal class WordLanguages : IWordLanguages
{
    private MsWord.Languages _languages;
    private bool _disposedValue;

    public int Count => _languages.Count;

    public IWordLanguage this[int index] => new WordLanguage(_languages[index]);

    internal WordLanguages(MsWord.Languages languages)
    {
        _languages = languages ?? throw new ArgumentNullException(nameof(languages));
        _disposedValue = false;
    }

    public IWordLanguage GetLanguageByID(int languageID)
    {
        try
        {
            var language = _languages[languageID];
            return language != null ? new WordLanguage(language) : null;
        }
        catch (COMException)
        {
            return null; // 语言不存在时返回null
        }
    }

    public bool Contains(int languageID)
    {
        try
        {
            // 尝试获取语言，如果抛出异常则说明不存在
            var language = _languages[languageID];
            return language != null;
        }
        catch (COMException)
        {
            return false;
        }
    }

    public IEnumerator<IWordLanguage> GetEnumerator()
    {
        for (int i = 1; i <= Count; i++)
        {
            yield return this[i];
        }
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _languages != null)
        {
            try
            {
                while (Marshal.ReleaseComObject(_languages) > 0) { }
            }
            catch { }
            _languages = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}