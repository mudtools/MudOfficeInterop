//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Vbe.Imp;

/// <summary>
/// VBE CodeModule 对象的二次封装实现类
/// 实现 IVbeCodeModule 接口
/// </summary>
internal class VbeCodeModule : IVbeCodeModule
{
    private MsVb.CodeModule _codeModule;
    private bool _disposedValue = false;

    internal VbeCodeModule(MsVb.CodeModule codeModule)
    {
        _codeModule = codeModule ?? throw new ArgumentNullException(nameof(codeModule));
    }

    #region 基础属性
    public object? Parent => _codeModule.Parent;

    public IVbeApplication Application => _codeModule.VBE != null ? new VbeApplication(_codeModule.VBE) : null;

    public string Name => _codeModule.Name;

    public int CountOfLines => _codeModule.CountOfLines;

    public int CountOfDeclarationLines => _codeModule.CountOfDeclarationLines;

    public string Language => "VB";
    #endregion

    #region 状态属性
    public bool IsEmpty => _codeModule.CountOfLines == 0;
    #endregion

    #region 代码访问
    public string GetLines(int startLine, int count = 1)
    {
        try
        {
            if (startLine < 1 || startLine > _codeModule.CountOfLines)
                return "";

            int actualCount = Math.Min(count, _codeModule.CountOfLines - startLine + 1);
            if (actualCount <= 0)
                return "";

            return _codeModule.get_Lines(startLine, actualCount);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error getting lines {startLine}-{startLine + count - 1}: {ex.Message}");
            return "";
        }
    }

    public void AddFromString(string codeText)
    {
        try
        {
            _codeModule.AddFromString(codeText);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error adding code: {ex.Message}");
        }
    }

    public string GetAllCode()
    {
        if (_codeModule.CountOfLines > 0)
        {
            return GetLines(1, _codeModule.CountOfLines);
        }
        return "";
    }
    #endregion

    #region 操作方法
    public void Select(bool replace = true)
    {
        try
        {
            var parentComponent = this.Parent as MsVb.VBComponent;
            parentComponent?.Activate();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error selecting CodeModule: {ex.Message}");
        }
    }

    public void DeleteLines(int startLine, int count = 1)
    {
        try
        {
            if (startLine >= 1 && count > 0 && (startLine + count - 1) <= _codeModule.CountOfLines)
            {
                _codeModule.DeleteLines(startLine, count);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error deleting lines {startLine}-{startLine + count - 1}: {ex.Message}");
        }
    }

    public void InsertLines(int lineNumber, string codeText)
    {
        try
        {
            if (lineNumber >= 1 && lineNumber <= (_codeModule.CountOfLines + 1))
            {
                _codeModule.InsertLines(lineNumber, codeText);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error inserting lines at {lineNumber}: {ex.Message}");
        }
    }

    public void ReplaceLines(int startLine, int count, string newCodeText)
    {
        DeleteLines(startLine, count);
        InsertLines(startLine, newCodeText);
    }

    public void AddCode(string codeText)
    {
        try
        {
            _codeModule.AddFromString(codeText);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error adding code: {ex.Message}");
        }
    }

    public void AddCodeFromFile(string fileName)
    {
        try
        {
            if (System.IO.File.Exists(fileName))
            {
                string code = System.IO.File.ReadAllText(fileName);
                AddCode(code);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error adding code from file '{fileName}': {ex.Message}");
        }
    }

    public void Clear()
    {
        try
        {
            if (_codeModule.CountOfLines > 0)
            {
                _codeModule.DeleteLines(1, _codeModule.CountOfLines);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error clearing code module: {ex.Message}");
        }
    }
    #endregion

    #region 查找和替换
    public bool Find(string target, int startLine = 1, int startColumn = 1,
                                       int endLine = int.MaxValue, int endColumn = int.MaxValue,
                                       bool wholeWord = false, bool matchCase = false, bool patternSearch = false)
    {
        try
        {
            endLine = Math.Min(endLine, _codeModule.CountOfLines);
            if (endLine <= 0 || startLine > endLine) return false;

            endColumn = Math.Min(endColumn, GetLines(endLine).Length); // Simplified end column check

            int foundLine = 0;
            int foundColumn = 0;
            bool found = _codeModule.Find(target, ref startLine,
                ref startColumn, ref endLine,
                ref endColumn,
                wholeWord, matchCase, patternSearch);
            return found;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error finding text '{target}': {ex.Message}");
        }
        return false;
    }
    #endregion

    #region 导出和导入
    public void Export(string fileName)
    {
        System.Diagnostics.Debug.WriteLine("Exporting CodeModule via parent VBComponent.");
        try
        {
            var parentComponent = this.Parent as MsVb.VBComponent;
            parentComponent?.Export(fileName);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error exporting CodeModule to '{fileName}': {ex.Message}");
        }
    }
    #endregion

    #region IDisposable Support
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            _codeModule = null;
            _disposedValue = true;
        }
    }

    ~VbeCodeModule()
    {
        Dispose(disposing: false);
    }

    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
    #endregion
}
