//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;
/// <summary>
/// PowerPoint Presentations 集合对象的二次封装实现类
/// 实现 IPowerPointPresentations 接口
/// </summary>
internal class PowerPointPresentations : IPowerPointPresentations
{
    private MsPowerPoint.Presentations _presentations;
    private bool _disposedValue = false;

    /// <summary>
    /// 初始化 PowerPointPresentations 实例
    /// </summary>
    /// <param name="presentations">要封装的 Microsoft.Office.Interop.PowerPoint.Presentations 对象</param>
    internal PowerPointPresentations(MsPowerPoint.Presentations presentations)
    {
        _presentations = presentations ?? throw new ArgumentNullException(nameof(presentations));
    }

    #region 基础属性
    public int Count => _presentations.Count;

    public IPowerPointPresentation this[int index] => new PowerPointPresentation(_presentations[index]);

    public object Parent => _presentations.Parent;

    public IPowerPointApplication Application => _presentations.Application != null ? new PowerPointApplication(_presentations.Application) : null;
    #endregion

    #region 创建和添加
    public IPowerPointPresentation Add(bool withWindow = true)
    {
        var presentation = _presentations.Add(withWindow ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse);
        return new PowerPointPresentation(presentation);
    }

    public IPowerPointPresentation Open(string fileName, bool readOnly = false, bool untitled = false, bool withWindow = true)
    {
        var presentation = _presentations.Open(
            fileName,
            readOnly ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse,
            untitled ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse,
            withWindow ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse
        );
        return new PowerPointPresentation(presentation);
    }

    public IPowerPointPresentation CreateFromTemplate(string templatePath, bool withWindow = true)
    {
        return Open(templatePath, false, false, withWindow);
    }

    public int OpenRange(string[] filePaths, bool readOnly = false)
    {
        int count = 0;
        foreach (var path in filePaths)
        {
            try
            {
                Open(path, readOnly);
                count++;
            }
            catch
            {
                // Log error or handle individual file open failure
            }
        }
        return count;
    }
    #endregion

    #region 查找和筛选
    public IPowerPointPresentation[] FindByName(string name, bool matchCase = false)
    {
        var results = new List<IPowerPointPresentation>();
        for (int i = 1; i <= Count; i++)
        {
            var pres = this[i];
            if (string.Compare(pres.Name, name, !matchCase) == 0)
            {
                results.Add(pres);
            }
        }
        return results.ToArray();
    }

    public IPowerPointPresentation[] FindByPath(string path, bool matchCase = false)
    {
        var results = new List<IPowerPointPresentation>();
        for (int i = 1; i <= Count; i++)
        {
            var pres = this[i];
            if (string.Compare(pres.FullName, path, !matchCase) == 0)
            {
                results.Add(pres);
            }
        }
        return results.ToArray();
    }

    public IPowerPointPresentation GetActivePresentation()
    {
        try
        {
            var activePres = _presentations.Application.ActivePresentation;
            if (activePres != null)
            {
                return new PowerPointPresentation(activePres);
            }
            return null;
        }
        catch
        {
            // Handle error
        }
        return null;
    }

    #endregion

    #region 操作方法
    public void Clear(bool saveChanges = true)
    {
        for (int i = Count; i >= 1; i--)
        {
            try
            {
                foreach (var presentation in this)
                    presentation?.Close();
            }
            catch
            {
            }
        }
    }


    public void Delete(string name)
    {
        try
        {
            var presentations = FindByName(name);
            foreach (var presentation in presentations)
                presentation?.Close();
        }
        catch
        {
        }
    }

    public void Delete(IPowerPointPresentation presentation, bool saveChanges = true)
    {
        if (presentation is PowerPointPresentation pptPres)
        {
            try
            {
                pptPres._presentation?.Close();
            }
            catch { /* Handle error */ }
        }
    }
    #endregion

    #region 导出和导入
    public int ExportToFolder(string folderPath, PpSaveAsFileType format = PpSaveAsFileType.ppSaveAsOpenXMLPresentation, string prefix = "presentation_")
    {
        if (!System.IO.Directory.Exists(folderPath)) return 0;

        int count = 0;
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var pres = this[i];
                string fileName = System.IO.Path.Combine(folderPath, $"{prefix}{pres.Name}.{format}");
                pres.SaveAs(fileName, format);
                count++;
            }
            catch
            {
                // Log error
            }
        }
        return count;
    }

    public int ImportFromFolder(string folderPath, string fileExtension = ".pptx")
    {
        if (!System.IO.Directory.Exists(folderPath)) return 0;

        int count = 0;
        string[] files = System.IO.Directory.GetFiles(folderPath, "*" + fileExtension);
        foreach (string file in files)
        {
            try
            {
                Open(file);
                count++;
            }
            catch
            {
                // Log error
            }
        }
        return count;
    }

    #endregion



    #region IEnumerable<IPowerPointPresentation> Support
    public IEnumerator<IPowerPointPresentation> GetEnumerator()
    {
        for (int i = 1; i <= _presentations.Count; i++)
        {
            yield return new PowerPointPresentation((MsPowerPoint.Presentation)_presentations[i]);
        }
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
    #endregion

    #region IDisposable Support
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            _presentations = null;
            _disposedValue = true;
        }
    }

    ~PowerPointPresentations()
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
