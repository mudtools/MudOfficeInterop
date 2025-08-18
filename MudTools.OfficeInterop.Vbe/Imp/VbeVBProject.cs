//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Vbe.Imp;
/// <summary>
/// VBE VBProject 对象的二次封装实现类
/// 实现 IVbeVBProject 接口
/// </summary>
internal class VbeVBProject : IVbeVBProject
{
    internal MsVb.VBProject _vbProject;
    private bool _disposedValue = false;

    internal VbeVBProject(MsVb.VBProject vbProject)
    {
        _vbProject = vbProject ?? throw new ArgumentNullException(nameof(vbProject));
    }

    #region 基础属性
    public string Name
    {
        get => _vbProject.Name;
        set => _vbProject.Name = value;
    }

    public vbext_ProjectType Type => (vbext_ProjectType)_vbProject.Type;

    public object Parent => _vbProject.Parent;

    public IVbeApplication Application => _vbProject.Application != null ? new VbeApplication(_vbProject.VBE) : null;

    public string FileName => _vbProject.FileName;

    public string Description
    {
        get => _vbProject.Description;
        set => _vbProject.Description = value;
    }

    public string HelpFile
    {
        get => _vbProject.HelpFile;
        set => _vbProject.HelpFile = value;
    }

    public int HelpContextID
    {
        get => _vbProject.HelpContextID;
        set => _vbProject.HelpContextID = value;
    }

    public vbext_VBAMode Mode => (vbext_VBAMode)_vbProject.Mode;

    public vbext_ProjectProtection Protection => (vbext_ProjectProtection)_vbProject.Protection; // vbext_ProjectProtection
    #endregion

    #region 状态属性
    public bool IsSaved => _vbProject.Saved;

    public bool IsProtected => _vbProject.Protection == MsVb.vbext_ProjectProtection.vbext_pp_locked;

    #endregion

    #region 核心对象集合
    public IVbeVBComponents VBComponents => _vbProject.VBComponents != null ? new VbeVBComponents(_vbProject.VBComponents) : null;

    public IVbeReferences References => _vbProject.References != null ? new VbeReferences(_vbProject.References) : null;
    #endregion

    #region 操作方法
    public void Select(bool replace = true)
    {
        try
        {
            _vbProject.VBE.ActiveVBProject = _vbProject;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error selecting VBProject: {ex.Message}");
        }
    }

    public void Delete()
    {
        if (this.Parent is MsVb.VBProjects parentCollection)
        {
            parentCollection.Remove(this._vbProject);
        }
    }

    public void Save()
    {
        try
        {
            if (!string.IsNullOrEmpty(_vbProject.FileName))
            {
                SaveAs(_vbProject.FileName);
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("VBProject.FileName is empty. Cannot save without a path. Use SaveAs.");
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error saving VBProject: {ex.Message}");
        }
    }

    public void SaveAs(string fileName)
    {
        try
        {
            _vbProject.SaveAs(fileName);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error saving VBProject as '{fileName}': {ex.Message}");
        }
    }


    public void Refresh()
    {
        System.Diagnostics.Debug.WriteLine("Refreshing VBProject (implies repainting project window or updating references).");
        // No direct Refresh method. Repaint happens automatically or by interacting with IDE.
        // Updating references might be part of Refresh logic.
        // _vbProject.References.Update(); // Example for updating references
    }
    #endregion

    #region 项目操作
    public void Compile()
    {
        System.Diagnostics.Debug.WriteLine("Compiling VBProject.");
        try
        {
            _vbProject.VBE.CommandBars.ExecuteMso("CompileProject");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error compiling VBProject: {ex.Message}");
        }
    }

    public void Run()
    {
        System.Diagnostics.Debug.WriteLine("Running VBProject (implies running Startup Object/Form).");
        try
        {
            _vbProject.VBE.CommandBars.ExecuteMso("RunSub");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error running VBProject: {ex.Message}");
        }
    }
    #endregion

    #region 引用管理
    public IVbeReference AddReference(object reference, string guid, int major, int minor)
    {
        System.Diagnostics.Debug.WriteLine("Adding reference to VBProject.");
        try
        {
            MsVb.Reference newRef = null;
            if (reference is string refString)
            {
                if (System.IO.File.Exists(refString))
                {
                    // Add from file
                    newRef = _vbProject.References.AddFromFile(refString);
                }
                else
                {
                    // Assume it's a GUID string
                    newRef = _vbProject.References.AddFromGuid(refString, major, minor);
                }
            }
            else if (reference is Guid refGuid)
            {
                newRef = _vbProject.References.AddFromGuid(refGuid.ToString(), major, minor);
            }

            return newRef != null ? new VbeReference(newRef) : null;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error adding reference: {ex.Message}");
            return null;
        }
    }

    public void RemoveReference(IVbeReference reference)
    {
        System.Diagnostics.Debug.WriteLine("Removing reference from VBProject.");
        try
        {
            if (reference is VbeReference vbeRef)
            {
                _vbProject.References.Remove(vbeRef._reference);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error removing reference: {ex.Message}");
        }
    }
    #endregion    

    #region 导出和转换
    public string GetCodeText()
    {
        System.Diagnostics.Debug.WriteLine("Getting all code text from VBProject.");
        try
        {
            var codeBuilder = new System.Text.StringBuilder();
            for (int i = 1; i <= _vbProject.VBComponents.Count; i++)
            {
                var component = _vbProject.VBComponents.Item(i);
                codeBuilder.AppendLine($"' === Component: {component.Name} ===");
                if (component.CodeModule.CountOfLines > 0)
                {
                    codeBuilder.AppendLine(component.CodeModule.Lines[1, component.CodeModule.CountOfLines]);
                }
                codeBuilder.AppendLine("' ==========================");
                codeBuilder.AppendLine();
            }
            return codeBuilder.ToString();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error getting all code text: {ex.Message}");
            return "";
        }
    }


    #endregion


    #region IDisposable Support
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            _vbProject = null; // Just nullify the reference

            _disposedValue = true;
        }
    }

    ~VbeVBProject()
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
