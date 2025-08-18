//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Vbe.Imp;
/// <summary>
/// VBE VBProjects 集合对象的二次封装实现类
/// 实现 IVbeVBProjects 接口
/// </summary>
internal class VbeVBProjects : IVbeVBProjects
{
    private MsVb.VBProjects _vbProjects;
    private bool _disposedValue = false;

    internal VbeVBProjects(MsVb.VBProjects vbProjects)
    {
        _vbProjects = vbProjects ?? throw new ArgumentNullException(nameof(vbProjects));
    }

    #region 基础属性
    public int Count => _vbProjects.Count;

    public IVbeVBProject this[int index] => new VbeVBProject(_vbProjects.Item(index));

    public object Parent => _vbProjects.Parent;

    public IVbeApplication Application => _vbProjects.VBE != null ? new VbeApplication(_vbProjects.VBE) : null;
    #endregion

    #region 创建和添加
    public IVbeVBProject Add(vbext_ProjectType projectType, string projectName = "")
    {
        MsVb.VBProject newProject = _vbProjects.Add((MsVb.vbext_ProjectType)projectType);
        if (!string.IsNullOrEmpty(projectName))
        {
            newProject.Name = projectName;
        }
        return new VbeVBProject(newProject);
    }

    public IVbeVBProject Open(string fileName)
    {
        MsVb.VBProject openedProject = _vbProjects.Open(fileName);
        return new VbeVBProject(openedProject);
    }


    public IVbeVBProject CreateFromTemplate(string templatePath, string projectName = "")
    {
        return Add(vbext_ProjectType.vbext_pt_HostProject, projectName);
    }
    #endregion

    #region 查找和筛选
    public IVbeVBProject[] FindByName(string name, bool matchCase = false)
    {
        var results = new List<IVbeVBProject>();
        for (int i = 1; i <= Count; i++)
        {
            var project = this[i];
            if (string.Compare(project.Name, name, !matchCase) == 0)
            {
                results.Add(project);
            }
        }
        return results.ToArray();
    }

    public IVbeVBProject[] FindByType(vbext_ProjectType projectType)
    {
        var results = new List<IVbeVBProject>();
        for (int i = 1; i <= Count; i++)
        {
            var project = this[i];
            if (project.Type == projectType)
            {
                results.Add(project);
            }
        }
        return results.ToArray();
    }

    public IVbeVBProject[] FindByPath(string path, bool matchCase = false)
    {
        var results = new List<IVbeVBProject>();
        for (int i = 1; i <= Count; i++)
        {
            var project = this[i];
            if (string.Compare(project.FileName, path, !matchCase) == 0)
            {
                results.Add(project);
            }
        }
        return results.ToArray();
    }

    public IVbeVBProject[] GetStandardExeProjects()
    {
        return FindByType(vbext_ProjectType.vbext_pt_StandAlone);
    }

    public IVbeVBProject[] GetDllProjects()
    {
        return FindByType(vbext_ProjectType.vbext_pt_HostProject);
    }

    public IVbeVBProject[] GetProtectedProjects()
    {
        var results = new List<IVbeVBProject>();
        for (int i = 1; i <= Count; i++)
        {
            var project = this[i];
            if (project.Protection == vbext_ProjectProtection.vbext_pp_locked)
            {
                results.Add(project);
            }
        }
        return results.ToArray();
    }
    #endregion

    #region 操作方法
    public void Clear()
    {
        for (int i = Count; i >= 1; i--)
        {
            try
            {
                Delete(i);
            }
            catch
            {
            }
        }
    }

    public void Delete(int index)
    {
        try
        {
            MsVb.VBProject projectToDelete = _vbProjects.Item(index) as MsVb.VBProject;
            _vbProjects.Remove(projectToDelete);
        }
        catch
        {

        }
    }

    public void Delete(string name)
    {
        try
        {
            MsVb.VBProject projectToDelete = _vbProjects.Item(name) as MsVb.VBProject;
            _vbProjects.Remove(projectToDelete);
        }
        catch
        {

        }
    }

    public void Delete(IVbeVBProject project)
    {
        if (project is VbeVBProject vbeProject)
        {
            try
            {
                _vbProjects.Remove(vbeProject._vbProject);
            }
            catch
            {
                // Handle error
            }
        }
    }

    public void DeleteRange(int[] indices)
    {
        var sortedIndices = new List<int>(indices);
        sortedIndices.Sort((a, b) => b.CompareTo(a));
        foreach (int index in sortedIndices)
        {
            Delete(index);
        }
    }

    #endregion

    #region 导出和导入
    public int ExportToFolder(string folderPath, string prefix = "project_")
    {
        if (!System.IO.Directory.Exists(folderPath)) return 0;

        int count = 0;
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var project = this[i];
                string fileName = System.IO.Path.Combine(folderPath, $"{prefix}{project.Name}.vbp");
                project.SaveAs(fileName);
                count++;
            }
            catch
            {
                // Log error
            }
        }
        return count;
    }

    public int ImportFromFolder(string folderPath)
    {
        if (!System.IO.Directory.Exists(folderPath)) return 0;

        int count = 0;
        string[] files = System.IO.Directory.GetFiles(folderPath, "*.vbp"); // Assuming .vbp for VB6 projects
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

    #region 高级功能
    public IVbeVBProject ActiveProject => _vbProjects.Parent != null ? new VbeVBProject(_vbProjects.Parent as MsVb.VBProject) : null;

    public void CompileAll()
    {
        System.Diagnostics.Debug.WriteLine("Compiling all VBProjects.");
        try
        {
            for (int i = 1; i <= Count; i++)
            {
                this[i].Compile();
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error compiling all projects: {ex.Message}");
        }
    }
    #endregion

    #region IEnumerable<IVbeVBProject> Support
    public IEnumerator<IVbeVBProject> GetEnumerator()
    {
        for (int i = 1; i <= _vbProjects.Count; i++)
        {
            yield return new VbeVBProject(_vbProjects.Item(i));
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
            _vbProjects = null;

            _disposedValue = true;
        }
    }

    ~VbeVBProjects()
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
