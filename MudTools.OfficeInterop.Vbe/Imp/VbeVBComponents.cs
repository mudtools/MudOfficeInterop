//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using Microsoft.Vbe.Interop;
using System.Runtime.InteropServices;

namespace MudTools.OfficeInterop.Vbe.Imp;
/// <summary>
/// VBE VBComponents 集合对象的二次封装实现类
/// 实现 IVbeVBComponents 接口
/// </summary>
internal class VbeVBComponents : IVbeVBComponents
{
    private MsVb.VBComponents _vbComponents;
    private bool _disposedValue = false;

    internal VbeVBComponents(MsVb.VBComponents vbComponents)
    {
        _vbComponents = vbComponents ?? throw new ArgumentNullException(nameof(vbComponents));
    }

    #region 基础属性
    public int Count => _vbComponents.Count;

    public IVbeVBComponent this[int index] => new VbeVBComponent(_vbComponents.Item(index));

    public IVbeVBComponent this[string name] => new VbeVBComponent(_vbComponents.Item(name));

    public object Parent => _vbComponents.Parent;
    #endregion

    #region 创建和添加
    public IVbeVBComponent Add(vbext_ComponentType componentType, string name = "")
    {
        MsVb.VBComponent newComponent = _vbComponents.Add((MsVb.vbext_ComponentType)componentType);
        if (!string.IsNullOrEmpty(name))
        {
            newComponent.Name = name;
        }
        return new VbeVBComponent(newComponent);
    }

    public IVbeVBComponent Import(string fileName)
    {
        MsVb.VBComponent importedComponent = _vbComponents.Import(fileName);
        return new VbeVBComponent(importedComponent);
    }


    public IVbeVBComponent CreateFromTemplate(string templatePath, string name = "")
    {
        return Add(vbext_ComponentType.vbext_ct_StdModule, name);
    }
    #endregion

    #region 查找和筛选
    public IVbeVBComponent[] FindByName(string name, bool matchCase = false)
    {
        var results = new List<IVbeVBComponent>();
        for (int i = 1; i <= Count; i++)
        {
            var component = this[i];
            if (string.Compare(component.Name, name, !matchCase) == 0)
            {
                results.Add(component);
            }
        }
        return results.ToArray();
    }

    public IVbeVBComponent[] FindByType(vbext_ComponentType componentType)
    {
        var results = new List<IVbeVBComponent>();
        for (int i = 1; i <= Count; i++)
        {
            var component = this[i];
            if (component.Type == componentType)
            {
                results.Add(component);
            }
        }
        return results.ToArray();
    }

    public IVbeVBComponent[] GetStandardModules()
    {
        return FindByType(Vbe.vbext_ComponentType.vbext_ct_StdModule);
    }

    public IVbeVBComponent[] GetClassModules()
    {
        return FindByType(Vbe.vbext_ComponentType.vbext_ct_ClassModule);
    }

    public IVbeVBComponent[] GetUserForms()
    {
        return FindByType(Vbe.vbext_ComponentType.vbext_ct_MSForm);
    }

    public IVbeVBComponent[] GetDocumentModules()
    {
        var results = new List<IVbeVBComponent>();
        for (int i = 1; i <= Count; i++)
        {
            var component = this[i];
            if (component.Type == Vbe.vbext_ComponentType.vbext_ct_Document)
            {
                results.Add(component);
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
            MsVb.VBComponent componentToDelete = _vbComponents.Item(index) as MsVb.VBComponent;
            _vbComponents.Remove(componentToDelete);
        }
        catch
        {
        }
    }

    public void Delete(string name)
    {
        try
        {
            MsVb.VBComponent componentToDelete = _vbComponents.Item(name) as MsVb.VBComponent;
            _vbComponents.Remove(componentToDelete);
        }
        catch
        {
        }
    }

    public void Delete(IVbeVBComponent component)
    {
        if (component is VbeVBComponent vbeComponent)
        {
            try
            {
                _vbComponents.Remove(vbeComponent._vbComponent);
            }
            catch
            {

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
    public int ExportToFolder(string folderPath, string prefix = "component_")
    {
        if (!System.IO.Directory.Exists(folderPath)) return 0;

        int count = 0;
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var component = this[i];
                string fileName = System.IO.Path.Combine(folderPath, $"{prefix}{component.Name}.bas"); // .bas for modules, .frm for forms etc.
                component.Export(fileName);
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
        string[] files = System.IO.Directory.GetFiles(folderPath, "*.*", System.IO.SearchOption.TopDirectoryOnly)
                                             .Where(f => f.EndsWith(".bas", StringComparison.OrdinalIgnoreCase) ||
                                                         f.EndsWith(".cls", StringComparison.OrdinalIgnoreCase) ||
                                                         f.EndsWith(".frm", StringComparison.OrdinalIgnoreCase))
                                             .ToArray();
        foreach (string file in files)
        {
            try
            {
                Import(file);
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



    #region IEnumerable<IVbeVBComponent> Support
    public IEnumerator<IVbeVBComponent> GetEnumerator()
    {
        for (int i = 1; i <= _vbComponents.Count; i++)
        {
            yield return new VbeVBComponent(_vbComponents.Item(i));
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
            if (_vbComponents != null)
            {
                foreach (VBComponent vbCommpontent in _vbComponents)
                {
                    Marshal.ReleaseComObject(vbCommpontent);
                }
                Marshal.ReleaseComObject(_vbComponents);
            }

            _vbComponents = null;

            _disposedValue = true;
        }
    }

    ~VbeVBComponents()
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
