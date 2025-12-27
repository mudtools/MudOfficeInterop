//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Vbe.Imp;
/// <summary>
/// VBE References 集合对象的二次封装实现类
/// 实现 IVbeReferences 接口
/// </summary>
internal class VbeReferences : IVbeReferences
{
    private MsVb.References _references;
    private bool _disposedValue = false;

    internal VbeReferences(MsVb.References references)
    {
        _references = references ?? throw new ArgumentNullException(nameof(references));
    }

    #region 基础属性
    public int Count => _references.Count;

    public IVbeReference this[int index] => new VbeReference((MsVb.Reference)_references.Item(index));

    public object? Parent => _references.Parent;

    public IVbeApplication Application => _references.VBE != null ? new VbeApplication(_references.VBE) : null;
    #endregion

    #region 创建和添加
    public IVbeReference AddFromGuid(string guid, int major, int minor)
    {
        MsVb.Reference newReference = _references.AddFromGuid(guid, major, minor);
        return new VbeReference(newReference);
    }

    public IVbeReference AddFromFile(string fileName)
    {
        MsVb.Reference newReference = _references.AddFromFile(fileName);
        return new VbeReference(newReference);
    }

    public IVbeReference CopyFrom(IVbeReference sourceReference)
    {
        try
        {
            if (sourceReference is VbeReference vbeRef)
            {
                var comRef = vbeRef._reference;
                return AddFromFile(sourceReference.FullPath);
            }
        }
        catch
        {
            // Handle error
        }
        return null;
    }
    #endregion

    #region 查找和筛选
    public IVbeReference[] FindByName(string name, bool matchCase = false)
    {
        var results = new List<IVbeReference>();
        for (int i = 1; i <= Count; i++)
        {
            var reference = this[i];
            if (string.Compare(reference.Name, name, !matchCase) == 0)
            {
                results.Add(reference);
            }
        }
        return results.ToArray();
    }

    public IVbeReference[] FindByGuid(string guid)
    {
        var results = new List<IVbeReference>();
        for (int i = 1; i <= Count; i++)
        {
            var reference = this[i];
            if (reference.Guid == guid)
            {
                results.Add(reference);
            }
        }
        return results.ToArray();
    }

    public IVbeReference[] FindByPath(string path, bool matchCase = false)
    {
        var results = new List<IVbeReference>();
        for (int i = 1; i <= Count; i++)
        {
            var reference = this[i];
            if (string.Compare(reference.FullPath, path, !matchCase) == 0)
            {
                results.Add(reference);
            }
        }
        return results.ToArray();
    }

    public IVbeReference[] GetBuiltInReferences()
    {
        var results = new List<IVbeReference>();
        for (int i = 1; i <= Count; i++)
        {
            var reference = this[i];
            if (reference.IsBuiltIn)
            {
                results.Add(reference);
            }
        }
        return results.ToArray();
    }

    public IVbeReference[] GetProjectReferences()
    {
        var results = new List<IVbeReference>();
        for (int i = 1; i <= Count; i++)
        {
            var reference = this[i];
            if (!reference.IsBuiltIn)
            {
                results.Add(reference);
            }
        }
        return results.ToArray();
    }

    public IVbeReference[] GetBrokenReferences()
    {
        var results = new List<IVbeReference>();
        for (int i = 1; i <= Count; i++)
        {
            var reference = this[i];
            if (reference.IsBroken)
            {
                results.Add(reference);
            }
        }
        return results.ToArray();
    }

    public IVbeReference[] GetValidReferences()
    {
        var results = new List<IVbeReference>();
        for (int i = 1; i <= Count; i++)
        {
            var reference = this[i];
            if (!reference.IsBroken)
            {
                results.Add(reference);
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
            MsVb.Reference referenceToDelete = _references.Item(index) as MsVb.Reference;
            _references.Remove(referenceToDelete);
        }
        catch
        {

        }
    }

    public void Delete(string name)
    {
        try
        {
            MsVb.Reference referenceToDelete = _references.Item(name) as MsVb.Reference;
            _references.Remove(referenceToDelete);
        }
        catch
        {

        }
    }

    public void Delete(IVbeReference reference)
    {
        if (reference is VbeReference vbeReference) // Check for real implementation
        {
            try
            {
                _references.Remove(vbeReference._reference);
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

    public int ImportFromFile(string filePath)
    {
        if (!System.IO.File.Exists(filePath)) return 0;

        int count = 0;
        try
        {
            string[] lines = System.IO.File.ReadAllLines(filePath);
            foreach (var line in lines)
            {
                string[] parts = line.Split('|');
                if (parts.Length >= 4)
                {
                    string name = parts[0];
                    string guid = parts[1];
                    string versionStr = parts[2];
                    string path = parts[3];
                    bool isBroken = parts.Length > 4 && bool.Parse(parts[4]); // Simplified

                    if (!isBroken)
                    {
                        if (!string.IsNullOrEmpty(guid) && versionStr.Contains('.'))
                        {
                            string[] verParts = versionStr.Split('.');
                            if (int.TryParse(verParts[0], out int major) && int.TryParse(verParts[1], out int minor))
                            {
                                try
                                {
                                    AddFromGuid(guid, major, minor);
                                    count++;
                                }
                                catch { /* Handle add error */ }
                            }
                        }
                        else if (!string.IsNullOrEmpty(path) && System.IO.File.Exists(path))
                        {
                            try
                            {
                                AddFromFile(path);
                                count++;
                            }
                            catch { /* Handle add error */ }
                        }
                    }
                }
            }
        }
        catch
        {
            // Log error
        }
        return count;
    }

    #endregion


    #region IEnumerable<IVbeReference> Support
    public IEnumerator<IVbeReference> GetEnumerator()
    {
        for (int i = 1; i <= _references.Count; i++)
        {
            yield return new VbeReference(_references.Item(i));
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
            _references = null;

            _disposedValue = true;
        }
    }

    ~VbeReferences()
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
