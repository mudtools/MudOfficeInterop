//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;


/// <summary>
/// PowerPoint 标签集合实现类
/// </summary>
internal class PowerPointTags : IPowerPointTags
{
    private readonly object _tags; // 使用 object 因为 Tags 可能有不同的 COM 类型
    private bool _disposedValue;

    /// <summary>
    /// 获取标签数量
    /// </summary>
    public int Count
    {
        get
        {
            try
            {
                if (_tags != null)
                {
                    var type = _tags.GetType();
                    var countProperty = type.GetProperty("Count");
                    return countProperty != null ? (int)countProperty.GetValue(_tags) : 0;
                }
                return 0;
            }
            catch
            {
                return 0;
            }
        }
    }

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object? Parent
    {
        get
        {
            try
            {
                if (_tags != null)
                {
                    var type = _tags.GetType();
                    var parentProperty = type.GetProperty("Parent");
                    return parentProperty?.GetValue(_tags);
                }
                return null;
            }
            catch
            {
                return null;
            }
        }
    }

    /// <summary>
    /// 根据索引获取标签值
    /// </summary>
    public string this[int index]
    {
        get
        {
            if (index < 1 || index > Count)
                throw new ArgumentOutOfRangeException(nameof(index), $"Index must be between 1 and {Count}.");

            try
            {
                if (_tags != null)
                {
                    var type = _tags.GetType();
                    var itemMethod = type.GetMethod("Item");
                    if (itemMethod != null)
                    {
                        var result = itemMethod.Invoke(_tags, new object[] { index });
                        return result?.ToString() ?? string.Empty;
                    }
                }
                return string.Empty;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to get tag value at index {index}.", ex);
            }
        }
    }

    /// <summary>
    /// 根据名称获取标签值
    /// </summary>
    public string this[string name]
    {
        get
        {
            if (string.IsNullOrEmpty(name))
                throw new ArgumentException("Tag name cannot be null or empty.", nameof(name));

            try
            {
                if (_tags != null)
                {
                    var type = _tags.GetType();
                    var itemMethod = type.GetMethod("Item");
                    if (itemMethod != null)
                    {
                        var result = itemMethod.Invoke(_tags, new object[] { name });
                        return result?.ToString() ?? string.Empty;
                    }
                }
                return string.Empty;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to get tag value with name '{name}'.", ex);
            }
        }
    }

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="tags">COM Tags 对象</param>
    internal PowerPointTags(object tags)
    {
        _tags = tags;
        _disposedValue = false;
    }


    /// <summary>
    /// 获取标签名称
    /// </summary>
    /// <param name="index">标签索引</param>
    /// <returns>标签名称</returns>
    public string Name(int index)
    {
        if (index < 1 || index > Count)
            throw new ArgumentOutOfRangeException(nameof(index), $"Index must be between 1 and {Count}.");

        try
        {
            if (_tags != null)
            {
                var type = _tags.GetType();
                var nameMethod = type.GetMethod("Name");
                if (nameMethod != null)
                {
                    var result = nameMethod.Invoke(_tags, new object[] { index });
                    return result?.ToString() ?? string.Empty;
                }
            }
            return string.Empty;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to get tag name at index {index}.", ex);
        }
    }

    /// <summary>
    /// 获取标签值
    /// </summary>
    /// <param name="index">标签索引</param>
    /// <returns>标签值</returns>
    public string Value(int index)
    {
        return this[index];
    }

    /// <summary>
    /// 添加标签
    /// </summary>
    /// <param name="name">标签名称</param>
    /// <param name="value">标签值</param>
    public void Add(string name, string value)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Tag name cannot be null or empty.", nameof(name));

        try
        {
            if (_tags != null)
            {
                var type = _tags.GetType();
                var addMethod = type.GetMethod("Add");
                if (addMethod != null)
                {
                    addMethod.Invoke(_tags, new object[] { name, value ?? string.Empty });
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to add tag '{name}'.", ex);
        }
    }

    /// <summary>
    /// 删除标签
    /// </summary>
    /// <param name="name">标签名称</param>
    public void Delete(string name)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Tag name cannot be null or empty.", nameof(name));

        try
        {
            if (_tags != null)
            {
                var type = _tags.GetType();
                var deleteMethod = type.GetMethod("Delete");
                if (deleteMethod != null)
                {
                    deleteMethod.Invoke(_tags, new object[] { name });
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to delete tag '{name}'.", ex);
        }
    }

    /// <summary>
    /// 清除所有标签
    /// </summary>
    public void Clear()
    {
        try
        {
            for (int i = Count; i >= 1; i--)
            {
                try
                {
                    var tagName = Name(i);
                    Delete(tagName);
                }
                catch
                {
                    continue;
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to clear all tags.", ex);
        }
    }

    /// <summary>
    /// 检查标签是否存在
    /// </summary>
    /// <param name="name">标签名称</param>
    /// <returns>是否存在</returns>
    public bool Contains(string name)
    {
        if (string.IsNullOrEmpty(name))
            return false;

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    var tagName = Name(i);
                    if (string.Equals(tagName, name, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
                catch
                {
                    continue;
                }
            }
            return false;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to check if tag '{name}' exists.", ex);
        }
    }

    /// <summary>
    /// 更新标签值
    /// </summary>
    /// <param name="name">标签名称</param>
    /// <param name="value">新标签值</param>
    public void Update(string name, string value)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Tag name cannot be null or empty.", nameof(name));

        try
        {
            if (Contains(name))
            {
                Delete(name);
            }
            Add(name, value ?? string.Empty);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to update tag '{name}'.", ex);
        }
    }

    /// <summary>
    /// 获取所有标签名称
    /// </summary>
    /// <returns>标签名称列表</returns>
    public IEnumerable<string> GetAllNames()
    {
        var names = new List<string>();
        try
        {
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    names.Add(Name(i));
                }
                catch
                {
                    continue;
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get all tag names.", ex);
        }
        return names;
    }

    /// <summary>
    /// 获取所有标签键值对
    /// </summary>
    /// <returns>标签键值对字典</returns>
    public IDictionary<string, string> GetAllTags()
    {
        var tags = new Dictionary<string, string>();
        try
        {
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    var name = Name(i);
                    var value = this[i];
                    tags[name] = value;
                }
                catch
                {
                    continue;
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get all tags.", ex);
        }
        return tags;
    }

    /// <summary>
    /// 导出标签到文件
    /// </summary>
    /// <param name="fileName">文件路径</param>
    public void ExportToFile(string fileName)
    {
        if (string.IsNullOrEmpty(fileName))
            throw new ArgumentException("File name cannot be null or empty.", nameof(fileName));

        try
        {
            var tags = GetAllTags();
            var lines = new List<string>();
            foreach (var kvp in tags)
            {
                lines.Add($"{kvp.Key}={kvp.Value}");
            }
            System.IO.File.WriteAllLines(fileName, lines);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to export tags to '{fileName}'.", ex);
        }
    }

    /// <summary>
    /// 从文件导入标签
    /// </summary>
    /// <param name="fileName">文件路径</param>
    public void ImportFromFile(string fileName)
    {
        if (string.IsNullOrEmpty(fileName))
            throw new ArgumentException("File name cannot be null or empty.", nameof(fileName));

        if (!System.IO.File.Exists(fileName))
            throw new System.IO.FileNotFoundException("Tag file not found.", fileName);

        try
        {
            var lines = System.IO.File.ReadAllLines(fileName);
            foreach (var line in lines)
            {
                if (!string.IsNullOrEmpty(line) && line.Contains("="))
                {
                    var parts = line.Split(new[] { '=' }, 2);
                    if (parts.Length == 2)
                    {
                        Update(parts[0].Trim(), parts[1].Trim());
                    }
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to import tags from '{fileName}'.", ex);
        }
    }

    /// <summary>
    /// 查找符合名称条件的标签
    /// </summary>
    /// <param name="namePredicate">名称条件</param>
    /// <returns>符合条件的标签列表</returns>
    public IEnumerable<KeyValuePair<string, string>> FindByName(Func<string, bool> namePredicate)
    {
        if (namePredicate == null)
            throw new ArgumentNullException(nameof(namePredicate));

        var results = new List<KeyValuePair<string, string>>();
        try
        {
            var allTags = GetAllTags();
            foreach (var kvp in allTags)
            {
                if (namePredicate(kvp.Key))
                {
                    results.Add(kvp);
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to find tags by name predicate.", ex);
        }
        return results;
    }

    /// <summary>
    /// 查找符合值条件的标签
    /// </summary>
    /// <param name="valuePredicate">值条件</param>
    /// <returns>符合条件的标签列表</returns>
    public IEnumerable<KeyValuePair<string, string>> FindByValue(Func<string, bool> valuePredicate)
    {
        if (valuePredicate == null)
            throw new ArgumentNullException(nameof(valuePredicate));

        var results = new List<KeyValuePair<string, string>>();
        try
        {
            var allTags = GetAllTags();
            foreach (var kvp in allTags)
            {
                if (valuePredicate(kvp.Value))
                {
                    results.Add(kvp);
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to find tags by value predicate.", ex);
        }
        return results;
    }

    /// <summary>
    /// 获取标签集合信息
    /// </summary>
    /// <returns>标签集合信息字符串</returns>
    public string GetTagsInfo()
    {
        try
        {
            return $"Tags - Count: {Count}";
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get tags info.", ex);
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
