//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word 文档自定义属性集合实现类
/// </summary>
internal class WordCustomProperties : IWordCustomProperties
{
    private readonly object _customProperties; // 使用 object 类型因为可能需要特殊处理
    private readonly IWordDocument _document;
    private bool _disposedValue;

    /// <summary>
    /// 获取自定义属性数量
    /// </summary>
    public int Count
    {
        get
        {
            try
            {
                // 由于 CustomDocumentProperties 是动态属性，需要特殊处理
                var properties = _customProperties as System.Collections.IEnumerable;
                if (properties == null) return 0;

                var count = 0;
                foreach (var prop in properties)
                {
                    count++;
                }
                return count;
            }
            catch
            {
                return 0;
            }
        }
    }

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="customProperties">COM CustomProperties 对象</param>
    /// <param name="document">关联的文档对象</param>
    internal WordCustomProperties(object customProperties, IWordDocument document)
    {
        _customProperties = customProperties ?? throw new ArgumentNullException(nameof(customProperties));
        _document = document ?? throw new ArgumentNullException(nameof(document));
        _disposedValue = false;
    }

    /// <summary>
    /// 根据索引获取自定义属性
    /// </summary>
    /// <param name="index">自定义属性索引</param>
    /// <returns>自定义属性对象</returns>
    public IWordCustomProperty Item(int index)
    {
        try
        {
            // 由于 CustomDocumentProperties 需要特殊处理，这里使用反射
            var properties = _customProperties as System.Collections.IEnumerable;
            if (properties == null)
                throw new InvalidOperationException("Invalid custom properties object.");

            var enumerator = properties.GetEnumerator();
            for (int i = 1; i <= index && enumerator.MoveNext(); i++)
            {
                if (i == index)
                {
                    return new WordCustomProperty(enumerator.Current);
                }
            }

            throw new ArgumentOutOfRangeException(nameof(index), $"Index must be between 1 and {Count}.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to get custom property at index {index}.", ex);
        }
    }

    /// <summary>
    /// 根据名称获取自定义属性
    /// </summary>
    /// <param name="name">自定义属性名称</param>
    /// <returns>自定义属性对象</returns>
    public IWordCustomProperty Item(string name)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Custom property name cannot be null or empty.", nameof(name));

        try
        {
            // 使用反射获取指定名称的属性
            var properties = _customProperties as System.Collections.IEnumerable;
            if (properties == null)
                throw new InvalidOperationException("Invalid custom properties object.");

            foreach (var prop in properties)
            {
                var nameProperty = prop.GetType().InvokeMember("Name", System.Reflection.BindingFlags.GetProperty, null, prop, null);
                if (nameProperty?.ToString() == name)
                {
                    return new WordCustomProperty(prop);
                }
            }

            throw new InvalidOperationException($"Custom property '{name}' not found.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to get custom property with name '{name}'.", ex);
        }
    }

    /// <summary>
    /// 添加自定义属性
    /// </summary>
    /// <param name="name">属性名称</param>
    /// <param name="value">属性值</param>
    /// <returns>自定义属性对象</returns>
    public IWordCustomProperty Add(string name, object value)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Custom property name cannot be null or empty.", nameof(name));

        try
        {
            // 使用反射调用 Add 方法
            var type = _customProperties.GetType();
            var parameters = new object[] { name, false, value, missing, missing };
            var property = type.InvokeMember("Add", System.Reflection.BindingFlags.InvokeMethod, null, _customProperties, parameters);
            return new WordCustomProperty(property);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to add custom property '{name}'.", ex);
        }
    }

    /// <summary>
    /// 删除自定义属性
    /// </summary>
    /// <param name="name">属性名称</param>
    public void Delete(string name)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Custom property name cannot be null or empty.", nameof(name));

        try
        {
            if (Exists(name))
            {
                var property = Item(name);
                property.Delete();
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to delete custom property '{name}'.", ex);
        }
    }

    /// <summary>
    /// 检查自定义属性是否存在
    /// </summary>
    /// <param name="name">属性名称</param>
    /// <returns>是否存在</returns>
    public bool Exists(string name)
    {
        if (string.IsNullOrEmpty(name))
            return false;

        try
        {
            var properties = _customProperties as System.Collections.IEnumerable;
            if (properties == null) return false;

            foreach (var prop in properties)
            {
                var nameProperty = prop.GetType().InvokeMember("Name", System.Reflection.BindingFlags.GetProperty, null, prop, null);
                if (nameProperty?.ToString() == name)
                {
                    return true;
                }
            }
            return false;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// 获取枚举器
    /// </summary>
    /// <returns>自定义属性枚举器</returns>
    public IEnumerator<IWordCustomProperty> GetEnumerator()
    {
        try
        {
            var properties = new List<IWordCustomProperty>();
            var customProps = _customProperties as System.Collections.IEnumerable;
            if (customProps == null) return properties.GetEnumerator();

            foreach (var prop in customProps)
            {
                try
                {
                    properties.Add(new WordCustomProperty(prop));
                }
                catch
                {
                    // 忽略获取失败的属性
                    continue;
                }
            }
            return properties.GetEnumerator();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to enumerate custom properties.", ex);
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

    private static readonly object missing = System.Reflection.Missing.Value;

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
