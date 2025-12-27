//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word 文档自定义属性实现类
/// </summary>
internal class WordCustomProperty : IWordCustomProperty
{
    private readonly object _customProperty;
    private bool _disposedValue;


    /// <summary>
    /// 获取属性名称
    /// </summary>
    public string Name
    {
        get
        {
            try
            {
                return _customProperty.GetType().InvokeMember("Name", System.Reflection.BindingFlags.GetProperty, null, _customProperty, null)?.ToString() ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
    }

    /// <summary>
    /// 获取或设置属性值
    /// </summary>
    public object Value
    {
        get
        {
            try
            {
                return _customProperty.GetType().InvokeMember("Value", System.Reflection.BindingFlags.GetProperty, null, _customProperty, null);
            }
            catch
            {
                return null;
            }
        }
        set
        {
            try
            {
                _customProperty.GetType().InvokeMember("Value", System.Reflection.BindingFlags.SetProperty, null, _customProperty, new object[] { value });
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to set custom property value.", ex);
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
                return _customProperty.GetType().InvokeMember("Parent", System.Reflection.BindingFlags.GetProperty, null, _customProperty, null);
            }
            catch
            {
                return null;
            }
        }
    }

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="customProperty">COM CustomProperty 对象</param>
    internal WordCustomProperty(object customProperty)
    {
        _customProperty = customProperty ?? throw new ArgumentNullException(nameof(customProperty));
        _disposedValue = false;
    }

    /// <summary>
    /// 删除自定义属性
    /// </summary>
    public void Delete()
    {
        try
        {
            _customProperty.GetType().InvokeMember("Delete", System.Reflection.BindingFlags.InvokeMethod, null, _customProperty, null);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to delete custom property.", ex);
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

