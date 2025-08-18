//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.ComponentModel;

namespace MudTools.OfficeInterop.Vbe.Imp;
/// <summary>
/// VBE VBComponent 对象的二次封装实现类
/// 实现 IVbeVBComponent 接口
/// </summary>
internal class VbeVBComponent : IVbeVBComponent
{

    internal MsVb.VBComponent _vbComponent;
    private bool _disposedValue = false;

    internal VbeVBComponent(MsVb.VBComponent vbComponent)
    {
        _vbComponent = vbComponent ?? throw new ArgumentNullException(nameof(vbComponent));
    }

    #region 基础属性
    public string Name
    {
        get => _vbComponent.Name;
        set => _vbComponent.Name = value;
    }


    public vbext_ComponentType Type => (vbext_ComponentType)_vbComponent.Type;


    public IVbeApplication Application => _vbComponent.VBE != null ? new VbeApplication(_vbComponent.VBE) : null;

    public IVbeCodeModule CodeModule => _vbComponent.CodeModule != null ? new VbeCodeModule(_vbComponent.CodeModule) : null;

    public object Designer => _vbComponent.Designer;

    public string DesignerID => _vbComponent.DesignerID;

    public bool IsSaved => _vbComponent.Saved;
    #endregion

    #region 操作方法
    public void Activate()
    {
        _vbComponent.Activate();
    }

    public void Export(string fileName)
    {
        _vbComponent.Export(fileName);
    }

    #endregion

    #region 代码操作

    public void ClearCode()
    {
        try
        {
            var codeMod = _vbComponent.CodeModule;
            if (codeMod != null)
            {
                codeMod.DeleteLines(1, codeMod.CountOfLines);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error clearing code: {ex.Message}");
        }
    }
    #endregion

    #region 格式设置
    public void SetProperty(string propertyName, object value)
    {
        try
        {
            if (Designer != null)
            {
                TypeDescriptor.GetProperties(Designer)[propertyName]?.SetValue(Designer, value);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error setting property '{propertyName}': {ex.Message}");
        }
    }

    public object GetProperty(string propertyName)
    {
        try
        {
            if (Designer != null)
            {
                return TypeDescriptor.GetProperties(Designer)[propertyName]?.GetValue(Designer);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error getting property '{propertyName}': {ex.Message}");
        }
        return null;
    }

    #endregion

    #region 导出和转换
    public string GetCodeText()
    {
        try
        {
            var codeMod = _vbComponent.CodeModule;
            if (codeMod != null && codeMod.CountOfLines > 0)
            {
                return codeMod.Lines[1, codeMod.CountOfLines];
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error getting code text: {ex.Message}");
        }
        return "";
    }

    public void SetCodeText(string codeText)
    {
        try
        {
            ClearCode();
            if (!string.IsNullOrEmpty(codeText))
            {
                _vbComponent.CodeModule.AddFromString(codeText);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error setting code text: {ex.Message}");
        }
    }

    public override string ToString()
    {
        return $"VBComponent: Name={this.Name}, Type={this.Type}, Saved={this.IsSaved}";
    }

    public byte[] GetBytes()
    {
        // VBA components are textual. Getting "bytes" usually means encoding the text.
        System.Diagnostics.Debug.WriteLine("Getting bytes for VBComponent (implies encoding text).");
        try
        {
            string code = GetCodeText();
            return System.Text.Encoding.UTF8.GetBytes(code);
        }
        catch
        {
            return new byte[0];
        }
    }
    #endregion


    #region IDisposable Support
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            _vbComponent = null;
            _disposedValue = true;
        }
    }

    ~VbeVBComponent()
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
