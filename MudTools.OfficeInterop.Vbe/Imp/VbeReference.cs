//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Vbe.Imp;
/// <summary>
/// VBE Reference 对象的二次封装实现类
/// 实现 IVbeReference 接口
/// </summary>
internal class VbeReference : IVbeReference
{
    internal MsVb.Reference _reference; // internal for VbeReferences.Delete(IVbeReference)
    private bool _disposedValue = false;

    internal VbeReference(MsVb.Reference reference)
    {
        _reference = reference ?? throw new ArgumentNullException(nameof(reference));
    }

    #region 基础属性
    public string Name => _reference.Name;


    public string FullPath => _reference.FullPath;

    public string Guid => _reference.Guid;

    public int Major => _reference.Major;

    public int Minor => _reference.Minor;

    public string Description => _reference.Description;

    public IVbeApplication Application => _reference.VBE != null ? new VbeApplication(_reference.VBE) : null;
    #endregion

    #region 状态属性
    public bool IsBuiltIn
    {
        get
        {
            try
            {
                string pathLower = this.FullPath.ToLowerInvariant();
                string guidLower = this.Guid.ToLowerInvariant();
                if (pathLower.Contains("\\vba") || pathLower.Contains("\\office") ||
                    guidLower.StartsWith("{000204ef-0000-0000-c000-000000000046}") || // VBA
                    guidLower.StartsWith("{2df8d04c-5bfa-101b-bde5-00aa0044de52}") || // Office
                    guidLower.StartsWith("{00020813-0000-0000-c000-000000000046}")    // Excel
                   )
                {
                    return true;
                }
            }
            catch { /* Ignore errors in heuristic */ }
            return false;
        }
    }

    public bool IsBroken => _reference.IsBroken;

    public bool IsProtected => false; // Placeholder, references themselves aren't typically "protected"

    #endregion


    #region IDisposable Support
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            _reference = null;

            _disposedValue = true;
        }
    }

    ~VbeReference()
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
