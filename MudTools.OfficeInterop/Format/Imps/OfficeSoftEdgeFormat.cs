//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imps;
/// <summary>
/// Excel SoftEdgeFormat COM 对象的封装实现类。
/// 负责管理 COM 对象生命周期，提供安全的属性访问和资源释放。
/// </summary>
internal class OfficeSoftEdgeFormat : IOfficeSoftEdgeFormat
{
    /// <summary>
    /// 内部持有的原始 COM 对象。
    /// </summary>
    internal MsCore.SoftEdgeFormat _softEdgeFormat;

    /// <summary>
    /// 标记对象是否已被释放。
    /// </summary>
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装类。
    /// </summary>
    /// <param name="softEdgeFormat">原始的 Excel SoftEdgeFormat COM 对象，不可为 null。</param>
    /// <exception cref="ArgumentNullException">当传入的 softEdgeFormat 为 null 时抛出。</exception>
    internal OfficeSoftEdgeFormat(MsCore.SoftEdgeFormat softEdgeFormat)
    {
        _softEdgeFormat = softEdgeFormat ?? throw new ArgumentNullException(nameof(softEdgeFormat));
        _disposedValue = false;
    }

    /// <summary>
    /// 释放资源的受保护虚方法，支持派生类重写。
    /// </summary>
    /// <param name="disposing">是否由用户代码显式调用释放。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            // 释放托管资源：释放 COM 对象
            if (_softEdgeFormat != null)
            {
                Marshal.ReleaseComObject(_softEdgeFormat);
                _softEdgeFormat = null;
            }
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 公开的 Dispose 方法，用于显式释放资源。
    /// 调用后对象不应再被使用。
    /// </summary>
    public void Dispose() => Dispose(true);


    /// <summary>
    /// 获取或设置柔化边缘的类型。
    /// 默认值为 <see cref="MsoSoftEdgeType.msoSoftEdgeTypeMixed"/>。
    /// 设置时自动转换为底层 COM 枚举类型。
    /// </summary>
    public MsoSoftEdgeType Type
    {
        get => _softEdgeFormat != null
            ? _softEdgeFormat.Type.EnumConvert(MsoSoftEdgeType.msoSoftEdgeTypeMixed)
            : MsoSoftEdgeType.msoSoftEdgeTypeMixed;

        set
        {
            if (_softEdgeFormat != null)
            {
                _softEdgeFormat.Type = value.EnumConvert(MsCore.MsoSoftEdgeType.msoSoftEdgeTypeMixed);
            }
        }
    }

    /// <summary>
    /// 获取或设置柔化边缘的半径（单位：磅）。
    /// 值必须 >= 0。值越大，边缘越模糊。
    /// 若 COM 对象为空，则设置无效，获取返回 0。
    /// </summary>
    public float Radius
    {
        get => _softEdgeFormat != null ? _softEdgeFormat.Radius : 0f;

        set
        {
            if (_softEdgeFormat != null)
            {
                _softEdgeFormat.Radius = value;
            }
        }
    }
}