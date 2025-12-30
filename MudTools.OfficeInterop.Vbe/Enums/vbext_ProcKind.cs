
namespace MudTools.OfficeInterop.Vbe;

/// <summary>
/// 指定过程类型
/// </summary>
public enum vbext_ProcKind
{
    /// <summary>
    /// 普通过程（Sub 或 Function）
    /// </summary>
    vbext_pk_Proc,

    /// <summary>
    /// Let 属性过程
    /// </summary>
    vbext_pk_Let,

    /// <summary>
    /// Set 属性过程
    /// </summary>
    vbext_pk_Set,

    /// <summary>
    /// Get 属性过程
    /// </summary>
    vbext_pk_Get
}