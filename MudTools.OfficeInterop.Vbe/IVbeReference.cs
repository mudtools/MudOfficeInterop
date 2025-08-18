//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Vbe;
/// <summary>
/// VBE Reference 对象的二次封装接口
/// 提供对 Microsoft.Vbe.Interop.Reference 的安全访问和操作
/// </summary>
public interface IVbeReference : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取引用的名称
    /// 对应 Reference.Name 属性
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取引用的完整路径
    /// 对应 Reference.FullPath 属性
    /// </summary>
    string FullPath { get; }

    /// <summary>
    /// 获取引用的 GUID
    /// 对应 Reference.Guid 属性
    /// </summary>
    string Guid { get; }

    /// <summary>
    /// 获取引用的主版本号
    /// 对应 Reference.Major 属性
    /// </summary>
    int Major { get; }

    /// <summary>
    /// 获取引用的次版本号
    /// 对应 Reference.Minor 属性
    /// </summary>
    int Minor { get; }

    /// <summary>
    /// 获取引用的描述
    /// 对应 Reference.Description 属性
    /// </summary>
    string Description { get; }

    /// <summary>
    /// 获取引用所在的Application对象（VBE 对象）
    /// 对应 Reference.Application 属性
    /// </summary>
    IVbeApplication Application { get; }
    #endregion

    #region 状态属性
    /// <summary>
    /// 获取引用是否为内置引用
    /// </summary>
    bool IsBuiltIn { get; }

    /// <summary>
    /// 获取引用是否已破损
    /// 对应 Reference.IsBroken 属性
    /// </summary>
    bool IsBroken { get; }

    /// <summary>
    /// 获取引用是否被保护
    /// </summary>
    bool IsProtected { get; }
    #endregion
}
