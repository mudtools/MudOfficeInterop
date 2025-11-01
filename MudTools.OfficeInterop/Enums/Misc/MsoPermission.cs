//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定 Office 文档的权限类型
/// </summary>
public enum MsoPermission
{

    /// <summary>
    /// 查看权限（与读取权限相同）
    /// </summary>
    msoPermissionView = 1,

    /// <summary>
    /// 读取权限（与查看权限相同）
    /// </summary>
    msoPermissionRead = 1,

    /// <summary>
    /// 编辑权限
    /// </summary>
    msoPermissionEdit = 2,

    /// <summary>
    /// 保存权限
    /// </summary>
    msoPermissionSave = 4,

    /// <summary>
    /// 提取权限
    /// </summary>
    msoPermissionExtract = 8,

    /// <summary>
    /// 更改权限（包含查看、编辑、保存和提取权限）
    /// </summary>
    msoPermissionChange = 15,

    /// <summary>
    /// 打印权限
    /// </summary>
    msoPermissionPrint = 16,

    /// <summary>
    /// 对象模型访问权限
    /// </summary>
    msoPermissionObjModel = 32,

    /// <summary>
    /// 完全控制权限
    /// </summary>
    msoPermissionFullControl = 64,

    /// <summary>
    /// 所有常见权限的组合
    /// </summary>
    msoPermissionAllCommon = 127
}