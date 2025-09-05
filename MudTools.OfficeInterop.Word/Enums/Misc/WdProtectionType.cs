//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// Word 文档保护类型枚举
/// 用于指定文档的保护方式
/// </summary>
public enum WdProtectionType
{
    /// <summary>
    /// 无保护 - 文档完全可编辑
    /// </summary>
    wdNoProtection = -1,
    
    /// <summary>
    /// 仅允许修订 - 用户只能添加或删除修订
    /// </summary>
    wdAllowOnlyRevisions = 0,
    
    /// <summary>
    /// 仅允许批注 - 用户只能插入或修改批注
    /// </summary>
    wdAllowOnlyComments = 1,
    
    /// <summary>
    /// 仅允许填写窗体 - 用户只能在窗体字段中输入内容
    /// </summary>
    wdAllowOnlyFormFields = 2,
    
    /// <summary>
    /// 只读模式 - 用户只能查看文档，不能进行任何修改
    /// </summary>
    wdAllowOnlyReading = 3
}