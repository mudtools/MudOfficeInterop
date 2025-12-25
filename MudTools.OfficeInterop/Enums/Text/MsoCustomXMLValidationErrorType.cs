//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指示将如何清除或生成验证错误
/// </summary>
public enum MsoCustomXMLValidationErrorType
{
    /// <summary>
    /// 指定如果有可用于自定义XML部件的非空架构集合且验证生效，则对该部件的任何更改都会导致验证错误
    /// </summary>
    msoCustomXMLValidationErrorSchemaGenerated,

    /// <summary>
    /// 指定只要对错误所绑定的节点进行任何更改，错误就会自动清除
    /// </summary>
    msoCustomXMLValidationErrorAutomaticallyCleared,

    /// <summary>
    /// 指定在调用 Microsoft.Office.Core.CustomXMLValidationError.Delete 方法之前，错误不会清除
    /// </summary>
    msoCustomXMLValidationErrorManual
}