//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;
/// <summary>
/// 表示 Office 中选取器结果的接口封装。
/// 该接口提供对单个选取器结果属性的访问。
/// </summary>
public interface IOfficePickerResult : IDisposable
{
    /// <summary>
    /// 获取或设置选取器结果的显示名称。
    /// </summary>
    string DisplayName { get; set; }

    /// <summary>
    /// 获取或设置选取器结果的类型。
    /// </summary>
    string Type { get; set; }

    /// <summary>
    /// 获取或设置选取器结果的 SIP ID。
    /// </summary>
    string SIPId { get; set; }

    /// <summary>
    /// 获取或设置子项集合。
    /// </summary>
    object? SubItems { get; set; }


    /// <summary>
    /// 获取重复结果集合。
    /// </summary>
    object? DuplicateResults { get; }


    /// <summary>
    /// 获取或设置与项关联的附加数据。
    /// </summary>
    object? ItemData { get; set; }


    /// <summary>
    /// 获取或设置选取器结果的字段集合。
    /// </summary>
    IOfficePickerFields? Fields { get; set; }
}