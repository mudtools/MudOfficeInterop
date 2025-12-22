//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示邮件合并中一个标准字段与数据源字段之间映射关系的二次封装接口。
/// 此接口允许获取或设置与预定义标准字段（如“姓氏”）关联的数据源字段名称 [[1]]。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordMappedDataField : IDisposable
{
    /// <summary>
    /// 获取此映射字段所属的 Word 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此映射字段的父对象（通常是 <see cref="IWordMappedDataFields"/> 集合）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取此映射所对应的标准字段的名称（例如 "LastName", "Address1"）。
    /// 这些是 Word 预定义的常量名称，为只读属性。
    /// </summary>
    string? Name { get; }

    /// <summary>
    /// 获取与该标准字段关联的数据源字段值。
    /// </summary>
    string Value { get; }


    /// <summary>
    /// 获取或设置与该标准字段关联的数据源字段索引。
    /// </summary>
    int DataFieldIndex { get; }

    /// <summary>
    /// 获取与该标准字段映射的数据源中的实际字段名称。
    /// </summary>
    string DataFieldName { get; }

}