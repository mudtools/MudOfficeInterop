//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示对 Microsoft.Office.Interop.Word.BuildingBlockType 对象的封装接口。
/// 用于表示构建基块的类型，例如“页眉”、“页脚”、“自动图文集”等。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordBuildingBlockType : IOfficeObject<IWordBuildingBlockType>, IDisposable
{
    /// <summary>
    /// 获取该对象关联的Word应用程序实例
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取该对象的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取构建基块类型的唯一索引（对应 WdBuiltinBuildingBlockTypes 枚举值）。
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取构建基块类型的显示名称（如“页眉”、“页脚”）。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取该类型下的所有类别（Categories）集合。
    /// 每个类别包含一组同名分组的构建基块（如“常规”、“公司专用”等）。
    /// </summary>
    IWordCategories? Categories { get; }

}