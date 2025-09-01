//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示对 Microsoft Word 中 Category（类别）对象的封装接口。
/// 用于表示构建基块或自动图文集中内容的分类，如“常规”、“页眉”等。
/// </summary>
public interface IWordCategory : IDisposable
{
    /// <summary>
    /// 获取类别的名称（例如：“常规”、“自定义类别1”等）。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取类别中的构建基块集合。
    /// 构建基块是可重复使用的内容单元，可以根据类别进行组织和管理。
    /// </summary>
    IWordBuildingBlocks BuildingBlocks { get; }

    /// <summary>
    /// 获取类别在集合中的索引位置。
    /// 索引从1开始计数，用于标识该类别在所有类别中的排序位置。
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取类别关联的构建基块类型。
    /// 表示该类别中包含的构建基块的类型信息。
    /// </summary>
    IWordBuildingBlockType Type { get; }
}