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