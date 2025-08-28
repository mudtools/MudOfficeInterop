namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示对 Microsoft.Office.Interop.Word.BuildingBlockType 对象的封装接口。
/// 用于表示构建基块的类型，例如“页眉”、“页脚”、“自动图文集”等。
/// </summary>
public interface IWordBuildingBlockType : IDisposable
{
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
    IWordCategories Categories { get; }

    /// <summary>
    /// 获取该类型下所有构建基块的总数（跨所有类别）。
    /// </summary>
    int TotalBlockCount { get; }
}