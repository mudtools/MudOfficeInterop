namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 中的单个“自动图文集”条目（AutoText Entry）的封装接口。
/// 自动图文集条目允许用户通过简短名称快速插入预定义内容（如段落、表格等）。
/// </summary>
public interface IWordAutoTextEntry : IDisposable
{
    /// <summary>
    /// 获取自动图文集条目的名称（如 "Address"）
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取或设置自动图文集条目的值（即插入的内容）
    /// </summary>
    string Value { get; set; }

    /// <summary>
    ///获取此自动图文集条目样式名
    /// </summary>
    string StyleName { get; }


    /// <summary>
    /// 从模板中删除此自动图文集条目
    /// </summary>
    void Delete();
}