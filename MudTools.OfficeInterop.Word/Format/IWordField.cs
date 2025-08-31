using Microsoft.Office.Interop.Word;

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Field 的接口，用于操作Word域对象。
/// </summary>
public interface IWordField : IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取域的类型。
    /// </summary>
    WdFieldType Type { get; }

    WdFieldKind Kind { get; }

    /// <summary>
    /// 获取域的结果范围。
    /// </summary>
    IWordRange? ResultRange { get; }

    /// <summary>
    /// 获取域的代码范围。
    /// </summary>
    IWordRange? CodeRange { get; }

    /// <summary>
    /// 获取域的父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取或设置域是否锁定。
    /// </summary>
    bool Locked { get; set; }

    /// <summary>
    /// 获取域的索引号。
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取域的数据。
    /// </summary>
    string Data { get; set; }

    /// <summary>
    /// 获取域的结果文本。
    /// </summary>
    string Result { get; set; }

    /// <summary>
    /// 获取域的代码文本。
    /// </summary>
    string Code { get; set; }

    /// <summary>
    /// 获取域是否显示结果。
    /// </summary>
    bool ShowCodes { get; set; }

    /// <summary>
    /// 获取域的下一个域。
    /// </summary>
    IWordField? NextField { get; }

    /// <summary>
    /// 获取域的上一个域。
    /// </summary>
    IWordField? PreviousField { get; }

    /// <summary>
    /// 获取域是否为链接域。
    /// </summary>
    bool IsLinked { get; }

    /// <summary>
    /// 获取域的链接格式。
    /// </summary>
    IWordLinkFormat? LinkFormat { get; }

    /// <summary>
    /// 获取域的OLE格式。
    /// </summary>
    IWordOLEFormat? OLEFormat { get; }

    /// <summary>
    /// 更新域。
    /// </summary>
    /// <returns>是否更新成功。</returns>
    bool Update();

    /// <summary>
    /// 取消域的链接。
    /// </summary>
    void Unlink();

    /// <summary>
    /// 删除域。
    /// </summary>
    void Delete();

    /// <summary>
    /// 选择域。
    /// </summary>
    void Select();

    /// <summary>
    /// 复制域。
    /// </summary>
    void Copy();

    /// <summary>
    /// 剪切域。
    /// </summary>
    void Cut();

    /// <summary>
    /// 手动更新域。
    /// </summary>
    void DoClick();

    /// <summary>
    /// 验证域代码是否有效。
    /// </summary>
    /// <returns>域代码是否有效。</returns>
    bool ValidateCode();

    /// <summary>
    /// 获取域的源文件路径（如果是链接域）。
    /// </summary>
    /// <returns>源文件路径。</returns>
    string GetSourcePath();

    /// <summary>
    /// 设置域代码。
    /// </summary>
    /// <param name="code">新的域代码。</param>
    void SetCode(string code);

    /// <summary>
    /// 设置域结果。
    /// </summary>
    /// <param name="result">新的域结果。</param>
    void SetResult(string result);

    /// <summary>
    /// 获取域是否为日期域。
    /// </summary>
    bool IsDateField { get; }

    /// <summary>
    /// 获取域是否为页码域。
    /// </summary>
    bool IsPageField { get; }

    /// <summary>
    /// 获取域是否为目录域。
    /// </summary>
    bool IsTOCField { get; }
}