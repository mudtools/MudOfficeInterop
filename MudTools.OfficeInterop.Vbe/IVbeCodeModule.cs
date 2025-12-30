//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Vbe;
/// <summary>
/// VBE CodeModule 对象的二次封装接口
/// 提供对 Microsoft.Vbe.Interop.CodeModule 的安全访问和操作
/// </summary>
[ComObjectWrap(ComNamespace = "MsVb")]
public interface IVbeCodeModule : IOfficeObject<IVbeCodeModule>, IDisposable
{
    /// <summary>
    /// 获取此代码模块的父对象（VB 组件）。
    /// </summary>
    IVbeVBComponent? Parent { get; }

    /// <summary>
    /// 获取表示 VBA 编辑器环境的 VBE 对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IVbeApplication? VBE { get; }
    /// <summary>
    /// 获取或设置代码模块的名称。
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 从字符串向代码模块添加代码。
    /// </summary>
    /// <param name="codeString">要添加的代码字符串。</param>
    void AddFromString(string codeString);

    /// <summary>
    /// 从文件向代码模块添加代码。
    /// </summary>
    /// <param name="fileName">包含代码的文件路径。</param>
    void AddFromFile(string fileName);

    /// <summary>
    /// 获取代码模块中的所有代码行。
    /// </summary>
    [MethodIndex]
    string? Lines(int startLine, int count = 0);

    /// <summary>
    /// 获取代码模块中的总行数。
    /// </summary>
    int CountOfLines { get; }

    /// <summary>
    /// 在指定行插入代码。
    /// </summary>
    /// <param name="line">要插入代码的行号。</param>
    /// <param name="codeString">要插入的代码字符串。</param>
    void InsertLines(int line, string codeString);

    /// <summary>
    /// 删除从指定行开始的若干行代码。
    /// </summary>
    /// <param name="startLine">起始行号。</param>
    /// <param name="count">要删除的行数，默认为 1。</param>
    void DeleteLines(int startLine, int count = 1);

    /// <summary>
    /// 替换指定行的代码。
    /// </summary>
    /// <param name="line">要替换的行号。</param>
    /// <param name="codeString">新的代码字符串。</param>
    void ReplaceLine(int line, string codeString);

    /// <summary>
    /// 获取当前过程的起始行号。
    /// </summary>
    [MethodIndex]
    int? ProcStartLine(string proName, vbext_ProcKind kind);
    /// <summary>
    /// 获取当前过程的总行数。
    /// </summary>
    [MethodIndex]
    int? ProcCountLines(string proName, vbext_ProcKind kind);

    /// <summary>
    /// 获取当前过程体的起始行号。
    /// </summary>
    [MethodIndex]
    int? ProcBodyLine(string proName, vbext_ProcKind kind);

    /// <summary>
    /// 获取指定行所在的过程名称。
    /// </summary>
    [MethodIndex]
    string? ProcOfLine(int line, out vbext_ProcKind kind);

    /// <summary>
    /// 获取声明部分的行数。
    /// </summary>
    int CountOfDeclarationLines { get; }

    /// <summary>
    /// 创建事件过程并返回新过程的起始行号。
    /// </summary>
    /// <param name="eventName">事件名称。</param>
    /// <param name="objectName">对象名称。</param>
    /// <returns>新创建事件过程的起始行号。</returns>
    int? CreateEventProc(string eventName, string objectName);

    /// <summary>
    /// 在代码模块中查找指定的文本。
    /// </summary>
    /// <param name="target">要查找的目标文本。</param>
    /// <param name="startLine">输入/输出参数：查找起始行号。</param>
    /// <param name="startColumn">输入/输出参数：查找起始列号。</param>
    /// <param name="endLine">输入/输出参数：查找结束行号。</param>
    /// <param name="endColumn">输入/输出参数：查找结束列号。</param>
    /// <param name="wholeWord">是否全词匹配，默认为 false。</param>
    /// <param name="matchCase">是否区分大小写，默认为 false。</param>
    /// <param name="patternSearch">是否使用模式匹配，默认为 false。</param>
    /// <returns>如果找到目标文本，则为 true；否则为 false。</returns>
    bool? Find(string target, int startLine, int startColumn, int endLine, int endColumn, bool wholeWord = false, bool matchCase = false, bool patternSearch = false);

    /// <summary>
    /// 获取与此代码模块关联的代码窗格。
    /// </summary>
    IVbeCodePane? CodePane { get; }
}
