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
public interface IVbeCodeModule : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取代码模块的父对象（通常是 VBComponent）
    /// 对应 CodeModule.Parent 属性
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取代码模块所在的Application对象（VBE 对象）
    /// 对应 CodeModule.Application 属性
    /// </summary>
    IVbeApplication Application { get; }

    /// <summary>
    /// 获取代码模块的名称（通常是其父 VBComponent 的名称）
    /// 对应 CodeModule.Name 属性
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取代码模块中的总行数
    /// 对应 CodeModule.CountOfLines 属性
    /// </summary>
    int CountOfLines { get; }

    /// <summary>
    /// 获取代码模块中声明部分的行数
    /// 对应 CodeModule.CountOfDeclarationLines 属性
    /// </summary>
    int CountOfDeclarationLines { get; }

    /// <summary>
    /// 获取代码模块的编程语言（例如，VB）
    /// </summary>
    string Language { get; }
    #endregion

    #region 状态属性
    /// <summary>
    /// 获取代码模块是否为空（无任何代码行）
    /// </summary>
    bool IsEmpty { get; }
    #endregion

    #region 代码访问
    /// <summary>
    /// 获取指定行号的代码文本
    /// 对应 CodeModule.Lines 属性 (索引器)
    /// </summary>
    /// <param name="startLine">起始行号 (从1开始)</param>
    /// <param name="count">要获取的行数</param>
    /// <returns>代码文本</returns>
    string GetLines(int startLine, int count = 1);

    /// <summary>
    /// 获取代码模块的所有代码文本
    /// </summary>
    /// <returns>完整的代码文本</returns>
    string GetAllCode();
    #endregion

    #region 操作方法

    /// <summary>
    /// 添加代码文本到模块末尾
    /// 对应 CodeModule.AddFromString 方法
    /// </summary>
    /// <param name="codeText">要添加的代码文本</param>
    void AddFromString(string codeText);
    /// <summary>
    /// 选择代码模块（通常在 VBE 中激活其父组件）
    /// </summary>
    /// <param name="replace">是否替换当前选择</param>
    void Select(bool replace = true);

    /// <summary>
    /// 删除指定范围的代码行
    /// 对应 CodeModule.DeleteLines 方法
    /// </summary>
    /// <param name="startLine">起始行号 (从1开始)</param>
    /// <param name="count">要删除的行数</param>
    void DeleteLines(int startLine, int count = 1);

    /// <summary>
    /// 在指定行插入代码文本
    /// 对应 CodeModule.InsertLines 方法
    /// </summary>
    /// <param name="lineNumber">插入位置的行号 (从1开始)</param>
    /// <param name="codeText">要插入的代码文本</param>
    void InsertLines(int lineNumber, string codeText);

    /// <summary>
    /// 替换指定范围的代码文本
    /// </summary>
    /// <param name="startLine">起始行号 (从1开始)</param>
    /// <param name="count">要替换的行数</param>
    /// <param name="newCodeText">新的代码文本</param>
    void ReplaceLines(int startLine, int count, string newCodeText);

    /// <summary>
    /// 添加代码文本到模块末尾
    /// 对应 CodeModule.AddFromString 方法
    /// </summary>
    /// <param name="codeText">要添加的代码文本</param>
    void AddCode(string codeText);

    /// <summary>
    /// 从文件添加代码文本到模块末尾
    /// 对应 CodeModule.AddFromFile 方法
    /// </summary>
    /// <param name="fileName">代码文件路径</param>
    void AddCodeFromFile(string fileName);

    /// <summary>
    /// 清除代码模块中的所有代码
    /// </summary>
    void Clear();
    #endregion

    #region 查找和替换
    /// <summary>
    /// 在代码模块中查找文本
    /// 对应 CodeModule.Find 方法
    /// </summary>
    /// <param name="target">要查找的文本</param>
    /// <param name="startLine">起始行号</param>
    /// <param name="startColumn">起始列号</param>
    /// <param name="endLine">结束行号</param>
    /// <param name="endColumn">结束列号</param>
    /// <param name="wholeWord">是否匹配整个单词</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <param name="patternSearch">是否使用模式搜索</param>
    /// <returns>找到的文本范围 (行号, 列号) 或 (-1, -1) 表示未找到</returns>
    bool Find(string target, int startLine = 1, int startColumn = 1,
                                int endLine = int.MaxValue, int endColumn = int.MaxValue,
                                bool wholeWord = false, bool matchCase = false, bool patternSearch = false);

    #endregion

    #region 导出和导入
    /// <summary>
    /// 将代码模块导出到文本文件
    /// 对应 CodeModule.Export 方法 (通过父 VBComponent)
    /// </summary>
    /// <param name="fileName">导出文件路径</param>
    void Export(string fileName);

    /// <summary>
    /// 将代码模块转换为字符串表示（包含所有代码）
    /// </summary>
    /// <returns>字符串表示</returns>
    string ToString();
    #endregion

}
