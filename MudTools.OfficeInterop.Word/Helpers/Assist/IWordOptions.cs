//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Microsoft Word 中的应用程序设置。
/// <para>注：使用 Application.Options 属性可返回 Options 对象。</para>
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordOptions : IDisposable
{
    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    #region 常用属性

    /// <summary>
    /// 获取或设置是否允许拖放文本功能
    /// 对应 Word 选项：高级 -> 编辑选项 -> 允许拖放式文字编辑
    /// </summary>
    bool AllowDragAndDrop { get; set; }

    /// <summary>
    /// 获取或设置是否在键入时自动创建画布
    /// 对应 Word 选项：高级 -> 编辑选项 -> 插入自选图形时自动创建绘图画布
    /// </summary>
    bool AutoCreateNewDrawings { get; set; }

    /// <summary>
    /// 获取或设置是否启用实时预览功能
    /// 对应 Word 选项：常规 -> 用户界面选项 -> 启用实时预览
    /// </summary>
    bool EnableLivePreview { get; set; }

    #endregion

    #region 视图与显示属性

    /// <summary>
    /// 获取或设置是否在文档中显示网格线。
    /// 对应 Word 选项：视图 -> 显示 -> 网格线
    /// </summary>
    bool DisplayGridLines { get; set; }

    #endregion

    #region 编辑与输入选项

    /// <summary>
    /// 获取或设置键入时是否替换选定的文本。
    /// 对应 Word 选项：文件 -> 选项 -> 高级 -> 编辑选项 -> 键入内容替换所选内容
    /// </summary>
    bool ReplaceSelection { get; set; }

    /// <summary>
    /// 获取或设置键入时是否自动应用标题样式。
    /// 对应 Word 选项：文件 -> 选项 -> 校对 -> 自动更正选项 -> 键入时自动套用格式 -> 标题
    /// </summary>
    bool AutoFormatAsYouTypeApplyHeadings { get; set; }

    #endregion

    #region 保存与备份选项

    /// <summary>
    /// 获取或设置关闭 Word 时是否提示保存 Normal 模板。
    /// 对应 Word 选项：文件 -> 选项 -> 高级 -> 保存 -> 提示保存 Normal 模板
    /// </summary>
    bool SaveNormalPrompt { get; set; }

    /// <summary>
    /// 获取或设置保存文档时是否提示输入文档属性。
    /// 对应 Word 选项：文件 -> 选项 -> 保存 -> 文档管理服务器文件 -> 保存时提示输入文档属性
    /// </summary>
    bool SavePropertiesPrompt { get; set; }

    #endregion

    #region 打印与输出选项

    /// <summary>
    /// 获取或设置打印时是否在文档中更新字段。
    /// 对应 Word 选项：文件 -> 选项 -> 显示 -> 打印选项 -> 打印前更新域
    /// </summary>
    bool UpdateFieldsAtPrint { get; set; }

    /// <summary>
    /// 获取或设置打印时是否更新链接数据。
    /// 对应 Word 选项：文件 -> 选项 -> 高级 -> 打印 -> 更新链接数据
    /// </summary>
    bool UpdateLinksAtPrint { get; set; }

    /// <summary>
    /// 获取或设置打印时是否打印隐藏文本。
    /// 对应 Word 选项：文件 -> 选项 -> 显示 -> 打印选项 -> 打印隐藏文字
    /// </summary>
    bool PrintHiddenText { get; set; }

    /// <summary>
    /// 获取或设置打印时是否使用草稿质量输出。
    /// 对应 Word 选项：文件 -> 选项 -> 高级 -> 打印 -> 使用草稿品质
    /// </summary>
    bool PrintDraft { get; set; }

    /// <summary>
    /// 获取或设置打印时是否逆页序打印。
    /// 对应 Word 选项：文件 -> 选项 -> 高级 -> 打印 -> 逆序打印页面
    /// </summary>
    bool PrintReverse { get; set; }

    #endregion

    #region 语言与校对选项

    /// <summary>
    /// 获取或设置是否检查拼写错误。
    /// 对应 Word 选项：文件 -> 选项 -> 校对 -> 在 Word 中更正拼写和语法时 -> 键入时检查拼写
    /// </summary>
    bool CheckSpellingAsYouType { get; set; }

    /// <summary>
    /// 获取或设置是否检查语法错误。
    /// 对应 Word 选项：文件 -> 选项 -> 校对 -> 在 Word 中更正拼写和语法时 -> 键入时标记语法错误
    /// </summary>
    bool CheckGrammarAsYouType { get; set; }

    /// <summary>
    /// 获取或设置是否忽略全部大写的单词。
    /// 对应 Word 选项：文件 -> 选项 -> 校对 -> 在 Word 中更正拼写和语法时 -> 忽略全部大写的单词
    /// </summary>
    bool IgnoreUppercase { get; set; }

    /// <summary>
    /// 获取或设置是否忽略包含数字的单词。
    /// 对应 Word 选项：文件 -> 选项 -> 校对 -> 在 Word 中更正拼写和语法时 -> 忽略含数字的单词
    /// </summary>
    bool IgnoreMixedDigits { get; set; }

    #endregion

    #region 高级选项

    /// <summary>
    /// 获取或设置是否启用声音反馈。
    /// 对应 Word 选项：文件 -> 选项 -> 轻松访问 -> 反馈选项 -> 提供声音反馈
    /// </summary>
    bool EnableSound { get; set; }

    #endregion    

    #region 修订与跟踪选项

    /// <summary>
    /// 获取或设置插入内容的修订标记颜色。
    /// 对应 Word 选项：审阅 -> 跟踪 -> 修订选项 -> 插入内容
    /// </summary>
    WdColorIndex InsertedTextColor { get; set; }

    /// <summary>
    /// 获取或设置删除内容的修订标记颜色。
    /// 对应 Word 选项：审阅 -> 跟踪 -> 修订选项 -> 删除内容
    /// </summary>
    WdColorIndex DeletedTextColor { get; set; }

    /// <summary>
    /// 获取或设置修订行的标记颜色。
    /// 对应 Word 选项：审阅 -> 跟踪 -> 修订选项 -> 修订行
    /// </summary>
    WdColorIndex RevisedLinesColor { get; set; }
    #endregion

    #region 高级编辑选项

    /// <summary>
    /// 获取或设置是否允许使用 Ctrl+单击跟踪超链接。
    /// 对应 Word 选项：文件 -> 选项 -> 高级 -> 编辑选项 -> 用 Ctrl+单击跟踪超链接
    /// </summary>
    bool CtrlClickHyperlinkToOpen { get; set; }

    /// <summary>
    /// 获取或设置是否允许法语中大写字母带重音符号。
    /// 对应 Word 选项：文件 -> 选项 -> 高级 -> 编辑选项 -> 允许法语中大写字母带重音符号
    /// </summary>
    bool AllowAccentedUppercase { get; set; }

    #endregion
}