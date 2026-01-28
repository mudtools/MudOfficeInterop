//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// PowerPoint Selection 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.PowerPoint.Selection 的安全访问和操作
/// </summary>
public interface IPowerPointSelection : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取 Selection 对象的父对象（通常是 Application 或 Window）
    /// 对应 Selection.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取 Selection 对象所在的 Application 对象
    /// 对应 Selection.Application 属性
    /// </summary>
    IPowerPointApplication Application { get; }

    /// <summary>
    /// 获取选择的类型
    /// 对应 Selection.Type 属性
    /// </summary>
    PpSelectionType Type { get; }

    /// <summary>
    /// 获取选择中项目的数量
    /// 对应 Selection.ShapeRange.Count 或 Selection.SlideRange.Count (取决于 Type)
    /// </summary>
    int Count { get; }
    #endregion

    #region 状态属性
    /// <summary>
    /// 获取选择是否为空 (无任何项目被选中)
    /// </summary>
    bool IsEmpty { get; }
    #endregion

    #region 核心对象 (根据选择类型动态返回)
    /// <summary>
    /// 获取选择的形状范围 (当 Type 为 ppSelectionShapes 或 ppSelectionText 时)
    /// 对应 Selection.ShapeRange 属性
    /// </summary>
    IPowerPointShapeRange ShapeRange { get; }

    /// <summary>
    /// 获取选择的文本范围 (当 Type 为 ppSelectionText 时)
    /// 对应 Selection.TextRange 属性
    /// </summary>
    IPowerPointTextRange TextRange { get; }

    /// <summary>
    /// 获取选择的幻灯片范围 (当 Type 为 ppSelectionSlides 时)
    /// 对应 Selection.SlideRange 属性
    /// </summary>
    IPowerPointSlideRange SlideRange { get; }
    #endregion

    #region 操作方法
    /// <summary>
    /// 取消当前选择
    /// 对应 Selection.Unselect 方法
    /// </summary>
    void Unselect();

    /// <summary>
    /// 选择所有内容 (通常指幻灯片上的所有形状)
    /// </summary>
    /// <param name="replace">是否替换当前选择</param>
    void SelectAll(bool replace = true);

    /// <summary>
    /// 复制选择的内容
    /// 对应 Selection.Copy 方法
    /// </summary>
    void Copy();

    /// <summary>
    /// 剪切选择的内容
    /// 对应 Selection.Cut 方法
    /// </summary>
    void Cut();

    /// <summary>
    /// 删除选择的内容
    /// 对应 Selection.Delete 方法
    /// </summary>
    void Delete();
    #endregion

    #region 文本操作 (当选择包含文本时)
    /// <summary>
    /// 在选择的文本范围内查找文本
    /// </summary>
    /// <param name="findWhat">要查找的文本</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <param name="matchWholeWord">是否匹配整个单词</param>
    /// <returns>找到的文本范围对象</returns>
    IPowerPointTextRange FindText(string findWhat, bool matchCase = false, bool matchWholeWord = false);

    /// <summary>
    /// 替换选择文本范围内的文本
    /// </summary>
    /// <param name="findWhat">要查找的文本</param>
    /// <param name="replaceWhat">替换为的文本</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <param name="matchWholeWord">是否匹配整个单词</param>
    /// <returns>替换的次数</returns>
    int ReplaceText(string findWhat, string replaceWhat, bool matchCase = false, bool matchWholeWord = false);

    /// <summary>
    /// 设置选择文本的字体
    /// </summary>
    /// <param name="fontName">字体名称</param>
    /// <param name="fontSize">字体大小</param>
    /// <param name="bold">是否加粗</param>
    /// <param name="italic">是否倾斜</param>
    void SetTextFont(string fontName = "", float fontSize = 0, bool bold = false, bool italic = false);

    /// <summary>
    /// 设置选择文本的颜色
    /// </summary>
    /// <param name="color">颜色值</param>
    void SetTextColor(int color);

    /// <summary>
    /// 设置选择文本的对齐方式
    /// </summary>
    /// <param name="alignment">对齐方式</param>
    void SetTextAlignment(PpParagraphAlignment alignment);
    #endregion

    #region 形状操作 (当选择包含形状时)
    /// <summary>
    /// 对齐选中的形状
    /// </summary>
    /// <param name="alignment">对齐方式</param>
    /// <param name="relativeTo">相对对象</param>
    void AlignShapes(MsoAlignCmd alignment, int relativeTo = 0);

    /// <summary>
    /// 分布选中的形状
    /// </summary>
    /// <param name="distribution">分布方式</param>
    void DistributeShapes(MsoDistributeCmd distribution);

    /// <summary>
    /// 组合选中的形状
    /// </summary>
    /// <returns>组合后的形状对象</returns>
    IPowerPointShape GroupShapes();

    /// <summary>
    /// 取消组合选中的形状
    /// </summary>
    void UngroupShapes();

    /// <summary>
    /// 设置选中形状的填充
    /// </summary>
    /// <param name="color">填充颜色</param>
    void SetShapeFill(int color);

    /// <summary>
    /// 设置选中形状的边框
    /// </summary>
    /// <param name="color">边框颜色</param>
    /// <param name="weight">边框粗细</param>
    void SetShapeBorder(int color, float weight = 1);
    #endregion

    #region 幻灯片操作 (当选择包含幻灯片时)
    /// <summary>
    /// 复制选中的幻灯片
    /// </summary>
    void CopySlides();

    /// <summary>
    /// 剪切选中的幻灯片
    /// </summary>
    void CutSlides();

    /// <summary>
    /// 删除选中的幻灯片
    /// </summary>
    void DeleteSlides();
    #endregion

}
