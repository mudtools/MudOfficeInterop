//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;
/// <summary>
/// PowerPoint 形状范围接口
/// </summary>
public interface IPowerPointShapeRange : IDisposable
{
    /// <summary>
    /// 获取形状数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置形状范围名称
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取形状范围的左边缘位置
    /// </summary>
    float Left { get; set; }

    /// <summary>
    /// 获取形状范围的上边缘位置
    /// </summary>
    float Top { get; set; }

    /// <summary>
    /// 获取形状范围的宽度
    /// </summary>
    float Width { get; set; }

    /// <summary>
    /// 获取形状范围的高度
    /// </summary>
    float Height { get; set; }

    /// <summary>
    /// 获取形状范围的旋转角度
    /// </summary>
    float Rotation { get; set; }

    /// <summary>
    /// 获取或设置形状范围的可见性
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取形状范围的锁定纵横比
    /// </summary>
    bool LockAspectRatio { get; set; }

    /// <summary>
    /// 获取形状范围的Z轴顺序
    /// </summary>
    int ZOrderPosition { get; }

    /// <summary>
    /// 获取形状范围的文本框
    /// </summary>
    IPowerPointTextFrame TextFrame { get; }

    /// <summary>
    /// 获取形状范围的填充格式
    /// </summary>
    IPowerPointFillFormat Fill { get; }

    /// <summary>
    /// 获取形状范围的线条格式
    /// </summary>
    IPowerPointLineFormat Line { get; }

    /// <summary>
    /// 获取形状范围的阴影格式
    /// </summary>
    IPowerPointShadowFormat Shadow { get; }

    /// <summary>
    /// 获取形状范围的三维格式
    /// </summary>
    IPowerPointThreeDFormat ThreeD { get; }

    /// <summary>
    /// 获取形状范围的动画设置
    /// </summary>
    IPowerPointAnimationSettings AnimationSettings { get; }

    /// <summary>
    /// 获取形状范围的标签集合
    /// </summary>
    IPowerPointTags Tags { get; }

    /// <summary>
    /// 获取形状范围内的所有形状
    /// </summary>
    IEnumerable<IPowerPointShape> Shapes { get; }

    /// <summary>
    /// 选择形状范围
    /// </summary>
    /// <param name="replace">是否替换当前选择</param>
    void Select(bool replace = true);

    /// <summary>
    /// 复制形状范围
    /// </summary>
    void Copy();

    /// <summary>
    /// 剪切形状范围
    /// </summary>
    void Cut();

    /// <summary>
    /// 删除形状范围
    /// </summary>
    void Delete();

    /// <summary>
    /// 水平翻转形状范围
    /// </summary>
    void FlipHorizontal();

    /// <summary>
    /// 垂直翻转形状范围
    /// </summary>
    void FlipVertical();

    /// <summary>
    /// 设置形状范围的Z轴顺序
    /// </summary>
    /// <param name="position">Z轴顺序位置</param>
    void ZOrder(int position);

    /// <summary>
    /// 组合形状范围
    /// </summary>
    /// <returns>组合后的形状</returns>
    IPowerPointShape Group();

    /// <summary>
    /// 取消组合形状范围
    /// </summary>
    /// <returns>取消组合后的形状范围</returns>
    IPowerPointShapeRange Ungroup();

    /// <summary>
    /// 对齐形状范围
    /// </summary>
    /// <param name="alignCmd">对齐命令</param>
    /// <param name="relativeToSlide">是否相对于幻灯片对齐</param>
    void Align(int alignCmd, bool relativeToSlide = false);

    /// <summary>
    /// 分布形状范围
    /// </summary>
    /// <param name="distributeCmd">分布命令</param>
    /// <param name="relativeToSlide">是否相对于幻灯片分布</param>
    void Distribute(int distributeCmd, bool relativeToSlide = false);

    /// <summary>
    /// 获取指定索引的形状
    /// </summary>
    /// <param name="index">形状索引</param>
    /// <returns>形状对象</returns>
    IPowerPointShape Item(int index);

    /// <summary>
    /// 获取指定名称的形状
    /// </summary>
    /// <param name="name">形状名称</param>
    /// <returns>形状对象</returns>
    IPowerPointShape Item(string name);

    /// <summary>
    /// 根据条件查找形状
    /// </summary>
    /// <param name="predicate">查找条件</param>
    /// <returns>符合条件的形状列表</returns>
    IEnumerable<IPowerPointShape> Find(Func<IPowerPointShape, bool> predicate);

    /// <summary>
    /// 获取形状范围的文本内容
    /// </summary>
    /// <returns>文本内容</returns>
    string GetText();

    /// <summary>
    /// 设置形状范围的文本内容
    /// </summary>
    /// <param name="text">文本内容</param>
    void SetText(string text);

    /// <summary>
    /// 替换形状范围中的文本
    /// </summary>
    /// <param name="findText">查找文本</param>
    /// <param name="replaceText">替换文本</param>
    /// <returns>替换次数</returns>
    int ReplaceText(string findText, string replaceText);

    /// <summary>
    /// 添加文本到形状范围
    /// </summary>
    /// <param name="text">要添加的文本</param>
    void AddText(string text);

    /// <summary>
    /// 清除形状范围的文本
    /// </summary>
    void ClearText();

    /// <summary>
    /// 设置形状范围的填充颜色
    /// </summary>
    /// <param name="color">颜色值</param>
    void SetFillColor(int color);

    /// <summary>
    /// 设置形状范围的线条颜色
    /// </summary>
    /// <param name="color">颜色值</param>
    void SetLineColor(int color);

    /// <summary>
    /// 设置形状范围的线条粗细
    /// </summary>
    /// <param name="weight">线条粗细</param>
    void SetLineWeight(float weight);

    /// <summary>
    /// 应用阴影效果
    /// </summary>
    /// <param name="shadowType">阴影类型</param>
    void ApplyShadow(int shadowType = 1);

    /// <summary>
    /// 应用三维效果
    /// </summary>
    /// <param name="depth">深度</param>
    void Apply3DEffect(float depth = 10);

    /// <summary>
    /// 应用动画效果
    /// </summary>
    /// <param name="effectType">效果类型</param>
    /// <param name="triggerType">触发类型</param>
    void ApplyAnimation(int effectType = 1, int triggerType = 1);
}