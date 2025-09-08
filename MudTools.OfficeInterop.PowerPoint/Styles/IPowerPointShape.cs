//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// PowerPoint 形状接口（精简版）
/// </summary>
public interface IPowerPointShape : IDisposable
{
    /// <summary>
    /// 获取或设置形状名称
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取形状索引
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取父对象
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取形状类型
    /// </summary>
    MsoShapeType Type { get; }

    /// <summary>
    /// 获取或设置左边缘位置
    /// </summary>
    double Left { get; set; }

    /// <summary>
    /// 获取或设置上边缘位置
    /// </summary>
    double Top { get; set; }

    /// <summary>
    /// 获取或设置宽度
    /// </summary>
    double Width { get; set; }

    /// <summary>
    /// 获取或设置高度
    /// </summary>
    double Height { get; set; }

    /// <summary>
    /// 获取或设置旋转角度
    /// </summary>
    double Rotation { get; set; }

    /// <summary>
    /// 获取或设置可见性
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置锁定纵横比
    /// </summary>
    bool LockAspectRatio { get; set; }

    /// <summary>
    /// 获取Z轴顺序位置
    /// </summary>
    int ZOrderPosition { get; }

    /// <summary>
    /// 获取文本框架
    /// </summary>
    IPowerPointTextFrame TextFrame { get; }

    IPowerPointOLEFormat OLEFormat { get; }

    /// <summary>
    /// 获取填充格式
    /// </summary>
    IPowerPointFillFormat Fill { get; }

    /// <summary>
    /// 获取线条格式
    /// </summary>
    IPowerPointLineFormat Line { get; }

    /// <summary>
    /// 获取阴影格式
    /// </summary>
    IPowerPointShadowFormat Shadow { get; }

    /// <summary>
    /// 获取三维格式
    /// </summary>
    IPowerPointThreeDFormat ThreeD { get; }

    /// <summary>
    /// 获取是否具有文本框架
    /// </summary>
    bool HasTextFrame { get; }

    /// <summary>
    /// 获取形状ID
    /// </summary>
    int ID { get; }

    /// <summary>
    /// 获取或设置替代文本
    /// </summary>
    string AlternativeText { get; set; }

    /// <summary>
    /// 获取是否为组形状
    /// </summary>
    bool IsGroup { get; }

    /// <summary>
    /// 选择形状
    /// </summary>
    /// <param name="replace">是否替换当前选择</param>
    void Select(bool replace = true);

    /// <summary>
    /// 复制形状
    /// </summary>
    void Copy();

    /// <summary>
    /// 剪切形状
    /// </summary>
    void Cut();

    /// <summary>
    /// 删除形状
    /// </summary>
    void Delete();

    /// <summary>
    /// 移动形状
    /// </summary>
    /// <param name="x">水平移动距离</param>
    /// <param name="y">垂直移动距离</param>
    void Move(double x, double y);

    /// <summary>
    /// 调整形状大小
    /// </summary>
    /// <param name="width">新宽度</param>
    /// <param name="height">新高度</param>
    void Scale(double width, double height);

    /// <summary>
    /// 旋转形状
    /// </summary>
    /// <param name="angle">旋转角度</param>
    void Rotate(double angle);

    /// <summary>
    /// 水平翻转
    /// </summary>
    void FlipHorizontal();

    /// <summary>
    /// 垂直翻转
    /// </summary>
    void FlipVertical();

    /// <summary>
    /// 设置Z轴顺序
    /// </summary>
    /// <param name="position">位置</param>
    void ZOrder(int position);

    /// <summary>
    /// 取消组合
    /// </summary>
    /// <returns>形状范围</returns>
    IPowerPointShapeRange Ungroup();

    /// <summary>
    /// 获取文本内容
    /// </summary>
    /// <returns>文本内容</returns>
    string GetText();

    /// <summary>
    /// 设置文本内容
    /// </summary>
    /// <param name="text">文本内容</param>
    void SetText(string text);

    /// <summary>
    /// 替换文本
    /// </summary>
    /// <param name="findText">查找文本</param>
    /// <param name="replaceText">替换文本</param>
    /// <returns>替换次数</returns>
    int ReplaceText(string findText, string replaceText);

    /// <summary>
    /// 设置填充颜色
    /// </summary>
    /// <param name="color">颜色值</param>
    void SetFillColor(int color);

    /// <summary>
    /// 设置线条颜色
    /// </summary>
    /// <param name="color">颜色值</param>
    void SetLineColor(int color);

    /// <summary>
    /// 设置线条粗细
    /// </summary>
    /// <param name="weight">粗细</param>
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
    /// 导出为图片
    /// </summary>
    /// <param name="fileName">文件名</param>
    /// <param name="filterName">格式</param>
    void Export(string fileName, int filterName = 2);

    /// <summary>
    /// 刷新显示
    /// </summary>
    void Refresh();

    /// <summary>
    /// 设置透明度
    /// </summary>
    /// <param name="transparency">透明度</param>
    void SetTransparency(float transparency);
}
