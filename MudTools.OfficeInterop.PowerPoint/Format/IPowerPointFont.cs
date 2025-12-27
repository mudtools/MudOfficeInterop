//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// PowerPoint 字体接口
/// </summary>
public interface IPowerPointFont : IDisposable
{
    /// <summary>
    /// 获取父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置字体名称
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取或设置字体大小
    /// </summary>
    float Size { get; set; }

    /// <summary>
    /// 获取或设置字体颜色
    /// </summary>
    int Color { get; set; }

    /// <summary>
    /// 获取或设置是否加粗
    /// </summary>
    bool Bold { get; set; }

    /// <summary>
    /// 获取或设置是否斜体
    /// </summary>
    bool Italic { get; set; }

    /// <summary>
    /// 获取或设置下划线类型
    /// </summary>
    int Underline { get; set; }


    /// <summary>
    /// 获取或设置上标
    /// </summary>
    bool Subscript { get; set; }

    /// <summary>
    /// 获取或设置下标
    /// </summary>
    bool Superscript { get; set; }


    /// <summary>
    /// 获取或设置阴影效果
    /// </summary>
    bool Shadow { get; set; }

    /// <summary>
    /// 获取或设置轮廓效果
    /// </summary>
    bool Emboss { get; set; }

    /// <summary>
    /// 获取或设置字体样式
    /// </summary>
    int Style { get; set; }

    /// <summary>
    /// 获取或设置字体填充格式
    /// </summary>
    IPowerPointFillFormat Fill { get; }

    /// <summary>
    /// 获取或设置字体轮廓格式
    /// </summary>
    IPowerPointLineFormat Line { get; }

    /// <summary>
    /// 获取或设置字体效果格式
    /// </summary>
    IPowerPointShadowFormat ShadowFormat { get; }

    /// <summary>
    /// 复制字体设置
    /// </summary>
    /// <returns>复制的字体对象</returns>
    IPowerPointFont Duplicate();

    /// <summary>
    /// 应用字体设置到指定文本范围
    /// </summary>
    /// <param name="textRange">目标文本范围</param>
    void ApplyTo(IPowerPointTextRange textRange);

    /// <summary>
    /// 重置字体设置为默认值
    /// </summary>
    void Reset();

    /// <summary>
    /// 设置字体基本属性
    /// </summary>
    /// <param name="fontName">字体名称</param>
    /// <param name="fontSize">字体大小</param>
    /// <param name="color">字体颜色</param>
    void SetBasicProperties(string fontName = null, float fontSize = 0, int color = 0);

    /// <summary>
    /// 设置字体样式
    /// </summary>
    /// <param name="bold">是否加粗</param>
    /// <param name="italic">是否斜体</param>
    /// <param name="underline">下划线类型</param>
    /// <param name="strikeThrough">删除线</param>
    void SetStyle(bool bold = false, bool italic = false, int underline = 0, bool strikeThrough = false);

    /// <summary>
    /// 设置字体效果
    /// </summary>
    /// <param name="shadow">阴影效果</param>
    /// <param name="emboss">轮廓效果</param>
    /// <param name="imprint">浮雕效果</param>
    /// <param name="subscript">下标</param>
    /// <param name="superscript">上标</param>
    void SetEffects(bool shadow = false, bool emboss = false, bool imprint = false, bool subscript = false, bool superscript = false);


    /// <summary>
    /// 应用主题字体
    /// </summary>
    /// <param name="themeFontIndex">主题字体索引</param>
    void ApplyThemeFont(int themeFontIndex = 1);

    /// <summary>
    /// 获取字体信息
    /// </summary>
    /// <returns>字体信息字符串</returns>
    string GetFontInfo();
}
