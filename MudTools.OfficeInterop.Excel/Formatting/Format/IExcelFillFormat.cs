//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel FillFormat 对象的二次封装接口
/// </summary>
public interface IExcelFillFormat : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取 ChartFormat 对象的父对象
    /// 父对象通常是 ChartArea, PlotArea, Series 等图表元素
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取 ChartFormat 对象所在的 Application 对象
    /// </summary>
    IExcelApplication Application { get; }
    #endregion

    /// <summary>
    /// 获取或设置填充类型
    /// </summary>
    MsoFillType Type { get; }

    /// <summary>
    /// 获取或设置前景色
    /// </summary>
    IExcelColorFormat ForeColor { get; set; }

    /// <summary>
    /// 获取或设置背景色
    /// </summary>
    IExcelColorFormat BackColor { get; set; }

    /// <summary>
    /// 获取或设置图案类型
    /// </summary>
    MsoPatternType Pattern { get; }

    /// <summary>
    /// 获取或设置透明度（0-100）
    /// </summary>
    int Transparency { get; set; }

    /// <summary>
    /// 获取或设置是否可见
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取渐变颜色类型
    /// </summary>
    MsoGradientColorType GradientColorType { get; }

    /// <summary>
    /// 获取渐变度数
    /// </summary>
    float GradientDegree { get; }

    /// <summary>
    /// 获取渐变样式
    /// </summary>
    MsoGradientStyle GradientStyle { get; }

    /// <summary>
    /// 获取渐变变体
    /// </summary>
    int GradientVariant { get; }

    /// <summary>
    /// 获取预设渐变类型
    /// </summary>
    MsoPresetGradientType PresetGradientType { get; }

    /// <summary>
    /// 获取预设纹理
    /// </summary>
    MsoPresetTexture PresetTexture { get; }

    /// <summary>
    /// 获取或设置纹理对齐方式
    /// </summary>
    MsoTextureAlignment TextureAlignment { get; set; }

    /// <summary>
    /// 获取纹理类型
    /// </summary>
    MsoTextureType TextureType { get; }

    /// <summary>
    /// 获取或设置纹理水平偏移量
    /// </summary>
    float TextureOffsetX { get; set; }

    /// <summary>
    /// 获取或设置纹理垂直偏移量
    /// </summary>
    float TextureOffsetY { get; set; }

    /// <summary>
    /// 获取或设置纹理水平缩放比例
    /// </summary>
    float TextureHorizontalScale { get; set; }

    /// <summary>
    /// 获取或设置纹理垂直缩放比例
    /// </summary>
    float TextureVerticalScale { get; set; }

    /// <summary>
    /// 获取或设置是否平铺纹理
    /// </summary>
    bool TextureTile { get; set; }
    /// <summary>
    /// 获取纹理名称
    /// </summary>
    string TextureName { get; }

    /// <summary>
    /// 设置用户图片作为填充
    /// </summary>
    /// <param name="PictureFile">图片文件路径</param>
    void UserPicture(string PictureFile);

    /// <summary>
    /// 设置用户纹理作为填充
    /// </summary>
    /// <param name="TextureFile">纹理文件路径</param>
    void UserTextured(string TextureFile);
}