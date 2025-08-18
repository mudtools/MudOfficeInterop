//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 中单个内嵌形状的封装接口
/// </summary>
public interface IWordInlineShape : IDisposable
{
    /// <summary>
    /// 获取内嵌形状的类型
    /// </summary>
    int Type { get; }

    /// <summary>
    /// 获取或设置内嵌形状的替代文本
    /// </summary>
    string AlternativeText { get; set; }

    /// <summary>
    /// 获取或设置内嵌形状的高度
    /// </summary>
    float Height { get; set; }

    /// <summary>
    /// 获取或设置内嵌形状的宽度
    /// </summary>
    float Width { get; set; }

    /// <summary>
    /// 获取内嵌形状的范围（伪代码）
    /// </summary>
    object Range { get; }

    /// <summary>
    /// 获取内嵌形状的父对象（伪代码）
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取内嵌形状的OLE格式（伪代码）
    /// </summary>
    object OLEFormat { get; }

    /// <summary>
    /// 获取内嵌形状的链接格式（伪代码）
    /// </summary>
    object LinkFormat { get; }

    /// <summary>
    /// 获取内嵌形状的填充属性（伪代码）
    /// </summary>
    object Fill { get; }

    /// <summary>
    /// 获取内嵌形状的线条属性（伪代码）
    /// </summary>
    object Line { get; }

    /// <summary>
    /// 删除当前内嵌形状
    /// </summary>
    void Delete();

    /// <summary>
    /// 复制当前内嵌形状
    /// </summary>
    void Copy();

    /// <summary>
    /// 剪切当前内嵌形状
    /// </summary>
    void Cut();

    /// <summary>
    /// 选择当前内嵌形状
    /// </summary>
    void Select();

    /// <summary>
    /// 将内嵌形状转换为浮动形状
    /// </summary>
    /// <returns>转换后的浮动形状对象</returns>
    IWordShape ConvertToShape();

    /// <summary>
    /// 更新链接的内嵌形状
    /// </summary>
    void Update();

    /// <summary>
    /// 锁定内嵌形状的比例
    /// </summary>
    bool LockAspectRatio { get; set; }

    /// <summary>
    /// 获取内嵌形状是否为图片类型
    /// </summary>
    bool IsPicture { get; }

    /// <summary>
    /// 获取内嵌形状是否为OLE对象类型
    /// </summary>
    bool IsOLEObject { get; }
}