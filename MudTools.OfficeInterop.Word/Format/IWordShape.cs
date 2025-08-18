//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
using System;

/// <summary>
/// 表示 Word 中单个形状的封装接口
/// </summary>
public interface IWordShape : IDisposable
{
    /// <summary>
    /// 获取或设置形状的名称
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取或设置形状的左边距
    /// </summary>
    float Left { get; set; }

    /// <summary>
    /// 获取或设置形状的上边距
    /// </summary>
    float Top { get; set; }

    /// <summary>
    /// 获取或设置形状的宽度
    /// </summary>
    float Width { get; set; }

    /// <summary>
    /// 获取或设置形状的高度
    /// </summary>
    float Height { get; set; }

    /// <summary>
    /// 获取或设置形状是否可见
    /// </summary>
    bool Visible { get; set; }


    /// <summary>
    /// 获取或设置形状的替代文本
    /// </summary>
    string AlternativeText { get; set; }

    /// <summary>
    /// 获取形状的文本框架（伪代码）
    /// </summary>
    object TextFrame { get; }

    /// <summary>
    /// 获取形状的填充属性（伪代码）
    /// </summary>
    object Fill { get; }

    /// <summary>
    /// 获取形状的线条属性（伪代码）
    /// </summary>
    object Line { get; }

    /// <summary>
    /// 删除当前形状
    /// </summary>
    void Delete();

    /// <summary>
    /// 选择当前形状
    /// </summary>
    /// <param name="replace">是否替换当前选择</param>
    void Select(bool replace = true);

    /// <summary>
    /// 获取形状的Z轴顺序位置
    /// </summary>
    int ZOrderPosition { get; }

    IWordOLEFormat OLEFormat { get; }
}