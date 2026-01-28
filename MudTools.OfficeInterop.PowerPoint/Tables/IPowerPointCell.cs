//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.PowerPoint;

using System;

/// <summary>
/// 表示 PowerPoint 表格中的一个单元格。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointCell : IDisposable
{
    /// <summary>
    /// 获取创建此单元格的 PowerPoint 应用程序对象。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="Application"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此单元格的父对象。
    /// </summary>
    /// <value>表示父对象的 <see cref="object"/>。</value>
    object Parent { get; }

    /// <summary>
    /// 获取与此单元格相关联的形状对象。
    /// </summary>
    /// <value>表示包含此单元格的表格形状的 <see cref="IPowerPointShape"/> 对象。</value>
    IPowerPointShape? Shape { get; }

    /// <summary>
    /// 获取此单元格的边框集合。
    /// </summary>
    /// <value>表示单元格边框的 <see cref="IPowerPointBorders"/> 对象。</value>
    IPowerPointBorders? Borders { get; }

    /// <summary>
    /// 将当前单元格与指定单元格合并。
    /// </summary>
    /// <param name="mergeTo">要合并到的目标单元格。</param>
    void Merge([ComNamespace("MsPowerPoint")] IPowerPointCell mergeTo);

    /// <summary>
    /// 将当前单元格拆分为指定数量的行和列。
    /// </summary>
    /// <param name="numRows">拆分后单元格应包含的行数。</param>
    /// <param name="numColumns">拆分后单元格应包含的列数。</param>
    void Split(int numRows, int numColumns);

    /// <summary>
    /// 选中当前单元格。
    /// </summary>
    void Select();

    /// <summary>
    /// 获取一个值，该值指示当前单元格是否被选中。
    /// </summary>
    /// <value>如果单元格被选中，则为 <see langword="true"/>；否则为 <see langword="false"/>。</value>
    bool Selected { get; }
}