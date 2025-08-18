//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel PivotField 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.PivotField 的安全访问和操作
/// </summary>
public interface IExcelPivotField : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取或设置数据透视表字段的名称
    /// 对应 PivotField.Name 属性
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取数据透视表字段的父对象 (通常是 PivotTable)
    /// 对应 PivotField.Parent 属性
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取数据透视表字段所在的Application对象
    /// 对应 PivotField.Application 属性
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取或设置数据透视表字段的方向 (行、列、页、数据或隐藏)
    /// 对应 PivotField.Orientation 属性
    /// </summary>
    XlPivotFieldOrientation Orientation { get; set; }

    /// <summary>
    /// 获取或设置数据透视表字段的位置 (在其方向内的顺序)
    /// 对应 PivotField.Position 属性
    /// </summary>
    int Position { get; set; }

    /// <summary>
    /// 获取或设置数据透视表字段的数字格式
    /// 对应 PivotField.NumberFormat 属性
    /// </summary>
    string NumberFormat { get; set; }

    /// <summary>
    /// 获取数据透视表字段的数据源范围
    /// 对应 PivotField.SourceName 属性 (可能需要解析)
    /// </summary>
    object SourceName { get; }

    /// <summary>
    /// 获取数据透视表字段的汇总函数
    /// 对应 PivotField.Function 属性
    /// </summary>
    XlConsolidationFunction Function { get; set; }

    string Formula { get; set; }
    #endregion   

    #region 状态属性
    /// <summary>
    /// 对应 PivotField.EnableItemSelection 属性 
    /// </summary>
    bool EnableItemSelection { get; }

    /// <summary>
    /// 获取数据透视表字段是否为已计算字段
    /// 对应 PivotField.IsCalculated 属性
    /// </summary>
    bool IsCalculated { get; }

    /// <summary>
    /// 对应 PivotField.IsMemberProperty 属性
    /// </summary>
    bool IsMemberProperty { get; }
    #endregion

    #region 图表元素 (子对象)
    /// <summary>
    /// 获取数据透视表字段的项目集合
    /// 对应 PivotField.PivotItems 属性
    /// </summary>
    IExcelPivotItems PivotItems { get; }

    /// <summary>
    /// 获取数据透视表字段的数据范围 (如果适用)
    /// 对应 PivotField.DataRange 属性
    /// </summary>
    IExcelRange DataRange { get; }

    /// <summary>
    /// 获取数据透视表字段的标签范围 (如果适用)
    /// 对应 PivotField.LabelRange 属性
    /// </summary>
    IExcelRange LabelRange { get; }
    #endregion

    #region 操作方法
    /// <summary>
    /// 选择数据透视表字段
    /// 对应 PivotField.Select 方法
    /// </summary>
    /// <param name="replace">是否替换当前选择</param>
    void Select(bool replace = true);

    /// <summary>
    /// 删除数据透视表字段 (通常意味着将其方向设为 xlHidden)
    /// </summary>
    void Delete();

    /// <summary>
    /// 清除数据透视表字段内容 (通常指从当前区域移除)
    /// </summary>
    void Clear();

    /// <summary>
    /// 复制数据透视表字段
    /// </summary>
    void Copy();

    /// <summary>
    /// 剪切数据透视表字段
    /// </summary>
    void Cut();
    #endregion
}
