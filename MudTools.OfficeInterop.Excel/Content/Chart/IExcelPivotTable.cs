//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel PivotTable 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.PivotTable 的安全访问和操作
/// </summary>
public interface IExcelPivotTable : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取或设置数据透视表的名称
    /// 对应 PivotTable.Name 属性
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取数据透视表的父对象 (通常是 Worksheet)
    /// 对应 PivotTable.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取数据透视表所在的Application对象
    /// 对应 PivotTable.Application 属性
    /// </summary>
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取数据透视表的源数据缓存
    /// 对应 PivotTable.PivotCache 属性
    /// </summary>
    IExcelPivotCache? PivotCache();

    /// <summary>
    /// 获取或设置数据透视表的源数据
    /// 对应 PivotTable.SourceData 属性 (只读)
    /// </summary>
    object? SourceData { get; }

    /// <summary>
    /// 获取数据透视表的版本
    /// 对应 PivotTable.Version 属性
    /// </summary>
    XlPivotTableVersionList Version { get; }
    #endregion

    #region 数据和字段
    /// <summary>
    /// 获取数据透视表的数据主体区域 (不包括页字段报告筛选器)
    /// 对应 PivotTable.DataBodyRange 属性
    /// </summary>
    IExcelRange? DataBodyRange { get; }

    /// <summary>
    /// 获取数据透视表的整个表格区域 (包括页字段报告筛选器)
    /// 对应 PivotTable.TableRange1 属性
    /// </summary>
    IExcelRange? TableRange1 { get; }

    /// <summary>
    /// 获取数据透视表的第二区域 (如果有页字段报告筛选器，则包含这些字段)
    /// 对应 PivotTable.TableRange2 属性
    /// </summary>
    IExcelRange? TableRange2 { get; }

    /// <summary>
    /// 获取数据透视表的页字段 (报告筛选器) 集合
    /// 对应 PivotTable.PageFields 属性
    /// </summary>
    IExcelPivotFields? PageFields { get; }

    /// <summary>
    /// 获取数据透视表的行字段集合
    /// 对应 PivotTable.RowFields 属性
    /// </summary>
    IExcelPivotFields? RowFields { get; }

    /// <summary>
    /// 获取数据透视表的列字段集合
    /// 对应 PivotTable.ColumnFields 属性
    /// </summary>
    IExcelPivotFields? ColumnFields { get; }

    /// <summary>
    /// 获取数据透视表的数据字段集合
    /// 对应 PivotTable.DataFields 属性
    /// </summary>
    IExcelPivotFields? DataFields { get; }

    /// <summary>
    /// 获取数据透视表的可见数据字段集合
    /// 对应 PivotTable.VisibleFields 属性
    /// </summary>
    IExcelPivotFields? VisibleFields { get; }

    /// <summary>
    /// 获取数据透视表的隐藏数据字段集合
    /// 对应 PivotTable.HiddenFields 属性
    /// </summary>
    IExcelPivotFields? HiddenFields { get; }
    #endregion

    #region 格式和布局
    /// <summary>
    /// 获取或设置数据透视表的表格样式
    /// 对应 PivotTable.TableStyle2 属性
    /// </summary>
    IExcelTableStyle? TableStyle { get; set; }

    /// <summary>
    /// 获取或设置数据透视表是否显示行条纹
    /// 对应 PivotTable.ShowTableStyleRowStripes 属性
    /// </summary>
    bool ShowTableStyleRowStripes { get; set; }

    /// <summary>
    /// 获取或设置数据透视表是否显示列条纹
    /// 对应 PivotTable.ShowTableStyleColumnStripes 属性
    /// </summary>
    bool ShowTableStyleColumnStripes { get; set; }


    /// <summary>
    /// 获取或设置数据透视表是否显示末列特殊样式
    /// 对应 PivotTable.ShowTableStyleLastColumn 属性
    /// </summary>
    bool ShowTableStyleLastColumn { get; set; }

    /// <summary>
    /// 获取或设置数据透视表是否显示行总计
    /// 对应 PivotTable.RowGrand 属性
    /// </summary>
    bool RowGrand { get; set; }

    /// <summary>
    /// 获取或设置数据透视表是否显示列总计
    /// 对应 PivotTable.ColumnGrand 属性
    /// </summary>
    bool ColumnGrand { get; set; }

    /// <summary>
    /// 获取或设置数据透视表在刷新或移动字段时是否自动设置格式
    /// 对应 PivotTable.HasAutoFormat 属性
    /// </summary>
    bool HasAutoFormat { get; set; }
    #endregion

    #region 状态属性 

    /// <summary>
    /// 获取数据透视表是否被保护
    /// 对应 PivotTable.EnableWizard, EnableDataValueEditing 等属性组合判断
    /// </summary>
    bool IsProtected { get; }
    #endregion

    #region 操作方法
    /// <summary>
    /// 获取数据透视表中的特定字段
    /// </summary>
    /// <param name="Index">要获取的字段索引或名称</param>
    /// <returns>对应的数据透视表字段，如果未找到则返回null</returns>
    IExcelPivotField? PivotFields(object Index);
    /// <summary>
    /// 获取数据透视表中所有字段的集合
    /// </summary>
    /// <returns>包含所有数据透视表字段的集合，如果出现错误则返回null</returns>
    IExcelPivotFields? PivotFields();

    /// <summary>
    /// 获取数据透视表中所有计算字段的集合
    /// 对应 PivotTable.CalculatedFields 属性
    /// </summary>
    /// <returns>包含所有计算字段的集合，如果出现错误则返回null</returns>
    IExcelCalculatedFields? CalculatedFields();
    /// <summary>
    /// 选择数据透视表
    /// 对应 PivotTable.Select 方法
    /// </summary>
    /// <param name="replace">是否替换当前选择</param>
    void Select(bool replace = true);

    /// <summary>
    /// 复制数据透视表
    /// 对应 PivotTable.Copy 方法
    /// </summary>
    void Copy();

    /// <summary>
    /// 剪切数据透视表
    /// 对应 PivotTable.Cut 方法
    /// </summary>
    void Cut();

    /// <summary>
    /// 删除数据透视表
    /// 对应 PivotTable.Delete 方法
    /// </summary>
    void Delete();
    #endregion

    #region 数据透视表操作
    /// <summary>
    /// 刷新数据透视表
    /// 对应 PivotTable.RefreshTable 方法
    /// </summary>
    void Refresh();

    /// <summary>
    /// 更新数据透视表 (通常与 Refresh 同义)
    /// </summary>
    void Update();


    /// <summary>
    /// 清除数据透视表内容 (清除数据，保留结构)
    /// </summary>
    void Clear();

    /// <summary>
    /// 清除数据透视表格式
    /// </summary>
    void ClearFormats();

    /// <summary>
    /// 清除数据透视表内容和格式
    /// </summary>
    void ClearAll();

    /// <summary>
    /// 应用自动格式
    /// </summary>
    /// <param name="format">自动格式编号</param>
    void ApplyAutoFormat(XlRangeAutoFormat format = XlRangeAutoFormat.xlRangeAutoFormatClassic1); // xlRangeAutoFormatClassic1 = 1
    #endregion

    #region 格式设置
    /// <summary>
    /// 设置数据透视表样式
    /// </summary>
    /// <param name="styleName">样式名称</param>
    void SetStyle(string styleName);

    #endregion

    #region 高级功能   

    /// <summary>
    /// 打印数据透视表
    /// </summary>
    /// <param name="preview">是否打印预览</param>
    void PrintOut(bool preview = false);
    #endregion
}
