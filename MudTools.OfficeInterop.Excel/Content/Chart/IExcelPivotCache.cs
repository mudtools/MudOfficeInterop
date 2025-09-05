//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel PivotCache 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.PivotCache 的安全访问和操作
/// </summary>
public interface IExcelPivotCache : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取数据透视表缓存的索引位置
    /// 对应 PivotCache.Index 属性
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取数据透视表缓存的父对象 (通常是 Workbook)
    /// 对应 PivotCache.Parent 属性
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取数据透视表缓存所在的Application对象
    /// 对应 PivotCache.Application 属性
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取数据透视表缓存的源数据类型
    /// 对应 PivotCache.SourceType 属性
    /// </summary>
    int SourceType { get; } // 使用 int 代表 XlPivotTableSourceType

    /// <summary>
    /// 获取数据透视表缓存的源数据
    /// 对应 PivotCache.SourceData 属性
    /// </summary>
    object SourceData { get; } // 可以是 string, Range, ListObject 等


    /// <summary>
    /// 获取数据透视表缓存的记录数
    /// 对应 PivotCache.RecordCount 属性
    /// </summary>
    int RecordCount { get; }

    /// <summary>
    /// 获取数据透视表缓存的版本
    /// 对应 PivotCache.Version 属性
    /// </summary>
    int Version { get; } // 使用 int 代表 XlPivotTableVersionList
    #endregion

    #region 操作方法
    /// <summary>
    /// 刷新数据透视表缓存
    /// 对应 PivotCache.Refresh 方法
    /// </summary>
    void Refresh();

    /// <summary>
    /// 更新数据透视表缓存 (通常与 Refresh 同义)
    /// </summary>
    void Update();
    #endregion
}
