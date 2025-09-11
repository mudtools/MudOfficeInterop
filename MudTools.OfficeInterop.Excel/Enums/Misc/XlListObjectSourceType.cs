//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 列表对象数据源类型枚举
/// 用于指定创建ListObject时的数据源类型
/// </summary>
public enum XlListObjectSourceType
{
    /// <summary>
    /// 外部数据源（如Microsoft SharePoint Foundation网站）
    /// </summary>
    xlSrcExternal,
    
    /// <summary>
    /// 范围数据源（工作表中的单元格区域）
    /// </summary>
    xlSrcRange,
    
    /// <summary>
    /// XML数据源
    /// </summary>
    xlSrcXml,
    
    /// <summary>
    /// 查询数据源（ODBC或OLEDB连接字符串）
    /// </summary>
    xlSrcQuery,
    
    /// <summary>
    /// PowerPivot数据模型源
    /// </summary>
    xlSrcModel
}