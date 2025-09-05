//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定在文件对话框中文件列表的排序方式
/// </summary>
public enum MsoSortBy
{
    /// <summary>
    /// 按文件名排序
    /// </summary>
    msoSortByFileName = 1,
    
    /// <summary>
    /// 按文件大小排序
    /// </summary>
    msoSortBySize,
    
    /// <summary>
    /// 按文件类型排序
    /// </summary>
    msoSortByFileType,
    
    /// <summary>
    /// 按最后修改时间排序
    /// </summary>
    msoSortByLastModified,
    
    /// <summary>
    /// 不排序
    /// </summary>
    msoSortByNone
}