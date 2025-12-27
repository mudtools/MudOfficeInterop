//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 提供Excel排序功能的接口
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelSort : IOfficeObject<IExcelSort>, IDisposable
{
    /// <summary>
    /// 获取父级工作表
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取条件值对象所在的Application对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取或设置排序范围
    /// </summary>
    IExcelRange? Rng { get; }

    /// <summary>
    /// 获取或设置排序头行
    /// </summary>
    XlYesNoGuess Header { get; set; }


    /// <summary>
    /// 获取或设置排序方法
    /// </summary>
    XlSortMethod SortMethod { get; set; }

    /// <summary>
    /// 获取排序字段集合
    /// </summary>
    IExcelSortFields? SortFields { get; }

    /// <summary>
    /// 获取或设置是否区分大小写
    /// </summary>
    bool MatchCase { get; set; }

    /// <summary>
    /// 获取或设置排序方向（数据方向）
    /// </summary>
    XlSortOrientation Orientation { get; set; }

    /// <summary>
    /// 设置排序范围
    /// </summary>
    /// <param name="range">要排序的单元格范围</param>
    void SetRange(IExcelRange range);

    /// <summary>
    /// 应用排序
    /// </summary>
    void Apply();
}
