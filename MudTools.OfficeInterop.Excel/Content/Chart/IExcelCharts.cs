//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel Charts 集合对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Charts 的安全访问和操作
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelCharts : IOfficeObject<IExcelCharts, MsExcel.Charts>, IEnumerable<IExcelChart?>, IDisposable
{
    #region 基础属性

    /// <summary>
    /// 获取图表集合中的图表数量
    /// 对应 Charts.Count 属性
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的图表对象
    /// 索引从1开始
    /// </summary>
    /// <param name="index">图表索引（从1开始）</param>
    /// <returns>图表对象</returns>
    [ComPropertyWrap(NeedConvert = true)]
    IExcelChart? this[int index] { get; }

    /// <summary>
    /// 获取指定名称的图表对象
    /// </summary>
    /// <param name="name">图表名称</param>
    /// <returns>图表对象</returns>
    [ComPropertyWrap(NeedConvert = true)]
    IExcelChart? this[string name] { get; }

    /// <summary>
    /// 获取图表集合所在的父对象（通常是工作表或工作簿）
    /// 对应 Charts.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取图表集合所在的Application对象
    /// 对应 Charts.Application 属性
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    #endregion

    #region 创建和添加
    /// <summary>
    /// 将工作表复制到工作簿的另一位置。
    /// </summary>
    /// <param name="before">可选 对象。 将要在其之前放置所复制工作表的工作表。 如果指定 Before ，则不能指定 After。</param>
    /// <param name="after">可选 对象。 将要在其之后放置所复制工作表的工作表。 如果指定 After ，则不能指定 Before。</param>
    void Copy(IExcelRange? before = null, IExcelRange? after = null);
    /// <summary>
    /// 将图表直接插入网格。
    /// </summary>
    /// <param name="after">指定工作表的对象，新建的工作表将置于此工作表之后。</param>
    /// <param name="before">指定工作表的对象，新建的工作表将置于此工作表之前。</param>
    /// <param name="count">要添加的工作表数。 默认值为 1。</param>
    /// <param name="newLayout">如果 NewLayout 为 True，则使用新的动态格式设置规则插入图表， (标题处于打开状态，并且仅当有多个系列) 时，图例才会打开。</param>
    /// <returns>新创建的图表对象</returns>
    IExcelChart? Add2(IExcelRange? before, IExcelRange? after, int? count = 1, bool? newLayout = null);

    /// <summary>
    /// 将工作表移到工作簿中的其他位置。如果未指定 Before 或 After，Microsoft Excel 将创建包含已移动工作表的新工作簿。
    /// </summary>
    /// <param name="before">可选 对象。 在其之前放置移动工作表的工作表。 如果指定 Before ，则不能指定 After。</param>
    /// <param name="after">可选 对象。 在其之后放置移动工作表的工作表。 如果指定 After ，则不能指定 Before。</param>
    void Move(IExcelRange? before, IExcelRange? after);
    #endregion


    #region 操作方法
    /// <summary>
    /// 打印对象。 To 和 From中的“页面”是指打印的页面，而不是工作表或工作簿中的整体页面。
    /// </summary>
    /// <param name="from">可选 对象。 打印的开始页号。 如果省略此参数，则从起始位置开始打印。</param>
    /// <param name="To">可选 对象。 打印的终止页号。 如果省略此参数，则打印至最后一页。</param>
    /// <param name="copies">可选 对象。 打印份数。 如果省略此参数，则只打印一份。</param>
    /// <param name="preview">可选 对象。 如果为 True，Microsoft Excel 将在打印对象之前调用打印预览。 如果为 False（或省略该参数），则立即打印对象。</param>
    /// <param name="activePrinter">可选 对象。 设置活动打印机的名称。</param>
    /// <param name="printToFile">可选 对象。 如果为 True，则打印到文件。 如果未 PrToFileName 指定 ，Microsoft Excel 会提示用户输入输出文件的名称。</param>
    /// <param name="collate">可选 对象。 如果为 True，则逐份打印多个副本。</param>
    /// <param name="prToFileName">可选 对象。 如果 PrintToFile 设置为 True，则此参数指定要打印到的文件的名称。</param>
    void PrintOut(int? from, int? To, int? copies = null,
        bool? preview = null, string? activePrinter = null,
        bool? printToFile = null, bool? collate = null, bool? prToFileName = null);

    /// <summary>
    /// 打印对象。 To 和 From中的“页面”是指打印的页面，而不是工作表或工作簿中的整体页面。
    /// </summary>
    /// <param name="from">可选 对象。 打印的开始页号。 如果省略此参数，则从起始位置开始打印。</param>
    /// <param name="To">可选 对象。 打印的终止页号。 如果省略此参数，则打印至最后一页。</param>
    /// <param name="copies">可选 对象。 打印份数。 如果省略此参数，则只打印一份。</param>
    /// <param name="preview">可选 对象。 如果为 True，Microsoft Excel 将在打印对象之前调用打印预览。 如果为 False（或省略该参数），则立即打印对象。</param>
    /// <param name="activePrinter">可选 对象。 设置活动打印机的名称。</param>
    /// <param name="printToFile">可选 对象。 如果为 True，则打印到文件。 如果未 PrToFileName 指定 ，Microsoft Excel 会提示用户输入输出文件的名称。</param>
    /// <param name="collate">可选 对象。 如果为 True，则逐份打印多个副本。</param>
    /// <param name="prToFileName">可选 对象。 如果 PrintToFile 设置为 True，则此参数指定要打印到的文件的名称。</param>
    public void PrintOut_2(int? from, int? To, int? copies = null,
        bool? preview = null, string? activePrinter = null,
        bool? printToFile = null, bool? collate = null, bool? prToFileName = null);

    /// <summary>
    /// 按对象打印后的外观效果显示对象的预览。
    /// </summary>
    /// <param name="enableChanges">可选 对象。 如果为 True ，则启用对指定图表的更改。</param>
    void PrintPreview(bool? enableChanges);

    /// <summary>
    /// 删除指定索引的图表
    /// </summary>
    void Delete();

    /// <summary>
    /// 选择所有图表
    /// </summary>
    /// <param name="replace">是否替换当前选择</param>
    void Select(bool? replace = true);
    #endregion
}
