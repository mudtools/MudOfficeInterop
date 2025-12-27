//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示 Excel 中注音符号集合的封装接口
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelPhonetics : IOfficeObject<IExcelPhonetics>, IEnumerable<IExcelPhonetic?>, IDisposable
{
    /// <summary>
    /// 获取当前COM对象的父对象。
    /// 对应 RecentFile.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取当前COM对象的Application对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取注音符号的数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的注音符号对象
    /// </summary>
    /// <param name="index">注音符号索引（从1开始）</param>
    /// <returns>注音符号对象</returns>
    [ComPropertyWrap(NeedConvert = true)]
    IExcelPhonetic? this[int index] { get; }

    /// <summary>
    /// 获取注音符号的字体属性。
    /// </summary>
    IExcelFont? Font { get; }

    /// <summary>
    /// 获取或设置注音符号的对齐方式
    /// </summary>
    int Alignment { get; set; }

    /// <summary>
    /// 获取或设置注音符号的字符类型
    /// </summary>
    int CharacterType { get; set; }

    /// <summary>
    /// 向集合中添加新的注音符号
    /// </summary>
    /// <param name="start">开始位置</param>
    /// <param name="length">长度</param>
    /// <param name="text">注音文本</param>
    /// <returns>新创建的注音符号对象</returns>
    void Add(int start, int length, string text);

    /// <summary>
    /// 删除所有注音符号
    /// </summary>
    void Delete();
}