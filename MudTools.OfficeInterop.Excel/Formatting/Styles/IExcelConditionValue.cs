//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel ConditionValue 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.ConditionValue 的安全访问和操作
/// ConditionValue 对象通常用于定义 ColorScale 和 Databar 的最小值、中间值（ColorScale）和最大值点。
/// </summary>
public interface IExcelConditionValue : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取条件值对象的父对象 (通常是 ColorScaleCriterion 或 DataBar)
    /// 对应 ConditionValue.Parent 属性
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取条件值对象所在的Application对象
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取或设置条件值的类型
    /// 对应 ConditionValue.Type 属性
    /// </summary>
    XlConditionValueTypes Type { get; }

    /// <summary>
    /// 获取或设置条件值
    /// 对应 ConditionValue.Value 属性
    /// </summary>
    object Value { get; }
    #endregion


    /// <summary>
    /// 修改条件值的类型和值
    /// </summary>
    /// <param name="newtype">新的条件值类型</param>
    /// <param name="newvalue">新的条件值，可以为 null</param>
    void Modify(XlConditionValueTypes newtype, object? newvalue = null);
}
