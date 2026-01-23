//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel ChartObject 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.ChartObject 的安全访问和操作
/// </summary>
public interface IExcelChartObject : IExcelComGraphObject, IDisposable
{
    #region 位置和大小

    /// <summary>
    /// 获取或设置图表对象的旋转角度
    /// 对应 ChartObject.Shape.Rotation 属性
    /// </summary>
    double Rotation { get; set; }

    #endregion

    #region 图表属性

    /// <summary>
    /// 获取图表对象的图表
    /// 对应 ChartObject.Chart 属性
    /// </summary>
    IExcelChart? Chart { get; }

    /// <summary>
    /// 获取图表对象是否为嵌入式图表
    /// </summary>
    bool IsEmbedded { get; }
    #endregion

    #region 操作方法
    /// <summary>
    /// 激活对象
    /// </summary>
    void Activate();

    /// <summary>
    /// 调整图表对象大小
    /// </summary>
    /// <param name="width">新宽度</param>
    /// <param name="height">新高度</param>
    /// <param name="keepAspectRatio">是否保持纵横比</param>
    void Resize(double width, double height, bool keepAspectRatio = false);

    /// <summary>
    /// 旋转图表对象
    /// </summary>
    /// <param name="angle">旋转角度</param>
    void Rotate(double angle);

    #endregion

    #region 导出和转换   

    /// <summary>
    /// 复制图表到新图表对象
    /// </summary>
    /// <param name="worksheetName">新图表对象名称</param>
    /// <returns>新创建的图表对象对象</returns>
    IExcelWorksheet? CopyToNewWorksheet(string worksheetName = "");
    #endregion
}
