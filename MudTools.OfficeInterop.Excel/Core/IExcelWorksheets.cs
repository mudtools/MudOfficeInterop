//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel
{
    /// <summary>
    /// Excel工作表集合接口 (适用于 Worksheets 对象)
    /// </summary>
    public interface IExcelWorksheets : IExcelCommonSheets
    {
        #region 创建和添加

        /// <summary>
        /// 向工作簿添加新的工作表
        /// </summary>
        /// <param name="before">添加到指定工作表之前</param>
        /// <param name="after">添加到指定工作表之后</param>
        /// <param name="count">添加的工作表数量</param>
        /// <param name="type">工作表类型</param>
        /// <returns>新创建的工作表对象</returns>
        IExcelWorksheet Add(IExcelWorksheet? before = null, IExcelWorksheet? after = null,
                           int count = 1, int type = 0);

        /// <summary>
        /// 批量添加工作表
        /// </summary>
        /// <param name="names">工作表名称数组</param>
        /// <param name="before">添加到指定工作表之前</param>
        /// <param name="after">添加到指定工作表之后</param>
        /// <returns>成功添加的工作表数量</returns>
        int AddRange(string[] names, IExcelWorksheet? before = null, IExcelWorksheet? after = null);

        /// <summary>
        /// 基于模板创建工作表
        /// </summary>
        /// <param name="templatePath">模板文件路径</param>
        /// <param name="name">工作表名称</param>
        /// <param name="before">添加到指定工作表之前</param>
        /// <param name="after">添加到指定工作表之后</param>
        /// <returns>新创建的工作表对象</returns>
        IExcelWorksheet? CreateFromTemplate(string templatePath, string name = "",
                                         IExcelWorksheet? before = null, IExcelWorksheet? after = null);

        #endregion

        #region 查找和筛选

        /// <summary>
        /// 根据可见性查找工作表
        /// </summary>
        /// <param name="visible">可见性状态</param>
        /// <returns>匹配的工作表数组</returns>
        IExcelWorksheet[] FindByVisibility(XlSheetVisibility visible);

        /// <summary>
        /// 获取可见的工作表
        /// </summary>
        /// <returns>可见工作表数组</returns>
        IExcelWorksheet[] GetVisibleWorksheets();

        /// <summary>
        /// 获取隐藏的工作表
        /// </summary>
        /// <returns>隐藏工作表数组</returns>
        IExcelWorksheet[] GetHiddenWorksheets();

        /// <summary>
        /// 获取受保护的工作表
        /// </summary>
        /// <returns>受保护工作表数组</returns>
        IExcelWorksheet[] GetProtectedWorksheets();

        /// <summary>
        /// 获取未受保护的工作表
        /// </summary>
        /// <returns>未受保护工作表数组</returns>
        IExcelWorksheet[] GetUnprotectedWorksheets();

        #endregion

        #region 操作方法

        /// <summary>
        /// 删除所有工作表（除了第一个）
        /// 对应 Worksheets.Delete 方法
        /// </summary>
        void Clear();

        /// <summary>
        /// 移动工作表
        /// </summary>
        /// <param name="worksheet">要移动的工作表</param>
        /// <param name="before">移动到指定工作表之前</param>
        /// <param name="after">移动到指定工作表之后</param>
        void Move(IExcelWorksheet worksheet, IExcelWorksheet? before = null, IExcelWorksheet? after = null);

        /// <summary>
        /// 复制工作表
        /// </summary>
        /// <param name="worksheet">要复制的工作表</param>
        /// <param name="before">复制到指定工作表之前</param>
        /// <param name="after">复制到指定工作表之后</param>
        /// <param name="newName">新工作表名称</param>
        /// <returns>复制的工作表对象</returns>
        IExcelWorksheet? Copy(IExcelWorksheet worksheet, IExcelWorksheet? before = null,
                            IExcelWorksheet? after = null, string newName = "");

        #endregion

        #region 排列和布局
        /// <summary>
        /// 按指定顺序排列工作表
        /// </summary>
        /// <param name="names">工作表名称顺序数组</param>
        void ArrangeInOrder(string[] names);

        /// <summary>
        /// 按名称排序工作表
        /// </summary>
        /// <param name="ascending">是否升序排列</param>
        void SortByName(bool ascending = true);
        #endregion


        #region 高级功能

        /// <summary>
        /// 隐藏所有工作表
        /// </summary>
        void HideAll();

        /// <summary>
        /// 显示所有工作表
        /// </summary>
        void ShowAll();

        #endregion
    }
}
