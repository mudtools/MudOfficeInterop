//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps
{
    internal class ExcelWorksheets : ExcelCommonSheets, IExcelWorksheets
    {
        /// <summary>
        /// 底层的 COM Worksheets 集合对象
        /// </summary>
        private MsExcel.Worksheets _worksheets;

        #region 构造函数和释放

        /// <summary>
        /// 初始化 ExcelWorksheets 实例
        /// </summary>
        /// <param name="worksheets">底层的 COM Worksheets 集合对象</param>
        internal ExcelWorksheets(MsExcel.Worksheets worksheets)
        {
            _worksheets = worksheets;
            _disposedValue = false;
        }

        /// <summary>
        /// 释放资源的核心方法
        /// </summary>
        /// <param name="disposing">是否为显式释放</param>
        protected override void Dispose(bool disposing)
        {
            if (_disposedValue) return;

            if (disposing)
            {
                try
                {
                    // 释放所有子工作表对象
                    for (int i = 1; i <= Count; i++)
                    {
                        var worksheet = this[i] as ExcelWorksheet;
                        worksheet?.Dispose();
                    }

                    // 释放底层COM对象
                    if (_worksheets != null)
                        Marshal.ReleaseComObject(_worksheets);
                }
                catch
                {
                    // 忽略释放过程中的异常
                }
                _worksheets = null;
            }

            _disposedValue = true;
            base.Dispose(disposing);
        }

        #endregion

        #region 基础属性

        /// <summary>
        /// 获取工作表集合中的工作表数量
        /// </summary>
        public override int Count => _worksheets?.Count ?? 0;

        /// <summary>
        /// 获取指定索引的工作表对象
        /// </summary>
        /// <param name="index">工作表索引（从1开始）</param>
        /// <returns>工作表对象</returns>
        public override IExcelWorksheet this[int index]
        {
            get
            {
                if (_worksheets == null || index < 1 || index > Count)
                    return null;

                try
                {
                    return _worksheets[index] is MsExcel.Worksheet worksheet ? new ExcelWorksheet(worksheet) : null;
                }
                catch
                {
                    return null;
                }
            }
        }


        /// <summary>
        /// 获取工作表集合所在的父对象
        /// </summary>
        public override object Parent => _worksheets?.Parent;

        protected override object NativeSheets => _worksheets;

        /// <summary>
        /// 获取工作表集合所在的Application对象
        /// </summary>
        public override IExcelApplication Application
        {
            get
            {
                var application = _worksheets?.Application as Microsoft.Office.Interop.Excel.Application;
                return application != null ? new ExcelApplication(application) : null;
            }
        }

        #endregion

        #region 创建和添加

        /// <summary>
        /// 向工作簿添加新的工作表
        /// </summary>
        /// <param name="before">添加到指定工作表之前</param>
        /// <param name="after">添加到指定工作表之后</param>
        /// <param name="count">添加的工作表数量</param>
        /// <param name="type">工作表类型</param>
        /// <returns>新创建的工作表对象</returns>
        public IExcelWorksheet Add(IExcelWorksheet before = null, IExcelWorksheet after = null,
                                 int count = 1, int type = 0)
        {
            if (_worksheets == null)
                return null;

            try
            {
                var beforeSheet = before as ExcelWorksheet;
                var afterSheet = after as ExcelWorksheet;

                var worksheet = _worksheets.Add(
                    beforeSheet?.Worksheet,
                    afterSheet?.Worksheet,
                    count,
                    (Microsoft.Office.Interop.Excel.XlSheetType)type
                ) as Microsoft.Office.Interop.Excel.Worksheet;

                return worksheet != null ? new ExcelWorksheet(worksheet) : null;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 批量添加工作表
        /// </summary>
        /// <param name="names">工作表名称数组</param>
        /// <param name="before">添加到指定工作表之前</param>
        /// <param name="after">添加到指定工作表之后</param>
        /// <returns>成功添加的工作表数量</returns>
        public int AddRange(string[] names, IExcelWorksheet before = null, IExcelWorksheet after = null)
        {
            if (_worksheets == null || names == null || names.Length == 0)
                return 0;

            int successCount = 0;
            foreach (string name in names)
            {
                try
                {
                    var worksheet = Add(before, after);
                    if (worksheet != null)
                    {
                        worksheet.Name = name;
                        successCount++;
                    }
                }
                catch
                {
                    // 忽略单个工作表添加异常
                }
            }
            return successCount;
        }

        /// <summary>
        /// 基于模板创建工作表
        /// </summary>
        /// <param name="templatePath">模板文件路径</param>
        /// <param name="name">工作表名称</param>
        /// <param name="before">添加到指定工作表之前</param>
        /// <param name="after">添加到指定工作表之后</param>
        /// <returns>新创建的工作表对象</returns>
        public IExcelWorksheet CreateFromTemplate(string templatePath, string name = "",
                                                IExcelWorksheet before = null, IExcelWorksheet after = null)
        {
            if (_worksheets == null || string.IsNullOrEmpty(templatePath))
                return null;

            try
            {
                var beforeSheet = before as ExcelWorksheet;
                var afterSheet = after as ExcelWorksheet;

                var worksheet = _worksheets.Add(
                    beforeSheet?.Worksheet,
                    afterSheet?.Worksheet,
                    Type.Missing,
                    templatePath
                ) as Microsoft.Office.Interop.Excel.Worksheet;

                if (worksheet != null)
                {
                    var excelWorksheet = new ExcelWorksheet(worksheet);
                    if (!string.IsNullOrEmpty(name))
                    {
                        excelWorksheet.Name = name;
                    }
                    return excelWorksheet;
                }
                return null;
            }
            catch
            {
                return null;
            }
        }

        #endregion

        #region 查找和筛选

        /// <summary>
        /// 根据可见性查找工作表
        /// </summary>
        /// <param name="visible">可见性状态</param>
        /// <returns>匹配的工作表数组</returns>
        public IExcelWorksheet[] FindByVisibility(XlSheetVisibility visible)
        {
            if (_worksheets == null || Count == 0)
                return [];

            List<IExcelWorksheet> result = [];
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    var worksheet = this[i];
                    if (worksheet != null && worksheet.Visible == visible)
                    {
                        result.Add(worksheet);
                    }
                }
                catch
                {
                    // 忽略单个工作表访问异常
                }
            }
            return result.ToArray();
        }

        /// <summary>
        /// 获取可见的工作表
        /// </summary>
        /// <returns>可见工作表数组</returns>
        public IExcelWorksheet[] GetVisibleWorksheets()
        {
            return FindByVisibility(XlSheetVisibility.xlSheetVisible);
        }

        /// <summary>
        /// 获取隐藏的工作表
        /// </summary>
        /// <returns>隐藏工作表数组</returns>
        public IExcelWorksheet[] GetHiddenWorksheets()
        {
            return FindByVisibility(XlSheetVisibility.xlSheetHidden);
        }

        /// <summary>
        /// 获取受保护的工作表
        /// </summary>
        /// <returns>受保护工作表数组</returns>
        public IExcelWorksheet[] GetProtectedWorksheets()
        {
            if (_worksheets == null || Count == 0)
                return [];

            var result = new System.Collections.Generic.List<IExcelWorksheet>();
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    var worksheet = this[i];
                    if (worksheet != null && worksheet.IsProtected)
                    {
                        result.Add(worksheet);
                    }
                }
                catch
                {
                    // 忽略单个工作表访问异常
                }
            }
            return result.ToArray();
        }

        /// <summary>
        /// 获取未受保护的工作表
        /// </summary>
        /// <returns>未受保护工作表数组</returns>
        public IExcelWorksheet[] GetUnprotectedWorksheets()
        {
            if (_worksheets == null || Count == 0)
                return new IExcelWorksheet[0];

            var result = new System.Collections.Generic.List<IExcelWorksheet>();
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    var worksheet = this[i];
                    if (worksheet != null && !worksheet.IsProtected)
                    {
                        result.Add(worksheet);
                    }
                }
                catch
                {
                    // 忽略单个工作表访问异常
                }
            }
            return result.ToArray();
        }

        #endregion

        #region 操作方法

        /// <summary>
        /// 删除所有工作表（除了第一个）
        /// </summary>
        public void Clear()
        {
            if (_worksheets == null || Count <= 1) return;

            try
            {
                // 从后往前删除，避免索引变化问题
                for (int i = Count; i >= 2; i--)
                {
                    try
                    {
                        var worksheet = _worksheets[i] as Microsoft.Office.Interop.Excel.Worksheet;
                        worksheet.Delete();
                    }
                    catch
                    {
                        // 忽略删除过程中的异常
                    }
                }
            }
            catch
            {
                // 忽略清空过程中的异常
            }
        }

        public override void Delete(int index)
        {
            if (_worksheets == null || index < 1 || index > Count)
                return;

            try
            {
                var worksheet = _worksheets[index] as Microsoft.Office.Interop.Excel.Worksheet;
                worksheet.Delete();
            }
            catch
            {
                // 忽略删除过程中的异常
            }
        }

        public override void Delete(string name)
        {
            if (_worksheets == null || string.IsNullOrEmpty(name))
                return;

            try
            {
                var worksheet = _worksheets[name] as Microsoft.Office.Interop.Excel.Worksheet;
                worksheet?.Delete();
            }
            catch
            {
                // 忽略删除过程中的异常
            }
        }

        public override void Delete(IExcelWorksheet worksheet)
        {
            if (_worksheets == null || worksheet == null)
                return;

            try
            {
                worksheet.Delete();
            }
            catch
            {
                // 忽略删除过程中的异常
            }
        }

        /// <summary>
        /// 移动工作表
        /// </summary>
        /// <param name="worksheet">要移动的工作表</param>
        /// <param name="before">移动到指定工作表之前</param>
        /// <param name="after">移动到指定工作表之后</param>
        public void Move(IExcelWorksheet worksheet, IExcelWorksheet before = null, IExcelWorksheet after = null)
        {
            if (_worksheets == null || worksheet == null)
                return;

            try
            {
                worksheet.Move(before, after);
            }
            catch
            {
                // 忽略移动过程中的异常
            }
        }

        /// <summary>
        /// 复制工作表
        /// </summary>
        /// <param name="worksheet">要复制的工作表</param>
        /// <param name="before">复制到指定工作表之前</param>
        /// <param name="after">复制到指定工作表之后</param>
        /// <param name="newName">新工作表名称</param>
        /// <returns>复制的工作表对象</returns>
        public IExcelWorksheet Copy(IExcelWorksheet worksheet, IExcelWorksheet? before = null,
                                  IExcelWorksheet? after = null, string newName = "")
        {
            if (_worksheets == null || worksheet == null)
                return null;

            try
            {
                worksheet.Copy(before, after);

                // 获取复制后的工作表
                var copiedWorksheet = this[Count] as ExcelWorksheet;
                if (copiedWorksheet != null && !string.IsNullOrEmpty(newName))
                {
                    copiedWorksheet.Name = newName;
                }

                return copiedWorksheet;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 选择多个工作表
        /// </summary>
        /// <param name="worksheetNames">工作表名称数组</param>
        public override void Select(params string[] worksheetNames)
        {
            if (_worksheets == null || worksheetNames == null || worksheetNames.Length == 0)
                return;

            try
            {
                object[] sheets = new object[worksheetNames.Length];
                for (int i = 0; i < worksheetNames.Length; i++)
                {
                    sheets[i] = worksheetNames[i];
                }
                _worksheets.Select(sheets);
            }
            catch
            {

            }
        }

        #endregion

        /// <summary>
        /// 按指定顺序排列工作表
        /// </summary>
        /// <param name="names">工作表名称顺序数组</param>
        public void ArrangeInOrder(string[] names)
        {
            if (_worksheets == null || names == null || names.Length == 0)
                return;

            try
            {
                // 按照指定顺序重新排列工作表
                for (int i = 0; i < names.Length; i++)
                {
                    try
                    {
                        if (_worksheets[names[i]] is MsExcel.Worksheet worksheet && worksheet.Index != i + 1)
                        {
                            if (i == 0)
                                worksheet.Move(_worksheets[1], Type.Missing);
                            else
                                worksheet.Move(Type.Missing, _worksheets[i]);
                        }
                    }
                    catch
                    {
                        // 忽略单个工作表排列异常
                    }
                }
            }
            catch
            {
                // 忽略排列过程中的异常
            }
        }

        /// <summary>
        /// 按名称排序工作表
        /// </summary>
        /// <param name="ascending">是否升序排列</param>
        public void SortByName(bool ascending = true)
        {
            if (_worksheets == null || Count <= 1)
                return;

            try
            {
                // 获取所有工作表名称
                var names = new string[Count];
                for (int i = 1; i <= Count; i++)
                {
                    names[i - 1] = this[i]?.Name ?? "";
                }

                // 排序名称
                if (ascending)
                    Array.Sort(names);
                else
                    Array.Sort(names, (x, y) => y.CompareTo(x));

                // 按排序后的顺序重新排列
                ArrangeInOrder(names);
            }
            catch
            {
                // 忽略排序过程中的异常
            }
        }

        #region 高级功能

        /// <summary>
        /// 获取活动工作表
        /// </summary>
        /// <returns>活动工作表对象</returns>
        public override IExcelWorksheet ActiveWorksheet
        {
            get
            {
                try
                {
                    var wb = _worksheets?.Parent as MsExcel.Workbook;
                    if (wb == null)
                        return null;
                    var activeSheet = wb.ActiveSheet as MsExcel.Worksheet;
                    return activeSheet != null ? new ExcelWorksheet(activeSheet) : null;
                }
                catch
                {
                    return null;
                }
            }
        }

        /// <summary>
        /// 打印所有工作表
        /// </summary>
        /// <param name="preview">是否打印预览</param>
        public override void PrintOutAll(bool preview = false)
        {
            if (_worksheets == null) return;

            try
            {
                if (preview)
                {
                    _worksheets.PrintPreview();
                }
                else
                {
                    _worksheets.PrintOut(
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing
                    );
                }
            }
            catch
            {
                // 忽略打印过程中的异常
            }
        }

        /// <summary>
        /// 计算所有工作表
        /// </summary>
        public override void Calculate()
        {
            if (_worksheets == null) return;

            try
            {
                for (int i = 1; i <= Count; i++)
                {
                    try
                    {
                        this[i]?.Calculate();
                    }
                    catch
                    {
                        // 忽略单个工作表计算异常
                    }
                }
            }
            catch
            {
                // 忽略计算过程中的异常
            }
        }

        /// <summary>
        /// 刷新所有工作表
        /// </summary>
        public override void RefreshAll()
        {
            if (_worksheets == null) return;

            try
            {
                for (int i = 1; i <= Count; i++)
                {
                    try
                    {
                        this[i]?.Recalculate();
                    }
                    catch
                    {
                        // 忽略单个工作表刷新异常
                    }
                }
            }
            catch
            {
                // 忽略刷新过程中的异常
            }
        }

        /// <summary>
        /// 隐藏所有工作表
        /// </summary>
        public void HideAll()
        {
            if (_worksheets == null || Count == 0) return;

            try
            {
                // 保留第一个工作表可见，隐藏其余工作表
                for (int i = 2; i <= Count; i++)
                {
                    try
                    {
                        this[i].Visible = XlSheetVisibility.xlSheetHidden;
                    }
                    catch
                    {
                        // 忽略单个工作表隐藏异常
                    }
                }
            }
            catch
            {
                // 忽略隐藏过程中的异常
            }
        }

        /// <summary>
        /// 显示所有工作表
        /// </summary>
        public void ShowAll()
        {
            if (_worksheets == null || Count == 0) return;

            try
            {
                for (int i = 1; i <= Count; i++)
                {
                    try
                    {
                        this[i].Visible = XlSheetVisibility.xlSheetVisible;
                    }
                    catch
                    {
                        // 忽略单个工作表显示异常
                    }
                }
            }
            catch
            {
                // 忽略显示过程中的异常
            }
        }

        public override IEnumerator<IExcelWorksheet> GetEnumerator()
        {
            for (int i = 0; i < Count; i++)
            {
                yield return this[i];
            }
        }

        #endregion
    }
}