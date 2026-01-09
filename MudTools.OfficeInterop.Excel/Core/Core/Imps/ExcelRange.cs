//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.Excel.Imps;

partial class ExcelRange
{

    ///  <inheritdoc/>
    [global::System.CodeDom.Compiler.GeneratedCode("Mud.ServiceCodeGenerator", "1.4.7")]
    public MudTools.OfficeInterop.Excel.IExcelRange? this[int? index1, int? index2]
    {
        get
        {
            if (_range == null)
                throw new ObjectDisposedException(nameof(_range));
            if (index1 == null)
                throw new ArgumentNullException(nameof(index1));
            if (_range == null)
                throw new ObjectDisposedException(nameof(_range));
            if (index2 == null)
                throw new ArgumentNullException(nameof(index2));

            try
            {
                var comElement = _range.Item[index1, index2];
                if (comElement is MsExcel.Range rComObj)
                    return new MudTools.OfficeInterop.Excel.Imps.ExcelRange(rComObj);
                else
                    return null;
            }
            catch (COMException ce)
            {
                throw new ExcelOperationException("根据双索引获取 MudTools.OfficeInterop.Excel.Imps.ExcelRange 对象失败: " + ce.Message, ce);
            }
            catch (Exception ex)
            {
                throw new ExcelOperationException("根据双索引获取 MudTools.OfficeInterop.Excel.Imps.ExcelRange 对象失败", ex);
            }
        }
        set
        {
            if (_range == null)
                throw new ObjectDisposedException(nameof(_range));
            if (index1 == null)
                throw new ArgumentNullException(nameof(index1));
            if (_range == null)
                throw new ObjectDisposedException(nameof(_range));
            if (index2 == null)
                throw new ArgumentNullException(nameof(index2));

            try
            {
                if (value == null)
                    throw new ArgumentNullException(nameof(value));

                _range.Item[index1, index2] = ((Imps.ExcelRange)value).InternalComObject;
            }
            catch (COMException ce)
            {
                throw new ExcelOperationException("根据双索引设置 MudTools.OfficeInterop.Excel.Imps.ExcelRange 对象失败: " + ce.Message, ce);
            }
            catch (Exception ex)
            {
                throw new ExcelOperationException("根据双索引设置 MudTools.OfficeInterop.Excel.Imps.ExcelRange 对象失败", ex);
            }
        }
    }

    ///  <inheritdoc/>
    [global::System.CodeDom.Compiler.GeneratedCode("Mud.ServiceCodeGenerator", "1.4.7")]
    public MudTools.OfficeInterop.Excel.IExcelRange? this[string? index1, string? index2]
    {
        get
        {
            if (_range == null)
                throw new ObjectDisposedException(nameof(_range));
            if (index1 == null)
                throw new ArgumentNullException(nameof(index1));
            if (_range == null)
                throw new ObjectDisposedException(nameof(_range));
            if (index2 == null)
                throw new ArgumentNullException(nameof(index2));

            try
            {
                var comElement = _range.Item[index1, index2];
                if (comElement is MsExcel.Range rComObj)
                    return new MudTools.OfficeInterop.Excel.Imps.ExcelRange(rComObj);
                else
                    return null;
            }
            catch (COMException ce)
            {
                throw new ExcelOperationException("根据双索引获取 MudTools.OfficeInterop.Excel.Imps.ExcelRange 对象失败: " + ce.Message, ce);
            }
            catch (Exception ex)
            {
                throw new ExcelOperationException("根据双索引获取 MudTools.OfficeInterop.Excel.Imps.ExcelRange 对象失败", ex);
            }
        }
        set
        {
            if (_range == null)
                throw new ObjectDisposedException(nameof(_range));
            if (index1 == null)
                throw new ArgumentNullException(nameof(index1));
            if (_range == null)
                throw new ObjectDisposedException(nameof(_range));
            if (index2 == null)
                throw new ArgumentNullException(nameof(index2));

            try
            {
                if (value == null)
                    throw new ArgumentNullException(nameof(value));
                _range.Item[index1, index2] = ((Imps.ExcelRange)value).InternalComObject;
            }
            catch (COMException ce)
            {
                throw new ExcelOperationException("根据双索引设置 MudTools.OfficeInterop.Excel.Imps.ExcelRange 对象失败: " + ce.Message, ce);
            }
            catch (Exception ex)
            {
                throw new ExcelOperationException("根据双索引设置 MudTools.OfficeInterop.Excel.Imps.ExcelRange 对象失败", ex);
            }
        }
    }

    ///  <inheritdoc/>
    [global::System.CodeDom.Compiler.GeneratedCode("Mud.ServiceCodeGenerator", "1.4.7")]
    public MudTools.OfficeInterop.Excel.IExcelRange? this[string index]
    {
        get
        {
            if (_range == null)
                throw new ObjectDisposedException(nameof(_range));
            if (string.IsNullOrEmpty(index))
                throw new ArgumentNullException(nameof(index));

            try
            {
                var comElement = _range.Item[index];
                MudTools.OfficeInterop.Excel.IExcelRange? result = null;
                if (comElement is MsExcel.Range rComObj)
                    result = new MudTools.OfficeInterop.Excel.Imps.ExcelRange(rComObj);
                if (result != null)
                    _disposableList.Add(result);
                return result;
            }
            catch (COMException ce)
            {
                throw new ExcelOperationException("根据索引检索 MudTools.OfficeInterop.Excel.Imps.ExcelRange 对象失败: " + ce.Message, ce);
            }
        }
    }

    ///  <inheritdoc/>
    public MudTools.OfficeInterop.Excel.IExcelRange? this[int index]
    {
        get
        {
            if (_range == null)
                throw new ObjectDisposedException(nameof(_range));
            if (index < 1)
                throw new IndexOutOfRangeException("第一个索引参数不能少于1");

            try
            {
                var comElement = _range.Item[index];
                MudTools.OfficeInterop.Excel.IExcelRange? result = null;
                if (comElement is MsExcel.Range rComObj)
                    result = new MudTools.OfficeInterop.Excel.Imps.ExcelRange(rComObj);
                if (result != null)
                    _disposableList.Add(result);
                return result;
            }
            catch (COMException ce)
            {
                throw new ExcelOperationException("根据索引检索 MudTools.OfficeInterop.Excel.Imps.ExcelRange 对象失败: " + ce.Message, ce);
            }
        }
    }

    public object[,]? ArrayValue
    {
        get
        {
            if (_range == null)
                throw new ObjectDisposedException(nameof(_range));
            var obj = _range.Value2;
            if (obj is object[,] objs && objs != null)
                return objs;
            return null;
        }
        set
        {
            if (_range == null)
                throw new ObjectDisposedException(nameof(_range));
            _range.Value2 = value;
        }
    }

    ///  <inheritdoc/>
    public object? GetValue(XlRangeValueDataType rangeValueDataType)
    {
        if (_range == null)
            throw new ObjectDisposedException(nameof(_range));
        var rangeValueDataTypeObj = rangeValueDataType.EnumConvert(MsExcel.XlRangeValueDataType.xlRangeValueDefault);

        try
        {
            var returnValue = _range?.Value[rangeValueDataTypeObj];
            return returnValue;
        }
        catch (COMException cx)
        {
            throw new ExcelOperationException("执行GetValue操作失败: " + cx.Message, cx);
        }
        catch (Exception ex)
        {
            throw new ExcelOperationException("执行GetValue操作失败", ex);
        }
    }

    ///  <inheritdoc/>
    public void SetValue(XlRangeValueDataType rangeValueDataType, object? value)
    {
        if (_range == null)
            throw new ObjectDisposedException(nameof(_range));
        var rangeValueDataTypeObj = rangeValueDataType.EnumConvert(MsExcel.XlRangeValueDataType.xlRangeValueDefault);
        var valueObj = value != null ? (object)value : global::System.Type.Missing;

        try
        {
            _range?.Value[rangeValueDataTypeObj] = valueObj;
        }
        catch (COMException cx)
        {
            throw new ExcelOperationException("执行SetValue操作失败: " + cx.Message, cx);
        }
        catch (Exception ex)
        {
            throw new ExcelOperationException("执行SetValue操作失败", ex);
        }
    }
}