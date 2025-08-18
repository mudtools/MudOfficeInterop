//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 提供COM组件互操作功能的工具类
/// </summary>
internal static class ComInterop
{
    /// <summary>
    /// 从已运行的COM对象获取接口指针
    /// </summary>
    /// <param name="rclsid">COM类的CLSID</param>
    /// <param name="pvReserved">保留参数，必须为IntPtr.Zero</param>
    /// <param name="ppunk">返回请求的COM对象接口</param>
    [DllImport("oleaut32.dll", PreserveSig = false)]
    private static extern void GetActiveObject(
        ref Guid rclsid,
        IntPtr pvReserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object ppunk
    );

    /// <summary>
    /// 将ProgID转换为对应的CLSID
    /// </summary>
    /// <param name="lpszProgID">COM组件的ProgID字符串</param>
    /// <param name="pclsid">返回对应的CLSID GUID值</param>
    /// <returns>HRESULT结果代码</returns>
    [DllImport("ole32.dll")]
    private static extern int CLSIDFromProgID(
        [MarshalAs(UnmanagedType.LPWStr)] string lpszProgID,
        out Guid pclsid
    );

    /// <summary>
    /// 根据ProgID获取正在运行的COM对象实例
    /// </summary>
    /// <param name="progId">COM组件的ProgID标识符</param>
    /// <returns>返回活动的COM对象实例</returns>
    /// <exception cref="COMException">当无法获取CLSID或COM对象时抛出</exception>
    public static object GetActiveObject(string progId)
    {
        int hr = CLSIDFromProgID(progId, out Guid clsid);
        if (hr < 0)
            throw new COMException($"Failed to get CLSID for {progId}", hr);

        object obj;
        try
        {
            GetActiveObject(ref clsid, IntPtr.Zero, out obj);
        }
        catch (Exception ex)
        {
            throw new COMException($"Failed to get active object: {progId}", ex);
        }
        return obj;
    }
}