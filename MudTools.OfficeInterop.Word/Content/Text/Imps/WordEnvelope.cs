namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Microsoft.Office.Interop.Word.Envelope COM 对象的托管包装器实现。
/// 负责与底层 COM 对象交互并管理其生命周期。
/// </summary>
internal class WordEnvelope : IWordEnvelope
{
    private MsWord.Envelope? _envelope;
    private bool _disposedValue;

    /// <summary>
    /// 初始化一个新的 WordEnvelope 实例。
    /// </summary>
    /// <param name="envelope">底层的 Word Envelope COM 对象。</param>
    /// <exception cref="ArgumentNullException">当 envelope 为 null 时抛出。</exception>
    internal WordEnvelope(MsWord.Envelope envelope)
    {
        _envelope = envelope ?? throw new ArgumentNullException(nameof(envelope));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc />
    public IWordRange? Address
    {
        get => _envelope != null ? new WordRange(_envelope.Address) : null;
    }

    /// <inheritdoc />
    public IWordRange? ReturnAddress
    {
        get => _envelope != null ? new WordRange(_envelope.ReturnAddress) : null;
    }


    /// <inheritdoc />
    public WdPaperTray FeedSource
    {
        get => _envelope?.FeedSource.EnumConvert(WdPaperTray.wdPrinterDefaultBin) ?? WdPaperTray.wdPrinterDefaultBin;
        set
        {
            if (_envelope != null)
                _envelope.FeedSource = value.EnumConvert(MsWord.WdPaperTray.wdPrinterDefaultBin);
        }
    }

    /// <inheritdoc />
    public float AddressFromLeft
    {
        get => _envelope?.AddressFromLeft ?? 0f;
        set
        {
            if (_envelope != null)
                _envelope.AddressFromLeft = value;
        }
    }

    /// <inheritdoc />
    public float AddressFromTop
    {
        get => _envelope?.AddressFromTop ?? 0f;
        set
        {
            if (_envelope != null)
                _envelope.AddressFromTop = value;
        }
    }

    /// <inheritdoc />
    public float ReturnAddressFromLeft
    {
        get => _envelope?.ReturnAddressFromLeft ?? 0f;
        set
        {
            if (_envelope != null)
                _envelope.ReturnAddressFromLeft = value;
        }
    }

    /// <inheritdoc />
    public float ReturnAddressFromTop
    {
        get => _envelope?.ReturnAddressFromTop ?? 0f;
        set
        {
            if (_envelope != null)
                _envelope.ReturnAddressFromTop = value;
        }
    }

    public float DefaultWidth
    {
        get => _envelope?.DefaultWidth ?? 0f;
        set
        {
            if (_envelope != null)
                _envelope.DefaultWidth = value;
        }
    }

    public float DefaultHeight
    {
        get => _envelope?.DefaultHeight ?? 0f;
        set
        {
            if (_envelope != null)
                _envelope.DefaultHeight = value;
        }
    }

    public bool DefaultPrintBarCode
    {
        get => _envelope?.DefaultPrintBarCode ?? false;
        set
        {
            if (_envelope != null)
                _envelope.DefaultPrintBarCode = value;
        }
    }

    public string DefaultSize
    {
        get => _envelope?.DefaultSize ?? string.Empty;
        set
        {
            if (_envelope != null)
                _envelope.DefaultSize = value;
        }
    }

    public bool Vertical
    {
        get => _envelope?.Vertical ?? false;
        set
        {
            if (_envelope != null)
                _envelope.Vertical = value;
        }
    }

    public bool DefaultFaceUp
    {
        get => _envelope?.DefaultFaceUp ?? false;
        set
        {
            if (_envelope != null)
                _envelope.DefaultFaceUp = value;
        }
    }

    public bool DefaultOmitReturnAddress
    {
        get => _envelope?.DefaultOmitReturnAddress ?? false;
        set
        {
            if (_envelope != null)
                _envelope.DefaultOmitReturnAddress = value;
        }
    }

    public WdEnvelopeOrientation DefaultOrientation
    {
        get => _envelope?.DefaultOrientation.EnumConvert(WdEnvelopeOrientation.wdLeftPortrait) ?? WdEnvelopeOrientation.wdLeftPortrait;
        set
        {
            if (_envelope != null)
                _envelope.DefaultOrientation = value.EnumConvert(MsWord.WdEnvelopeOrientation.wdLeftPortrait);
        }
    }

    public IWordStyle? ReturnAddressStyle
    {
        get => _envelope != null ? new WordStyle(_envelope.ReturnAddressStyle) : null;
    }

    public IWordStyle? AddressStyle
    {
        get => _envelope != null ? new WordStyle(_envelope.AddressStyle) : null;
    }

    public float RecipientNamefromLeft
    {
        get => _envelope?.RecipientNamefromLeft ?? 0f;
        set
        {
            if (_envelope != null)
                _envelope.RecipientNamefromLeft = value;
        }
    }

    public float RecipientNamefromTop
    {
        get => _envelope?.RecipientNamefromTop ?? 0f;
        set
        {
            if (_envelope != null)
                _envelope.RecipientNamefromTop = value;
        }
    }

    public float RecipientPostalfromLeft
    {
        get => _envelope?.RecipientPostalfromLeft ?? 0f;
        set
        {
            if (_envelope != null)
                _envelope.RecipientPostalfromLeft = value;
        }
    }

    public float RecipientPostalfromTop
    {
        get => _envelope?.RecipientPostalfromTop ?? 0f;
        set
        {
            if (_envelope != null)
                _envelope.RecipientPostalfromTop = value;
        }
    }

    public float SenderNamefromLeft
    {
        get => _envelope?.SenderNamefromLeft ?? 0f;
        set
        {
            if (_envelope != null)
                _envelope.SenderNamefromLeft = value;
        }
    }

    public float SenderNamefromTop
    {
        get => _envelope?.SenderNamefromTop ?? 0f;
        set
        {
            if (_envelope != null)
                _envelope.SenderNamefromTop = value;
        }
    }

    public float SenderPostalfromLeft
    {
        get => _envelope?.SenderPostalfromLeft ?? 0f;
        set
        {
            if (_envelope != null)
                _envelope.SenderPostalfromLeft = value;
        }
    }

    public float SenderPostalfromTop
    {
        get => _envelope?.SenderPostalfromTop ?? 0f;
        set
        {
            if (_envelope != null)
                _envelope.SenderPostalfromTop = value;
        }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc />
    public void Options()
    {
        if (_envelope != null)
            _envelope.Options();
    }

    /// <inheritdoc />
    public void Insert(
        string? address = null,
        string? returnAddress = null,
        string? autoText = null,
        bool omitReturnAddress = false,
        bool printBarcode = false,
        bool printFIMA = false,
        string? size = null,
        int? feedSource = null)
    {
        EnsureNotDisposed();

        // 将可空参数转换为 object 类型，因为 COM 方法需要 ref object
        var addressObj = (object?)address ?? Type.Missing;
        var returnAddressObj = (object?)returnAddress ?? Type.Missing;
        var autoTextObj = (object?)autoText ?? Type.Missing;
        var omitReturnAddressObj = (object)omitReturnAddress;
        var printBarcodeObj = (object)printBarcode;
        var printFIMAObj = (object)printFIMA;
        var sizeObj = (object?)size ?? Type.Missing;
        var feedSourceObj = (object?)(feedSource ?? (int?)null) ?? Type.Missing;

        try
        {
            _envelope.Insert(
                ref addressObj,
                ref returnAddressObj,
                ref autoTextObj,
                ref omitReturnAddressObj,
                ref printBarcodeObj,
                ref printFIMAObj,
                ref sizeObj,
                ref feedSourceObj);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法在文档中插入信封。", ex);
        }
    }

    /// <inheritdoc />
    public void PrintOut(
        string? address = null,
        string? returnAddress = null,
        string? autoText = null,
        bool omitReturnAddress = false,
        bool printBarcode = false,
        bool printFIMA = false,
        string? size = null,
        int? feedSource = null)
    {
        EnsureNotDisposed();

        // 将可空参数转换为 object 类型，因为 COM 方法需要 ref object
        var addressObj = (object?)address ?? Type.Missing;
        var returnAddressObj = (object?)returnAddress ?? Type.Missing;
        var autoTextObj = (object?)autoText ?? Type.Missing;
        var omitReturnAddressObj = (object)omitReturnAddress;
        var printBarcodeObj = (object)printBarcode;
        var printFIMAObj = (object)printFIMA;
        var sizeObj = (object?)size ?? Type.Missing;
        var feedSourceObj = (object?)(feedSource ?? (int?)null) ?? Type.Missing;

        try
        {
            _envelope.PrintOut(
                ref addressObj,
                ref returnAddressObj,
                ref autoTextObj,
                ref omitReturnAddressObj,
                ref printBarcodeObj,
                ref printFIMAObj,
                ref sizeObj,
                ref feedSourceObj);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法打印信封。", ex);
        }
    }

    /// <inheritdoc />
    public void UpdateDocument()
    {
        EnsureNotDisposed();
        try
        {
            _envelope.UpdateDocument();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法更新文档中的信封内容。", ex);
        }
    }

    /// <summary>
    /// 检查对象是否已被释放，如果已被释放则抛出异常。
    /// </summary>
    /// <exception cref="ObjectDisposedException">当对象已被释放时抛出。</exception>
    private void EnsureNotDisposed()
    {
        if (_disposedValue)
            throw new ObjectDisposedException(nameof(WordEnvelope));
        if (_envelope == null)
            throw new ObjectDisposedException(nameof(WordEnvelope));
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放由 WordEnvelope 使用的资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            if (disposing && _envelope != null)
            {
                // 释放对 COM 对象的引用
                Marshal.ReleaseComObject(_envelope);
                _envelope = null;
            }
            _disposedValue = true;
        }
    }

    /// <inheritdoc />
    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }

    #endregion
}