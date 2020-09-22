VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPNGparser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' PNG Parser & PNG to 32bpp converter
' The PNG will be parsed using the following resources if they are available
' and in the following order.
' 1) If GDI+ is available, the entire PNG will be processed via GDI+
' 2) If zLIB.DLL is available, the PNG will be decompressed via zLIB
' 3) If none of the above, the PNG will be decompressed with pure VB

' No APIs are declared public. This is to prevent possibly, differently
' declared APIs, or different versions of the same API, from conflciting
' with any APIs you declared in your project. Same rule for UDTs.
' Note: I did take some liberties in several API declarations throughout

' following are used for saving dib to PNG (testing phase only)
Private Declare Function GdipImageRotateFlip Lib "GDIPlus" (ByVal Image As Long, ByVal rfType As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "GDIPlus" (ByVal Width As Long, ByVal Height As Long, ByVal stride As Long, ByVal PixelFormat As Long, scan0 As Any, BITMAP As Long) As Long
Private Declare Function GdipCreateBitmapFromGdiDib Lib "GDIPlus" (gdiBitmapInfo As BITMAPINFO, gdiBitmapData As Any, BITMAP As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal psString As Any) As Long
Private Declare Function GdipGetImageEncodersSize Lib "GDIPlus" (numEncoders As Long, Size As Long) As Long
Private Declare Function GdipGetImageEncoders Lib "GDIPlus" (ByVal numEncoders As Long, ByVal Size As Long, Encoders As Any) As Long
Private Declare Function GdipSaveImageToFile Lib "GDIPlus" (ByVal hImage As Long, ByVal sFilename As String, clsidEncoder As Any, encoderParams As Any) As Long
Private Declare Function GdipSaveImageToStream Lib "GDIPlus" (ByVal Image As Long, ByVal Stream As IUnknown, clsidEncoder As Any, encoderParams As Any) As Long
Private Enum EncoderParameterValueType
    [EncoderParameterValueTypeByte] = 1
    [EncoderParameterValueTypeASCII] = 2
    [EncoderParameterValueTypeShort] = 3
    [EncoderParameterValueTypeLong] = 4
    [EncoderParameterValueTypeRational] = 5
    [EncoderParameterValueTypeLongRange] = 6
    [EncoderParameterValueTypeUndefined] = 7
    [EncoderParameterValueTypeRationalRange] = 8
End Enum
Private Type EncoderParameter
    GUID(0 To 3)   As Long
    NumberOfValues As Long
    Type           As EncoderParameterValueType
    Value          As Long
End Type
'-- Encoder Parameters structure
Private Type EncoderParameters
    count     As Long
    Parameter As EncoderParameter
End Type
Private Type ImageCodecInfo
    ClassID(0 To 3)   As Long
    FormatID(0 To 3)  As Long
    CodecName         As Long
    DllName           As Long
    FormatDescription As Long
    FilenameExtension As Long
    MimeType          As Long
    Flags             As Long
    Version           As Long
    SigCount          As Long
    SigSize           As Long
    SigPattern        As Long
    SigMask           As Long
End Type


' Used to create a return DIB section
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hDC As Long, ByRef pBitmapInfo As Any, ByVal un As Long, ByRef lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Type SafeArrayBound
    cElements As Long
    lLbound As Long
End Type
Private Type SafeArray1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    rgSABound As SafeArrayBound
End Type
Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As Long
End Type

Private Const PixelFormat32bppARGB As Long = &H26200A
Private Const PixelFormat32bppPARGB As Long = &HE200B

' used to see if DLL exported function exists
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

' used for workaround of VB not exposing IStream interface
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GetHGlobalFromStream Lib "ole32" (ByVal ppstm As Long, hGlobal As Long) As Long


' Following are used only if PNG file is being manually decompressed with pure VB
Private Type CodesType
    Length() As Long
    code() As Long
End Type
Private OutPos As Long
Private Inpos As Long
Private ByteBuff As Long
Private BitNum As Long
Private BitMask() As Long
Private Pow2() As Long

Private LCodes As CodesType
Private DCodes As CodesType
Private LitLen As CodesType
Private Dist As CodesType
Private TempLit As CodesType
Private TempDist As CodesType
Private LenOrder() As Long

Private inStream() As Byte  ' overlay only; nevery initialized
Private outStream() As Byte ' overlay only; nevery initialized


' Following are used only if PNG file is being processed by GDI+
Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type
Private Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "GDIPlus" (ByVal token As Long)
Private Declare Function GdipLoadImageFromFile Lib "GDIPlus" (ByVal FileName As String, hImage As Long) As Long
Private Declare Function GdipLoadImageFromStream Lib "GDIPlus" (ByVal Stream As IUnknown, Image As Long) As Long
Private Declare Function GdipGetImageBounds Lib "gdiplus.dll" (ByVal nImage As Long, srcRect As RECTF, srcUnit As Long) As Long
Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "GDIPlus" (ByVal hDC As Long, hGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "GDIPlus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "GDIPlus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal CallbackData As Long = 0) As Long
Private Declare Function GdipGetImagePixelFormat Lib "GDIPlus" (ByVal Image As Long, PixelFormat As Long) As Long
Private Const UnitPixel As Long = 2&
Private Type RECTF
    nLeft As Single
    nTop As Single
    nWidth As Single
    nHeight As Single
End Type


' Following are used if PNG will be decompressed by zLIB
Private Declare Function Zuncompress Lib "zlib.dll" Alias "uncompress" (ByRef Dest As Any, ByRef destLen As Long, ByRef Source As Any, ByVal sourceLen As Long) As Long
Private Declare Function Zcrc32 Lib "zlib.dll" Alias "crc32" (ByVal crc As Long, ByRef buf As Any, ByVal Length As Long) As Long

' following are the actual PNG image properties, exposed via class properties
' (**)Not all are translated until called from the appropriate class property
Private m_Width As Long                 ' image width
Private m_Height As Long                ' image height
Private m_BitDepth As Byte              ' image bit depth/count: 1,2,4,8,16
Private m_ColorType As Byte             ' image color type: 0,2,3,4,6
Private m_Interlacing As Byte           ' interlaced: 0,1
Private m_Palette() As Byte             ' image palette information
Private m_TransSimple() As Byte         ' image simple transparency information
Private m_TransColor As Long  ' translated simple transparency color (BGR or index value)

' matrix/lookup tables
Private pow2x8() As Long                ' a look up table for bit shifting (1,2,4 bit pixels)
' arrays to help process interlacing more quickly & are in relation to the 8x8 matrix
Private MatrixRow(1 To 8) As Byte       ' row where each pass starts within interlace matrix
Private MatrixCol(1 To 8) As Byte       ' column where each pass starts within interlace matrix
Private MatrixRowAdd(1 To 8) As Byte    ' gaps between each row withiin each pass
Private MatrixColAdd(1 To 8) As Byte    ' gaps between each column within each pass
' following are used for progressive/pixellated display only
Private MatrixCY(1 To 8) As Byte        ' height of each pixel in a scanline
Private MatrixCX(1 To 8) As Byte        ' width of each pixel in a scanline

Private Enum eColorTypes    ' internal use only
    clrGrayScale = 0
    clrTrueColor = 2
    clrPalette = 3
    clrGrayAlpha = 4
    clrTrueAlpha = 6
End Enum

' PNG chunk names & their numerical equivalent (those used in this class)
' Per png specs; using alpha chars is a no-no should system not support those characters
' http://www.libpng.org/pub/png/spec/1.1/PNG-Chunks.html
Private Const chnk_IHDR As Long = &H52444849 'Image header
Private Const chnk_IDAT As Long = &H54414449 'Image data
Private Const chnk_IEND As Long = &H444E4549 'End of Image
Private Const chnk_PLTE As Long = &H45544C50 'Palette
Private Const chnk_tRNS As Long = &H534E5274 'Simple Transparency

Private Const png_Signature1 As Long = 1196314761
Private Const png_Signature2 As Long = 169478669
'^^ Complete signature is 8 bytes: 137 80 78 71 13 10 26 10

Private pngStream() As Byte ' overlay of bytes when loading via array, else individual chunk bytes
Private cHost As c32bppDIB  ' owner of 32bpp destination image

Public Function LoadStream(Stream() As Byte, dibClass As c32bppDIB, _
            Optional ByVal streamOffset As Long = 0, _
            Optional ByVal streamLength As Long = 0) As Boolean

    ' PURPOSE: Determine if passed array is a PNG & if it is, then convert it to
    ' a 32bpp owned by dibClass
    
    ' Parameters.
    ' Stream() :: a byte array containing the possible PNG image
    ' dibClass :: an initialized c32bppDIB class
    ' streamOffset :: array position for 1st byte in the stream
    ' streamLength :: size of stream that contains the image
    '   - If zero, then size is UBound(inStream)-streamOffset+1

    ' IMPORTANT: the array offset & length are not checked in this class.
    '   They were checked before this class was called. If this class is to
    '   be pulled out and put in another project, ensure you include the
    '   validation shown in c32bppDIB.LoadPicture_Stream
    
    Dim tSA As SafeArray1D
    With tSA    ' prepare to overlay. Overlay prevents VB copying bytes into another array for processing
        .cbElements = 1     ' as byte array
        .cDims = 1          ' 1 dimensional
        .pvData = VarPtr(Stream(streamOffset))
        .rgSABound.cElements = streamLength
    End With
    CopyMemory ByVal VarPtrArray(pngStream), VarPtr(tSA), 4& ' establish overlay

    Set cHost = dibClass
    LoadStream = LoadPNG(vbNullString, streamLength)
    CopyMemory ByVal VarPtrArray(pngStream), 0&, 4& ' remove overlay
    Set cHost = Nothing
    
End Function

Public Function LoadFile(ByVal FileName As String, dibClass As c32bppDIB) As Boolean

    ' PURPOSE: Determine if passed file is a PNG & if it is, then convert it to
    ' a 32bpp owned by dibClass
    
    ' Parameters.
    ' FileName :: full path and file
    ' dibClass :: an initialized c32bppDIB class

    ' IMPORTANT: the existance and validity of the filename is not checked here.
    '   They were checked before this class was called. If this class is to
    '   be pulled out and put in another project, ensure you include the
    '   validation shown in c32bppDIB.LoadPicture_File

    Set cHost = dibClass
    LoadFile = LoadPNG(FileName, 0&)
    Set cHost = Nothing

End Function

Private Function LoadPNG(FileName As String, streamLength As Long) As Boolean

    ' PURPOSE: Determine if passed file is a PNG & if it is, then convert it to
    ' a 32bpp owned by dibClass
    
    ' Parameters.
    ' FileName :: full path and file
    ' dibClass :: an initialized c32bppDIB class

    Dim ptrLoc As Long          ' used to ensure parsing doesn't go past EOF of corrupted file
    Dim ptrArray As Long
    Dim FileNumber As Long      ' the file handle
    Dim gpLong As Long          ' general purpose long value
    Dim lenIDAT As Long         ' running total of the png data size (compressed)
    
    Dim ChunkName As Long       ' name of the chunk
    Dim ChunkLen As Long        ' length of the chunk
    
    Dim RawPNGdata() As Byte    ' uncompressed png data
    Dim IDATdata() As Byte      ' compressed png data
    
    Dim uncmprssSize As Long    ' calculated size of uncompressed PNG data
    Dim lError As Long
    
    Dim bCRCchecks As Boolean   ' whether or not to use CRC checks on chunks
    Dim crc32value As Long      ' if CRC checks applied, the the CRC value
    
    ' reset class' only key property
    m_TransColor = -1
    
    ' attempt to open the file with read access
    If FileName = vbNullString Then
    
        ptrLoc = 7              ' counter to prevent overflow of array
        ptrArray = 8            ' current position in passed array
        If IsPNG() = False Then
            Exit Function
        Else
            LoadPNG = True ' & process it using GDI+ if available
            If GDIplusLoad(FileName) = True Then Exit Function
        End If
        
    Else
        On Error Resume Next
        FileNumber = FreeFile
        Open FileName For Binary Access Read As FileNumber
        If Err Then
            Close #FileNumber
            Err.Clear
            Exit Function
        End If
        
        ' validate we are looking at a png file
        streamLength = LOF(FileNumber) + 1
        If streamLength > 56 Then ' minimal (signature=8;header=13,3 rqd chunks=36 min)
            ReDim pngStream(0 To 57)
            Get FileNumber, 1, pngStream()
            If IsPNG() = True Then
                LoadPNG = True
            Else
                Close #FileNumber
                Exit Function
            End If
        End If
        On Error GoTo 0
        ' process using GDI+ if available
        If GDIplusLoad(FileName) = True Then
            Close #FileNumber
            LoadPNG = True
            Exit Function
        End If
        ptrArray = -4
        ptrLoc = 9                  ' next position in the file
        Seek FileNumber, ptrLoc     ' move to that position
   End If
    
    ReDim IDATdata(0 To streamLength \ 2)   ' array to hold compressed data; start with arbritrary length
    bCRCchecks = ValidateDLL("zlib.dll", "crc32") ' determine if CRC checks will be made

    Do ' read & pre-process the png file
    
        ' Chunks consist of 4 bytes for the length of the chunk
        '                 + n bytes for the chunk
        '                 + 4 bytes for a CRC value
        If FileName = vbNullString Then
            CopyMemory gpLong, pngStream(ptrArray), &H4 ' length of the current chunk
        Else
            Get FileNumber, , gpLong ' number of bytes for the chunk
        End If
        ChunkLen = iparseReverseLong(gpLong) ' longs are big endian, need little endian for Windows
        
        ' track position of pointer in the file
        ptrLoc = ptrLoc + ChunkLen + 12 ' 12 = 4byte name + 4byte CRC + 4byte chunk count
        If ptrLoc > streamLength Then
            ' corrupted file; abort
            lError = 1
            Exit Do
        End If
        
        ' read chunk name & chunk data, read CRC separately
        If FileName = vbNullString Then
            If bCRCchecks = True Then CopyMemory crc32value, pngStream(ptrArray + ChunkLen + 8), &H4
        Else
            ReDim pngStream(0 To ChunkLen + 3)
            Get FileNumber, , pngStream
            Get FileNumber, , gpLong             ' read the CRC value (big endian)
            If bCRCchecks = True Then crc32value = gpLong
        End If
        CopyMemory ChunkName, pngStream(ptrArray + 4), &H4 ' extract the chunk name
        
        If ChunkLen < 1 Then
            ' should never be <0; however can be zero at anytime
            If ChunkName = chnk_IEND Then Exit Do
        Else
            
            ' each of the chunk parsing routines will be in a separate function.
            ' This is so that they can easily be modified without affecting any
            ' of the other code. Additionally, it is possible that chunk types will
            ' increase as PNG continues to evolve. Except IDAT, all chunks
            ' are in their own routines; even sRGB which is only a 1 byte chunk.
            Select Case ChunkName
            
            Case chnk_IDAT ' UCase chunk names are critical - CRC check
                ' compressed, filtered image data
                On Error Resume Next
                ' error? what error? all precautions taken in zChunk_IHDR function
                ' However; no predicting "Out of Memory" errors
                If Not crc32value = 0 Then lError = Not (iparseReverseLong(crc32value) = Zcrc32(0, pngStream(ptrArray + 4), ChunkLen + 4))
                If Not lError = 0 Then Exit Do
                
                gpLong = lenIDAT + ChunkLen ' size of array needed
                If gpLong > UBound(IDATdata) Then ' test length & increment if needed
                    ReDim Preserve IDATdata(0 To gpLong + streamLength \ 4)
                End If
                CopyMemory IDATdata(lenIDAT), pngStream(ptrArray + 8), ChunkLen ' & append the new data
                lenIDAT = gpLong  ' cache number of compressed bytes so far
                If Err Then
                    lError = 1
                    Exit Do
                End If
                On Error GoTo 0
            
            Case chnk_PLTE ' UCase chunk names are critical - CRC check
                lError = zChunk_PLTE(ChunkLen, ptrArray + 4, crc32value)
                If Not lError = 0 Then Exit Do
            
            Case chnk_tRNS ' simple transparency option
                ' CRC checked 'cause if invalid, we could generate an out of bounds
                ' error in one of the other routines that reference this array
                lError = zChunk_tRNS(ChunkLen, ptrArray + 4, crc32value)
                If Not lError = 0 Then Exit Do
                
            Case chnk_IHDR ' UCase chunk names are critical - CRC check
                ' Note: the zChunk_IHDR routine also calculates uncompressed size
                lError = zChunk_IHDR(ChunkLen, ptrArray + 4, uncmprssSize, crc32value)
                If Not lError = 0 Then Exit Do
                
            Case chnk_IEND ' UCase chunk names are critical - CRC check
                ' should CRC check for corrupted file; but why? we're at end of image
                Exit Do
                
            End Select
    
        End If
        If FileName = vbNullString Then ptrArray = ptrArray + ChunkLen + 12 ' move to next position in the array
    Loop

ExitRoutine:
    ' clean up
    If Not FileName = vbNullString Then
        Close #FileNumber
        Erase pngStream()
    End If
    
    If lenIDAT = 0 Or lError <> 0 Then  ' invalid png image
        If Err Then Err.Clear
        
    Else
        ' process the compressed data
        Call PostLoadPNG(IDATdata(), lenIDAT, uncmprssSize)
    End If

End Function

Private Function PostLoadPNG(IDATdata() As Byte, lenIDAT As Long, uncmprssSize As Long) As Boolean

    ' Purpose: Uncompress compressed bytes and send to the un-filtering routines
    Dim RawPNGdata() As Byte
    Dim bUncompressed As Boolean
    Dim lRtn As Long

    On Error Resume Next
    ' we need to uncompress our PNG file
    ReDim RawPNGdata(0 To uncmprssSize - 1)
    
    ' if zLIB is available, let it uncompress; faster than pure VB
    If ValidateDLL("zlib.dll", "uncompress") = True Then
        bUncompressed = (Zuncompress(RawPNGdata(0), uncmprssSize, IDATdata(0), lenIDAT) = 0)
    End If
    
    If Not bUncompressed Then
        ' either zLib returned an error or it wasn't available, uncompress by hand
        bUncompressed = vbDecompress(RawPNGdata(), IDATdata(), uncmprssSize)
        If Err Then Err.Clear
    End If
    Erase IDATdata()
    
    If Not bUncompressed Then
        ' failed to uncompress & shouldn't happen 'cause if I calculated uncmprssSize
        ' wrong, then other calculations in this routine are wrong too
        ' See: CalcUncompressedWidth
        Exit Function
    End If

    Call InitializePalette  ' if PNG is palettized, create palette
    cHost.InitializeDIB m_Width, m_Height    ' create 32bpp DIB to hold PNG
    
    ' call function to begin converting PNG to Bitmap
    If m_Interlacing = 0 Then
        lRtn = UnfilterNI(RawPNGdata()) ' non-interlaced image
    Else
        lRtn = UnfilterInterlaced(RawPNGdata()) ' interlaced image
    End If

    ' return results
    If lRtn = 0 Then
        cHost.DestroyDIB ' failure decoding the PNG
    Else
        If m_ColorType > 0 And (m_ColorType Mod 2) = 0 Then
            cHost.Alpha = True
        ElseIf m_TransColor <> -1 Then
            cHost.Alpha = True
        ElseIf Not iparseIsArrayEmpty(Not m_TransSimple) Then
            cHost.Alpha = True
        End If
        cHost.ImageType = imgPNG
        PostLoadPNG = True
    End If

End Function

Private Function CalcUncompressedWidth() As Long

    Dim uncompressedWidth As Long, iBitPP As Byte
    Dim Pass As Long, passWidth As Long, passHeight As Long

    On Error GoTo NoLoad

    InitializeMatrix    ' build the interlacing matrix; also used for non-interlaced too

    ' get the actual bits per pixel the png is using
    ' (i.e., 16bitdepth png @ ColorType 6 = 64bits per pixel)
    GetDepthInfo 0, 0, iBitPP, 0
    
    If m_Interlacing = 0 Then ' no interlacing
        ' uncompressed width will be byte aligned width + 1 for filter byte
        ' multiplied by the height
        passWidth = GetBytesPerPixel(m_Width, iBitPP)
        uncompressedWidth = passWidth * m_Height + m_Height
    Else
        ' interlaced will also be byte aligned but per scanline width
        ' Each of the 7 passes can have different widths + 1 filter byte per line
        For Pass = 1 To 7
            ' calculate number of pixels per scan line
            passWidth = m_Width \ MatrixColAdd(Pass) - (m_Width Mod MatrixColAdd(Pass) > MatrixCol(Pass))
            ' determine number of bytes needed for each scanline
            passWidth = GetBytesPerPixel(passWidth, iBitPP)
            ' calculate number of rows for this scan's pass
            passHeight = m_Height \ MatrixRowAdd(Pass) - (m_Height Mod MatrixRowAdd(Pass) > MatrixRow(Pass))
            ' now get the total bytes needed for the entire pass,
            ' adding 1 filter byte for each line in the pass:  i.e., + passHeight
            uncompressedWidth = uncompressedWidth + passWidth * passHeight + passHeight
        Next
        
    End If
    
    CalcUncompressedWidth = uncompressedWidth

NoLoad:
End Function

Private Sub InitializeMatrix()
        
    ' a quick look up table for bit shifting operations
    If m_ColorType = clrGrayScale Or m_ColorType = clrPalette Then
        Dim i As Integer
        ReDim pow2x8(0 To 8)
        pow2x8(0) = 1
        For i = 1 To 8
            pow2x8(i) = pow2x8(i - 1) * 2
        Next
    End If
    
    ' initialize interlacing matrix, used in the ConvertPNGtoBMP routine and
    ' also used to calculate the uncompressed size of the compressed PNG data
    
    ' Non-interlaced images are considered Pass#8, where interlaced images always
    ' contain 7 passes (1 thru 7).
    
    ' determines what row in the interlaced image, the current pass begins at
    CopyMemory MatrixRow(1), 262144, &H4
    CopyMemory MatrixRow(5), 65538, &H4 'Array(0, 0, 4, 0, 2, 0, 1, 0)
    ' determines what column in the interlaced image, the current pass begins at
    CopyMemory MatrixCol(1), 33555456, &H4
    CopyMemory MatrixCol(5), 256&, &H4 'Array(0, 4, 0, 2, 0, 1, 0, 0)
    ' determines the row interval of the current pass
    CopyMemory MatrixRowAdd(1), 67635208, &H4
    CopyMemory MatrixRowAdd(5), 16908804, &H4 'Array(8, 8, 8, 4, 4, 2, 2, 1)
    ' determines the column interval of the current pass
    CopyMemory MatrixColAdd(1), 67373064, &H4
    CopyMemory MatrixColAdd(5), 16843266, &H4 'Array(8, 8, 4, 4, 2, 2, 1, 1)
    
    ' 1st 7 elements of next 2 arrays used for pixellated interlaced images
    ' determines the width of each pixellated pixel for the current pass
    CopyMemory MatrixCX(1), 33817608, &H4
    CopyMemory MatrixCX(5), 16843010, &H4 'Array(8, 4, 4, 2, 2, 1, 1, 1)
    ' determines the height of each pixellated pixel for the current pass
    CopyMemory MatrixCY(1), MatrixColAdd(1), &H8 'Array(8, 8, 4, 4, 2, 2, 1, 1)

End Sub


Private Function ConvertPNGtoBMP(rawBytes() As Byte, ByVal scanPass As Byte, _
                    ByVal scanCY As Long, ByVal rBPRow As Long, _
                    Optional ByVal startOffset As Long = 0) As Boolean

' // LaVolpe, Dec 1 thru 10 - Added from Scratch
' 19 Dec -- rewrote to simply. More & more options made the routine too disorganized
'       added the bNeedAlpha, bNeedRGB, bPackByte flags to categorize needed actions

' This routine (quite big due to various conversion issues) will convert
' PNG bits/bytes to BMP bits/bytes. Common issues to deal with include the
' byte-aligned PNG data may not be the same width as DWord-aligned BMP data. Also,
' BMPs do not support 48bpp nor 2bpp bitmaps. And the biggie:
' PNGs are RGB while DIBs are BGR; otherwise, displaying PNGs would be super quick.
                    
' rawBytes() are the png scanline bytes
' scanPass() is a valid number btwn 1-8 (8 indicates non-interlaced)
' scanCY is number of scanlines (btwn 1 and image height)
' rBPRow is the number of bytes per scanline of the rayBytes array (byte aligned)

    Dim rRow As Long, rColumn As Long ' current row/column of the png image
    Dim dRow As Long, dColumn As Long ' current row/column of destination bitmap
    Dim dIndex As Long ' position of dRow in the destBytes() array
    
    Dim dBitPP As Byte, dBytePP As Byte ' nr of bits/bytes per pixel of bitmap
    Dim rBytePP As Byte, rBitPP As Byte ' nr of bytes per pixel in png image
    
    Dim tColor As Long  ' color value when copying 3 or 4 bytes to a 3 or 4 byte array
    Dim palOffset As Long
    Dim alphaModifyer As Long
    Dim pIndex As Byte ' alpha related variables
    
    Dim destBytes() As Byte ' placeholders for DIB bits (DMA)
    Dim tSA As SafeArray1D  ' array overlays for DIB bits (DMA)
    Dim dRowWidth As Long   ' DWord aligned width of bitmap scanline/row
    
    Dim bPltBlockCopy As Boolean
    Dim maskShift As Long, maskInit As Long

    On Error GoTo err_h
    
    ' use direct memory access (DMA) to reference the DIB pixel data
    With tSA
        .cDims = 1                          ' Number of dimensions
        .cbElements = 1                     ' Size of data elements
        .pvData = cHost.BitsPointer         ' Data address
        .rgSABound.lLbound = 0              ' Low bound of array
        .rgSABound.cElements = m_Width * m_Height * 4 ' Nr of Elements
    End With
    CopyMemory ByVal VarPtrArray(destBytes), VarPtr(tSA), 4

    ' determine the bits/bytes of the png and bitmap images
    GetDepthInfo dBitPP, dBytePP, rBitPP, rBytePP

    ' get location of BMP scanline we are processing from PNG scanline
    If startOffset = 0 Then
        dRow = MatrixRow(scanPass)
    Else
        dRow = startOffset
    End If
    dRowWidth = (UBound(destBytes) + 1) \ m_Height

    For rRow = startOffset To scanCY - 1 ' < in relation to rawBytes() array
        
        dIndex = (m_Height - dRow - 1) * dRowWidth ' < destBytes array pointer for 1st pixel in scanline
        rColumn = rRow * rBPRow + 1  ' < rawBytes array pointer for 1st pixel in scanline
        
        dColumn = MatrixCol(scanPass) ' column in relation to the destBytes() array
        maskShift = 8 - m_BitDepth
        
        Do While dColumn < m_Width
        
            ' we need the alpha value for transfer to 32 bit BMP
            ' but for use with AlphaBlend, need to pre-multiply rgb elements
            Select Case m_ColorType
            
            Case clrTrueAlpha ' true color with alpha
                If m_BitDepth = 16 Then
                    alphaModifyer = rawBytes(rColumn + 6)
                    If alphaModifyer = &HFF Then
                        tColor = rawBytes(rColumn + 4) Or rawBytes(rColumn + 2) * &H100& Or rawBytes(rColumn) * &H10000
                    Else
                        tColor = ((alphaModifyer * rawBytes(rColumn + 4)) \ &HFF) Or ((alphaModifyer * rawBytes(rColumn + 2)) \ &HFF) * &H100& Or ((alphaModifyer * rawBytes(rColumn)) \ &HFF) * &H10000
                    End If
                Else
                    alphaModifyer = rawBytes(rColumn + 3)
                    If alphaModifyer = &HFF Then
                        tColor = rawBytes(rColumn + 2) Or rawBytes(rColumn + 1) * &H100& Or rawBytes(rColumn) * &H10000
                    ElseIf Not alphaModifyer = 0 Then
                        tColor = ((alphaModifyer * rawBytes(rColumn + 2)) \ &HFF) Or ((alphaModifyer * rawBytes(rColumn + 1)) \ &HFF) * &H100& Or ((alphaModifyer * rawBytes(rColumn)) \ &HFF) * &H10000
                    End If
                End If
                
            Case clrGrayAlpha  ' grayscale with alpha
                If Not m_BitDepth = 16 Then alphaModifyer = rawBytes(rColumn + 1) Else alphaModifyer = rawBytes(rColumn + 2)
                If alphaModifyer = &HFF Then
                    tColor = rawBytes(rColumn) Or (rawBytes(rColumn) * &H100&) Or (rawBytes(rColumn) * &H10000)
                ElseIf Not alphaModifyer = 0 Then
                    tColor = ((alphaModifyer * rawBytes(rColumn)) \ &HFF)
                    tColor = tColor Or (tColor * &H100&) Or (tColor * &H10000)
                End If
                
            Case clrPalette ' paletted + simple transparency in its own palette-alpha table
                tColor = GetPaletteValue(maskShift, rawBytes(rColumn), pIndex)
                If m_TransColor = -1 Then
                    alphaModifyer = &HFF
                Else ' no transparency in the png
                    If Not m_TransSimple(pIndex) = 0 Then
                        alphaModifyer = m_TransSimple(pIndex)
                        palOffset = pIndex * 3
                        tColor = ((alphaModifyer * m_Palette(palOffset + 2)) \ &HFF) Or ((alphaModifyer * m_Palette(palOffset + 1)) \ &HFF) * &H100& Or ((alphaModifyer * m_Palette(palOffset)) \ &HFF) * &H10000
                    End If
                End If
                
            Case clrTrueColor ' true color + simple transparency
                If m_BitDepth > 8 Then ' 24,32,48bpp image
                    tColor = rawBytes(rColumn + 4) Or rawBytes(rColumn + 2) * &H100& Or rawBytes(rColumn) * &H10000
                Else
                    tColor = rawBytes(rColumn + 2) Or rawBytes(rColumn + 1) * &H100& Or rawBytes(rColumn) * &H10000
                End If
                If m_TransColor = -1 Then   ' no transparency used
                    alphaModifyer = &HFF
                Else
                    If Not tColor = m_TransColor Then ' m_Transcolor is BGR vs RGB
                        alphaModifyer = &HFF
                        If m_BitDepth = 16 Then
                            If (tColor = (m_TransSimple(4) Or (m_TransSimple(2) * &H100&) Or (m_TransSimple(0) * &H10000))) Then alphaModifyer = 0
                        End If
                    End If
                End If
                
            Case clrGrayScale ' grayscale with simple transparency
                tColor = GetPaletteValue(maskShift, rawBytes(rColumn), pIndex)
                If Not m_TransColor = pIndex Then
                    alphaModifyer = &HFF
                    If m_BitDepth = 16 Then alphaModifyer = -alphaModifyer * (rawBytes(rColumn) <> m_TransSimple(0))
                End If
            End Select
            
            If Not alphaModifyer = 0 Then ' only modify DIB bits if needed
                destBytes(dIndex + dColumn * 4 + 3) = alphaModifyer
                CopyMemory destBytes(dIndex + dColumn * 4), tColor, &H3
                alphaModifyer = 0
            End If
            
            ' ensure our source byte pointer is moved along appropriately
            If m_BitDepth < 8 Then
                If maskShift = 0 Then
                    rColumn = rColumn + 1
                    maskShift = 8 - m_BitDepth
                Else
                    maskShift = maskShift - m_BitDepth
                End If
            Else
                rColumn = rColumn + rBytePP      ' else increment per source byte pp
            End If
            dColumn = dColumn + MatrixColAdd(scanPass)
        
        Loop
        dRow = dRow + MatrixRowAdd(scanPass)
        
    Next

    ' clean up & return result
    CopyMemory ByVal VarPtrArray(destBytes), 0&, 4
    ConvertPNGtoBMP = True

    Exit Function

err_h:  ' should never get here
Err.Clear
If tSA.cDims Then CopyMemory ByVal VarPtrArray(destBytes), 0&, 4

End Function


Private Function GetPaletteValue(ByVal PixelPos As Long, ByVal PixelValue As Byte, _
                Optional ByRef RtnIndex As Byte) As Long

    ' // LaVolpe, Dec 1 thru 10 - added from scratch
    ' Returns a palette index and palette color from a compressed byte
    Dim arrayLoc As Long
    RtnIndex = (PixelValue \ pow2x8(PixelPos)) And (pow2x8(m_BitDepth) - 1)
    arrayLoc = RtnIndex * 3
    GetPaletteValue = m_Palette(arrayLoc + 2) Or (m_Palette(arrayLoc + 1) * &H100&) Or (m_Palette(arrayLoc) * &H10000)
    
End Function

Private Function GetBytesPerPixel(totalWidth As Long, btsPerPixel As Byte) As Long
    ' // LaVolpe, Dec 1 thru 10
    ' returns number of bytes required to display n pixels at p color depth (byte aligned)
    GetBytesPerPixel = (totalWidth * btsPerPixel + 7) \ 8

End Function

Private Sub GetDepthInfo(destBitPP As Byte, destBytePP As Byte, _
                        rawBitsPP As Byte, rawBytesPP As Byte)
    
' returns the bits per pixel & bytes per pixel for the destination bitmap
' and also the respective values for the png image
    
' PNG > DIB bmp (per pixel) conversion chart I use throughout the routines:
'Color Type     bit depth   PNG bits/bytes per pixel  BMP bits/bytes pp (ignore alpha)
'----------     ---------   ------------------------- --------------------------------
'0 gray scale   1                   1   1                   1   1   (?  ?)
'               2                   2   1                   4   1   (?  ?)
'               4                   4   1                   4   1   (?  ?)
'               8                   8   1                   8   1   (8  1)
'               16                  16  2                   8   1   (8  1)

'2 true color   8                   24  3                   24  3   (24 3)
'               16                  48  6                   24  3   (24 3)

'3 palette      1                   1   1                   1   1   (?  ?)
'               2                   2   1                   4   1   (?  ?)
'               4                   4   1                   4   1   (?  ?)
'               8                   8   1                   8   1   (8  1)
'               16                  16  2                   8   1   (8  1)

'4 gray+alpha   8                   16  2                   32  4   (24 3)
'               16                  32  4                   32  4   (24 3)

'6 true+alpha   8                   32  4                   32  4   (24 3)
'               16                  64  8                   32  4   (24 3)

'any bit depth that uses simple transparency (trns chunk)   32  4   (n/a)
'--------------------------------------------------------------------------

    
    Select Case m_ColorType
    
    Case clrTrueAlpha ' true color w/alpha (only 8,16 bpp pngs)
        rawBytesPP = 4 * (m_BitDepth \ 8): rawBitsPP = m_BitDepth * 4
        
    Case clrGrayAlpha: ' grayscale w/alpha (only 8,16 bpp pngs)
        rawBytesPP = 2 * (m_BitDepth \ 8): rawBitsPP = m_BitDepth * 2
        
    Case clrTrueColor: ' true color (rgb triples) (8,16 bpp pngs)
        rawBytesPP = 3 * (m_BitDepth \ 8): rawBitsPP = m_BitDepth * 3
        
    Case clrGrayScale ' grayscale images (all bit depths)
        If m_BitDepth = 2 Then ' special case as MS bitmaps don't do 2bpp
            rawBytesPP = 1: rawBitsPP = 2
        ElseIf m_BitDepth > 4 Then ' (8,16 bpp pngs)
            rawBytesPP = m_BitDepth \ 8: rawBitsPP = m_BitDepth
        Else ' (1,4 bpp pngs)
            rawBytesPP = 1: rawBitsPP = m_BitDepth
        End If
        
    Case clrPalette: ' palette entries (1,2,4,8 bpp pngs)
        rawBytesPP = 1: rawBitsPP = m_BitDepth
        
    End Select
    
    ' our DIB will always be 32bpp
    destBitPP = 32: destBytePP = 4

End Sub

Private Function PaethPredictor(ByVal Left As Integer, ByVal Above As Integer, ByVal UpperLeft As Integer) As Integer

    ' // LaVolpe, Dec 1 thru 10 - rewrote for understanding & commented/linked
    
    ' http://www.w3.org/TR/PNG/#9-table91
    ' algorithm is used for both encoding & decoding the png image's filter
    ' based off of the formula created by Alan W. Paeth & provided fully in url above

    Dim pa As Integer, pb As Integer, pC As Integer, p As Integer
    p = (0 + Left + Above - UpperLeft)
    pa = Abs(p - Left)
    pb = Abs(p - Above)
    pC = Abs(p - UpperLeft)
    
    ' tie breaker
    ' The order in which the comparisons are performed is critical and shall not be altered
    If (pa <= pb) And (pa <= pC) Then
        PaethPredictor = Left
    ElseIf pb <= pC Then
        PaethPredictor = Above
    Else
        PaethPredictor = UpperLeft
    End If

End Function

Private Sub DecodeFilter_Avg(Filtered() As Byte, ByVal RowNr As Long, ByVal ScanLine As Long, ByVal stepVal As Integer)

    ' // LaVolpe, Dec 1 thru 10 - rewrote for understanding; nothing wrong with original
    'http://www.w3.org/TR/PNG/#9-table91
    'Filters may use the original values of the following bytes to generate the new byte value:
    '
    'x  the byte being filtered;
    'a  the byte corresponding to x in the pixel immediately before the pixel containing x (or the byte immediately before x, when the bit depth is less than 8);
    'b  the byte corresponding to x in the previous scanline;
    'c  the byte corresponding to b in the pixel immediately before the pixel containing b (or the byte immediately before b, when the bit depth is less than 8).
    
    ' algorithm: Recon(x) = Filt(x) + floor((Recon(a) + Recon(b)) / 2)
    ' Unsigned arithmetic modulo 256 is used, so that both the inputs and outputs fit into bytes

    Dim X As Long, startByte As Long
    
    startByte = RowNr * ScanLine + 1
    
    On Error GoTo EH
    ' break out for faster loops, removing IF statements/combinations
    If RowNr = 0 Then   ' 1st row; there will be no Top row to get data from
        ' if png is encoded properly, shouldn't get here
        ' now process the 2nd pixel on, to finish the scanline
        For X = startByte + stepVal To startByte + ScanLine - 2
            Filtered(X) = (0 + Filtered(X) + (Filtered(X - stepVal) \ 2)) Mod 256
        Next
        
    Else    ' 2nd or subsequent rows
        ' process the 1st n bytes (1st pixel only)
        For X = startByte To startByte + stepVal - 1
            Filtered(X) = (0 + Filtered(X) + (Filtered(X - ScanLine) \ 2)) Mod 256
        Next
        ' now process the 2nd pixel on, to finish the scanline
        For X = X To startByte + ScanLine - 2
            Filtered(X) = (0 + Filtered(X) + (0 + Filtered(X - stepVal) + Filtered(X - ScanLine)) \ 2) Mod 256
        Next
    End If
EH:
End Sub

Private Sub DecodeFilter_Paeth(Filtered() As Byte, ByVal RowNr As Long, ByVal ScanLine As Long, ByVal stepVal As Integer)

    ' // LaVolpe, Dec 1 thru 10 - rewrote for understanding; nothing wrong with original
    'http://www.w3.org/TR/PNG/#9-table91
    
    'Filters may use the original values of the following bytes to generate the new byte value:
    '
    'x  the byte being filtered;
    'a  the byte corresponding to x in the pixel immediately before the pixel containing x (or the byte immediately before x, when the bit depth is less than 8);
    'b  the byte corresponding to x in the previous scanline;
    'c  the byte corresponding to b in the pixel immediately before the pixel containing b (or the byte immediately before b, when the bit depth is less than 8).
    
    ' algorithm: Recon(x) = Filt(x) + PaethPredictor(Recon(a), Recon(b), Recon(c))
    ' Unsigned arithmetic modulo 256 is used, so that both the inputs and outputs fit into bytes

    Dim X As Long, startByte As Long

    startByte = RowNr * ScanLine + 1
    
    ' break out for faster loops, removing IF statements/combinations
    On Error GoTo EH
    
    If RowNr = 0 Then    ' 1st row; there will be no Top row to get data from
        ' if png is encoded properly, shouldn't get here
        ' now process the 2nd pixel on, to finish the scanline
        For X = startByte + stepVal To startByte + ScanLine - 2
            Filtered(X) = (0 + Filtered(X) + Filtered(X - stepVal)) Mod 256
        Next
    
    Else    ' 2nd or subsequent rows
        ' process the 1st n bytes (1st pixel only)
        For X = startByte To startByte + stepVal - 1
            Filtered(X) = (0 + Filtered(X) + Filtered(X - ScanLine)) Mod 256
        Next
        ' now process the 2nd pixel on, to finish the scanline
        For X = X To startByte + ScanLine - 2
            Filtered(X) = (0 + Filtered(X) + PaethPredictor(Filtered(X - stepVal), Filtered(X - ScanLine), Filtered(X - ScanLine - stepVal))) Mod 256
        Next

    End If
EH:
End Sub

Private Sub DecodeFilter_Sub(Filtered() As Byte, ByVal RowNr As Long, ByVal ScanLine As Long, ByVal stepVal As Integer)

    ' // LaVolpe, Dec 1 thru 10 - rewrote for understanding; nothing wrong with original
    'http://www.w3.org/TR/PNG/#9-table91
    
    'Filters may use the original values of the following bytes to generate the new byte value:
    '
    'x  the byte being filtered;
    'a  the byte corresponding to x in the pixel immediately before the pixel containing x (or the byte immediately before x, when the bit depth is less than 8);
    'b  the byte corresponding to x in the previous scanline;
    'c  the byte corresponding to b in the pixel immediately before the pixel containing b (or the byte immediately before b, when the bit depth is less than 8).
    
    ' algorithm: Recon(x) = Filt(x) + Recon(a)
    ' Unsigned arithmetic modulo 256 is used, so that both the inputs and outputs fit into bytes

    Dim startByte As Long
    Dim n As Long, X As Long
    
    startByte = RowNr * ScanLine + 1
    
    On Error GoTo EH
    ' 1st n bytes for 1st pixel are unfiltered
    For n = startByte + stepVal To startByte + ScanLine - 2 Step stepVal
        For X = n To n + stepVal - 1
            Filtered(X) = (0 + Filtered(X) + Filtered(X - stepVal)) Mod 256
        Next
    Next
EH:
End Sub

Private Sub DecodeFilter_Up(Filtered() As Byte, ByVal RowNr As Long, ByVal ScanLine As Long, ByVal stepVal As Integer)

    ' // LaVolpe, Dec 1 thru 10 - rewrote for understanding; nothing wrong with original
    'http://www.w3.org/TR/PNG/#9-table91
    
    'Filters may use the original values of the following bytes to generate the new byte value:
    '
    'x  the byte being filtered;
    'a  the byte corresponding to x in the pixel immediately before the pixel containing x (or the byte immediately before x, when the bit depth is less than 8);
    'b  the byte corresponding to x in the previous scanline;
    'c  the byte corresponding to b in the pixel immediately before the pixel containing b (or the byte immediately before b, when the bit depth is less than 8).
    
    ' algorithm:  Recon(x) = Filt(x) + Recon(b)
    ' Unsigned arithmetic modulo 256 is used, so that both the inputs and outputs fit into bytes
    
    Dim startByte As Long, X As Long
    
    On Error GoTo EH
    If Not RowNr = 0 Then    ' 1st row; there will be no Top row to get data from
        startByte = RowNr * ScanLine + 1
        For X = startByte To startByte + ScanLine - 2
            Filtered(X) = (0 + Filtered(X) + Filtered(X - ScanLine)) Mod 256
        Next
    End If
EH:
End Sub

Private Function UnfilterInterlaced(Filtered() As Byte) As Boolean

    ' // LaVolpe, Dec 1 thru 10 - built from scratch
    ' http://www.libpng.org/pub/png/spec/1.2/PNG-DataRep.html#DR.Interlaced-data-order
    
    ' Progressive display/scan order per 8 pixel blocks (64 total pixels)
    '   1 6 4 6 2 6 4 6     ' 1st scan: 1 pixel (@col 0), row 0 [1/64 of total image]
    '   7 7 7 7 7 7 7 7     ' 2nd scan: 1 pixel (@col 4), row 0 [1/32 of image shown]
    '   5 6 5 6 5 6 5 6     ' 3rd scan: 2 pixels (@cols 0:4), row 4 [1/16 of image]
    '   7 7 7 7 7 7 7 7     ' 4th scan: 4 pixels (@cols 2:6), rows 0:4 [1/8]
    '   3 6 4 6 3 6 4 6     ' 5th scan: 8 pixels (@cols 0:2:4:6), rows 2:6 [1/4]
    '   7 7 7 7 7 7 7 7     ' 6th scan: 16 pixels (@cols 1:3:5:7), rows 0:2:4:6 [1/2]
    '   5 6 5 6 5 6 5 6     ' 7th scan: 32 pixels (@cols all), rows 1:3:5:7 [100%]
    '   7 7 7 7 7 7 7 7                 64 pixels, 15 scanlines over 7 passes
    
    ' Note : all logic in this routine is based off of the above grid.
    ' Scanline widths are only guaranteed to be same for each scanline in the same pass.
    ' Scanlines can be padded both horizontally & vertically if the image doesn't fit into
    '   a nice 8x8 grid evenly.
    ' Each scanline in interlaced image is also filtered, but they are filtered in relation
    ' to only the other scanlines in the same pass, different than non-interlaced images.
    ' Think of non-interlaced images as single-pass interlaced images.

    ' counter variables
    Dim Pass As Byte, srcRow As Long
    ' sizing/bit alignment variables
    Dim nr8wide As Long, nr8high As Long
    Dim nrBytes As Long, passPtr As Long
    Dim InterlacePass() As Byte  ' unfiltered progressive display (used 7x for 7 passes)
    ' bytes and bits per pixel values
    Dim bytesPP As Byte, BPRow As Long, bitPP As Byte
    
    ' need bit & byte information
    GetDepthInfo 0, 0, bitPP, bytesPP
    
    ' oversize array for "pass" bytes to prevent reszing array on each pass
    BPRow = GetBytesPerPixel((m_Width \ MatrixColAdd(7)), bitPP)
    ' how many bytes are needed for the final pass; largest pass size in bytes
    nrBytes = (BPRow + 1) * (m_Height \ MatrixRowAdd(7))
    ReDim InterlacePass(0 To nrBytes - 1)

    ' interlaced images always come in 7 passes; although not all passes may be used
    For Pass = 1 To 7
        ' ensure bounds are valid. If image is smaller than 8x8 not all passes are valid/used
        ' Tested with images as small as 1x1
    
        ' calculate nr of pixels for this pass that will fit in width of image
        nr8wide = m_Width \ MatrixColAdd(Pass) - (m_Width Mod MatrixColAdd(Pass) > MatrixCol(Pass))
        If nr8wide > 0 Then
            
            ' calcuate nr of rows for this pass that will fit in height of image
            nr8high = m_Height \ MatrixRowAdd(Pass) - (m_Height Mod MatrixRowAdd(Pass) > MatrixRow(Pass))
            If nr8high > 0 Then
    
                ' calculate row bytes for the interlaced image, byte aligned
                BPRow = GetBytesPerPixel(nr8wide, bitPP) + 1
                ' how many bytes are needed for the complete pass, less filter byte?
                nrBytes = BPRow * nr8high
                '^^ the filter routines expect the filter byte to be in its parameters, so add it
                
                ' unfilter the scanlines
                CopyMemory InterlacePass(0), Filtered(passPtr), nrBytes
                For srcRow = 0 To nr8high - 1
                    Select Case Filtered(BPRow * srcRow + passPtr)
                    Case 0: ' no filtering
                    Case 1: ' sub filter
                        DecodeFilter_Sub InterlacePass, srcRow, BPRow, bytesPP
                    Case 2: ' up filter
                        DecodeFilter_Up InterlacePass, srcRow, BPRow, 0
                    Case 3: ' average filter
                        DecodeFilter_Avg InterlacePass, srcRow, BPRow, bytesPP
                    Case 4: ' paeth filter
                        DecodeFilter_Paeth InterlacePass, srcRow, BPRow, bytesPP
                    Case Else
                        ' If we got here, there is a different filtering mechanism at large
                        Exit Function
                    End Select
                Next
        
                ' offset the filtered array pointer to account for the 1byte filter flag per scanline
                ' This will point to the next pass's X,Y position in the Unfiltered() array
                passPtr = passPtr + nrBytes
            
                ' send unfiltered array to be transfered to the DIB
                If ConvertPNGtoBMP(InterlacePass(), Pass, nr8high, BPRow) = False Then Exit Function
            
            End If ' check for nr8high < 1
        End If ' check for nr8wide < 1
    
    Next Pass

    UnfilterInterlaced = True
    
End Function

Private Function UnfilterNI(FilteredData() As Byte) As Boolean

    ' // LaVolpe, Dec 1 thru 10 - completely rewritten to remove excess large array usage
    ' http://www.w3.org/TR/PNG/#9-table91

    Dim Row As Long, BPRow As Long
    Dim lBpp As Byte, stepVal As Byte
    
    GetDepthInfo 0, 0, lBpp, stepVal
    BPRow = GetBytesPerPixel(m_Width, lBpp) + 1
    '^^ the filtered row contains an extra byte (1st byte of each row)
    '   that identifies the filter algorithm used for that row
    
    For Row = 0 To m_Height - 1

        Select Case FilteredData(BPRow * Row)
        Case 0 'no filtering
        Case 1 'Sub
            DecodeFilter_Sub FilteredData, Row, BPRow, stepVal
        Case 2 'Up
            DecodeFilter_Up FilteredData, Row, BPRow, 0
        Case 3 'Average
            DecodeFilter_Avg FilteredData, Row, BPRow, stepVal
        Case 4 'Paeth
            DecodeFilter_Paeth FilteredData, Row, BPRow, stepVal
        Case Else
            ' invalid filter type; no action
        End Select
        
    Next Row
    UnfilterNI = ConvertPNGtoBMP(FilteredData(), 8, Row, BPRow, 0)
    
End Function

Private Function zChunk_IHDR(bufLen As Long, streamOffset As Long, cmprSize As Long, crcValue As Long) As Long
                
    ' IHDR structure
    '    Width As Long              << cannot be negative
    '    Height As Long             << cannot be negative
    '    BitDepth As Byte           << must be 1,2,4,8,16
    '    ColorType As Byte          << must be 0,2,3,4,6
    '    Compression As Byte        << must be zero
    '    Filter As Byte             << must be zero
    '    Interlacing As Byte        << must be zero or one
    
    On Error Resume Next
    Dim lRtn As Long, lValue As Long
    
    If Not crcValue = 0 Then
        lRtn = Not (iparseReverseLong(crcValue) = Zcrc32(0, pngStream(streamOffset), bufLen + 4))
    End If
    If lRtn = 0 Then
        
        CopyMemory m_Width, pngStream(streamOffset + 4), 4
        m_Width = iparseReverseLong(m_Width)
        CopyMemory m_Height, pngStream(streamOffset + 8), 4
        m_Height = iparseReverseLong(m_Height)
        
        If m_Width < 1 Or m_Height < 1 Then
            lRtn = 1 'Corrupted Image Header. Cannot continue.
        
        Else
            
            If pngStream(streamOffset + 14) > 0 Then
                lRtn = 1 ' Invalid Compression Flag in Header. Cannot continue.
            Else
                If pngStream(streamOffset + 15) > 0 Then
                    lRtn = 1 'Invalid Filter Flag in Header. Cannot continue.
                Else
                    
                    m_BitDepth = pngStream(streamOffset + 12)
                    Select Case m_BitDepth
                    Case 1, 2, 4, 8, 16
                        ' it is a valid bit depth
                        m_ColorType = pngStream(streamOffset + 13)
                        Select Case m_ColorType
                        Case 0, 2, 3, 4, 6
                            ' it is a valid color type
                            m_Interlacing = pngStream(streamOffset + 16)
                            If m_Interlacing > 1 Then
                                lRtn = 1 'Invalid Interlacing Flag in Header. Cannot continue.
                            End If
                        Case Else
                            lRtn = 1 'Invalid Color Type Flag in Header. Cannot continue.
                        End Select
                    Case Else
                        lRtn = 1 'Invalid Bit Depth Flag in Header. Cannot continue.
                    End Select
                    
                End If  ' Filter flag
            End If  ' Compression flag
        End If  ' Dimensions
        
        If lRtn = 0 Then
            ' check for png sizes that would cause overflow errors in other calculations...
            ' This has 2 basic checks
            ' check DWord width alignment * height first are within bounds
            lValue = ((((m_Width * 32) + &H1F) And Not &H1F&) \ &H8) * m_Height
            ' see if uncompress png data is too long
            If Not Err Then cmprSize = CalcUncompressedWidth()
            If cmprSize = 0 Then
                If Err Then Err.Clear
                ' either the dWord aligned bytes required for the bitmap
                ' are too large (Long value) or the size of the total
                ' uncompressed pixel array is too large (Long value).
                ' Image is way too big to process anyway! would require GBs of memory
                lRtn = 1 'Image is too large to process. Cannot continue.
            End If
        End If
    End If

    zChunk_IHDR = lRtn

End Function

Private Function zChunk_PLTE(bufLen As Long, streamOffset As Long, crcValue As Long) As Long

    ' http://www.w3.org/TR/PNG/#11PLTE
    If m_ColorType = 0 Or m_ColorType = 4 Then Exit Function
    '^^ per specs, palettes shall not appear for those color types
    '   Since we can ignore the palette, we won't trigger a critcal error
    
    Dim lRtn As Long
    If Not crcValue = 0 Then
        lRtn = Not (iparseReverseLong(crcValue) = Zcrc32(0, pngStream(streamOffset), bufLen + 4))
    End If
    If lRtn = 0 Then
        
        ' per png specs, palette must be divisible by 3
        If (bufLen \ 3) * 3 = bufLen Then
            ReDim m_Palette(bufLen - 1)
            CopyMemory m_Palette(0), pngStream(streamOffset + 4), bufLen
        Else ' error
            lRtn = 1 'Invalid Palette. Cannot continue.
        End If
    End If
    zChunk_PLTE = lRtn

End Function

Private Function zChunk_tRNS(bufLen As Long, streamOffset As Long, crcValue As Long) As Long
'http://www.w3.org/TR/PNG/#11tRNS
    
    If m_ColorType > clrPalette Then Exit Function
    ' Per specs, the tRNS chunk shall not be used for Color Types 4 and 6

    On Error GoTo ExitMe
    Dim UB As Long, palIndex As Byte, lRtn As Long
    
    If Not crcValue = 0 Then
        lRtn = Not (iparseReverseLong(crcValue) = Zcrc32(0, pngStream(streamOffset), bufLen + 4))
    End If
    If lRtn = 0 Then
        
        ' we will ensure the passed array is dimensioned properly and also cache
        ' the simple transparency color for easier reference while processing
        
        ReDim m_TransSimple(0 To bufLen - 1)
        CopyMemory m_TransSimple(0), pngStream(streamOffset + 4), bufLen
    
        If m_ColorType = clrGrayScale Then ' grayscale with simple transparency
            ' least significant bits used. Tweak array to hold only those bits in byte format
            If m_BitDepth = 16 Then
                m_TransColor = m_TransSimple(1) ' color-index value not a color
            Else ' color-index value not a color
                GetPaletteValue 8 \ m_BitDepth - 1, m_TransSimple(1), palIndex
                m_TransColor = palIndex
            End If
            
        ElseIf m_ColorType = clrTrueColor Then ' rgb triple (true color)
            ' save as BGR to be compared against PNG samples
            m_TransColor = m_TransSimple(5) Or m_TransSimple(3) * 256& Or m_TransSimple(1) * 65536
            ' for 16bpp PNGs, the 0,2,4 array elements are needed also but will be tested in ConvertPngToBmp
            
        ElseIf m_ColorType = clrPalette Then ' TransSimple() is an array
            ' This array is directly related to the Palette. Each palette entry
            ' will have a related TransSimple() entry. Exception: When Palette entries
            ' are sorted (in ascending order of alpha value), then any Palette entries
            ' that have alpha values of 255 probably will not be in that related array.
            ' In these cases, we will fake it & provide the missing entries.
        
            ' to prevent out of bounds errors, ensure array is 255
            If UBound(m_TransSimple) < 255 Then ' pngs are not required to provide all
                UB = UBound(m_TransSimple)
                ReDim Preserve m_TransSimple(0 To 255)    ' prevent out ouf bounds errors
                FillMemory m_TransSimple(UB + 1), 255 - UB, 255
            End If
            m_TransColor = 0 ' simply a flag > -1, has no other meaning
        End If
        
        If Err Then
            Err.Clear   ' an error regarding the TransSimple() array
            m_TransColor = -1 ' no transparency color
        End If
    
    End If

ExitMe:
End Function

Private Sub InitializePalette()

    ' Purpose: Create a palette for the PNG file, if needed
    ' The colors from the palette will be transfered to the 32bpp image

    If m_ColorType = clrPalette Or m_ColorType = clrGrayScale Then
    
        Dim nrEntries As Long, stepVal As Long
        Dim X As Long, Index As Long, Color As Long

        ' PNG grayscale palettes are not provided, they are assumed...
        If iparseIsArrayEmpty(Not m_Palette) Then
            
            ReDim m_Palette(0 To 767)
            If m_ColorType = clrGrayScale Then
            
                nrEntries = pow2x8(m_BitDepth) - 1      ' number grayscale palette entries
                stepVal = 255 \ (pow2x8(m_BitDepth) - 1) ' step value for the palette
                For X = 1 To nrEntries
                    Color = X * stepVal
                    Index = X * 3
                    m_Palette(Index) = Color
                    m_Palette(Index + 1) = Color
                    m_Palette(Index + 2) = Color
                Next
                
            ElseIf m_BitDepth = 1 Then              ' fix up 2 color palette
                CopyMemory m_Palette(4), vbWhite, &H3
            End If
        End If
    
    End If
    
End Sub

Private Function IsPNG() As Boolean

    ' Purpose: Determine if PNG magic number exists in 1st 8 bytes of the file/array
    ' Note: array was already validated as not empty when class LoadStream/LoadFile was called
    Dim gpLong As Long
    
    ' validate we are looking at a png file
    CopyMemory gpLong, pngStream(LBound(pngStream)), &H4
    If gpLong = png_Signature1 Then
        CopyMemory gpLong, pngStream(LBound(pngStream) + 4), &H4
        IsPNG = (gpLong = png_Signature2)
    End If

End Function

Private Function GDIplusLoad(FileName As String) As Boolean
    
    ' Purpose: Use GDI+ to load a PNG either by fileName or by array/stream
    ' FileName :: if vbNullString, then the pngStream() array contains the
    '   PNG file else FileName is full path & name of the PNG file
    
    Dim gdiSI As GdiplusStartupInput
    Dim gdiToken As Long
    
    ' does the system have GDI+ on it?
    If ValidateDLL("gdiplus.dll", "GdiplusStartup") = False Then Exit Function
    
    ' if so, try starting it up
    gdiSI.GdiplusVersion = 1
    GdiplusStartup gdiToken, gdiSI
    If gdiToken = 0 Then Exit Function  ' failure; can't use GDI+
     
    Dim hImage As Long, hGraphics As Long
    Dim hDib As Long, tDC As Long, lRtn As Long
    Dim rDimensions As RECTF, pStream As IUnknown
    
    On Error GoTo ExitRoutine
    
    If FileName = vbNullString Then ' we need an array vs file name
        ' hack of my own. Create an IUnknown Stream that has the same properties
        ' and minimum methods needed as the IStream interface which VB does not
        ' expose. Once the stream is created, we have GDI+ load from it
        Set pStream = CreateStream(pngStream())
        If Not pStream Is Nothing Then Call GdipLoadImageFromStream(pStream, hImage)

    Else    ' we use the passed file name; have GDI+ load the file
        Call GdipLoadImageFromFile(StrConv(FileName, vbUnicode), hImage)
    End If
    
    If Not hImage = 0 Then
        ' get size of PNG
        lRtn = GdipGetImageBounds(hImage, rDimensions, UnitPixel)
        If lRtn = 0 Then

            ' cache sizes for other functions in this class, build 32bpp
            m_Width = CLng(rDimensions.nWidth)
            m_Height = CLng(rDimensions.nHeight)
            cHost.InitializeDIB m_Width, m_Height

            ' wrap a GDI+ DC around our DIB's DC
            tDC = cHost.LoadDIBinDC(True)
            lRtn = GdipCreateFromHDC(tDC, hGraphics)
            If lRtn = 0 Then
                ' now draw the PNG into our 32bpp. GDI+ is nice enough to pre-multiply
                ' the RGB values for us during the rendering
                With rDimensions
                    GdipDrawImageRectRectI hGraphics, hImage, 0, 0, .nWidth, .nHeight, .nLeft, .nTop, .nWidth, .nHeight, UnitPixel, 0, 0, 0
                End With
                GdipDeleteGraphics hGraphics    ' remove the GDI+ DC wrapper
                hGraphics = 0
            End If
            cHost.LoadDIBinDC False ' unselect our DIB
        End If

        If lRtn = 0 Then                    ' return results
            GDIplusLoad = True
            Call GdipGetImagePixelFormat(hImage, lRtn)
            GdipDisposeImage hImage             ' destroy the GDI+ image
            cHost.Alpha = (lRtn = PixelFormat32bppARGB Or lRtn = PixelFormat32bppPARGB)
            cHost.ImageType = imgPNG
        Else
            GdipDisposeImage hImage             ' destroy the GDI+ image
            cHost.DestroyDIB
        End If
        hImage = 0
    End If
        
ExitRoutine:
    If Not gdiToken = 0 Then
        If Not hGraphics = 0 Then GdipDeleteGraphics hGraphics
        If Not hImage = 0 Then GdipDisposeImage hImage
        GdiplusShutdown gdiToken   ' stop GDI+
    End If
End Function

Private Function ValidateDLL(DllName As String, dllProc As String) As Boolean
    
    ' PURPOSE: Test a DLL for a specific function.
    
    Dim lb As Long, pa As Long
    
    'attempt to open the DLL to be checked
    lb = LoadLibrary(DllName)
    If lb Then
        'if so, retrieve the address of one of the function calls
        pa = GetProcAddress(lb, dllProc)
        ' free references
        FreeLibrary lb
    End If
    ValidateDLL = (Not (lb = 0 Or pa = 0))
    
End Function

Private Function CreateStream(bytContent() As Byte, Optional byteOffset As Long = 0) As stdole.IUnknown
    
    ' Purpose: Create an IStream-compatible IUnknown interface containing the
    ' passed byte aray. This IUnknown interface can be passed to GDI+ functions
    ' that expect an IStream interface -- neat hack
    
    On Error GoTo HandleError
    Dim o_lngLowerBound As Long
    Dim o_lngByteCount  As Long
    Dim o_hMem As Long
    Dim o_lpMem  As Long
     
    If iparseIsArrayEmpty(Not bytContent) Then    ' create a growing stream as needed
         Call CreateStreamOnHGlobal(0, 1, CreateStream)
    Else                                    ' create a fixed stream
         o_lngByteCount = UBound(bytContent) - byteOffset + 1
         o_hMem = GlobalAlloc(&H2, o_lngByteCount)
         If o_hMem <> 0 Then
             o_lpMem = GlobalLock(o_hMem)
             If o_lpMem <> 0 Then
                 CopyMemory ByVal o_lpMem, bytContent(byteOffset), o_lngByteCount
                 Call GlobalUnlock(o_hMem)
                 Call CreateStreamOnHGlobal(o_hMem, 1, CreateStream)
             End If
         End If
     End If
    
HandleError:
End Function


' decompression using pure VB, this is only run if system
' does not have GDI+ nor does it have zLib.dll
' source by: alfred.koppold@freenet.de

Private Function vbDecompress(UnCompressedArray() As Byte, CompressedArray() As Byte, ByVal UncompressedSize As Long, Optional ZIP64 As Boolean = False) As Boolean
    
    Dim IsLastBlock As Boolean
    Dim CompType As Long
    Dim Char As Long
    Dim Nubits As Long
    Dim L1 As Long
    Dim L2 As Long
    Dim X As Long
    Dim lRtn As Long
    Dim MinLLength As Long
    Dim MaxLLength As Long
    Dim MinDLength As Long
    Dim MaxDLength As Long
    Dim IsStaticBuild As Boolean
    
    On Error GoTo Stop_Decompression
    
    Dim tSAIN As SafeArray1D, tSAOUT As SafeArray1D
    With tSAIN
        .cbElements = 1
        .cDims = 1
        .pvData = VarPtr(CompressedArray(0))
        .rgSABound.cElements = UBound(CompressedArray) + 1
    End With
    CopyMemory ByVal VarPtrArray(inStream), VarPtr(tSAIN), &H4
    
    With tSAOUT
        .cbElements = 1
        .cDims = 1
        .pvData = VarPtr(UnCompressedArray(0))
        .rgSABound.cElements = UncompressedSize
    End With
    CopyMemory ByVal VarPtrArray(outStream), VarPtr(tSAOUT), &H4
    

    'InStream = ByteArray
    Call vbInit_Decompress(UncompressedSize)
    Inpos = 2
    Do
        IsLastBlock = vbGetBits(1)    ' last compressed block?
        CompType = vbGetBits(2)       ' compression type used for block
        If CompType = 0 Then
            If Inpos + 4 > UBound(inStream) Then
                ' ensure 4 more bytes exist for L1 and L2 below
                lRtn = -1
                Exit Do
            End If
        
            If BitNum >= 8 Then
                Inpos = Inpos - (BitNum \ 8)
                BitNum = BitNum - ((BitNum \ 8) * 8)
            End If
            
            L1 = inStream(Inpos) Or (inStream(Inpos + 1) * &H100&)
            L2 = inStream(Inpos + 2) Or (inStream(Inpos + 3) * &H100&)
            
            Inpos = Inpos + 4
            If L1 - (Not (L2) And &HFFFF&) Then
                lRtn = -2
                Exit Do
            End If
            If Inpos + L1 - 1 > UBound(inStream) Then
                lRtn = -1
                Exit Do
            End If
            
            If OutPos + L1 - 1 > UBound(outStream) Then
                lRtn = -1
                Exit Do
            End If
            
            CopyMemory outStream(OutPos), inStream(Inpos), L1
            OutPos = OutPos + L1
            Inpos = Inpos + L1
            ByteBuff = 0
            BitNum = 0
            
        ElseIf CompType = 3 Then
            lRtn = -1
            Exit Do
            
        Else
            If CompType = 1 Then
                If vbCreate_Static_Tree(MinLLength, MaxLLength, MinDLength, MaxDLength, IsStaticBuild) <> 0 Then
                    lRtn = 9
                    Exit Do
                End If
            Else
                If vbCreate_Dynamic_Tree(MinLLength, MaxLLength, MinDLength, MaxDLength) <> 0 Then
                    lRtn = 9
                    Exit Do
                End If
            End If
 
            Do
                vbNeedBits MaxLLength
                Nubits = MinLLength
                Do While LitLen.Length(ByteBuff And BitMask(Nubits)) <> Nubits
                    Nubits = Nubits + 1
                Loop
 
                Char = LitLen.code(ByteBuff And BitMask(Nubits))
                vbDropBits Nubits
                
                If Char < 256 Then
                    outStream(OutPos) = Char
                    OutPos = OutPos + 1
                    
                ElseIf Char > 256 Then
                    Char = Char - 257
                    L1 = LCodes.code(Char) + vbGetBits(LCodes.Length(Char))
                    If (L1 = 258) And ZIP64 Then L1 = vbGetBits(16) + 3
                    vbNeedBits MaxDLength
                    Nubits = MinDLength

                    Do While Dist.Length(ByteBuff And BitMask(Nubits)) <> Nubits
                        Nubits = Nubits + 1
                    Loop
 
                    Char = Dist.code(ByteBuff And BitMask(Nubits))
                    vbDropBits Nubits
                    L2 = DCodes.code(Char) + vbGetBits(DCodes.Length(Char))
    
                    For X = 1 To L1
                        If OutPos > UncompressedSize Then
                            OutPos = UncompressedSize
                            GoTo Stop_Decompression
                        End If
                         outStream(OutPos) = outStream(OutPos - L2)
                         OutPos = OutPos + 1
                    Next X
                
                End If

            Loop While Char <> 256 'EOF
        End If
    Loop While Not IsLastBlock
    
Stop_Decompression:

CopyMemory ByVal VarPtrArray(inStream), 0&, 4
CopyMemory ByVal VarPtrArray(outStream), 0&, 4

Erase BitMask
Erase Pow2
Erase LCodes.code
Erase LCodes.Length
Erase DCodes.code
Erase DCodes.Length
Erase LitLen.code
Erase LitLen.Length
Erase Dist.code
Erase Dist.Length
Erase LenOrder

If Err Then
    lRtn = Err.Number
    Err.Clear
End If

vbDecompress = (lRtn = 0)

End Function

Private Function vbCreate_Static_Tree(MinLLength As Long, MaxLLength As Long, MinDLength As Long, MaxDLength As Long, IsStaticBuild As Boolean) As Long

    Dim X As Long
    Dim Length(287) As Long

    If IsStaticBuild = False Then
        ' quick fill the tree (tile Blt)
        For X = 0 To 7: Length(X) = 8: Next ' 0 to 143 elements = 8
        For X = X To 143 Step 8
            CopyMemory Length(X), Length(0), 32
        Next
        For X = X To X + 7: Length(X) = 9: Next ' 144 to 255 elements = 9
        For X = X To 255 Step 8
            CopyMemory Length(X), Length(144), 32
        Next
        For X = X To X + 7: Length(X) = 7: Next ' 256 to 279 elements = 7
        For X = X To 279 Step 8
            CopyMemory Length(X), Length(256), 32
        Next
        For X = X To X + 3: Length(X) = 8: Next ' 280 to 287 elements = 8
        CopyMemory Length(X), Length(280), 16
        
        If vbCreate_Codes(TempLit, Length, 287, MaxLLength, MinLLength) <> 0 Then
             vbCreate_Static_Tree = -1
            Exit Function
        End If
        
        For X = 0 To 7: Length(X) = 5: Next  ' reset 0 to 32 to 5's
        For X = X To 31 Step 8
            CopyMemory Length(X), Length(0), 32
        Next
        vbCreate_Static_Tree = vbCreate_Codes(TempDist, Length, 31, MaxDLength, MinDLength)
        IsStaticBuild = True
        
    Else
        MinLLength = 7
        MaxLLength = 9
        MinDLength = 5
        MaxDLength = 5
    End If

    LitLen = TempLit
    Dist = TempDist
End Function

Private Function vbCreate_Dynamic_Tree(MinLLength As Long, MaxLLength As Long, MinDLength As Long, MaxDLength As Long) As Long
    Dim Length() As Long
    Dim Bl_Tree As CodesType
    Dim MinBL As Long
    Dim MaxBL As Long
    Dim NumLen As Long
    Dim Numdis As Long
    Dim NumCod As Long
    Dim Char As Long
    Dim Nubits As Long
    Dim LN As Long
    Dim Pos As Long
    Dim X As Long

    NumLen = vbGetBits(5) + 257
    Numdis = vbGetBits(5) + 1
    NumCod = vbGetBits(4) + 4

    ReDim Length(18)
    For X = 0 To NumCod - 1
        Length(LenOrder(X)) = vbGetBits(3)
    Next

    For X = NumCod To 18
        Length(LenOrder(X)) = 0
    Next

    If vbCreate_Codes(Bl_Tree, Length, 18, MaxBL, MinBL) <> 0 Then
        vbCreate_Dynamic_Tree = -1
        Exit Function
    End If
    
    ReDim Length(NumLen + Numdis)
    Pos = 0
    Do While Pos < NumLen + Numdis
        vbNeedBits MaxBL
        Nubits = MinBL
        
        Do While Bl_Tree.Length(ByteBuff And BitMask(Nubits)) <> Nubits
             Nubits = Nubits + 1
        Loop

        Char = Bl_Tree.code(ByteBuff And BitMask(Nubits))
        vbDropBits Nubits

        If Char < 16 Then
            Length(Pos) = Char
            Pos = Pos + 1
        Else
            If Char = 16 Then
                If Pos = 0 Then
                    vbCreate_Dynamic_Tree = -5
                    Exit Function
                End If
                LN = Length(Pos - 1)
                Char = 3 + vbGetBits(2)
            ElseIf Char = 17 Then
                 Char = 3 + vbGetBits(3)
                 LN = 0
            Else
                Char = 11 + vbGetBits(7)
                LN = 0
            End If

            If Pos + Char > NumLen + Numdis Then
                vbCreate_Dynamic_Tree = -6
                Exit Function
            End If

            Do While Char > 0
                Char = Char - 1
                Length(Pos) = LN
                Pos = Pos + 1
            Loop
        End If
    Loop

    If vbCreate_Codes(LitLen, Length, NumLen - 1, MaxLLength, MinLLength) <> 0 Then
        vbCreate_Dynamic_Tree = -1
        Exit Function
    End If

    For X = 0 To Numdis
        Length(X) = Length(X + NumLen)
    Next
    
    vbCreate_Dynamic_Tree = vbCreate_Codes(Dist, Length, Numdis - 1, MaxDLength, MinDLength)

End Function

Private Function vbCreate_Codes(tree As CodesType, Lengths() As Long, NumCodes As Long, MaxBits As Long, Minbits As Long) As Long
    
    Dim Bits(16) As Long
    Dim next_code(16) As Long
    Dim code As Long
    Dim LN As Long
    Dim X As Long

    Minbits = 16
    For X = 0 To NumCodes
        Bits(Lengths(X)) = Bits(Lengths(X)) + 1
        If Lengths(X) > MaxBits Then MaxBits = Lengths(X)
        If Lengths(X) < Minbits And Lengths(X) > 0 Then Minbits = Lengths(X)
    Next

    LN = 1
    For X = 1 To MaxBits
        LN = LN + LN
        LN = LN - Bits(X)
        If LN < 0 Then
            vbCreate_Codes = LN
            Exit Function
        End If
    Next

    vbCreate_Codes = LN
    ReDim tree.code(2 ^ MaxBits - 1)
    ReDim tree.Length(2 ^ MaxBits - 1)
    code = 0
    Bits(0) = 0
    
    For X = 1 To MaxBits
        code = (code + Bits(X - 1)) * 2
        next_code(X) = code
    Next

    For X = 0 To NumCodes
        LN = Lengths(X)
        If LN <> 0 Then
            code = vbBit_Reverse(next_code(LN), LN)
            tree.Length(code) = LN
            tree.code(code) = X
            next_code(LN) = next_code(LN) + 1
        End If
    Next

End Function

Private Function vbBit_Reverse(ByVal Value As Long, ByVal Numbits As Long)

    Do While Numbits > 0
        vbBit_Reverse = vbBit_Reverse * 2 + (Value And 1)
        Numbits = Numbits - 1
        Value = Value \ 2
    Loop
    
End Function

Private Sub vbInit_Decompress(UncompressedSize As Long)
    
    Dim Temp()
    Dim X As Long
    ReDim outStream(UncompressedSize)

    Erase LitLen.code
    Erase LitLen.Length
    Erase Dist.code
    Erase Dist.Length

    ReDim LCodes.code(31)
    ReDim LCodes.Length(31)
    ReDim DCodes.code(31)
    ReDim DCodes.Length(31)
    ReDim LenOrder(0 To 18)

    Temp() = Array(16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15)
    For X = 0 To UBound(Temp): LenOrder(X) = Temp(X): Next
     Temp() = Array(3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258)
    For X = 0 To UBound(Temp): LCodes.code(X) = Temp(X): Next
     Temp() = Array(0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0)
    For X = 0 To UBound(Temp): LCodes.Length(X) = Temp(X): Next
     Temp() = Array(1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577, 32769, 49153)
    For X = 0 To UBound(Temp): DCodes.code(X) = Temp(X): Next
     Temp() = Array(0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13, 14, 14)
    For X = 0 To UBound(Temp): DCodes.Length(X) = Temp(X): Next
    
    ReDim Pow2(0 To 16)
    ReDim BitMask(0 To 16)
    Pow2(0) = 1
    For X = 1 To 16
     Pow2(X) = Pow2(X - 1) * 2
     BitMask(X) = Pow2(X) - 1
    Next
    OutPos = 0
    Inpos = 0
    ByteBuff = 0
    BitNum = 0

End Sub

Private Sub vbNeedBits(Numbits As Long)
    Do While BitNum < Numbits
        If Inpos > UBound(inStream) Then Exit Do
        ByteBuff = ByteBuff + (inStream(Inpos) * Pow2(BitNum))
        BitNum = BitNum + 8
        Inpos = Inpos + 1
    Loop
End Sub

Private Sub vbDropBits(Numbits As Long)
    ByteBuff = ByteBuff \ Pow2(Numbits)
    BitNum = BitNum - Numbits
End Sub

Private Function vbGetBits(Numbits As Long) As Long
    
    While BitNum < Numbits
        ByteBuff = ByteBuff + (inStream(Inpos) * Pow2(BitNum))
        BitNum = BitNum + 8
        Inpos = Inpos + 1
    Wend
    vbGetBits = ByteBuff And BitMask(Numbits)
    ByteBuff = ByteBuff \ Pow2(Numbits)
    BitNum = BitNum - Numbits

End Function


' this is not a final routine yet; testing phase only
Public Function SaveTo(FileName As String, outStream() As Byte, cHost As c32bppDIB) As Boolean

    ' this is in testing. not finished yet, but working

    Dim gdiSI As GdiplusStartupInput
    Dim gdiToken As Long
    
    If cHost.Handle = 0 Then Exit Function
    
    ' does the system have GDI+ on it?
    If ValidateDLL("gdiplus.dll", "GdiplusStartup") = False Then Exit Function
    
    ' if so, try starting it up
    gdiSI.GdiplusVersion = 1
    GdiplusStartup gdiToken, gdiSI
    If gdiToken = 0 Then Exit Function  ' failure; can't use GDI+

    Dim hImg         As Long
    Dim uEncCLSID(0 To 3) As Long
    Dim tBMPI As BITMAPINFO
    Dim IIStream As IUnknown
    
    
    If cHost.Alpha = True Then
        Call GdipCreateBitmapFromScan0(cHost.Width, cHost.Height, cHost.ScanWidth, PixelFormat32bppPARGB, ByVal cHost.BitsPointer, hImg)
        ' dib is bottom up, scan0 does top down, so flip it
        GdipImageRotateFlip hImg, 6 ' flip vertically
    Else
        With tBMPI.bmiHeader
            .biBitCount = 32
            .biHeight = cHost.Height
            .biWidth = cHost.Width
            .biPlanes = 1
            .biSize = 40
        End With
        GdipCreateBitmapFromGdiDib tBMPI, ByVal cHost.BitsPointer, hImg
    End If
    If Not hImg = 0 Then
        Call pvGetEncoderClsID("image/png", uEncCLSID)
        If FileName = vbNullString Then
            Erase outStream
            Set IIStream = CreateStream(outStream)
            If GdipSaveImageToStream(hImg, IIStream, uEncCLSID(0), ByVal 0&) = 0 Then
                ArrayFromStream IIStream, outStream()
                SaveTo = True
            End If
        Else
            On Error Resume Next
            If Len(Dir$(FileName, vbArchive Or vbHidden Or vbReadOnly Or vbSystem)) > 0 Then
                SetAttr FileName, vbNormal
                Kill FileName
                If Err Then
                    Err.Clear
                    Exit Function
                End If
            End If
            SaveTo = (GdipSaveImageToFile(hImg, StrConv(FileName, vbUnicode), uEncCLSID(0), ByVal 0&) = 0)
        End If
        GdipDisposeImage hImg
    End If
    GdiplusShutdown gdiToken   ' stop GDI+
    
End Function

Private Function pvPtrToStrW(ByVal lpsz As Long) As String
  ' supporting routine for SaveToFile
  Dim sOut As String
  Dim lLen As Long

    lLen = lstrlenW(lpsz)

    If (lLen > 0) Then
        sOut = StrConv(String$(lLen, vbNullChar), vbUnicode)
        Call CopyMemory(ByVal sOut, ByVal lpsz, lLen * 2)
        pvPtrToStrW = StrConv(sOut, vbFromUnicode)
    End If
End Function

Private Function pvGetEncoderClsID(strMimeType As String, ClassID() As Long) As Long
  ' supporting routine for SaveToFile
  Dim Num      As Long
  Dim Size     As Long
  Dim lIdx     As Long
  Dim ICI()    As ImageCodecInfo
  Dim Buffer() As Byte
    
    pvGetEncoderClsID = -1 ' Failure flag
    
    '-- Get the encoder array size
    Call GdipGetImageEncodersSize(Num, Size)
    If (Size = 0) Then Exit Function ' Failed!
    
    '-- Allocate room for the arrays dynamically
    ReDim ICI(1 To Num) As ImageCodecInfo
    ReDim Buffer(1 To Size) As Byte
    
    '-- Get the array and string data
    Call GdipGetImageEncoders(Num, Size, Buffer(1))
    '-- Copy the class headers
    Call CopyMemory(ICI(1), Buffer(1), (Len(ICI(1)) * Num))
    
    '-- Loop through all the codecs
    For lIdx = 1 To Num
        '-- Must convert the pointer into a usable string
        If (StrComp(pvPtrToStrW(ICI(lIdx).MimeType), strMimeType, vbTextCompare) = 0) Then
            CopyMemory ClassID(0), ICI(lIdx).ClassID(0), 16& ' Save the Class ID
            pvGetEncoderClsID = lIdx      ' Return the index number for success
            Exit For
        End If
    Next lIdx
    '-- Free the memory
    Erase ICI
    Erase Buffer
End Function

Private Function ArrayFromStream(Stream As IUnknown, arrayBytes() As Byte) As Boolean

    Dim o_hMem As Long, o_lpMem As Long
    Dim o_lngByteCount As Long
    
    If Not Stream Is Nothing Then
    
        If GetHGlobalFromStream(ByVal ObjPtr(Stream), o_hMem) = 0 Then
            o_lngByteCount = GlobalSize(o_hMem)
            If o_lngByteCount > 0 Then
                o_lpMem = GlobalLock(o_hMem)
                If o_lpMem <> 0 Then
                    ReDim arrayBytes(0 To o_lngByteCount - 1)
                    CopyMemory arrayBytes(0), ByVal o_lpMem, o_lngByteCount
                    GlobalUnlock o_hMem
                    ArrayFromStream = True
                End If
            End If
        End If

    End If

End Function

