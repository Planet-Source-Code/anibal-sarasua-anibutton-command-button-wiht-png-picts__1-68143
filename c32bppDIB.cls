VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c32bppDIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Not truly Win95 compatible. Win95 compatiblity will take some work. Carles P.V.
' posted a very good stretch interpolation routine on PSC. That routine would be
' needed to stretch an alpha image with great quality, but then per pixel merging
' of the alpha image to the destination DC is also needed. If image is not alpha,
' then StretchBlt could be used for Win95, which is how this class will fall back to.

' About 32bpp pre-multiplied RGB (pARGB) bitmaps, if you are not aware.
'   - These are used specifically for the AlphaBlend API & are GDI+ compatible
'   Advantages:
'       - Images can be per-pixel alpha blended
'       - Opacity can be simultaneously adjusted during rendering
'       - AlphaBlend does both BitBlt & StretchBlt for pARGB images.
'       - Speed: AlphaBlend & GDI+ are pretty quick APIs vs manual blending
'   Disadvantages:
'       - The original RGB values are permanently destroyed during pre-multiplying
'           -- there is no way to convert pARGB back to non-premultiplied RGB values
'           -- the formula would be: reconstructedRed=(preMultipliedRed * 255) \ Alpha.
'               but because of integer division when pre-multiplying the result is only close
'               and if this should be premultiplied again & converted again, the results can get worse
'       - Displaying a pre-multiplied bitmap without AlphaBlend will not result in
'           the image being displayed as expected.
'       - Not ideal for saving due to its size: SizeOf= W x H x 4
'           -- better to save source image instead or compress the DIB bytes using favorite compression utility
'           -- with GDI+, image can be converted to PNG for storage
'       - AlphaBlend & GDI+ are not supported on Win95

' This class holds the 32bpp image. It also marshals any new image thru
' the battery of parsers to determine best method for converting the image
' to a 32bpp alpha-compatible image.

' The parser order is very important for fastest/best results...
' cPNGparser :: will convert PNG, all bit depths; aborts quickly if not PNG
' cGIFparser :: will convert non-transparent/transparent GIFs; aborts quickly
' cICOpraser :: will convert XP-Alpha, paletted, true color, & Vista PNG icons
'               -- can also convert most non-animated cursors
' cBMPparser :: will convert bitmaps, wmf/emf & jpgs

' The parsers are efficient. Most image formats have a magic number that give
'   a hint to what type of image the file/stream is. However, checks need to
'   be employed because other files could feasibly have those same magic
'   numbers. If the image is determined not to be one the parser is designed
'   to handle, the parser rejects it and the next parser takes over.  The
'   icon parser is slightly different because PNG files can be included into
'   a Vista ico file. When this occurs, the icon parser will pass off the
'   PNG format to the PNG parser automatically.
' And last but not least, the parsers have no advanced knowledge of the image
' format; as far as they are concerned anything passed is just a byte array

' No APIs are declared public. This is to prevent possibly, differently
' declared APIs, or different versions of the same API, from conflciting
' with any APIs you declared in your project. Same rule for UDTs.
' Note: I did take some liberties in several API declarations throughout

' CHANGE HISTORY
' 6 Jan 07:
'   - Added SaveFormat parameter to LoadPicture_File & LoadPicture_Stream
'       -- option makes class cache the original bytes of the image if loaded
'       -- the 32bpp DIB will always be larger than the source bytes and for
'           usercontrols, saving the original bytes takes less space than
'           saving the DIB bytes.
'   - Added GetOrginalFormat to retrieve bytes when SaveFormat was passed as True
'   - Added SetOriginalFormat used when copying one DIB class to another
'   - Tweaked cPNGparser.SaveTo to properly save when FileName is passed as null
'   - Bug fixed: Didn't read the operating system correctly during Initialize
'   - Also calculating Win95 in .Render & .isWin95Alpha was wrong; fixed
' 4 Jan 07:
'   - Added LoadPicture_ByHandle, LoadPicture_StdPicture, ScaleImage & CopyImageTo
'   - Added cPNGparser.SaveTo (testing). Requires GDI+ but will save 32bpp to PNG file or stream
'       -- not accessible, right now, from c32bppDIB. Must initialize cPNGparser to use it.
'   - Modified cICOparser's GetBestMatch algorithm
'   - Added imgPNGicon as an image type to distinguish PNG in Vista Icon vs standard .PNG file
'   - Bug found: removing overlays in cGIFparser.ConvertGIFto32bpp; forgot ByVal VarPtrArray(...)
'       which could cause crash when compiled. Fixed & double checked everywhere else too
' 1 Jan 07:
'   - Added SaveToFile & SaveToStream methods
'   - cBMPparser could possibly try to query unauthorized memory; fixed
'   - Methodology changed a bit when parsers return results. If image is definitely one
'       that the parser is responsible for & the image is invalid, the parser will return
'       True to prevent other parsers from handling the image. The c32bppDIB.Handle is used
'       to determine true success or failure.
'       -- cGIFparser when recognizing improperly formatted GIF would allow image to continue to
'           other parsers which then may cause those parsers to lock up.
' 26 Dec 06: First version


' Used to determine operating system
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As Any) As Long
Private Const VER_PLATFORM_WIN32_WINDOWS As Long = &H1
Private Type OSVERSIONINFOEX
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128 ' up to here is OSVERSIONINFO vs EX
   wServicePackMajor As Integer ' 8 bytes larger than OSVERSIONINFO
   wServicePackMinor As Integer
   wSuiteMask As Integer
   wProductType As Byte
   wReserved As Byte
End Type

' APIs used to manage the 32bpp DIB
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hDC As Long, ByRef pBitmapInfo As Any, ByVal un As Long, ByRef pointer As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal lBlendFunction As Long) As Long
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictDesc, riid As Any, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long
Private Declare Function GetObjectType Lib "gdi32.dll" (ByVal hgdiobj As Long) As Long
Private Declare Function GetIconInfo Lib "user32.dll" (ByVal hIcon As Long, ByRef piconinfo As ICONINFO) As Long
Private Const STRETCH_HALFTONE As Long = 4
Private Const OBJ_BITMAP As Long = 7
Private Const OBJ_METAFILE As Long = 9
Private Const OBJ_ENHMETAFILE As Long = 13


' used when saving an image or part of the image
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Type SafeArrayBound
    cElements As Long
    lLbound As Long
End Type
Private Type SafeArray
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    rgSABound(0 To 1) As SafeArrayBound ' reusable UDT for 1 & 2 dim arrays
End Type

Private Type PictDesc
    Size As Long
    Type As Long
    hHandle As Long
    hPal As Long
End Type
Private Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
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
    bmiPalette As Long
End Type

Private Const AC_SRC_OVER = &H0
Private Const AC_SRC_ALPHA = &H1

Public Enum eImageFormat    ' source image format
    imgNone = 0    ' no image loaded
    imgBitmap = 1  ' standard bitmap or jpg
    imgIcon = 3    ' standard icon
    imgWMF = 2     ' windows meta file
    imgEMF = 4     ' enhanced WMF
    imgCursor = 5  ' standard cursor
    imgBmpARGB = 6  ' 32bpp bitmap where RGB is not pre-multiplied
    imgBmpPARGB = 7 ' 32bpp bitmap where RGB is pre-multiplied
    imgIconARGB = 8 ' XP-type icon; 32bpp ARGB
    imgGIF = 9      ' gif; if class.Alpha=True, then transparent GIF
    imgPNG = 10     ' PNG image
    imgPNGicon = 11 ' PNG in icon file (Vista)
    imgCursorARGB = 12 ' alpha blended cursors? do they exist yet?
End Enum

Public Enum eScaleOptions
    scaleToSize = 0         ' [Default] will always scale
    scaleDownAsNeeded = 1   ' will only scale down if image won't fit
    ScaleStretch = 2        ' wll always stretch/distort
End Enum

Private m_Handle As Long        ' handle to 32bpp DIB
Private m_Pointer As Long       ' pointer to DIB bits
Private m_Height As Long        ' height of DIB
Private m_Width As Long         ' width of DIB
Private m_hDC As Long           ' DC if self-managing one
Private m_prevObj As Long       ' object deselected from DC when needed
Private m_os9x As Long          ' 1=win9x, 3=win95
Private m_Format As eImageFormat ' type of source image
Private m_ManageDC As Boolean   ' does class manage its own DC
Private m_AlphaImage As Boolean ' does the DIB contain alpha/transparency
Private m_ImageBytes() As Byte  ' should you want the DIB class to cache original bytes
' ^^ N/A if image is loaded by handle, stdPicture, or resource

Public Function LoadPicture_File(ByVal FileName As String, _
                                Optional ByVal iconCx As Long, _
                                Optional ByVal iconCy As Long, _
                                Optional ByVal SaveFormat As Boolean) As Boolean

    ' PURPOSE: Convert passed image file into a 32bpp image
    
    ' Parameters.
    ' FileName :: full path of file. Validation occurs before we continue
    ' iconCx :: desired width of icon if file is an icon file. Default is 32x32
    ' iconCy :: desired height of icon if file is an icon file. Default is 32x32
    ' SaveFormat :: if true, then the image will be cached as a byte array only
    '   if the image was successfully loaded. Call GetOrginalFormat to retrieve them.
    
    ' Why would you want to save the bytes? If this is being used in a usercontrol,
    ' saving the bytes will almost always be less size than saving the 32bit DIB.
    ' Additionally, these classes have the ability to get different sizes from
    ' the original source (i.e., WMF, icon, cursors) if available, but if the
    ' 32bit DIB is saved, it is a constant size. The potential of different sizes
    ' could allow better resizing of the image vs stretching the DIB.

    On Error Resume Next
    If Len(Dir$(FileName, vbArchive Or vbReadOnly Or vbSystem Or vbHidden)) = 0 Then Exit Function
    If FileLen(FileName) < 57 Then Exit Function
    ' no image file/stream can be less than 57 bytes and still be an image
    If Err Then
        Err.Clear
        Exit Function
    End If
    Dim aDIB() As Byte  ' dummy array
    LoadPicture_File = LoadPictureEx(FileName, aDIB(), iconCx, iconCy, 0, 0, SaveFormat)
    
End Function

Public Function LoadPicture_Stream(inStream() As Byte, _
                                    Optional ByVal iconCx As Long, _
                                    Optional ByVal iconCy As Long, _
                                    Optional ByVal streamStart As Long = 0, _
                                    Optional ByVal streamLength As Long = 0, _
                                    Optional ByVal SaveFormat As Boolean) As Boolean
    
    ' PURPOSE: Convert passed array into a 32bpp image
    
    ' Parameters.
    ' inStream:: byte stream containing the image. Validation occurs below
    ' iconCx :: desired width of icon if file is an icon file. Default is 32x32
    ' iconCy :: desired height of icon if file is an icon file. Default is 32x32
    ' streamStart :: array position of 1st byte of the image file. Validated.
    ' streamLength :: total length of the image file. Validated.
    ' SaveFormat :: if true, then the image will be cached as a byte array only
    '   if the image was successfully loaded. Call GetOrginalFormat to retrieve them.
    
    ' Why would you want to save the bytes? If this is being used in a usercontrol,
    ' saving the bytes will almost always be less size than saving the 32bit DIB.
    ' Additionally, these classes have the ability to get different sizes from
    ' the original source (i.e., WMF, icon, cursors) if available, but if the
    ' 32bit DIB is saved, it is a constant size. The potential of different sizes
    ' could allow better resizing of the image vs stretching the DIB.
    
    If iparseIsArrayEmpty(Not inStream) Then Exit Function
    If streamStart < LBound(inStream) Then streamStart = LBound(inStream)
    If streamLength = 0 Then streamLength = UBound(inStream) - streamStart + 1
    If streamLength < 57 Then Exit Function
    ' no image file/stream can be less than 57 bytes and still be an image
    LoadPicture_Stream = LoadPictureEx(vbNullString, inStream, iconCx, iconCy, streamStart, streamLength, SaveFormat)

End Function

Public Function LoadPicture_Resource(ByVal ResIndex As Variant, ByVal ResSection As Variant, _
                            Optional VbGlobal As IUnknown, _
                            Optional ByVal iconCx As Long, _
                            Optional ByVal iconCy As Long, _
                            Optional ByVal streamStart As Long = 0, _
                            Optional ByVal streamLength As Long = 0) As Boolean

    ' PURPOSE: Convert passed resource array into a 32bpp image
    
    ' Parameters.
    ' ResIndex :: the resource file index (i.e., 101)
    ' ResSection :: one of the selections or String value of your resource section
    '       - i.e., vbResBitmap, vbResIcon, "Custom", etc
    ' VbGlobal :: pass as VB.GLOBAL of the project containing the resource file
    '       - Allows class to be mobile; can exist in DLL or OCX
    '       - if not provided, class will use resource from existing workspace
    '       - For example, if this class was in a compiled OCX, then the only way
    '           to use the host's resource file is passing the host's VB.Global reference
    ' iconCx :: desired width of icon if file is an icon file. Default is 32x32
    ' iconCy :: desired height of icon if file is an icon file. Default is 32x32
    ' streamStart :: array position of 1st byte of the image file. Validated.
    ' streamLength :: total length of the image file. Validated.
    '   -- See LoadPicture_Stream for the validation
    
    ' Tips:
    ' 1) Store 32bpp bitmaps in the "Custom" resource always. Storing in the
    '       Bitmap resource can change color depth depending on your screen settings
    ' 2) Icons, normal bitmaps, & cursors are generally stored in their own sections
    ' 3) All other types of images are normally stored in the "Custom" section
    
    On Error GoTo ExitRoutine
    
    DestroyDIB
    
    Dim oWorkSpace As VB.Global, tPic As StdPicture
    
    If VbGlobal Is Nothing Then
        Set oWorkSpace = VB.Global
    ElseIf TypeOf VbGlobal Is VB.Global Then
        Set oWorkSpace = VbGlobal
    Else
        Set oWorkSpace = VB.Global
    End If
    
    If VarType(ResSection) = vbString Then
        Dim inStream() As Byte
        ' could be anything, PNG,icon,gif,32bpp bitmap,wmf, etc
        inStream = oWorkSpace.LoadResData(ResIndex, ResSection)
        LoadPicture_Resource = LoadPicture_Stream(inStream, iconCx, iconCy, streamStart, streamLength)
    Else
        ' can only be single icon, bitmap or cursor
        Set tPic = oWorkSpace.LoadResPicture(ResIndex, ResSection)
        LoadPicture_StdPicture tPic
    End If
    LoadPicture_Resource = Not (m_Handle = 0)
    
ExitRoutine:
    If Err Then Err.Clear
End Function

Public Function LoadPicture_StdPicture(Picture As StdPicture) As Boolean

    ' PURPOSE: Convert passed stdPicture into a 32bpp image
    
    If Not Picture Is Nothing Then
        If Picture.Type = vbPicTypeBitmap Then
            Dim cBMP As New cBMPparser
            Call cBMP.ConvertstdPicTo32bpp(Picture, Me, 0)
            Set cBMP = Nothing
        Else
            Dim cICO As New cICOparser
            Call cICO.ConvertstdPicTo32bpp(Picture, Me)
            Set cICO = Nothing
        End If
        LoadPicture_StdPicture = Not (m_Handle = 0)
    End If
    

End Function

Public Function LoadPicture_ByHandle(Handle As Long) As Boolean

    ' PURPOSE: Convert passed image handle into a 32bpp image

    Dim icoInfo As ICONINFO, tPic As StdPicture
    If Not Handle = 0 Then
        Select Case GetObjectType(Handle)
        Case OBJ_BITMAP, OBJ_METAFILE, OBJ_ENHMETAFILE
            Set tPic = HandleToStdPicture(Handle, vbPicTypeBitmap)
        Case Else
            If Not GetIconInfo(Handle, icoInfo) = 0 Then
                If Not icoInfo.hbmColor = 0 Then DeleteObject icoInfo.hbmColor
                If Not icoInfo.hbmMask = 0 Then DeleteObject icoInfo.hbmMask
                Set tPic = HandleToStdPicture(Handle, vbPicTypeIcon)
            End If
        End Select
        If Not tPic Is Nothing Then
            LoadPicture_ByHandle = LoadPicture_StdPicture(tPic)
        End If
    End If
    
End Function

Public Sub ScaleImage(ByVal destWidth As Long, ByVal destHeight As Long, NewWidth As Long, NewHeight As Long, Optional ByVal ScaleMode As eScaleOptions = scaleDownAsNeeded)
                            
    ' Purpose: Returns the width and height needed to draw the image to the requested dimensions.
    ' Function should be called before .Render should you want to scale the image.  Additionally,
    ' scaling can assist in positioning image too, i.e., centering.
    
    ' destWidth [in]:: the width of the target canvas (drawing area)
    ' destHeight [in]:: the height the target canvas
    ' NewWidth [out]:: returns the width to use for the supplied ScaleMode
    ' NewHeight [out]:: returns the height to use for the supplied ScaleMode
    ' ScaleMode [in]::
    '   scaleToSize [Default] - will always proportionally stretch the image to the target canvas size
    '   scaleDownAsNeeded - will only shrink the image if needed; otherwise the original image size is passed
    '   scaleStretch - the return value is always the canvas width and height; image distortion occurs
                            
    If m_Handle = 0& Then Exit Sub
    
    Dim RatioX As Single, RatioY As Single
    ' calculate scale and offsets
    Select Case ScaleMode
    
    Case scaleDownAsNeeded, scaleToSize: ' scaled
        RatioX = destWidth / m_Width
        RatioY = destHeight / m_Height
        If ScaleMode = scaleDownAsNeeded Then
            If RatioX > 1! And RatioY > 1! Then
                RatioX = 1!: RatioY = RatioX
            End If
        End If
        If RatioX > RatioY Then RatioX = RatioY
        NewWidth = Int(RatioX * m_Width)
        NewHeight = Int(RatioX * m_Height)
    
        ' to center your image in the target canvas. Use the passed & returned parameters like so:
        ' canvasX = (destWidth - NewWidth) \ 2 + any Left offset you may be using
        ' canvasY = (destHeight - NewHeight) \ 2 + any Top offset you may be using
        ' canvasX and canvasY would then be passed to .Render as .Render's X,Y parameters
        
    Case ScaleStretch
        NewWidth = destWidth
        NewHeight = destHeight
    
    Case Else
        NewWidth = m_Width
        NewHeight = m_Height
        
    End Select

End Sub

Public Sub CopyImageTo(cDIBclass As c32bppDIB, Optional CopyOriginalFormat As Boolean = False)
    
    ' Function replicates the the current image to another DIB class
    
    ' If CopyOriginalFormat = True then, and only, if class loaded its image
    ' with the optional SaveFormat=True, then the original image bytes
    ' were cached and will be copied to the target cDIBclass also
    ' See LoadPicture_File & LoadPicture_Stream for more info
    
    If Not m_Handle = 0 Then                ' do we have an image to copy?
        If Not cDIBclass Is Nothing Then    ' was a valid ref passed?
            With cDIBclass
                .DestroyDIB                 ' kill ref's old DIB if any
                .InitializeDIB m_Width, m_Height    ' create new one & copy data
                    If .BitsPointer = 0 Then Stop
                CopyMemory ByVal .BitsPointer, ByVal m_Pointer, m_Width * 4& * m_Height
                .Alpha = m_AlphaImage
                .ImageType = m_Format
                If CopyOriginalFormat = True Then
                    If Not iparseIsArrayEmpty(Not m_ImageBytes) Then
                        Call .SetOriginalFormat(m_ImageBytes)
                    End If
                End If
            End With
        End If
    End If
    
End Sub

Public Function GetOrginalFormat(outStream() As Byte) As Boolean

    ' If SaveFormat is true when LoadPicture_Stream or LoadPicture_File was
    ' called the original bytes were cached when the image was successfully
    ' loaded. Call this to return those original bytes

    ' If there are no original bytes, the function returns False
    
    If Not iparseIsArrayEmpty(Not m_ImageBytes) Then
        outStream() = m_ImageBytes()
        GetOrginalFormat = True
    End If

End Function

Friend Sub SetOriginalFormat(inStream() As Byte)

    ' Purpose: Pass the original image file/bytes to this DIB from another DIB
    ' This is only called by the CopyTo function.
    
    m_ImageBytes() = inStream()

End Sub

Private Function LoadPictureEx(FileName As String, aStream() As Byte, _
                            cX As Long, cy As Long, _
                            streamOffset As Long, streamLength As Long, _
                            Optional SaveFormat As Boolean = False) As Boolean
    
    ' PURPOSE: Marshal passed file/array to image classes for conversion to 32bpp image
    ' For parameter information, see LoadPicture_File & LoadPicture_Stream
    
    Me.DestroyDIB
    
    If Not FileName = vbNullString Then ' file name was passed
        
        Dim FileNum As Integer
        On Error Resume Next
        
        FileNum = FreeFile()    ' attempt to open file with read access only
        Open FileName For Binary Access Read As #FileNum
        ' if successful, we will use it later
        If Err Then
            Close #FileNum
            Err.Clear
            Exit Function
        End If
    End If
    
    ' various image parsers, in order of precedence
    ' All 4 recognize transparency
    Dim cPNG As cPNGparser  ' very fast to abort if not a PNG file
    Dim cGIF As cGIFparser  ' very fast to abort if not a GIF file
    Dim cICO As cICOparser  ' must parse key parts of a file. handles icons & Vista PNG Icons
    Dim cBMP As cBMPparser  ' catchall. Handles bitmaps, wmf, emf & jpgs
    
    Dim bReturn As Boolean  ' function return value
    
    ' validate passed desired icon sizes
    If cX < 0 Then cX = 0
    If cy < 0 Then cy = 0
    
    Set cPNG = New cPNGparser   ' see if image is a PNG; aborts quickly if not
    If FileName = vbNullString Then
        bReturn = cPNG.LoadStream(aStream(), Me, streamOffset, streamLength)
    Else     ' note: processing from file is slightly faster than via array
        bReturn = cPNG.LoadFile(FileName, Me)
        If bReturn = True Then Close #FileNum         ' close the file
    End If
    Set cPNG = Nothing
    If Not bReturn Then
        If Not FileName = vbNullString Then
            streamOffset = 0&
            streamLength = LOF(FileNum) ' cache length of file
            ReDim aStream(streamOffset To streamLength - 1)
            Get #FileNum, , aStream()   ' populate our stream with the file contents
            Close #FileNum
        End If
        Set cGIF = New cGIFparser ' what about a GIF; aborts quickly if not
        bReturn = cGIF.LoadStream(aStream(), Me, streamOffset, streamLength)
        Set cGIF = Nothing
        If Not bReturn Then
            Set cICO = New cICOparser   ' will process Vista PNG icon if needed
            bReturn = cICO.LoadStream(aStream(), cX, cy, Me, streamOffset, streamLength)
            Set cICO = Nothing
            If Not bReturn Then ' check for bmp, emf, wmf & jpg << last chance
                Set cBMP = New cBMPparser
                bReturn = cBMP.LoadStream(aStream(), Me, streamOffset, streamLength)
                Set cBMP = Nothing
            End If
        End If
    End If
    If Not m_Handle = 0 Then
        If SaveFormat = True Then ' we will cache the original bytes
            If iparseIsArrayEmpty(Not aStream) = True And Not FileName = vbNullString Then
                ' we loaded the image from the file and not a stream, need to get stream
                FileNum = FreeFile()
                ReDim m_ImageBytes(0 To LOF(FileNum - 1))
                Open FileName For Binary Access Read As #FileNum
                Get #FileNum, 1, FileName
                Close #FileNum
            Else
                m_ImageBytes() = aStream()
            End If
        End If
        LoadPictureEx = True
    End If
    
End Function

Friend Property Let Alpha(isAlpha As Boolean)
    m_AlphaImage = isAlpha  ' determines the flags used for AlphaBlend API
    ' this flag is set by the various image parsers; setting it yourself
    ' can produce less than desirable effects.
    ' Used in .Render, cPNGparser.SaveTo & .isWin95Alpha
End Property
Public Property Get Alpha() As Boolean
    Alpha = m_AlphaImage
End Property

Public Property Get ImageType() As eImageFormat
    ImageType = m_Format    ' returns image format of the source image
End Property
Friend Property Let ImageType(iType As eImageFormat)
    m_Format = iType    ' set by the various image parsers. This is not used
    ' anywhere in these classes, you can do with it what you want -- for now.
End Property

Public Property Get Width() As Long
    Width = m_Width     ' width of image in pixels
End Property
Public Property Get Height() As Long
    Height = m_Height   ' height of image in pixels
End Property
Public Property Get BitsPointer() As Long
    BitsPointer = m_Pointer ' pointer to the bits of the image
End Property
Public Property Get ScanWidth() As Long
    ScanWidth = m_Width * 4&    ' number of bytes per scan line
End Property
Public Property Get Handle() As Long
    Handle = m_Handle   ' the picture handle of the image
End Property

Public Function LoadDIBinDC(ByVal bLoad As Boolean) As Long

    ' Purpose: Select/Unselect the DIB into a DC.
    ' Returns the DC handle when image is loaded
    ' Called by image parser if it needs to paint the image into the DIB
       
    If bLoad = True Then
        Dim tDC As Long
        If Not m_Handle = 0 Then    ' do we have an image?
            If m_hDC = 0 Then       ' do we have a DC?
                tDC = GetDC(0&)     ' if not create one
                m_hDC = CreateCompatibleDC(tDC)
                ReleaseDC 0&, tDC
            End If
            m_prevObj = SelectObject(m_hDC, m_Handle)
            LoadDIBinDC = m_hDC
        End If
    Else
        If Not m_prevObj = 0 Then
            SelectObject m_hDC, m_prevObj
            If m_ManageDC = False Then
                DeleteObject m_hDC
                m_hDC = 0
            End If
            m_prevObj = 0
        End If
    End If
End Function

Public Property Let ManageOwnDC(bManage As Boolean)
    ' Determines whether or not this class will manage its own DC
    ' If false, then a DC is created each time the image needs to be Rendered
    Dim tDC As Long
    If bManage = False Then     ' removing management of DC
        If Not m_hDC = 0 Then   ' DC does exist, destroy it
            ' first remove the dib, if one exists
            If Not m_Handle = 0 Then SelectObject m_hDC, m_prevObj
            m_prevObj = 0
        End If
        DeleteDC m_hDC
        m_hDC = 0
    Else                        ' allowing creation of dc
        If m_hDC = 0 Then       ' create DC only if we have a dib to put in it
            If Not m_Handle = 0 Then
                tDC = GetDC(0&)
                m_hDC = CreateCompatibleDC(tDC)
                ReleaseDC 0&, tDC
            End If
        End If
    End If
    m_ManageDC = bManage
End Property
Public Property Get ManageOwnDC() As Boolean
    ManageOwnDC = m_ManageDC
End Property

Public Property Get isWin95Alpha() As Boolean
    ' available if you wish to render alpha image manually.
    isWin95Alpha = ((m_os9x = 1) And m_AlphaImage = True)
    ' if property returns true & you will process the image for Win95, then...
    ' 1) Call BitsPointer to retrieve the pre-multiplied bits pointer
    ' 2) Process the image for Win95 using following algorithm
    '   - If the DIBalphaByte (every 4th byte) = 0 then
    '       skip the blending for that pixel
    '   - Else (remember DIBs are BGR order, not RGB & DIB is bottom up
    '       If the DIBalphaByte = 255 then
    '           Set the destination pixel = DIB pixel, converting BGR to RGB as needed
    '       Else
    '       - Set lAlpha = 255&-DIBalphaByte to determine alpha of the destination pixel
    '       - Then use following formula
    '           desDCpixelRed = (lAlpha * desDCpixelRed) \ &HFF + DIBpixelRed
    '           desDCpixelBlue = (lAlpha * desDCpixelBlue) \ &HFF + DIBpixelBlue
    '           desDCpixelGreen = (lAlpha * desDCpixelGreen) \ &HFF + DIBpixelGreen
    
End Property


Public Function InitializeDIB(ByVal Width As Long, ByVal Height As Long) As Boolean

    ' Creates a blank (all black, all transparent) DIB of requested height & width
    
    Dim tBMPI As BITMAPINFO, tDC As Long
    
    DestroyDIB ' clear any pre-existing dib
    
    If Width < 0 Then Exit Function
    If Height < 0 Then
        Height = Abs(Height) ' no top-down dibs
    Else
        If Height = 0 Then Exit Function
    End If
    
    On Error Resume Next
    With tBMPI.bmiHeader
        .biBitCount = 32
        .biHeight = Height
        .biWidth = Width
        .biPlanes = 1
        .biSize = 40&
        .biSizeImage = .biHeight * .biWidth * 4&
    End With
    If Err Then
        Err.Clear
        ' only possible error would be that Width*Height*4& is absolutely huge
        Exit Function
    End If
    
    If m_hDC = 0 Then
        ' create a DC if class is managing its own & one isn't created yet
        tDC = GetDC(0&) ' get screen DC regardless
        If m_ManageDC = True Then m_hDC = CreateCompatibleDC(tDC)
    Else
        tDC = m_hDC ' use the class' DC, since we have it
    End If
    m_Handle = CreateDIBSection(tDC, tBMPI, 0, m_Pointer, 0, 0)
    ' release the screen DC if we captured it
    If Not tDC = m_hDC Then ReleaseDC 0&, tDC
    
    If Not m_Handle = 0 Then    ' let's hope system resources allowed DIB creation
        m_Width = Width
        m_Height = Height
        InitializeDIB = True
    End If

End Function

Public Sub DestroyDIB()
    
    ' PURPOSE: Destroy any existing image
    If Not m_hDC = 0 Then   ' do we have a DC?
        ' do we have an image; if so get it out of the DC
        If Not m_prevObj = 0 Then SelectObject m_hDC, m_prevObj
        ' destroy our DC, no point in keeping it w/o image
        DeleteObject m_hDC
        m_hDC = 0
    End If
    ' if we do have an image, destroy it now
    If Not m_Handle = 0 Then
        DeleteObject m_Handle
        Erase m_ImageBytes
    End If
    ' reset other image attributes
    m_Width = 0
    m_Height = 0
    m_Handle = 0
    m_Pointer = 0
    m_prevObj = 0
    m_AlphaImage = False
    m_Format = imgNone
End Sub

Public Function Render(ByVal destinationDC As Long, _
                Optional ByVal X As Long, Optional ByVal Y As Long, _
                Optional dX As Long, Optional ByVal dY As Long, _
                Optional ByVal Opacity As Long = 100) As Boolean

    ' PURPOSE: Render an existing 32bpp DIB to a target DC
    ' Note: Scaling, if needed, must be done before routine is called
    ' Not Win95 compatible for alpha images. See isWin95Alpha for more.
    
    ' Parameters.
    ' destinationDC :: target DC to draw to
    ' X, Y :: the top/left coordinates to draw to, default is 0,0
    ' dX, dY :: the width and height to draw to, default is the image's width & height
    ' Opacity :: how opaque to draw the image, default is 100% opaque
    '       -- not applicable if system is Win95

    Dim lBlendFunc As Long, tDC As Long, hOldImage As Long
    Dim lStretchMode As Long
    
    ' validate a few things
    If m_Handle = 0& Then
        Exit Function
    ElseIf destinationDC = 0& Then
        Exit Function
    End If
    
    If Opacity = 0 Then
        Render = True
        Exit Function   ' pointless if image is 100% transparent
    Else
        Opacity = Abs(Opacity) Mod 100
        If Opacity = 0 Then Opacity = 100
    End If
    
    If dX < 1 Then dX = m_Width
    If dY < 1 Then dY = m_Height
    
    If m_hDC = 0& Then  ' do we have a DC to select our image into?
        tDC = GetDC(0&) ' if not create one, if ManageOwnDC=True, we will have one
        m_hDC = CreateCompatibleDC(tDC)
        ReleaseDC 0&, tDC
        hOldImage = SelectObject(m_hDC, m_Handle)
    Else
        ' we have a DC, but is the image selected into it?
        If m_prevObj = 0 Then hOldImage = SelectObject(m_hDC, m_Handle)
    End If
    
    If m_os9x = 1 Then ' win95, can't use AlphaBlend & Opacity is N/A
        Render = Not (StretchBlt(destinationDC, X, Y, dX, dY, m_hDC, 0, 0, m_Width, m_Height, vbSrcInvert) = 0)  'vbSrcCopy
        
    Else
        ' Stretch_Halftone not compatible with win9x
        If m_os9x = 0 Then lStretchMode = SetStretchBltMode(destinationDC, STRETCH_HALFTONE)

        ' calcualte the opacity required & add it to the BlendFunction variable
        lBlendFunc = AC_SRC_OVER Or ((CLng(255 * (Opacity / 100))) * &H10000)
        ' if the image has transparency, then we add the AC_SRC_ALPHA flag too
        If m_AlphaImage = True Then lBlendFunc = lBlendFunc Or (AC_SRC_ALPHA * &H1000000)
        Render = Not (AlphaBlend(destinationDC, X, Y, dX, dY, m_hDC, 0, 0, m_Width, m_Height, lBlendFunc) = 0)
        
        If m_os9x = 0 Then SetStretchBltMode destinationDC, lStretchMode
    
    End If
    
    ' remove the image from the DC if necessary
    If Not hOldImage = 0 Then SelectObject m_hDC, hOldImage
    If Not tDC = 0& Then    ' if we created a DC, let's destroy it now
        DeleteDC m_hDC
        m_hDC = 0&
    End If
    
End Function

Public Function SaveToFile(FileName As String, Optional PromptOverwrite As Boolean = True) As Boolean

    ' Should you want to save a 32bpp image to a file
    ' Did you know? A 32bpp "XP-icon" saved in bitmap format is actually smaller
    '               than saving it in an icon format....
    '   BMP Format: 14byte header + 40byte BitmapInfo + 32bpp image bytes
    '   ICO Format: 22byte header + 40byte BitmapInfo + 32bpp image bytes + 1bpp mask bytes
    
    ' FileName :: full path & name of file to be created
    ' PromptOverwrite :: if True, the user will be offered an option to abort
    '    if the target file already exists
    
    If FileName = vbNullString Then Exit Function
    If m_Handle = 0 Then Exit Function
    
    On Error GoTo ExitRoutine
    If Len(Dir(FileName, vbArchive Or vbHidden Or vbReadOnly Or vbSystem)) > 0 Then
        If PromptOverwrite = True Then
            If MsgBox("Overwrite current file?", vbYesNo + vbDefaultButton2 + vbQuestion, "Overwrite Confirmation") = vbNo Then
                Exit Function
            End If
        End If
        SetAttr FileName, vbNormal ' in case the file was hidden, read only
        Kill FileName
    End If
    
    Dim FileNum As Integer, tBMPI As BITMAPINFO
    Dim dibBits() As Byte
    
    With tBMPI.bmiHeader
        .biHeight = m_Height
        .biPlanes = 1
        .biSize = 40
        .biWidth = m_Width
        .biBitCount = 32
        .biSizeImage = .biWidth * .biHeight * 4&
        ReDim dibBits(1 To .biSizeImage)
        CopyMemory dibBits(1), ByVal m_Pointer, .biSizeImage
    End With
    
    FileNum = FreeFile()
    Open FileName For Binary As #FileNum
    Put #FileNum, , CInt(&H4D42) ' bmp magic number
    Put #FileNum, , CLng(54& + tBMPI.bmiHeader.biSizeImage) ' overall size of image
    ' ^^ 54 = 14 byte bmp header + 40 for the tBMPI structure
    Put #FileNum, , 0& ' reserved bytes
    Put #FileNum, , 54& ' image offset from beginning of file
    Put #FileNum, , tBMPI
    Put #FileNum, , dibBits()
    Close #FileNum
    SaveToFile = True
    
ExitRoutine:
If Not FileNum = 0 Then Close #FileNum
If Err Then Err.Clear

End Function

Public Function SaveToStream(outStream() As Byte) As Boolean

    ' Should you want to serialize the 32bpp DIB
    
    On Error GoTo ExitRoutine   ' should out of memory occur?
    
    If m_Handle = 0 Then Exit Function
    
    Dim tBMPI As BITMAPINFO
    
    With tBMPI.bmiHeader
        .biBitCount = 32
        .biHeight = m_Height
        .biPlanes = 1
        .biSize = 40
        .biSizeImage = m_Width * m_Height * 4&
        .biWidth = m_Width
    End With
    
    ReDim outStream(0 To 54 + tBMPI.bmiHeader.biSizeImage - 1)
    
    CopyMemory outStream(0), &H4D42, 2& ' bmp magic number
    CopyMemory outStream(2), CLng(54 + tBMPI.bmiHeader.biSizeImage), 4& ' overall size of image
    ' ^^ 54 = 14 byte bmp header + 40 for the tBMPI structure
    CopyMemory outStream(10), 54&, 4& ' image offset from beginning of file
    CopyMemory outStream(14), tBMPI, 40&
    CopyMemory outStream(54), ByVal m_Pointer, tBMPI.bmiHeader.biSizeImage
    
    SaveToStream = True
    
ExitRoutine:
    If Err Then
        Err.Clear
        Erase outStream()
    End If
End Function

Private Function HandleToStdPicture(ByVal hImage As Long, ByVal imgType As Long) As IPicture

    ' function creates a stdPicture object from a image handle (bitmap or icon)
    ' Called by LoadPicture_ByHandle
    
    Dim lpPictDesc As PictDesc, aGUID(0 To 3) As Long
    With lpPictDesc
        .Size = Len(lpPictDesc)
        .Type = imgType
        .hHandle = hImage
        .hPal = 0
    End With
    ' IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    aGUID(0) = &H7BF80980
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
    ' create stdPicture
    Call OleCreatePictureIndirect(lpPictDesc, aGUID(0), True, HandleToStdPicture)
    
End Function

Private Sub Class_Initialize()

    ' Determine operating system for compatibility of 32bpp images
    
   Dim lValue As Long, osType As OSVERSIONINFOEX
   ' Retrieve version data for OS.
   osType.dwOSVersionInfoSize = Len(osType)
   If GetVersionEx(osType) = 0 Then
      ' The OSVERSIONINFOEX structure is only supported
      ' in NT4/SP6+ and NT5.x, so we're likely running
      ' on an earlier version of Windows. Revert structure
      ' size to OSVERSIONINFO and try again.
      osType.dwOSVersionInfoSize = Len(osType) - 8
      Call GetVersionEx(osType)
   End If
   
   ' Trim CSDVersion string at first null
   lValue = InStr(osType.szCSDVersion, vbNullChar)
   If lValue > 1 Then
      osType.szCSDVersion = Left(osType.szCSDVersion, lValue - 1)
   ElseIf lValue = 1 Then
      osType.szCSDVersion = ""
   End If
   
   If osType.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
        m_os9x = 1  ' this is a Win9x system
         ' get LoWord of the build number
        lValue = osType.dwBuildNumber
        If (lValue And &HFFFF&) > &H7FFF Then
            lValue = (lValue And &HFFFF&) - &H10000
        Else
            lValue = lValue And &HFFFF&
        End If
        If (osType.dwMinorVersion >= 10) Or _
           (lValue >= 3000) Then m_os9x = m_os9x Or 2 ' win98 or ME
    End If

End Sub

Private Sub Class_Terminate()
    DestroyDIB ' simply clean up
End Sub
