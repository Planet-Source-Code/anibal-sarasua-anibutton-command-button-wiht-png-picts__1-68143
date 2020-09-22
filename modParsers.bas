Attribute VB_Name = "modParsers"
Option Explicit

' No APIs are declared public. This is to prevent possibly, differently
' declared APIs or different versions, of the same API, from conflciting
' with any APIs you declared in your project. Same rule for UDTs.

Private Type SafeArrayBound
    cElements As Long
    lLbound As Long
End Type
Private Type SafeArray1D        ' used as DMA overlay on a DIB
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    rgSABound As SafeArrayBound
End Type
Private Type SAFEARRAY2D
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As Long
    rgSABound(0 To 1)  As SafeArrayBound
End Type

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

' used to create a stdPicture from a byte array
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GetHGlobalFromStream Lib "ole32" (ByVal ppstm As Long, hGlobal As Long) As Long
Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long

' GDI32 APIs
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ExtCreateRegion Lib "gdi32" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long
Private Declare Function GetRegionData Lib "gdi32.dll" (ByVal hRgn As Long, ByVal dwCount As Long, ByRef lpRgnData As Any) As Long
Private Declare Function GetRgnBox Lib "gdi32.dll" (ByVal hRgn As Long, ByRef lpRect As RECT) As Long
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
' User32 APIs
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Function iparseCreateShapedRegion(cHost As c32bppDIB) As Long

    '*******************************************************
    ' FUNCTION RETURNS A HANDLE TO A REGION IF SUCCESSFUL.
    ' If unsuccessful, function retuns zero.
    ' The fastest region from bitmap routines around, custom
    ' designed by LaVolpe. This version modified to create
    ' regions from alpha masks.
    '*******************************************************
    
    ' Useful should a region want to be created around the alpha image.
    ' Tip. Check c32bppDIB.Alpha property before calling this routine.
    '   -- If Alpha=False: When applying to a Window handle,
    '       then SetWindowRgn hWnd,0&,True uses no extra system resources
    '       vs setting the window region to a rectangular region
    
    ' declare bunch of variables...
    Dim rgnRects() As RECT ' array of rectangles comprising region
    Dim rectCount As Long ' number of rectangles & used to increment above array
    Dim rStart As Long ' pixel that begins a new regional rectangle
    
    Dim X As Long, Y As Long ' loop counters
    Dim lScanWidth As Long ' used to size the DIB bit array
    
    Dim bDib() As Byte  ' the DIB bit array
    Dim tSA As SAFEARRAY2D ' array overlay
    Dim rtnRegion As Long ' region handle returned by this function
    Dim Width As Long, Height As Long
    
    On Error GoTo CleanUp
      
    ' Simple sanity checks
    Width = cHost.Width
    If Width < 1& Then Exit Function
    Height = cHost.Height
    
    ' The following bottom up/top down checks are rem'd out
    ' unnecessary, the host DIB should always be bottom up
    ' but left in should you want to adapt the code for another purpose
'    If Height < 1& Then
'        If Height = 0& Then Exit Function
'        bIsTopDownDIB = (Height < 0&)
'        Height = Abs(Height) ' per MSDN, it is possible a DIB could be top down
'                             ' however, typical formatting is bottom up
'    End If
    
    If cHost.Alpha = False Then
        iparseCreateShapedRegion = CreateRectRgn(0&, 0&, Width, Height)
        Exit Function
    End If
    
    lScanWidth = Width * 4& ' how many bytes per bitmap line?
    With tSA                ' prepare array overlay
        .cbElements = 1     ' byte elements
        .cDims = 2          ' two dim array
        .pvData = cHost.BitsPointer  ' data location
        .rgSABound(0).cElements = Height
        .rgSABound(1).cElements = lScanWidth
    End With
    ' overlay now
    CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
    
     ' start with an arbritray number of rectangles
    ReDim rgnRects(0 To Width * 3&)
    ' reset flag
    rStart = -1&
    
    ' begin pixel by pixel comparisons
    For Y = 0& To Height - 1&
        ' the alpha byte is every 4th byte
        For X = 3& To lScanWidth - 1& Step 4&
        
            ' test to see if next pixel is 100% transparent
            If bDib(X, Y) = 0 Then
                If rStart > -1& Then ' we're currently tracking a rectangle,
                                    ' so let's close it
                    ' see if array needs to be resized
                   If rectCount + 1& = UBound(rgnRects) Then _
                       ReDim Preserve rgnRects(0 To UBound(rgnRects) + Width * 3&)
                    
                    ' add the rectangle to our array
'                    If bIsTopDownDIB Then
'                        SetRect rgnRects(rectCount + 2&), rStart, Y, X \ 4, Y + 1
'                    Else
                        SetRect rgnRects(rectCount + 2&), rStart, Height - Y - 1&, X \ 4, Height - Y
'                    End If
                    rStart = -1& ' reset flag
                    rectCount = rectCount + 1&     ' keep track of nr in use
                End If
            
            Else
                ' not a target color
                If rStart = -1& Then rStart = X \ 4 ' set start point
            
            End If
        Next X
        If rStart > -1& Then
            ' got to end of bitmap without hitting another transparent pixel
            ' but we're tracking so we'll close rectangle now
           
                ' see if array needs to be resized
           If rectCount + 1& = UBound(rgnRects) Then _
               ReDim Preserve rgnRects(0 To UBound(rgnRects) + Width * 3&)
                ' add the rectangle to our array
'            If bIsTopDownDIB Then
'                SetRect rgnRects(rectCount + 2), rStart, Y, X \ 4, Y + 1
'            Else
                SetRect rgnRects(rectCount + 2), rStart, Height - Y - 1, X \ 4, Height - Y
'            End If
            rStart = -1& ' reset flag
            rectCount = rectCount + 1&     ' keep track of nr in use
        End If
    Next Y

    ' remove the array overlay
    CopyMemory ByVal VarPtrArray(bDib()), 0&, 4&
        
    On Error Resume Next
    ' check for failure & engage backup plan if needed
    If Not rectCount = 0 Then
        ' there were rectangles identified, try to create the region in one step
        rtnRegion = CreatePartialRegion(rgnRects(), 2&, rectCount + 1&, 0&, Width)
        
        ' ok, now to test whether or not we are good to go...
        ' if less than 2000 rectangles, region should have been created & if it didn't
        ' it wasn't due O/S restrictions -- failure
        If rtnRegion = 0& Then
            If rectCount > 2000& Then
                ' Win98 has limitation of approximately 4000 regional rectangles
                ' In cases of failure, we will create the region in steps of
                ' 2000 vs trying to create the region in one step
                rtnRegion = CreateWin98Region(rgnRects, rectCount + 1&, 0&, Width)
            End If
        End If
    End If

CleanUp:
    Erase rgnRects()
    
    If Err Then ' failure; probably low on resources
        If Not rtnRegion = 0& Then DeleteObject rtnRegion
        Err.Clear
    Else
        iparseCreateShapedRegion = rtnRegion
    End If


End Function

Private Function CreatePartialRegion(rgnRects() As RECT, lIndex As Long, uIndex As Long, leftOffset As Long, cX As Long) As Long
    ' Helper function for CreateShapedRegion & CreateWin98Region
    ' Called to create a region in its entirety or stepped (see CreateWin98Region)

    On Error Resume Next
    ' Note: Ideally contiguous rectangles of equal height & width should be combined
    ' into one larger rectangle. However, thru trial & error I found that Windows
    ' does this for us and taking the extra time to do it ourselves
    ' is too cumbersome & slows down the results.
    
    ' the first 32 bytes of a region is the header describing the region.
    ' Well, 32 bytes equates to 2 rectangles (16 bytes each), so I'll
    ' cheat a little & use rectangles to store the header
    With rgnRects(lIndex - 2) ' bytes 0-15
        .Left = 32&                     ' length of region header in bytes
        .Top = 1&                       ' required cannot be anything else
        .Right = uIndex - lIndex + 1&   ' number of rectangles for the region
        .Bottom = .Right * 16&          ' byte size used by the rectangles; can be zero
    End With
    With rgnRects(lIndex - 1&) ' bytes 16-31 bounding rectangle identification
        .Left = leftOffset                  ' left
        .Top = rgnRects(lIndex).Top         ' top
        .Right = leftOffset + cX            ' right
        .Bottom = rgnRects(uIndex).Bottom   ' bottom
    End With
    ' call function to create region from our byte (RECT) array
    CreatePartialRegion = ExtCreateRegion(ByVal 0&, (rgnRects(lIndex - 2&).Right + 2&) * 16&, rgnRects(lIndex - 2&))
    If Err Then Err.Clear

End Function

Private Function CreateWin98Region(rgnRects() As RECT, rectCount As Long, leftOffset As Long, cX As Long) As Long
    ' Fall-back routine when a very large region fails to be created.
    ' Win98 has problems with regional rectangles over 4000
    ' So, we'll try again in case this is the prob with other systems too.
    ' We'll step it at 2000 at a time which is stil very quick

    Dim X As Long, Y As Long ' loop counters
    Dim win98Rgn As Long     ' partial region
    Dim rtnRegion As Long    ' combined region & return value of this function
    Const RGN_OR As Long = 2&
    Const scanSize As Long = 2000&

    ' we start with 2 'cause first 2 RECTs are the header
    For X = 2& To rectCount Step scanSize
    
        If X + scanSize > rectCount Then
            Y = rectCount
        Else
            Y = X + scanSize
        End If
        
        ' attempt to create partial region, scanSize rects at a time
        win98Rgn = CreatePartialRegion(rgnRects(), X, Y, leftOffset, cX)
        
        If win98Rgn = 0& Then    ' failure
            ' cleaup combined region if needed
            If Not rtnRegion = 0& Then DeleteObject rtnRegion
            Exit For ' abort; system won't allow us to create the region
        Else
            If rtnRegion = 0& Then ' first time thru
                rtnRegion = win98Rgn
            Else ' already started
                ' use combineRgn, but only every scanSize times
                CombineRgn rtnRegion, rtnRegion, win98Rgn, RGN_OR
                DeleteObject win98Rgn
            End If
        End If
    Next
    ' done; return result
    CreateWin98Region = rtnRegion
    
End Function

Public Function iparseIsArrayEmpty(ByVal lArrayPointer As Long) As Boolean
  ' test to see if an array has been initialized
  iparseIsArrayEmpty = (lArrayPointer = -1&)
End Function

Public Function iparseByteAlignOnWord(ByVal bitDepth As Byte, ByVal Width As Long) As Long
    ' function to align any bit depth on dWord boundaries
    iparseByteAlignOnWord = (((Width * bitDepth) + &H1F&) And Not &H1F&) \ &H8&
End Function

Public Function iparseCreateStdPicFromArray(inArray() As Byte, Offset As Long, Size As Long) As IPicture
    
    ' function creates a stdPicture from the passed array
    ' Note: The array was already validated as not empty when calling class' LoadStream was called
    
    Dim o_hMem  As Long
    Dim o_lpMem  As Long
    Dim aGUID(0 To 3) As Long
    Dim IIStream As IUnknown
    
    aGUID(0) = &H7BF80980    ' GUID for stdPicture
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
    
    o_hMem = GlobalAlloc(&H2&, Size)
    If Not o_hMem = 0& Then
        o_lpMem = GlobalLock(o_hMem)
        If Not o_lpMem = 0& Then
            CopyMemory ByVal o_lpMem, inArray(Offset), Size
            Call GlobalUnlock(o_hMem)
            If CreateStreamOnHGlobal(o_hMem, 1&, IIStream) = 0& Then
                  Call OleLoadPicture(ByVal ObjPtr(IIStream), 0&, 0&, aGUID(0), iparseCreateStdPicFromArray)
            End If
        End If
    End If

End Function

Public Function iparseReverseLong(ByVal inLong As Long) As Long

    ' fast function to reverse a long value from big endian to little endian
    ' PNG files contain reversed longs
    Dim b1 As Long
    Dim b2 As Long
    Dim b3 As Long
    Dim b4 As Long
    Dim lHighBit As Long
    
    lHighBit = inLong And &H80000000
    If lHighBit Then
      inLong = inLong And Not &H80000000
    End If
    
    b1 = inLong And &HFF
    b2 = (inLong And &HFF00&) \ &H100&
    b3 = (inLong And &HFF0000) \ &H10000
    If lHighBit Then
      b4 = inLong \ &H1000000 Or &H80&
    Else
      b4 = inLong \ &H1000000
    End If
    
    If b1 And &H80& Then
      iparseReverseLong = ((b1 And &H7F&) * &H1000000 Or &H80000000) Or _
          b2 * &H10000 Or b3 * &H100& Or b4
    Else
      iparseReverseLong = b1 * &H1000000 Or _
          b2 * &H10000 Or b3 * &H100& Or b4
    End If

End Function
