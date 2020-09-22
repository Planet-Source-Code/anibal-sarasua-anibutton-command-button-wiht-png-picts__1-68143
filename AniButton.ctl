VERSION 5.00
Begin VB.UserControl AniButton 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   PropertyPages   =   "AniButton.ctx":0000
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Timer OverTimer 
      Interval        =   3
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "AniButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum Tamano
    [32x] = 32
    [48x] = 48
    [64x] = 64
    [128x] = 128
    [256x] = 256
End Enum

Dim mCaption As String
Dim mImagen As c32bppDIB
Dim mTamano As Tamano
Property Get Tamaño() As Tamano
    Tamaño = mTamano
End Property
Property Let Tamaño(xTam As Tamano)
    mTamano = xTam
    PropertyChanged "Tamaño"
End Property
Property Get Imagen() As Byte()
    Dim mBytes() As Byte
    Test = mImagen.SaveToStream(mBytes())
    Imagen = mBytes()
End Property
Property Let Imagen(mB() As Byte)
    Call mImagen.LoadPicture_Stream(mB())
    PropertyChanged "Images"
End Property
Property Get Caption() As String
    Caption = mCaption
End Property
Property Let Caption(xCap As String)
    mCaption = xCap
    If mCaption <> "" Then lblCaption.Caption = mCaption
    PropertyChanged "Caption"
End Property

Private Sub UserControl_Initialize()
    Set mImagen = New c32bppDIB
End Sub

Private Sub UserControl_Paint()
   UserControl.Cls
   mImagen.Render UserControl.hDC, 10, 10, mTamano, mTamano
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   Dim mDatos() As Byte
   With PropBag
      mCaption = .ReadProperty("CAPTION", "")
      mTamano = .ReadProperty("Tamaño", 32)
      If mCaption <> "" Then lblCaption.Caption = mCaption
      mDatos() = .ReadProperty("Images")
      Call mImagen.LoadPicture_Stream(mDatos())
    End With

End Sub

Private Sub UserControl_Terminate()
    Set mImagen = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim mDatos() As Byte
    With PropBag
        .WriteProperty "CAPTION", mCaption, ""
        .WriteProperty "Tamaño", mTamano, 32
         Test = mImagen.SaveToStream(mDatos())
         'MsgBox ("Antes de grabar " & Str(UBound(mDatos)))
        .WriteProperty "Images", mDatos
    End With
End Sub

'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
    Caption = "UserControl1"
    'MsgBox ("Lo inicializo :-(")
    Call mImagen.LoadPicture_File(App.Path & "\" & "test.png")
End Sub

