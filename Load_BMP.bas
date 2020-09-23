Attribute VB_Name = "loadBMP"
Option Explicit
Dim w As Long, h As Long

Private Function loadBMP(ByVal Filename As String, ByRef TextureImg() As GLubyte) As Boolean
    ' The file should be BMP with pictures 64x64,128x128,256x256 .....
    Dim X As Long, Y As Long, temp As Long
    Dim bi24BitInfo As BITMAPINFO
    
    If Dir(Filename) = "" Then End
    
    'Loading Picture into PictureBox
    With Stage.Pict
        .Picture = LoadPicture(Filename)
        'with small modifications images can be obtained from resource too
        '.Picture = LoadResPicture(number) and parameter FileName can be removed
        .Refresh
        w = .ScaleWidth
        h = .ScaleHeight
        ' Create the array as needed for the image.
        ReDim TextureImg(2, w - 1, h - 1)
    End With
    
    'Getting data from PictureBox directly
    With bi24BitInfo.bmiHeader
        .biBitCount = 24
        .biCompression = 0 ' BI_RGB
        .biPlanes = 1
        .biSize = Len(bi24BitInfo.bmiHeader)
        .biWidth = w
        .biHeight = h
    End With
    GetDIBits Stage.Pict.hDC, Stage.Pict.Image, 0, h, TextureImg(0, 0, 0), bi24BitInfo, 0
    
    'Swap BGR->RGB
    For X = 0 To w - 1
        For Y = 0 To h - 1
        temp = TextureImg(0, X, Y)
        TextureImg(0, X, Y) = TextureImg(2, X, Y)
        TextureImg(2, X, Y) = temp
        Next Y
    Next X

    loadBMP = True

End Function

Function LoadGLTextures() As Boolean
    Dim Status As Boolean
    Dim TextureImage() As GLbyte
    
    Status = False    ' Status Indicator

    If loadBMP(App.Path & "\tunneltexture.BMP", TextureImage()) Then
        Status = True
        glTexEnvi GL_TEXTURE_ENV, GL_TEXTURE_ENV_MODE, GL_MODULATE  'Texture blends with object background
        glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
        glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR
        glTexImage2D GL_TEXTURE_2D, 0, 3, w, h, 0, GL_RGB, GL_UNSIGNED_BYTE, TextureImage(0, 0, 0)
        glTexImage2D GL_TEXTURE_2D, 0, 3, w, h, 0, GL_RGB, GL_UNSIGNED_BYTE, TextureImage(0, 0, 0)
    End If

    Erase TextureImage   ' Free the texture image memory
    LoadGLTextures = Status
End Function
