VERSION 5.00
Begin VB.Form Stage 
   BackColor       =   &H00800000&
   Caption         =   "Tunnel"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   ScaleHeight     =   330
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   488
   StartUpPosition =   3  'Windows Default
   Tag             =   "101"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Pict 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   1800
      ScaleHeight     =   79
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   95
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4800
      Top             =   4320
   End
End
Attribute VB_Name = "Stage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author        : Jan Tosovsky
'Email         : j.tosovsky@tiscali.cz
'Website       : http://nio.astronomy.cz
'Date          : 13 September 2003
'Version       : 1.0
'Description   : OpenGL Plasma tunnel

'This code was converted from Delphi source http://www.sulaco.co.za/
'Some improvements were made for loading bitmaps

'Program requires OpenGL Type Library from Patrice Scribe
'http://is6.pacific.net.hk/~edx/tlb.htm
'With library you needn't declare any used OpenGL functions or constants
'Copy library to system directory and then register it:
'regsvr32 "C:\Windows\System\vbogl.tlb" where path may vary
'In Project>References... in VB menu check item VB OpenGL API 1.2 (ANSI)

Option Explicit

Private Type glCoord
    X As GLfloat
    Y As GLfloat
    Z As GLfloat
End Type

Private Type glVertex
    U As GLfloat
    V As GLfloat
End Type

Dim angle As Single
Dim FPSCount As Single, DemoStart As Long, ElapsedTime As Long
Dim Tunnels(32, 32) As glCoord
Dim Texcoord(32, 32) As glVertex

Private Sub Form_Load()
    EnableOpenGL Stage.hDC
    DrawInit
End Sub

Private Sub DrawInit()
    glClearColor 0, 0, 0, 0
    glShadeModel GL_SMOOTH
    glClearDepth 1#
    glEnable GL_DEPTH_TEST
    glDepthFunc GL_LESS
    glHint GL_PERSPECTIVE_CORRECTION_HINT, GL_NICEST
    glEnable GL_TEXTURE_2D
    LoadGLTextures
    CreateTunnel
End Sub

Sub MainLoop()
    Dim lasttime As Long
    DemoStart = GetTickCount
    Do
        ElapsedTime = GetTickCount() - DemoStart
        ElapsedTime = (lasttime + ElapsedTime) \ 2
        angle = (ElapsedTime + lasttime) / 32
        lasttime = ElapsedTime
        DoEvents
        Render
        FPSCount = FPSCount + 1
    Loop
End Sub

Sub Render()
    Dim I As Long, J As Long, c As GLfloat
    glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
    glLoadIdentity
    
    glTranslatef 0#, 0#, -4
     
    For I = 0 To 32
        For J = 0 To 32
            Texcoord(I, J).U = I / 32 + Cos((angle + 8 * J) / 60) / 2 '     //cos((zrot + 8*J)/60)/2
            Texcoord(I, J).V = J / 32 + (angle + J) / 120             '     //cos((zrot + 8*J)/60)/4
        Next J
    Next I

    ' draw tunnel "cylinder"
    For J = 0 To 31
        If J > 24 Then
            c = 1# - (J - 24) / 10
        Else
            c = 1#
        End If
        glColor3f c, c, c
        glBegin GL_QUADS
            For I = 0 To 31
                glTexCoord2f Texcoord(I, J).U, Texcoord(I, J).V
                glVertex3f Tunnels(I, J).X, Tunnels(I, J).Y, Tunnels(I, J).Z
                glTexCoord2f Texcoord(I + 1, J).U, Texcoord(I + 1, J).V
                glVertex3f Tunnels(I + 1, J).X, Tunnels(I + 1, J).Y, Tunnels(I + 1, J).Z
                glTexCoord2f Texcoord(I + 1, J + 1).U, Texcoord(I + 1, J + 1).V
                glVertex3f Tunnels(I + 1, J + 1).X, Tunnels(I + 1, J + 1).Y, Tunnels(I + 1, J + 1).Z
                glTexCoord2f Texcoord(I, J + 1).U, Texcoord(I, J + 1).V
                glVertex3f Tunnels(I, J + 1).X, Tunnels(I, J + 1).Y, Tunnels(I, J + 1).Z
            Next I
        glEnd
    Next J
    DoEvents
    SwapBuffers Stage.hDC

End Sub

Sub CreateTunnel()
    Dim pi As Double, I As Long, J As Long
    pi = Atn(1) * 4
    For I = 0 To 32
        For J = 0 To 32
            Tunnels(I, J).X = (3 - J / 12) * Cos(2 * pi / 32 * I)
            Tunnels(I, J).Y = (3 - J / 12) * Sin(2 * pi / 32 * I)
            Tunnels(I, J).Z = -J
        Next J
    Next I
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then End
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    End
End Sub

Private Sub Form_Resize()
    Dim w As Long, h As Long
    w = Me.ScaleWidth: h = Me.ScaleHeight
    If h = 0 Then h = 1
    glViewport 0, 0, w, h
    glMatrixMode GL_PROJECTION
    glLoadIdentity
    gluPerspective 45#, w / h, 1#, 100#
    glMatrixMode GL_MODELVIEW
    glLoadIdentity
    DoEvents
    MainLoop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DisableOpenGL
End Sub

Private Sub Timer1_Timer()
    DoEvents
    Me.Caption = "Tunnel (" & FPSCount & " FPS)"
    FPSCount = 0
End Sub
