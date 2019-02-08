VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Image Processing Project"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Pic3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3600
      Left            =   10080
      ScaleHeight     =   238
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   318
      TabIndex        =   34
      Top             =   6240
      Width           =   4800
   End
   Begin VB.TextBox Center 
      Height          =   615
      Left            =   11880
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox B_Get 
      Height          =   375
      Left            =   6000
      MultiLine       =   -1  'True
      TabIndex        =   24
      Text            =   "Edge_detection.frx":0000
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox G_Get 
      Height          =   375
      Left            =   6000
      MultiLine       =   -1  'True
      TabIndex        =   23
      Text            =   "Edge_detection.frx":0006
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CheckBox inv_color 
      Caption         =   "Invert_color"
      Height          =   495
      Left            =   840
      TabIndex        =   22
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CheckBox BW 
      Caption         =   "BW_Image"
      Height          =   495
      Left            =   4920
      TabIndex        =   21
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CheckBox B_color 
      Caption         =   "B_color"
      Height          =   495
      Left            =   4920
      TabIndex        =   20
      Top             =   9120
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox G_color 
      Caption         =   "G_color"
      Height          =   495
      Left            =   4920
      TabIndex        =   19
      Top             =   8520
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox R_color 
      Caption         =   "R_color"
      Height          =   495
      Left            =   4920
      TabIndex        =   18
      Top             =   7920
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.ListBox lstEdgeOptions 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2505
      Left            =   720
      TabIndex        =   17
      Top             =   7800
      Width           =   2655
   End
   Begin VB.TextBox R_Get 
      Height          =   375
      Left            =   6000
      MultiLine       =   -1  'True
      TabIndex        =   15
      Text            =   "Edge_detection.frx":000C
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox Coordinate 
      Height          =   495
      Left            =   7560
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   7440
      Width           =   1815
   End
   Begin VB.TextBox RGB_1 
      Height          =   495
      Left            =   7560
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   8400
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   495
      Left            =   12000
      TabIndex        =   9
      Top             =   1800
      Width           =   1215
   End
   Begin VB.ListBox List 
      Height          =   1230
      Left            =   3480
      TabIndex        =   8
      Top             =   5040
      Width           =   1815
   End
   Begin VB.HScrollBar Brightness 
      Height          =   255
      Left            =   360
      Max             =   255
      TabIndex        =   6
      Top             =   6120
      Width           =   2415
   End
   Begin VB.HScrollBar Zoom 
      Height          =   255
      Left            =   360
      Max             =   3
      Min             =   1
      TabIndex        =   4
      Top             =   5040
      Value           =   1
      Width           =   2415
   End
   Begin VB.CommandButton Setting 
      Caption         =   "Setting"
      Height          =   495
      Left            =   12000
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton start 
      Caption         =   "Start"
      Height          =   495
      Left            =   12000
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3600
      Left            =   6480
      ScaleHeight     =   238
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   318
      TabIndex        =   1
      Top             =   600
      Width           =   4800
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3600
      Left            =   480
      ScaleHeight     =   238
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   318
      TabIndex        =   0
      Top             =   600
      Width           =   4800
   End
   Begin VB.Frame Frame1 
      Caption         =   "Zoom"
      Height          =   855
      Left            =   240
      TabIndex        =   5
      Top             =   4680
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      Caption         =   "Brightness"
      Height          =   855
      Left            =   240
      TabIndex        =   7
      Top             =   5760
      Width           =   2775
   End
   Begin VB.Frame Frame3 
      Caption         =   "RGB of Video"
      Height          =   855
      Left            =   7440
      TabIndex        =   11
      Top             =   8160
      Width           =   2055
   End
   Begin VB.Frame Frame4 
      Caption         =   "Application Box"
      Height          =   1935
      Left            =   3240
      TabIndex        =   12
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Frame Frame5 
      Caption         =   "Coordinate"
      Height          =   855
      Left            =   7440
      TabIndex        =   14
      Top             =   7200
      Width           =   2055
   End
   Begin VB.Frame Frame6 
      Caption         =   "Choosed RGB-1"
      Height          =   1695
      Left            =   5880
      TabIndex        =   16
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Frame Frame7 
      Caption         =   "Processed Video"
      Height          =   4215
      Left            =   240
      TabIndex        =   25
      Top             =   240
      Width           =   5295
   End
   Begin VB.Frame Frame8 
      Caption         =   "Source video"
      Height          =   4215
      Left            =   6240
      TabIndex        =   26
      Top             =   240
      Width           =   5295
   End
   Begin VB.Frame Frame9 
      Caption         =   "Color_Filter_Mode"
      Height          =   2775
      Left            =   4680
      TabIndex        =   27
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Frame Frame10 
      Caption         =   "Coustom_Filter_Mode"
      Height          =   3615
      Left            =   600
      TabIndex        =   28
      Top             =   6960
      Width           =   2895
   End
   Begin VB.TextBox B_Get2 
      Height          =   375
      Left            =   7800
      MultiLine       =   -1  'True
      TabIndex        =   29
      Text            =   "Edge_detection.frx":0012
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox G_Get2 
      Height          =   375
      Left            =   7800
      MultiLine       =   -1  'True
      TabIndex        =   30
      Text            =   "Edge_detection.frx":0018
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox R_Get2 
      Height          =   375
      Left            =   7800
      MultiLine       =   -1  'True
      TabIndex        =   31
      Text            =   "Edge_detection.frx":001E
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Frame Frame11 
      Caption         =   "Choosed RGB-2"
      Height          =   1695
      Left            =   7680
      TabIndex        =   32
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Frame Frame12 
      Caption         =   "Color  Seperating  Video"
      Height          =   4215
      Left            =   9840
      TabIndex        =   35
      Top             =   5880
      Width           =   5295
   End
   Begin VB.Frame Frame13 
      Caption         =   "Coordinate of Center of color"
      Height          =   975
      Left            =   11760
      TabIndex        =   36
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Frame Frame14 
      Height          =   975
      Left            =   10200
      TabIndex        =   37
      Top             =   4680
      Width           =   4455
      Begin VB.Label Label1 
         Caption         =   "Programing by :  S.E.Mirsadeghi ""Option Explicit is not an Option"""
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   38
         Top             =   120
         Width           =   3855
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
        x_pos As Long
        y_pos As Long
End Type

Const WM_CAP_DRIVER_CONNECT As Long = 1034
Const WM_CAP_DRIVER_DISCONNECT As Long = 1035
Const WM_CAP_GRAB_FRAME As Long = 1084
Const WM_CAP_EDIT_COPY As Long = 1054
Const WM_CAP_DLG_VIDEOFORMAT As Long = 1065
Const WM_CAP_DLG_VIDEOSOURCE As Long = 1066
Const WM_CLOSE = &H10


Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal Y As Long, _
ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal nID As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, _
ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long

Public mCapHwnd As Long
Dim I As Long, j As Long

Dim FM() As Long

Dim click As Boolean, invert As Boolean
Dim clk As Byte
Dim z As Integer
Dim Br As Byte
Dim RGB_G(1 To 3) As Integer
Dim Rg As Long, Gg As Long, Bg As Long
Dim Rg2 As Long, Gg2 As Long, Bg2 As Long
Dim minX As Integer
Dim maxX As Integer
Dim miny As Integer
Dim maxy As Integer
Dim minX1 As Integer
Dim maxX1 As Integer
Dim miny1 As Integer
Dim maxy1 As Integer
Dim Xcor
Dim Ycor



Private Sub Command1_Click()
    SendMessage mCapHwnd, WM_CAP_DRIVER_DISCONNECT, 0, 0
    SendMessage mCapHwnd, WM_CLOSE, 0, 0
    End
End Sub

Private Sub Form_Load()

    List.AddItem "Custom Filter"
    List.AddItem "Color detection"
    List.AddItem "Color Filter"
    List.AddItem "Color Tracking"
    List.AddItem "Negative"
    'List.AddItem "Other"
    'List.AddItem "Taking Photo"
    List.ListIndex = 2
    lstEdgeOptions.AddItem "Prewitt Horizontal"
    lstEdgeOptions.AddItem "Prewitt Vertical"
    lstEdgeOptions.AddItem "Sobel Horizontal"
    lstEdgeOptions.AddItem "Sobel Vertical"
    lstEdgeOptions.AddItem "Laplacian"
    lstEdgeOptions.AddItem "Hilite"
    lstEdgeOptions.AddItem "Canny"
    lstEdgeOptions.AddItem "DemonSpectre Linear"
    lstEdgeOptions.AddItem "DemonSpectre Cubic"
    lstEdgeOptions.AddItem "Emboss"
    lstEdgeOptions.AddItem "Sharpnen"
    lstEdgeOptions.AddItem "Custom"
    lstEdgeOptions.ListIndex = 4

        Dim Msg, Style, Title, Help, Ctxt, Response, MyString
        click = True
        z = 1
        Msg = "Do you want to see both source and processed video"
        Style = vbYesNo + vbQuestion + vbDefaultButton3
        Title = "Question?"
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
            
            If Response = vbYes Then
                    picBack.Visible = True
                Else
                    picBack.Visible = False
            End If

    mCapHwnd = capCreateCaptureWindow("IntraDream", 0, 0, 0, 0, 0, 0, 0)
    SendMessage mCapHwnd, WM_CAP_DRIVER_CONNECT, 0, 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
    SendMessage mCapHwnd, WM_CAP_DRIVER_DISCONNECT, 0, 0
    SendMessage mCapHwnd, WM_CLOSE, 0, 0
    End
End Sub




Private Sub Setting_Click()

SendMessage mCapHwnd, WM_CAP_DLG_VIDEOSOURCE, 0, 0
SendMessage mCapHwnd, WM_CAP_DLG_VIDEOFORMAT, 0, 0

End Sub

Private Sub start_Click()
    
    click = Not (click)
        If click = True Then
                start.Caption = " Processing..."
            Else
                start.Caption = " Start"
        End If

    While click = True
        On Error Resume Next
        SendMessage mCapHwnd, WM_CAP_GRAB_FRAME, 0, 0
        SendMessage mCapHwnd, WM_CAP_EDIT_COPY, 0, 0
        picBack.Picture = Clipboard.GetData
        'z = Zoom.Value
        
        Dim L As String
        Select Case List.Text
    
            Case "Custom Filter"
                    lstEdgeOptions.Enabled = True
                    invert = inv_color.Value
                    L = lstEdgeOptions.List(lstEdgeOptions.ListIndex)
            
            If L = "Prewitt Horizontal" Then
                
                    ReDim FM(-1 To 1, -1 To 1) As Long
                    FM(-1, -1) = -1
                    FM(-1, 0) = -1
                    FM(-1, 1) = -1
                    FM(1, -1) = 1
                    FM(1, 0) = 1
                    FM(1, 1) = 1
                    DoFilter Me.picBack, Me.picMain, 3, 1, 0, z, invert
            ElseIf L = "Prewitt Vertical" Then
        
                    ReDim FM(-1 To 1, -1 To 1) As Long
                    FM(-1, -1) = 1
                    FM(0, -1) = 1
                    FM(1, -1) = 1
                    FM(-1, 1) = -1
                    FM(0, 1) = -1
                    FM(1, 1) = -1
                    DoFilter Me.picBack, Me.picMain, 3, 1, 0, z, invert
            ElseIf L = "Sobel Horizontal" Then
                
                    ReDim FM(-1 To 1, -1 To 1) As Long
                    FM(-1, -1) = -1
                    FM(-1, 0) = -2
                    FM(-1, 1) = -1
                    FM(1, -1) = 1
                    FM(1, 0) = 2
                    FM(1, 1) = 1
                    DoFilter Me.picBack, Me.picMain, 3, 1, 0, z, invert
            ElseIf L = "Sobel Vertical" Then
                
                    ReDim FM(-1 To 1, -1 To 1) As Long
                    FM(-1, -1) = 1
                    FM(0, -1) = 2
                    FM(1, -1) = 1
                    FM(-1, 1) = -1
                    FM(0, 1) = -2
                    FM(1, 1) = -1
                    DoFilter Me.picBack, Me.picMain, 3, 1, 0, z, invert
            ElseIf L = "Laplacian" Then
                
                    ReDim FM(-1 To 1, -1 To 1) As Long
                    FM(-1, 0) = -1
                    FM(0, -1) = -1
                    FM(0, 1) = -1
                    FM(1, 0) = -1
                    FM(0, 0) = 4
                    DoFilter Me.picBack, Me.picMain, 3, 1, 0, z, invert
            ElseIf L = "Hilite" Then
                
                    ReDim FM(-1 To 1, -1 To 1) As Long
                    FM(-1, -1) = -4
                    FM(-1, 0) = -2
                    FM(0, -1) = -2
                    FM(1, -1) = -1
                    FM(-1, 1) = -1
                    FM(0, 0) = 10
                    DoFilter Me.picBack, Me.picMain, 3, 1, 0, z, invert
                    
            ElseIf L = "Canny" Then
                
                    ReDim FM(-1 To 1, -1 To 1) As Long
                    'FM(-2, -2) = 2
                    'FM(-2, -1) = 4
                    'FM(-2, 0) = 5
                    'FM(-2, 1) = 4
                    'FM(-2, 2) = 2
                    'FM(-1, -2) = 4
                    FM(-1, -1) = -5
                    'FM(-1, 0) = 12
                    'FM(-1, 1) = 9
                    'FM(-1, 2) = 4
                    'FM(0, -2) = 5
                    'FM(0, -1) = 12
                    'FM(0, 0) = 0
                    'FM(0, 1) = 12
                    'FM(0, 2) = 5
                    'FM(1, -2) = 4
                    'FM(1, -1) = 9
                    'FM(1, 0) = 12
                    FM(1, 1) = 5
                    'FM(1, 2) = 4
                    'FM(2, -2) = 2
                    'FM(2, -1) = 4
                    'FM(2, 0) = 5
                    'FM(2, 1) = 4
                    'FM(2, 2) = 2
                    DoFilter Me.picBack, Me.picMain, 3, 1, 0, z, invert
            ElseIf L = "DemonSpectre Linear" Then
                
                    ReDim FM(-1 To 1, -1 To 1) As Long
                    FM(-1, -1) = -1
                    FM(-1, 1) = -1
                    FM(1, -1) = -1
                    FM(1, 1) = -1
                    FM(0, 0) = 4
                    DoFilter Me.picBack, Me.picMain, 3, 1, 0, z, invert
            ElseIf L = "Emboss" Then
            
            ReDim FM(-1 To 1, -1 To 1) As Long
                    FM(0, -1) = -1
                    FM(0, 1) = 1
                    
                    DoFilter Me.picBack, Me.picMain, 3, 1, 127, z, invert
            
            ElseIf L = "Sharpnen" Then
            
            ReDim FM(-1 To 1, -1 To 1) As Long
                    FM(-1, 0) = -1
                    FM(0, -1) = -1
                    FM(0, 1) = -1
                    FM(1, 0) = -1
                    FM(0, 0) = 5
                    DoFilter Me.picBack, Me.picMain, 3, 1, 0, z, invert
            ElseIf L = "Custom" Then
                
                    ReDim FM(-1 To 1, -1 To 1) As Long
                    FM(0, 0) = 1
                    FM(0, 1) = 1
                    FM(1, 0) = 1
                    FM(1, 1) = 1
                    DoFilter Me.picBack, Me.picMain, 3, 10, 10, z, invert
            
            Else

                    ReDim FM(-2 To 2, -2 To 2) As Long
                    FM(-1, -2) = 1
                    FM(-2, 1) = 1
                    FM(1, 2) = 1
                    FM(2, -1) = 1
                    FM(0, 0) = -4
                    DoFilter Me.picBack, Me.picMain, 5, 1, 0, z, invert
            End If
                
            'Case "Taking Photo"
                'Photo.Picture = picBack.Picture
    
            Case "Color detection"
                color_detect picBack, picMain
                
            Case "Color Tracking"
                Color_Tracking picBack, picMain
            
            Case "Negative"
            'picMain.Cls
                'picMain.Picture = picBack.Picture
            'picMain.PaintPicture picBack.Picture, 0, 0, OpCode:=vbSrcErase
            picMain.PaintPicture picBack.Picture, 0, 0, OpCode:=vbNotSrcCopy
            'picMain.Cls
        'Dim mouse_pos As POINTAPI
        'GetCursorPos mouse_pos
        'Form1.Caption = mouse_pos.x_pos - 688 & " , " & mouse_pos.y_pos - 33 & " , " & mouse_pos.x_pos & " , " & mouse_pos.y_pos
            Case "Other"
                picMain.Cls
                picMain.PaintPicture picBack.Picture, 0, 1, OpCode:=vbNotSrcErase
            Case "Color Filter"
                Color_filter Me.picBack, Me.picMain
        
            End Select
       
        DoEvents
    Wend

End Sub
Private Sub picBack_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If List.Text = "Color detection" Or List.Text = "Color Tracking" Then
        If Button = 1 Then
        Dim pic As Long
        
            pic = GetPixel(picMain.hdc, x, Y)
            Rg = pic Mod 256
            Gg = (pic Mod (256 ^ 2)) / 256
            Bg = pic / (256 ^ 2)
        End If
        
        If Button = 2 Then
        
        pic = GetPixel(picMain.hdc, x, Y)
            Rg2 = pic Mod 256
            Gg2 = (pic Mod (256 ^ 2)) / 256
            Bg2 = pic / (256 ^ 2)
        End If
        'clk = clk + 1
        'If clk = 1 Then
            'RGB_G(1) = R
            'RGB_G(2) = G
            'RGB_G(3) = B
        'Else
        'color_detect picBack, picMain, Rg, Gg, Bg
            '& "    " & RGB_G(1) & " , " & RGB_G(2) & " , " & RGB_G(3)
            'clk = 0
        'end if
        
        R_Get.Text = Rg
        G_Get.Text = Gg
        B_Get.Text = Bg
        
        R_Get2.Text = Rg2
        G_Get2.Text = Gg2
        B_Get2.Text = Bg2
        
            
    End If
End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
    If List.Text = "Color detection" Or List.Text = "Color Filter" Or List.Text = "Negative" Then
    
        Dim R As Integer
        Dim G As Integer
        Dim B As Integer
        Dim pic As Long
    
            pic = GetPixel(picMain.hdc, x, Y)
            R = pic Mod 256
            G = (pic Mod (256 ^ 2)) / 256
            B = pic / (256 ^ 2)
        
        RGB_1.Text = R & " , " & G & " , " & B
        Coordinate.Text = x & " , " & Y
    
    End If
    
End Sub

Public Sub DoFilter(srcPic As PictureBox, dstPic As PictureBox, FilterSize As Long, FilterWeight As Long, FilterBias As Long, zoom_factor As Integer, Optional ByVal InvertResult As Boolean = False)

    Dim ImageData() As Byte
    Dim tData() As Byte
    Dim x As Long, Y As Long
    Dim iWidth As Long, iHeight As Long
    Dim fDraw As New FastDrawing
    Dim R As Long, G As Long, B As Long
    Dim c As Long, d As Long
    Dim CalcX As Long, CalcY As Long
    Dim CalcVar As Long
    Dim TempRef As Long
    Dim FilterWeightTemp As Long
    Dim QuickX As Long
    
    iWidth = fDraw.GetImageWidth(srcPic) - 1
    iHeight = fDraw.GetImageHeight(srcPic) - 1
    fDraw.GetImageData2D srcPic, ImageData()
    fDraw.GetImageData2D srcPic, tData()

    CalcVar = (FilterSize \ 2)

    For x = 0 To iWidth
        QuickX = x * 3
    For Y = 0 To iHeight

        R = 0
        G = 0
        B = 0
        FilterWeightTemp = FilterWeight

        For c = x - CalcVar To x + CalcVar
            TempRef = c * 3
        For d = Y - CalcVar To Y + CalcVar
            
            CalcX = c - x
            CalcY = d - Y

            If FM(CalcX, CalcY) = 0 Then GoTo 7

            If c < 0 Or d < 0 Or c > iWidth Or d > iHeight Then
                FilterWeightTemp = FilterWeight - FM(CalcX, CalcY)
                GoTo 7
            End If

            R = R + (ImageData(TempRef + 2, d) * FM(CalcX, CalcY))
            G = G + (ImageData(TempRef + 1, d) * FM(CalcX, CalcY))
            B = B + (ImageData(TempRef, d) * FM(CalcX, CalcY))

7       Next d
        Next c

        If FilterWeight <> 1 Then
            R = R \ FilterWeightTemp
            G = G \ FilterWeightTemp
            B = B \ FilterWeightTemp
        End If
        
        If FilterBias <> 0 Then
            R = R + FilterBias
            G = G + FilterBias
            B = B + FilterBias
        End If
        
        ByteMe R
        ByteMe G
        ByteMe B
        
        If InvertResult = True Then
            R = 255 - R
            G = 255 - G
            B = 255 - B
        End If
        
        tData(QuickX + 2, Y) = CByte(R)
        tData(QuickX + 1, Y) = CByte(G)
        tData(QuickX, Y) = CByte(B)
    
    Next Y
    Next x

    'picMain.Width = iWidth * zoom_factor
    'picMain.Height = iHeight * zoom_factor
    fDraw.SetImageData2D dstPic, iWidth, iHeight, tData()
    
End Sub

Public Sub Color_filter(srcPic As PictureBox, dstPic As PictureBox)

    Dim ImageData() As Byte
    Dim tData() As Byte
    Dim x As Long, Y As Long
    Dim iWidth As Long, iHeight As Long
    Dim fDraw As New FastDrawing
    Dim R As Long, G As Long, B As Long
    Dim QuickX As Long
    Dim Br As Integer
    Dim R_V As Integer, G_V As Integer, B_V As Integer

    iWidth = fDraw.GetImageWidth(srcPic) - 1
    iHeight = fDraw.GetImageHeight(srcPic) - 1
    fDraw.GetImageData2D srcPic, ImageData()
    fDraw.GetImageData2D srcPic, tData()
    
    R_V = R_color.Value
    G_V = G_color.Value
    B_V = B_color.Value

    For x = 0 To iWidth
        QuickX = x * 3
    For Y = 0 To iHeight
            
            R = (ImageData(QuickX + 2, Y))
            G = (ImageData(QuickX + 1, Y))
            B = (ImageData(QuickX, Y))
            
        'If List.Text = "Color Filter" Then
        'R_color.Enabled = True
        'G_color.Enabled = True
        'B_color.Enabled = True
        
        If R_V = 0 Then
            R = 0
        End If
            
        If G_V = 0 Then
            G = 0
        End If
        
        If B_V = 0 Then
            B = 0
        End If
        
        If R_V = 0 And G_V = 0 Then
                R = 0
                G = 0
        End If
        
        If R_V = 0 And B_V = 0 Then
                R = 0
                B = 0
        End If
        
        If G_V = 0 And B_V = 0 Then
                G = 0
                B = 0
        End If
        
        If BW.Value = 1 Then
            'R = R + 20
            'G = G + 20
            'B = B + 20
            'ByteMe R
            'ByteMe G
            'ByteMe B
            
            If R < 128 And G < 128 And B < 128 Then
                    
                    R = 0
                    G = 0
                    B = 0
            Else
                     
                    R = 255
                    G = 255
                    B = 255

        End If
    End If

        
        tData(QuickX + 2, Y) = R
        tData(QuickX + 1, Y) = G
        tData(QuickX, Y) = B
    
    Next Y
    Next x
    
    'picMain.Width = iWidth * z
    'picMain.Height = iHeight * z
    fDraw.SetImageData2D dstPic, iWidth, iHeight, tData()

End Sub


Public Sub color_detect(srcPic As PictureBox, dstPic As PictureBox)
        
Dim ImageData() As Byte
Dim tData() As Byte
Dim tData2() As Byte
Dim x As Long, Y As Long
Dim iWidth As Long, iHeight As Long
Dim fDraw As New FastDrawing
Dim R As Long, G As Long, B As Long
Dim QuickX As Long
Dim Xcenter As Integer
Dim Ycenter As Integer
Dim Xcenter1 As Integer
Dim Ycenter1 As Integer
    minX = 10000
    maxX = 0
    miny = 10000
    maxy = 0
    Xcor = 0
    Ycor = 0
    minX1 = 10000
    maxX1 = 0
    miny1 = 10000
    maxy1 = 0

Dim cnt As Long


    iWidth = fDraw.GetImageWidth(srcPic) - 1
    iHeight = fDraw.GetImageHeight(srcPic) - 1
    fDraw.GetImageData2D srcPic, ImageData()
    fDraw.GetImageData2D srcPic, tData()

    For x = 0 To iWidth
        QuickX = x * 3
    For Y = 0 To iHeight

            R = (ImageData(QuickX + 2, Y))
            G = (ImageData(QuickX + 1, Y))
            B = (ImageData(QuickX, Y))
            
            If Abs(R - Rg) < 40 And Abs(G - Gg) < 40 And Abs(B - Bg) < 40 Then
                R = 255
                G = 255
                B = 0
                    
                    'Pixel_num = Pixel_num + 1
                    'Xcor = x + Xcor
                    'Ycor = Y + Ycor
                    If x < minX Then
                        minX = x
                    End If
                    
                    If x > maxX Then
                        maxX = x
                    End If
                    
                    If Y < miny Then
                        miny = Y
                    End If
                    
                    If Y > maxy Then
                        maxy = Y
                    End If
                    
            End If
            
            If Abs(R - Rg2) < 30 And Abs(G - Gg2) < 30 And Abs(B - Bg2) < 30 Then
                R = 255
                G = 0
                B = 255
                    If x < minX1 Then
                        minX1 = x
                    End If
                    
                    If x > maxX1 Then
                        maxX1 = x
                    End If
                    
                    If Y < miny1 Then
                        miny1 = Y
                    End If
                    
                    If Y > maxy1 Then
                        maxy1 = Y
                    End If
            End If
 
        tData(QuickX + 2, Y) = CByte(R)
        tData(QuickX + 1, Y) = CByte(G)
        tData(QuickX, Y) = CByte(B)

    Next Y
    Next x
        miny = Abs(240 - miny)
    maxy = Abs(240 - maxy)
        miny1 = Abs(240 - miny1)
    maxy1 = Abs(240 - maxy1)
    Xcenter = (minX + maxX) / 2
    Ycenter = (miny + maxy) / 2
    Xcenter1 = (minX1 + maxX1) / 2
    Ycenter1 = (miny1 + maxy1) / 2
    picMain.DrawWidth = 4
    
    fDraw.SetImageData2D dstPic, iWidth, iHeight, tData()
    picMain.Line (Xcenter, Ycenter)-(Xcenter1, Ycenter1), vbRed
End Sub

Public Sub Color_Tracking(srcPic As PictureBox, dstPic As PictureBox)
picMain.Cls

Dim ImageData() As Byte
Dim tData() As Byte
Dim x As Long, Y As Long
Dim iWidth As Long, iHeight As Long
Dim fDraw As New FastDrawing
Dim R As Long, G As Long, B As Long
Dim QuickX As Long
Dim xg As Long, yg As Long
Dim xg2 As Long, yg2 As Long
Dim cnt As Long, cnt2 As Long
Dim I1 As Long, I2 As Long
Dim Xcenter As Integer
Dim Ycenter As Integer

'Dim Pixel_num
        minX = 10000
    maxX = 0
    miny = 10000
    maxy = 0
    Xcor = 0
    Ycor = 0

    iWidth = fDraw.GetImageWidth(srcPic) - 1
    iHeight = fDraw.GetImageHeight(srcPic) - 1
    fDraw.GetImageData2D srcPic, ImageData()
    fDraw.GetImageData2D srcPic, tData()

    For x = 0 To iWidth
        QuickX = x * 3
    For Y = 0 To iHeight

            R = (ImageData(QuickX + 2, Y))
            G = (ImageData(QuickX + 1, Y))
            B = (ImageData(QuickX, Y))
            
             'If Abs(R - Rg) < 30 And Abs(G - Gg) < 30 And Abs(B - Bg) < 30 Then
             
                'If R > G Then
                    'G = 0
                'End If
                
                'If R > B Then
                    'B = 0
                'End If
                
                'If G > R Then
                    'R = 0
                'End If
                
                'If G > B Then
                    'B = 0
                'End If
                
                'If B > R Then
                    'R = 0
                'End If
                
                'If B > G Then
                    'G = 0
                'End If
             
             'End If
            If Abs(R - Rg) < 10 And Abs(G - Gg) < 10 And Abs(B - Bg) < 10 Then
                    
                    'Pixel_num = Pixel_num + 1
                    'Xcor = x + Xcor
                    'Ycor = Y + Ycor
                    If x < minX Then
                        minX = x
                    End If
                    
                    If x > maxX Then
                        maxX = x
                    End If
                    
                    If Y < miny Then
                        miny = Y
                    End If
                    
                    If Y > maxy Then
                        maxy = Y
                    End If
                    
                End If
            
            'If Abs(R - Rg) < 30 And Abs(G - Gg) < 30 And Abs(B - Bg) < 30 Then
                'I1 = (Rg + Gg + Bg) / 3
                'xg = (I1 * X) + xg
                'yg = (I1 * Y) + yg
                'cnt = cnt + 1
                'R = 255
                'G = 255
                'B = 0
            'End If
            
            'If Abs(R - Rg2) < 30 And Abs(G - Gg2) < 30 And Abs(B - Bg2) < 30 Then
                'I2 = (Rg2 + Gg2 + Bg2) / 3
                'xg2 = (I2 * X) + xg2
                'yg2 = (I2 * Y) + yg2
                'cnt2 = cnt2 + 1
                ' R = 0
                'G = 255
                'B = 255
            'End If
        
        'tData(QuickX + 2, Y) = CByte(R)
        'tData(QuickX + 1, Y) = CByte(G)
        'tData(QuickX, Y) = CByte(B)
            
    
    Next Y
    Next x
    
    miny = Abs(240 - miny)
    maxy = Abs(240 - maxy)
    Xcenter = (minX + maxX) / 2
    Ycenter = (miny + maxy) / 2
    
    'xg = xg / (cnt * 150)
    'yg = yg / (cnt * 150)
    'xg2 = xg2 / (cnt2 * 150)
    'yg2 = yg2 / (cnt2 * 150)
    Center.Text = Xcenter & " , " & Ycenter
    picMain.DrawWidth = 4
    'picBack.Line (xg, yg)-(xg2, yg2), vbBlack
    fDraw.SetImageData2D dstPic, iWidth, iHeight, ImageData()
    picMain.Circle (Xcenter, Ycenter), 5, vbRed
    picMain.Line (minX, miny)-(minX, maxy), vbBlack
    picMain.Line (minX, miny)-(maxX, miny), vbBlack
    picMain.Line (maxX, miny)-(maxX, maxy), vbBlack
    picMain.Line (minX, maxy)-(maxX, maxy), vbBlack
    
End Sub
Public Sub ByteMe(ByRef TempVar As Long)
    If TempVar > 255 Then TempVar = 255
    If TempVar < 0 Then TempVar = 0
End Sub
