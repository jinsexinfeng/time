VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "RASZ1901时间显示"
   ClientHeight    =   2010
   ClientLeft      =   14850
   ClientTop       =   7350
   ClientWidth     =   3810
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2160
      Top             =   960
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   2760
      Top             =   960
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3360
      Top             =   960
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      DrawMode        =   7  'Invert
      Height          =   100
      Left            =   1680
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "本节课还剩:88:88"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   3495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "88"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "88"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "88:88:88"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim a As Integer
Dim b As Integer
Dim auto As Integer
Dim sx As Integer
Dim sy As Integer
Dim t2 As String
Dim t3 As String
Dim st As String
Dim p1 As String

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Dim BolIsMove As Boolean, MousX As Long, MousY As Long

'透明函数SetLayeredWindowAttributes
'使用bai这个函数，可以轻松的实现半透明窗体。按照微软的要求，透明窗体窗体在创建时应使用WS_EX_LAYERED参数（用CreateWindowEx），或者在创建后设置该参数（用SetWindowLong），我选用后者。全部函数、常量声明如下：
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
'其中hwnd是透明窗体的句柄，crKey为颜色值，bAlpha是透明度，取值范围是[0,255]，dwFlags是透明方式，可以取两个值：当取值为LWA_ALPHA时，crKey参数无效，bAlpha参数有效；当取值为LWA_COLORKEY时，bAlpha参数有效而窗体中的所有颜色为crKey的地方将变为透明－－这个功能很有用：我们不必再为建立不规则形状的窗体而调用一大堆区域分析、创建、合并函数了，只需指定透明处的颜色值即可，哈哈哈哈！请看具体代码。
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
'半透明代码结束
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
'用于将CreateRoundRectRgn创建的圆角区域赋给窗体
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
'用于创建一个圆角矩形，该矩形由X1，Y1-X2，Y2确定，并由X3，Y3确定的椭圆描述圆角弧度。
'参数 类型及说明：
'X1,Y1 Long，矩形左上角的X，Y坐标
'X2,Y2 Long，矩形右下角的X，Y坐标
'X3 Long，圆角椭圆的宽。其范围从0（没有圆角）到矩形宽（全圆）
'Y3 Long，圆角椭圆的高。其范围从0（没有圆角）到矩形高（全圆）
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'将CreateRoundRectRgn创建的区域删除，这是必要的，否则不必要的占用电脑内存
Dim outrgn As Long
'接下来声明一个全局变量,用来获得区域句柄
Private Sub Form_Activate() '窗体Activate()事件
Call rgnform(Me, 30, 30) '调用子过程
End Sub

Private Sub Form_Unload(Cancel As Integer) '窗体Unload事件
DeleteObject outrgn '将圆角区域使用的所有系统资源释放
End Sub
Private Sub rgnform(ByVal frmbox As Form, ByVal fw As Long, ByVal fh As Long) '子过程，改变参数fw和fh的值可实现圆角
Dim w As Long, h As Long
w = frmbox.ScaleX(frmbox.Width, vbTwips, vbPixels)
h = frmbox.ScaleY(frmbox.Height, vbTwips, vbPixels)
outrgn = CreateRoundRectRgn(0, 0, w, h, fw, fh)
Call SetWindowRgn(frmbox.hwnd, outrgn, True)
End Sub
'圆角结束

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then BolIsMove = True
MousX = X
MousY = Y
End Sub
 
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim CurrX As Long, CurrY As Long
If BolIsMove Then
 CurrX = Me.Left - MousX + X
 CurrY = Me.Top - MousY + Y
 Me.Move CurrX, CurrY
End If
End Sub
 
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
BolIsMove = False
End Sub
'鼠标移动代码

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If t2 <> "e" Then Cancel = True
'取消关闭
End Sub

Private Sub Form_Load()
'窗口置顶代码
SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 2 Or 1
sx = Screen.Width
sy = Screen.Height
b = 0
st = 1
auto = 1
Form1.BackColor = RGB(228, 228, 229)
'半透明
Dim rtn As Long
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 230, LWA_ALPHA
End Sub

Private Sub Label1_Click()
Form1.Left = 18400 / 19200 * sx
'点击缩小
End Sub

Private Sub Label2_Click()
Form1.Left = 14500 / 19200 * sx
'点击放大
End Sub

Private Sub Label3_Click()
Form1.Left = 14500 / 19200 * sx
'点击放大
End Sub

Private Sub Label4_DblClick()
t2 = Val(i \ 60 + 1)
If t2 > 10 Then
t2 = InputBox("请输入设定时间（分）" & Chr(13) & "RASZ1901JZB制作", "时间设定", t2)
Else
t2 = InputBox("请输入设定时间（分）" & Chr(13) & "RASZ1901JZB制作", "时间设定", "40")
End If
If t2 = "s" Then
st = 0
Exit Sub
End If
If t2 = "e" Then
Unload Form1
Exit Sub
End If
If t2 = "a" Then
If auto = 1 Then
auto = 0
MsgBox ("已关闭自动时间设置")
Else
auto = 1
MsgBox ("已开启自动时间设置")
End If
Exit Sub
End If
For z = 1 To Len(t2)
t3 = Mid(t2, z, 1)
If Asc(t3) > 57 Or Asc(t3) < 48 Then
MsgBox ("输入错误")
Exit Sub
End If
Next z
If t2 <> "" Then
i = Val(t2) * 60
st = 1
End If
'自定义定时
End Sub

Private Sub Timer1_Timer()
t = Time()
Label1.Caption = t
If Len(t) = 7 Then t = 0 & t
Label2.Caption = Mid(t, 1, 2)
Label3.Caption = Mid(t, 4, 2)
'时间显示
If y0 <> Form1.Top Or x0 <> Form1.Left Then
If Form1.Top > 8400 / 10800 * sy Then
Form1.Top = 8400 / 10800 * sy
End If
If Form1.Left > 17000 / 19200 * sx Then
Label1.Visible = False
Form1.Width = 850
Form1.Left = sx - 950
Label2.Visible = True
Label3.Visible = True
Label4.Alignment = 0
a = 1
End If
If Form1.Left < 16990 / 19200 * sx Then
Label1.Visible = True
Form1.Width = 3900
Label2.Visible = False
Label3.Visible = False
Label4.Alignment = 2
a = 0
End If
y0 = Form1.Top
x0 = Form1.Left
End If
'缩小放大参数调整
If st = 0 Then
Timer2.Enabled = False
Timer3.Enabled = False
If a = 0 Then Label4.Caption = "已禁用下课提醒"
If a = 1 Then Label4.Caption = "禁"
Else
Timer2.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
If auto <> 0 Then
If Time() = "7:25:00" Or Time() = "8:15:00" Or Time() = "9:20:00" Or Time() = "10:10:00" Or Time() = "11:05:00" Or Time() = "13:45:00" Or Time() = "14:40:00" Or Time() = "15:30:00" Or Time() = "16:20:00" Then
i = 2400
End If
If Time() = "18:00:00" Then
i = 6300
End If
If Time() = "20:00:00" Then
i = 5400
End If
End If
'开始上课时定时
i = i - 1
If i >= 0 Then
Else
If a = 0 Then Label4.Caption = "提示:非课堂时间"
If a = 1 Then Label4.Caption = "停"
Exit Sub
End If
If i > 600 Or i <= 0 Then
If a = 0 Then
p1 = "本节课还剩:" & i \ 60 & ":" & i Mod 60
Label4.Caption = p1
End If
If a = 1 Then
Label4.Caption = i \ 60
End If
Timer3.Enabled = False
Else
Timer3.Enabled = True
End If
If i = 60 And a = 1 Then
Form1.Left = 14500 / 19200 * sx
End If
'课堂倒计时
End Sub

Private Sub Timer3_Timer()
If b = 0 Then
If a = 0 Then
p1 = "本节课还剩:" & i \ 60 & ":" & i Mod 60
Label4.Caption = p1
End If
If a = 1 Then
Label4.Caption = i \ 60
End If
b = 1
Else
Label4.Caption = ""
b = 0
End If
'最后时间闪烁
End Sub
