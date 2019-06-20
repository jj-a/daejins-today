VERSION 5.00
Begin VB.Form frm_8 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "대진의 하루"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   Icon            =   "frm_8.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Label lbl_re 
      BackStyle       =   0  '투명
      Height          =   495
      Left            =   120
      MouseIcon       =   "frm_8.frx":8732
      MousePointer    =   99  '사용자 정의
      TabIndex        =   1
      Top             =   5340
      Width           =   1815
   End
   Begin VB.Label lbl 
      BackStyle       =   0  '투명
      Height          =   735
      Left            =   2640
      MouseIcon       =   "frm_8.frx":8A3C
      MousePointer    =   99  '사용자 정의
      TabIndex        =   0
      Top             =   2520
      Width           =   3180
   End
   Begin VB.Image img_8 
      Height          =   6000
      Left            =   0
      MouseIcon       =   "frm_8.frx":8D46
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_8.frx":9050
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8250
   End
End
Attribute VB_Name = "frm_8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frm_8->frm_end
'8:종례

Dim pagenum


Private Sub Form_Activate()
If res8 = 2 Then
    img_8 = LoadPicture(App.Path & "\img\img8_02.jpg")
ElseIf res8 = 1 Then
    img_8 = LoadPicture(App.Path & "\img\img8_02_over.jpg")
ElseIf rt8 = 1 Then
    pagenum = 2
    rt8 = 0
Else
    img_8 = LoadPicture(App.Path & "\img\img8_00.jpg")
    pagenum = 1
End If
End Sub


Sub pagemove()

If res8 = 2 Then
    res8 = 0
    frm_ending.Show
    frm_8.Hide
ElseIf res8 = 1 Then
Else
    Select Case pagenum
    Case 1: img_8 = LoadPicture(App.Path & "\img\img8_01_0.jpg"): pagenum = 2
: frm_8_stt.Show: frm_8.Hide
    Case 2: img_8 = LoadPicture(App.Path & "\img\img8_01_1.jpg"): pagenum = 3
    Case 3: img_8 = LoadPicture(App.Path & "\img\img8_01_2.jpg"): pagenum = 4
: img_8.MouseIcon = LoadPicture(App.Path & "\img\mus_normal.cur")
: lbl_re.MouseIcon = LoadPicture(App.Path & "\img\mus_normal.cur")
End Select
End If

End Sub

Private Sub img_8_Click()
Call pagemove
End Sub


Private Sub lbl_Click()
If pagenum >= 4 Then
Call sndplay("eff01.wav")
pagenum = 0
img_8.MouseIcon = LoadPicture(App.Path & "\img\mus_click.cur")
lbl_re.MouseIcon = LoadPicture(App.Path & "\img\mus_click.cur")
frm_8_game.Show
frm_8.Hide
Else
Call pagemove
End If
End Sub

Private Sub lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If pagenum >= 4 Then
img_8 = LoadPicture(App.Path & "\img\img8_01_2_1.jpg")
End If
End Sub

Private Sub img_8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If pagenum >= 4 Then
img_8 = LoadPicture(App.Path & "\img\img8_01_2.jpg")
ElseIf res8 = 1 Then
img_8 = LoadPicture(App.Path & "\img\img8_02_over.jpg")
img_8.MouseIcon = LoadPicture(App.Path & "\img\mus_normal.cur")
lbl.MouseIcon = LoadPicture(App.Path & "\img\mus_normal.cur")
End If
End Sub

Private Sub lbl_re_Click()
If res8 = 1 Then
Call sndplay("eff01.wav")
res8 = 0
img_8.MouseIcon = LoadPicture(App.Path & "\img\mus_click.cur")
lbl.MouseIcon = LoadPicture(App.Path & "\img\mus_click.cur")
lbl_re.MouseIcon = LoadPicture(App.Path & "\img\mus_click.cur")
Call Form_Activate
Else
Call pagemove
End If
End Sub

Private Sub lbl_re_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If res8 = 1 Then
img_8 = LoadPicture(App.Path & "\img\img8_02_over_1.jpg")
End If
End Sub





