VERSION 5.00
Begin VB.Form frm_1 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "대진의 하루"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   Icon            =   "frm_1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frm_1.frx":8732
   ScaleHeight     =   6000
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Label lbl_re 
      BackStyle       =   0  '투명
      Height          =   495
      Left            =   120
      MouseIcon       =   "frm_1.frx":E30F
      MousePointer    =   99  '사용자 정의
      TabIndex        =   1
      Top             =   5340
      Width           =   1815
   End
   Begin VB.Label lbl 
      BackStyle       =   0  '투명
      Height          =   735
      Left            =   2640
      MouseIcon       =   "frm_1.frx":E619
      MousePointer    =   99  '사용자 정의
      TabIndex        =   0
      Top             =   2520
      Width           =   3180
   End
   Begin VB.Image img_1 
      Height          =   6000
      Left            =   0
      MouseIcon       =   "frm_1.frx":E923
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_1.frx":EC2D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8250
   End
End
Attribute VB_Name = "frm_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frm_1->frm_2
'1교시:컴퓨터구조

Dim pagenum


Private Sub Form_Activate()
If res1 = 2 Then
    img_1 = LoadPicture(App.Path & "\img\img1_02.jpg")
ElseIf res1 = 1 Then
    img_1 = LoadPicture(App.Path & "\img\img1_02_over.jpg")
ElseIf rt1 = 1 Then
    pagenum = 2
    rt1 = 0
Else
    img_1 = LoadPicture(App.Path & "\img\img1_00.jpg")
    pagenum = 1
End If
End Sub


Sub pagemove()

If res1 = 2 Then
    frm_2.Show
    frm_1.Hide
ElseIf res1 = 1 Then
Else
    Select Case pagenum
    Case 1: img_1 = LoadPicture(App.Path & "\img\img1_01_0.jpg"): pagenum = 2
: frm_1_stt.Show: frm_1.Hide
    Case 2: img_1 = LoadPicture(App.Path & "\img\img1_01_1.jpg"): pagenum = 3
    Case 3: img_1 = LoadPicture(App.Path & "\img\img1_01_2.jpg"): pagenum = 4
: img_1.MouseIcon = LoadPicture(App.Path & "\img\mus_normal.cur")
: lbl_re.MouseIcon = LoadPicture(App.Path & "\img\mus_normal.cur")
End Select
End If

End Sub

Private Sub img_1_Click()
Call pagemove
End Sub


Private Sub lbl_Click()
If pagenum >= 4 Then
Call sndplay("eff01.wav")
pagenum = 0
img_1.MouseIcon = LoadPicture(App.Path & "\img\mus_click.cur")
lbl_re.MouseIcon = LoadPicture(App.Path & "\img\mus_click.cur")
frm_1_game.Show
frm_1.Hide
Else
Call pagemove
End If
End Sub

Private Sub lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If pagenum >= 4 Then
img_1 = LoadPicture(App.Path & "\img\img1_01_2_1.jpg")
End If
End Sub

Private Sub img_1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If pagenum >= 4 Then
img_1 = LoadPicture(App.Path & "\img\img1_01_2.jpg")
ElseIf res1 = 1 Then
img_1 = LoadPicture(App.Path & "\img\img1_02_over.jpg")
img_1.MouseIcon = LoadPicture(App.Path & "\img\mus_normal.cur")
lbl.MouseIcon = LoadPicture(App.Path & "\img\mus_normal.cur")
End If
End Sub

Private Sub lbl_re_Click()
If res1 = 1 Then
Call sndplay("eff01.wav")
res1 = 0
img_1.MouseIcon = LoadPicture(App.Path & "\img\mus_click.cur")
lbl.MouseIcon = LoadPicture(App.Path & "\img\mus_click.cur")
lbl_re.MouseIcon = LoadPicture(App.Path & "\img\mus_click.cur")
Call Form_Activate
Else
Call pagemove
End If
End Sub

Private Sub lbl_re_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If res1 = 1 Then
img_1 = LoadPicture(App.Path & "\img\img1_02_over_1.jpg")
End If
End Sub





