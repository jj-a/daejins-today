VERSION 5.00
Begin VB.Form frm_3 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "대진의 하루"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   Icon            =   "frm_3.frx":0000
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
      MouseIcon       =   "frm_3.frx":8732
      MousePointer    =   99  '사용자 정의
      TabIndex        =   1
      Top             =   5340
      Width           =   1815
   End
   Begin VB.Label lbl 
      BackStyle       =   0  '투명
      Height          =   735
      Left            =   2640
      MouseIcon       =   "frm_3.frx":8A3C
      MousePointer    =   99  '사용자 정의
      TabIndex        =   0
      Top             =   2520
      Width           =   3180
   End
   Begin VB.Image img_3 
      Height          =   6000
      Left            =   0
      MouseIcon       =   "frm_3.frx":8D46
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_3.frx":9050
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8250
   End
End
Attribute VB_Name = "frm_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frm_3->frm_4
'3교시:ASP

Dim pagenum


Private Sub Form_Activate()
If res3 = 2 Then
    img_3 = LoadPicture(App.Path & "\img\img3_02.jpg")
ElseIf res3 = 1 Then
    img_3 = LoadPicture(App.Path & "\img\img3_02_over.jpg")
ElseIf rt3 = 1 Then
    pagenum = 2
    rt3 = 0
Else
    img_3 = LoadPicture(App.Path & "\img\img3_00.jpg")
    pagenum = 1
End If
End Sub


Sub pagemove()

If res3 = 2 Then
    frm_4.Show
    frm_3.Hide
ElseIf res3 = 1 Then
Else
    Select Case pagenum
    Case 1: img_3 = LoadPicture(App.Path & "\img\img3_01_0.jpg"): pagenum = 2
: frm_3_stt.Show: frm_3.Hide
    Case 2: img_3 = LoadPicture(App.Path & "\img\img3_01_1.jpg"): pagenum = 3
    Case 3: img_3 = LoadPicture(App.Path & "\img\img3_01_2.jpg"): pagenum = 4
: img_3.MouseIcon = LoadPicture(App.Path & "\img\mus_normal.cur")
: lbl_re.MouseIcon = LoadPicture(App.Path & "\img\mus_normal.cur")
End Select
End If

End Sub

Private Sub img_3_Click()
Call pagemove
End Sub


Private Sub lbl_Click()
If pagenum >= 4 Then
Call sndplay("eff01.wav")
pagenum = 0
img_3.MouseIcon = LoadPicture(App.Path & "\img\mus_click.cur")
lbl_re.MouseIcon = LoadPicture(App.Path & "\img\mus_click.cur")
frm_3_game.Show
frm_3.Hide
Else
Call pagemove
End If
End Sub

Private Sub lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If pagenum >= 4 Then
img_3 = LoadPicture(App.Path & "\img\img3_01_2_1.jpg")
End If
End Sub

Private Sub img_3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If pagenum >= 4 Then
img_3 = LoadPicture(App.Path & "\img\img3_01_2.jpg")
ElseIf res3 = 1 Then
img_3 = LoadPicture(App.Path & "\img\img3_02_over.jpg")
img_3.MouseIcon = LoadPicture(App.Path & "\img\mus_normal.cur")
lbl.MouseIcon = LoadPicture(App.Path & "\img\mus_normal.cur")
End If
End Sub

Private Sub lbl_re_Click()
If res3 = 1 Then
Call sndplay("eff01.wav")
res3 = 0
img_3.MouseIcon = LoadPicture(App.Path & "\img\mus_click.cur")
lbl.MouseIcon = LoadPicture(App.Path & "\img\mus_click.cur")
lbl_re.MouseIcon = LoadPicture(App.Path & "\img\mus_click.cur")
Call Form_Activate
Else
Call pagemove
End If
End Sub

Private Sub lbl_re_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If res3 = 1 Then
img_3 = LoadPicture(App.Path & "\img\img3_02_over_1.jpg")
End If
End Sub
