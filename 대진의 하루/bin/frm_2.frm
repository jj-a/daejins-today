VERSION 5.00
Begin VB.Form frm_2 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "대진의 하루"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   Icon            =   "frm_2.frx":0000
   LinkTopic       =   "Form2"
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
      MouseIcon       =   "frm_2.frx":8732
      MousePointer    =   99  '사용자 정의
      TabIndex        =   1
      Top             =   5340
      Width           =   1815
   End
   Begin VB.Label lbl 
      BackStyle       =   0  '투명
      Height          =   735
      Left            =   2520
      MouseIcon       =   "frm_2.frx":8A3C
      MousePointer    =   99  '사용자 정의
      TabIndex        =   0
      Top             =   2400
      Width           =   3180
   End
   Begin VB.Image img_2 
      Height          =   6000
      Left            =   0
      MouseIcon       =   "frm_2.frx":8D46
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_2.frx":9050
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8250
   End
End
Attribute VB_Name = "frm_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frm_2->frm_3
'2교시:체육

Dim pagenum


Private Sub Form_Activate()
If res2 = 2 Then
    img_2 = LoadPicture(App.Path & "\img\img2_02.jpg")
ElseIf res2 = 1 Then
    img_2 = LoadPicture(App.Path & "\img\img2_02_over.jpg")
    ElseIf rt2 = 1 Then
    pagenum = 2
    rt2 = 0
Else
    img_2 = LoadPicture(App.Path & "\img\img2_00.jpg")
    pagenum = 1
End If
End Sub


Sub pagemove()

If res2 = 2 Then
    frm_3.Show
    frm_2.Hide
ElseIf res2 = 1 Then
Else
    Select Case pagenum
    Case 1: img_2 = LoadPicture(App.Path & "\img\img2_01_0.jpg"): pagenum = 2
: frm_2_stt.Show: frm_2.Hide
    Case 2: img_2 = LoadPicture(App.Path & "\img\img2_01_1.jpg"): pagenum = 3
    Case 3: img_2 = LoadPicture(App.Path & "\img\img2_01_2.jpg"): pagenum = 4
    Case 4: img_2 = LoadPicture(App.Path & "\img\img2_01_3.jpg"): pagenum = 5
    Case 5: img_2 = LoadPicture(App.Path & "\img\img2_01_4.jpg"): pagenum = 6
    Case 6: img_2 = LoadPicture(App.Path & "\img\img2_01_5.jpg"): pagenum = 7
    Case 7: img_2 = LoadPicture(App.Path & "\img\img2_01_6.jpg"): pagenum = 8
    Case 8: img_2 = LoadPicture(App.Path & "\img\img2_01_7.jpg"): pagenum = 9
    Case 9: img_2 = LoadPicture(App.Path & "\img\img2_01_8.jpg"): pagenum = 10
: img_2.MouseIcon = LoadPicture(App.Path & "\img\mus_normal.cur")
: lbl_re.MouseIcon = LoadPicture(App.Path & "\img\mus_normal.cur")
End Select
End If

End Sub


Private Sub img_2_Click()
Call pagemove
End Sub


Private Sub lbl_Click()
If pagenum >= 10 Then
Call sndplay("eff01.wav")
pagenum = 0
img_2.MouseIcon = LoadPicture(App.Path & "\img\mus_click.cur")
lbl_re.MouseIcon = LoadPicture(App.Path & "\img\mus_click.cur")
frm_2_game.Show
frm_2.Hide
Else
Call pagemove
End If
End Sub

Private Sub lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If pagenum >= 10 Then
img_2 = LoadPicture(App.Path & "\img\img2_01_8_1.jpg")
End If
End Sub

Private Sub img_2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If pagenum >= 10 Then
img_2 = LoadPicture(App.Path & "\img\img2_01_8.jpg")
ElseIf res2 = 1 Then
img_2 = LoadPicture(App.Path & "\img\img2_02_over.jpg")
img_2.MouseIcon = LoadPicture(App.Path & "\img\mus_normal.cur")
lbl.MouseIcon = LoadPicture(App.Path & "\img\mus_normal.cur")
End If
End Sub


Private Sub lbl_re_Click()
If res2 = 1 Then
Call sndplay("eff01.wav")
res2 = 0
img_2.MouseIcon = LoadPicture(App.Path & "\img\mus_click.cur")
lbl.MouseIcon = LoadPicture(App.Path & "\img\mus_click.cur")
lbl_re.MouseIcon = LoadPicture(App.Path & "\img\mus_click.cur")
Call Form_Activate
Else
Call pagemove
End If
End Sub

Private Sub lbl_re_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If res2 = 1 Then
img_2 = LoadPicture(App.Path & "\img\img2_02_over_1.jpg")
End If
End Sub





