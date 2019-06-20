VERSION 5.00
Begin VB.Form frm_8_game 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "대진의 하루"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   Icon            =   "frm_8_game.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7800
      Top             =   120
   End
   Begin VB.Label lbl 
      BackStyle       =   0  '투명
      Height          =   495
      Index           =   4
      Left            =   2400
      MouseIcon       =   "frm_8_game.frx":8732
      MousePointer    =   99  '사용자 정의
      TabIndex        =   4
      Top             =   3360
      Width           =   4935
   End
   Begin VB.Label lbl 
      BackStyle       =   0  '투명
      Height          =   495
      Index           =   3
      Left            =   2400
      MouseIcon       =   "frm_8_game.frx":8A3C
      MousePointer    =   99  '사용자 정의
      TabIndex        =   3
      Top             =   2760
      Width           =   4935
   End
   Begin VB.Label lbl 
      BackStyle       =   0  '투명
      Height          =   495
      Index           =   2
      Left            =   2400
      MouseIcon       =   "frm_8_game.frx":8D46
      MousePointer    =   99  '사용자 정의
      TabIndex        =   2
      Top             =   2160
      Width           =   4935
   End
   Begin VB.Label lbl 
      BackStyle       =   0  '투명
      Height          =   495
      Index           =   1
      Left            =   2400
      MouseIcon       =   "frm_8_game.frx":9050
      MousePointer    =   99  '사용자 정의
      TabIndex        =   1
      Top             =   1560
      Width           =   4935
   End
   Begin VB.Label lbl_q 
      BackStyle       =   0  '투명
      Height          =   615
      Left            =   2400
      MouseIcon       =   "frm_8_game.frx":935A
      MousePointer    =   99  '사용자 정의
      TabIndex        =   0
      Top             =   360
      Width           =   4935
   End
   Begin VB.Image img 
      Height          =   6000
      Left            =   0
      MouseIcon       =   "frm_8_game.frx":9C24
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_8_game.frx":A4EE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8250
   End
End
Attribute VB_Name = "frm_8_game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim q() As Variant
Dim a(10, 4) As Variant
Dim r() As Variant
Dim i, j As Integer
Dim c As Integer

Private Sub Form_Activate()
img = LoadPicture(App.Path & "\img\img8_01_3.jpg")
i = 1
Call quiz
End Sub

Sub quiz()

If i = 11 Then
res8 = 2
frm_8.Show
frm_8_game.Hide

ElseIf i = 0 Then
res8 = 1
frm_8.Show
frm_8_game.Hide

Else

q() = Array(, "'소가 웃는 소리'를 세글자로 하면?", "진짜 문제 투성이인 것은?", _
"재수 없는 게 좋은 사람은?", "타이타닉의 구명보트에는 몇 명이 탈수 있을까?", _
"텔레토비에서 뽀가 떠나면?", "뜨거운 복숭아는?", _
"사람이 일생동안 가장 많이 하는 소리는?", "슈렉 엄마를 뭐라고 부를까?", _
"아홉명의 자식을 석자로 줄이면?", "개가 사람을 가르친다는 뜻의 네글자 말은?")

a(1, 1) = "이히히"
a(1, 2) = "으흐흐"
a(1, 3) = "우하하"
a(1, 4) = "킥킥킥"

a(2, 1) = "학교"
a(2, 2) = "시험지"
a(2, 3) = "집"
a(2, 4) = "나"

a(3, 1) = "선생님"
a(3, 2) = "부모님"
a(3, 3) = "동생"
a(3, 4) = "고3수험생"

a(4, 1) = "2명"
a(4, 2) = "9명"
a(4, 3) = "11명"
a(4, 4) = "15명"

a(5, 1) = "즐겁다"
a(5, 2) = "뽀빠이"
a(5, 3) = "뽀뽀틴"
a(5, 4) = "항의글 쓴다"

a(6, 1) = "천도복숭아"
a(6, 2) = "백도복숭아"
a(6, 3) = "황도복숭아"
a(6, 4) = "복숭아가 아니다"

a(7, 1) = "칭찬"
a(7, 2) = "비난"
a(7, 3) = "숨소리"
a(7, 4) = "잔소리"

a(8, 1) = "빨간색어머니"
a(8, 2) = "녹색어머니"
a(8, 3) = "우리엄마"
a(8, 4) = "시어머니"

a(9, 1) = "내팔자"
a(9, 2) = "상팔자"
a(9, 3) = "어이구"
a(9, 4) = "아이구"

a(10, 1) = "개인지도"
a(10, 2) = "개발학습"
a(10, 3) = "개량한복"
a(10, 4) = "건방진개"


r() = Array(, 3, 2, 4, 2, 2, 1, 3, 2, 4, 1)


lbl_q = i & ". " & q(i)
For j = 1 To 4
lbl(j) = j & ". " & a(i, j)
lbl(j).Enabled = True
Next

End If

End Sub


Sub next_to()
i = i + 1
Call quiz
End Sub

Private Sub lbl_Click(Index As Integer)

If Index = r(i) Then
Call sndplay("eff02.wav")
tmr.Enabled = True
Else
i = 0
Call quiz
End If

End Sub



Private Sub tmr_Timer()
Select Case c
Case 0
img = LoadPicture(App.Path & "\img\img8_01_4.jpg")
tmr.Interval = 2000
c = 1
lbl_q = ""
For j = 1 To 4
    lbl(j) = ""
    lbl(j).Enabled = False
Next
Case 1
img = LoadPicture(App.Path & "\img\img8_01_3.jpg")
tmr.Enabled = False
tmr.Interval = 500
c = 0
Call next_to
End Select
End Sub

