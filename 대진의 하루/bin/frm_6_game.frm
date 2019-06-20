VERSION 5.00
Begin VB.Form frm_6_game 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "대진의 하루"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   Icon            =   "frm_6_game.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Image img_ht 
      Height          =   720
      Left            =   120
      MouseIcon       =   "frm_6_game.frx":8732
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_6_game.frx":8A3C
      Top             =   5160
      Width           =   720
   End
   Begin VB.Label lbl 
      BackStyle       =   0  '투명
      Height          =   495
      Index           =   4
      Left            =   2040
      MouseIcon       =   "frm_6_game.frx":AE36
      MousePointer    =   99  '사용자 정의
      TabIndex        =   4
      Top             =   3440
      Width           =   5895
   End
   Begin VB.Label lbl 
      BackStyle       =   0  '투명
      Height          =   495
      Index           =   3
      Left            =   2040
      MouseIcon       =   "frm_6_game.frx":B140
      MousePointer    =   99  '사용자 정의
      TabIndex        =   3
      Top             =   2800
      Width           =   5895
   End
   Begin VB.Label lbl 
      BackStyle       =   0  '투명
      Height          =   495
      Index           =   2
      Left            =   2040
      MouseIcon       =   "frm_6_game.frx":B44A
      MousePointer    =   99  '사용자 정의
      TabIndex        =   2
      Top             =   2160
      Width           =   5895
   End
   Begin VB.Label lbl 
      BackStyle       =   0  '투명
      Height          =   495
      Index           =   1
      Left            =   2040
      MouseIcon       =   "frm_6_game.frx":B754
      MousePointer    =   99  '사용자 정의
      TabIndex        =   1
      Top             =   1560
      Width           =   5895
   End
   Begin VB.Label lbl_q 
      BackStyle       =   0  '투명
      Height          =   735
      Left            =   2160
      MouseIcon       =   "frm_6_game.frx":BA5E
      MousePointer    =   99  '사용자 정의
      TabIndex        =   0
      Top             =   240
      Width           =   5655
   End
   Begin VB.Image img 
      Height          =   6000
      Left            =   0
      MouseIcon       =   "frm_6_game.frx":C328
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_6_game.frx":CBF2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8250
   End
End
Attribute VB_Name = "frm_6_game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim q() As Variant
Dim a(5, 4) As Variant
Dim r() As Variant
Dim i, j As Integer

Private Sub Form_Activate()
If ch > 0 Then
Else
img = LoadPicture(App.Path & "\img\img6_01_3.jpg")
i = 1
Call quiz
End If
End Sub

Sub quiz()

If i = 6 Then
res6 = 2
ch = 0
frm_6.Show
frm_6_game.Hide

ElseIf i = 0 Then
res6 = 1
ch = 0
frm_6.Show
frm_6_game.Hide

Else

q() = Array(, "'Dreams come true' 를 우리말로 해석하시오.", _
"대학수학능력시험 외국어영역 시간은?", "단어와 뜻이 알맞게 짝지어진 것은?", _
"현재완료의 기본 형태로 옳은 것은?", _
"'It is necessary to advertise yourself.' 에서 진주어를 찾으시오.")

a(1, 1) = "꿈은 올 것이다."
a(1, 2) = "꿈이 오는 건 사실이다."
a(1, 3) = "꿈은 이루어진다."
a(1, 4) = "꿈을 꾸고 싶다."

a(2, 1) = "70분"
a(2, 2) = "80분"
a(2, 3) = "100분"
a(2, 4) = "120분"

a(3, 1) = "always 자주"
a(3, 2) = "aften 전혀~없다"
a(3, 3) = "sometimes 때때로"
a(3, 4) = "never 항상"

a(4, 1) = "have+p.p"
a(4, 2) = "be+~ing"
a(4, 3) = "to+v"
a(4, 4) = "had+been+p.p"

a(5, 1) = "It"
a(5, 2) = "to advertise yourself"
a(5, 3) = "to advertise"
a(5, 4) = "necessary"


r() = Array(, 3, 1, 3, 1, 2)


lbl_q = i & ". " & q(i)
For j = 1 To 4
lbl(j) = j & ". " & a(i, j)
Next

End If

End Sub

Private Sub lbl_Click(Index As Integer)

If Index = r(i) Then
Call sndplay("eff02.wav")
i = i + 1
Call quiz
Else
i = 0
Call quiz
End If

End Sub


Private Sub img_ht_Click()
Call sndplay("eff01.wav")
frm_6_hint.Show
ch = ch + 1
End Sub


