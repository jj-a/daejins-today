VERSION 5.00
Begin VB.Form frm_7_game 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "대진의 하루"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   Icon            =   "frm_7_game.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7800
      Top             =   120
   End
   Begin VB.Image img_ht 
      Height          =   720
      Left            =   120
      MouseIcon       =   "frm_7_game.frx":8732
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_7_game.frx":8A3C
      Top             =   5160
      Width           =   720
   End
   Begin VB.Label lbl 
      BackStyle       =   0  '투명
      Height          =   495
      Index           =   4
      Left            =   3240
      MouseIcon       =   "frm_7_game.frx":AE36
      MousePointer    =   99  '사용자 정의
      TabIndex        =   4
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label lbl 
      BackStyle       =   0  '투명
      Height          =   495
      Index           =   3
      Left            =   3240
      MouseIcon       =   "frm_7_game.frx":B140
      MousePointer    =   99  '사용자 정의
      TabIndex        =   3
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label lbl 
      BackStyle       =   0  '투명
      Height          =   495
      Index           =   2
      Left            =   3240
      MouseIcon       =   "frm_7_game.frx":B44A
      MousePointer    =   99  '사용자 정의
      TabIndex        =   2
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label lbl 
      BackStyle       =   0  '투명
      Height          =   495
      Index           =   1
      Left            =   3240
      MouseIcon       =   "frm_7_game.frx":B754
      MousePointer    =   99  '사용자 정의
      TabIndex        =   1
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label lbl_q 
      BackStyle       =   0  '투명
      Height          =   615
      Left            =   3120
      MouseIcon       =   "frm_7_game.frx":BA5E
      MousePointer    =   99  '사용자 정의
      TabIndex        =   0
      Top             =   480
      Width           =   4215
   End
   Begin VB.Image img 
      Height          =   6000
      Left            =   0
      MouseIcon       =   "frm_7_game.frx":C328
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_7_game.frx":CBF2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8250
   End
End
Attribute VB_Name = "frm_7_game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim q() As Variant
Dim a(5, 4) As Variant
Dim r() As Variant
Dim i, j As Integer
Dim c As Integer

Private Sub Form_Activate()
If ch > 0 Then
Else
img = LoadPicture(App.Path & "\img\img7_01_4.jpg")
i = 1
Call quiz
End If
End Sub

Sub quiz()

If i = 6 Then
res7 = 2
ch = 0
frm_7.Show
frm_7_game.Hide

ElseIf i = 0 Then
res7 = 1
ch = 0
frm_7.Show
frm_7_game.Hide

Else

q() = Array(, "프로그래밍 언어의 종류 중 고급언어로 옳지 않은 것은?", _
"컴파일러와 인터프리터를 비교한 것으로 옳지 않은 것은?", "다음 중 객체지향언어인 것은?", _
"객체지향 프로그래밍의 특징으로 옳지 않은 것은?", "반복문의 대표적인 문법은?")

a(1, 1) = "코볼"
a(1, 2) = "C언어"
a(1, 3) = "자바"
a(1, 4) = "어셈블리어"

a(2, 1) = "번역방법 : 컴파일러-전체번역  인터프리터-행단위번역"
a(2, 2) = "번역속도 : 컴파일러-느림  인터프리터-빠름"
a(2, 3) = "실행속도 : 컴파일러-빠름  인터프리터-느림"
a(2, 4) = "목적프로그램 : 컴파일러-생성안함  인터프리터-생성"

a(3, 1) = "C언어"
a(3, 2) = "베이직"
a(3, 3) = "자바"
a(3, 4) = "포트란"

a(4, 1) = "캡슐화"
a(4, 2) = "순차적"
a(4, 3) = "상속성"
a(4, 4) = "다형성"

a(5, 1) = "For~next문"
a(5, 2) = "Go to문"
a(5, 3) = "If문"
a(5, 4) = "Select case문"


r() = Array(, 4, 4, 3, 2, 1)


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
tmr.Enabled = True
Call sndplay("eff02.wav")
Else
i = 0
Call quiz
End If

End Sub


Private Sub tmr_Timer()
Select Case c
Case 0
img = LoadPicture(App.Path & "\img\img7_01_3.jpg")
tmr.Interval = 2000
c = 1
lbl_q = ""
For j = 1 To 4
    lbl(j) = ""
    lbl(j).Enabled = False
Next
Case 1
img = LoadPicture(App.Path & "\img\img7_01_4.jpg")
tmr.Enabled = False
tmr.Interval = 500
c = 0
Call next_to
End Select
End Sub


Private Sub img_ht_Click()
Call sndplay("eff01.wav")
frm_7_hint.Show
ch = ch + 1
End Sub



