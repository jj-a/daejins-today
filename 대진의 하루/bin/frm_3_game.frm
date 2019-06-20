VERSION 5.00
Begin VB.Form frm_3_game 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "대진의 하루"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   Icon            =   "frm_3_game.frx":0000
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
      Left            =   7440
      MouseIcon       =   "frm_3_game.frx":8732
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_3_game.frx":8A3C
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lbl_q 
      BackStyle       =   0  '투명
      Height          =   735
      Left            =   2400
      MouseIcon       =   "frm_3_game.frx":AE36
      MousePointer    =   99  '사용자 정의
      TabIndex        =   4
      Top             =   1080
      Width           =   5055
   End
   Begin VB.Label lbl 
      BackStyle       =   0  '투명
      Height          =   375
      Index           =   4
      Left            =   2640
      MouseIcon       =   "frm_3_game.frx":B700
      MousePointer    =   99  '사용자 정의
      TabIndex        =   3
      Top             =   4800
      Width           =   5055
   End
   Begin VB.Label lbl 
      BackStyle       =   0  '투명
      Height          =   375
      Index           =   3
      Left            =   2640
      MouseIcon       =   "frm_3_game.frx":BA0A
      MousePointer    =   99  '사용자 정의
      TabIndex        =   2
      Top             =   4080
      Width           =   5055
   End
   Begin VB.Label lbl 
      BackStyle       =   0  '투명
      Height          =   375
      Index           =   2
      Left            =   2640
      MouseIcon       =   "frm_3_game.frx":BD14
      MousePointer    =   99  '사용자 정의
      TabIndex        =   1
      Top             =   3480
      Width           =   5055
   End
   Begin VB.Label lbl 
      BackStyle       =   0  '투명
      Height          =   375
      Index           =   1
      Left            =   2640
      MouseIcon       =   "frm_3_game.frx":C01E
      MousePointer    =   99  '사용자 정의
      TabIndex        =   0
      Top             =   2880
      Width           =   5055
   End
   Begin VB.Image img 
      Height          =   6000
      Left            =   0
      MouseIcon       =   "frm_3_game.frx":C328
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_3_game.frx":CBF2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8250
   End
End
Attribute VB_Name = "frm_3_game"
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
img = LoadPicture(App.Path & "\img\img3_01_3.jpg")
i = 1
Call quiz
End If
End Sub

Sub quiz()

If i = 6 Then
res3 = 2
ch = 0
frm_3.Show
frm_3_game.Hide

ElseIf i = 0 Then
res3 = 1
ch = 0
frm_3.Show
frm_3_game.Hide

Else

q() = Array(, "ASP의 주요 객체가 아닌 것은?", "내 이름은 뭐지?", _
"다음 URL을 보고 서버에 전달하기 위해 사용된 데이터 전달 방식을 고르시오." & vbCrLf & vbCrLf & "http://localhost/good/asp/daejin.asp?year=2010＆name=jja", _
"ⓐ전역변수 사용시와 ⓑ지역변수 사용시 쓰이는 객체는?", _
"사용자(Client) 개인(방문자) 정보를 파일 형태로 저장해둘 수 있는 컬렉션은?")

a(1, 1) = "Request객체"
a(1, 2) = "Application객체"
a(1, 3) = "Response객체"
a(1, 4) = "Global객체"

a(2, 1) = "송대관"
a(2, 2) = "송세근"
a(2, 3) = "손계정"
a(2, 4) = "송수금"

a(3, 1) = "Response방식"
a(3, 2) = "Post방식"
a(3, 3) = "Cookies방식"
a(3, 4) = "Get방식"

a(4, 1) = "ⓐ Application  ⓑ Session"
a(4, 2) = "ⓐ Server  ⓑ Session"
a(4, 3) = "ⓐ Session  ⓑ Application"
a(4, 4) = "ⓐ Request  ⓑ Server"

a(5, 1) = "Form"
a(5, 2) = "Cookies"
a(5, 3) = "Servervariable"
a(5, 4) = "Querystring"


r() = Array(, 4, 2, 4, 1, 2)


lbl_q = i & ". " & q(i)
For j = 1 To 4
lbl(j) = j & ". " & a(i, j)
Next

End If

End Sub


Private Sub img_ht_Click()
Call sndplay("eff01.wav")
frm_3_hint.Show
ch = ch + 1
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


