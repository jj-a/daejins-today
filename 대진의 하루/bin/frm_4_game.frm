VERSION 5.00
Begin VB.Form frm_4_game 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "대진의 하루"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   Icon            =   "frm_4_game.frx":0000
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
      MouseIcon       =   "frm_4_game.frx":8732
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_4_game.frx":8A3C
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lbl 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Height          =   495
      Index           =   4
      Left            =   2640
      MouseIcon       =   "frm_4_game.frx":AE36
      MousePointer    =   99  '사용자 정의
      TabIndex        =   4
      Top             =   5160
      Width           =   4935
   End
   Begin VB.Label lbl 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Height          =   495
      Index           =   3
      Left            =   2640
      MouseIcon       =   "frm_4_game.frx":B140
      MousePointer    =   99  '사용자 정의
      TabIndex        =   3
      Top             =   4440
      Width           =   4935
   End
   Begin VB.Label lbl 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Height          =   495
      Index           =   2
      Left            =   2640
      MouseIcon       =   "frm_4_game.frx":B44A
      MousePointer    =   99  '사용자 정의
      TabIndex        =   2
      Top             =   3600
      Width           =   4935
   End
   Begin VB.Label lbl 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Height          =   495
      Index           =   1
      Left            =   2640
      MouseIcon       =   "frm_4_game.frx":B754
      MousePointer    =   99  '사용자 정의
      TabIndex        =   1
      Top             =   2760
      Width           =   4935
   End
   Begin VB.Label lbl_q 
      BackStyle       =   0  '투명
      Height          =   735
      Left            =   2400
      MouseIcon       =   "frm_4_game.frx":BA5E
      MousePointer    =   99  '사용자 정의
      TabIndex        =   0
      Top             =   960
      Width           =   5295
   End
   Begin VB.Image img 
      Height          =   6000
      Left            =   0
      MouseIcon       =   "frm_4_game.frx":C328
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_4_game.frx":CBF2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8250
   End
End
Attribute VB_Name = "frm_4_game"
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
img = LoadPicture(App.Path & "\img\img4_01_3.jpg")
i = 1
Call quiz
End If
End Sub

Sub quiz()

If i = 6 Then
res4 = 2
ch = 0
frm_4.Show
frm_4_game.Hide

ElseIf i = 0 Then
res4 = 1
ch = 0
frm_4.Show
frm_4_game.Hide

Else

q() = Array(, "언어의 특성이 아닌 것은?", "표준말의 대원칙으로 옳은 것은?", _
"다음 문장 중 바르게 쓴 것은?", _
"다음 문장의 어절과 음절을 고르시오." & vbCrLf & vbCrLf & "<마지막 수업을 하면서 우리는 국어의 소중함을 깨닫게 됩니다.>", _
"한글 맞춤법의 대원칙으로 알맞은 것을 고르시오.")

a(1, 1) = "사회성"
a(1, 2) = "역사성"
a(1, 3) = "자립성"
a(1, 4) = "독립성"

a(2, 1) = "교양있는 사람들이 두루 쓰는 고대 서울말"
a(2, 2) = "교양있는 사람들이 두루 쓰는 현대 서울말"
a(2, 3) = "모든 사람에게 두루 쓰이는 현대 서울말"
a(2, 4) = "모든 사람에게 두루 쓰이는 현대 순우리말"

a(3, 1) = "나는 어제 떡볶기를 맛있게 먹었다."
a(3, 2) = "깡총깡총 잘도 뛴다."
a(3, 3) = "나는 어제 자장면 곱빼기를 먹었다."
a(3, 4) = "이것이 내 돐사진이다."

a(4, 1) = "어절7 음절25"
a(4, 2) = "어절8 음절25"
a(4, 3) = "어절8 음절24"
a(4, 4) = "어절8 음절26"

a(5, 1) = "한글맞춤법은 표준어를 소리대로 적되 어법에 맞도록 함을 원칙으로 한다."
a(5, 2) = "한글맞춤법은 표준어를 된소리로 적되 어법에 맞도록 함을 원칙으로 한다."
a(5, 3) = "한글맞춤법은 표준어를 소리대로 적음을 원칙으로 한다."
a(5, 4) = "한글맞춤법은 사투리를 소리대로 적되 된소리를 쓰지 않음을 원칙으로 한다."


r() = Array(, 4, 2, 3, 2, 1)


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
frm_4_hint.Show
ch = ch + 1
End Sub



