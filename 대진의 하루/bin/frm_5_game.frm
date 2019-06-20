VERSION 5.00
Begin VB.Form frm_5_game 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "대진의 하루"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   Icon            =   "frm_5_game.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.TextBox txt 
      Alignment       =   2  '가운데 맞춤
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5400
      MouseIcon       =   "frm_5_game.frx":8732
      MousePointer    =   99  '사용자 정의
      TabIndex        =   0
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Image img_ht 
      Height          =   720
      Left            =   120
      MouseIcon       =   "frm_5_game.frx":8A3C
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_5_game.frx":8D46
      Top             =   5160
      Width           =   720
   End
   Begin VB.Label lbl_q 
      BackStyle       =   0  '투명
      Height          =   615
      Left            =   3480
      MouseIcon       =   "frm_5_game.frx":B140
      MousePointer    =   99  '사용자 정의
      TabIndex        =   1
      Top             =   720
      Width           =   4215
   End
   Begin VB.Image img 
      Height          =   6000
      Left            =   0
      MouseIcon       =   "frm_5_game.frx":BA0A
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_5_game.frx":C2D4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8250
   End
End
Attribute VB_Name = "frm_5_game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'타이머 이용해서 시간 제한 두기


Dim q() As Variant
Dim a(5) As Variant
Dim r() As Variant
Dim i, j As Integer


Private Sub Form_Activate()
If ch > 0 Then
Else
img = LoadPicture(App.Path & "\img\img5_01_2.jpg")
i = 1
Call quiz
End If
End Sub


Sub quiz()

If i = 6 Then
res5 = 2
ch = 0
frm_5.Show
frm_5_game.Hide

ElseIf i = 0 Then
res5 = 1
ch = 0
frm_5.Show
frm_5_game.Hide

Else

q() = Array(, "다양한 정보미디어를 디지털화하여 사용자에게 대화형으로 연결하여 제공하는 것" & vbCrLf & "--- □□□□□", _
"멀티미디어의 기본요소 중 하나를 쓰시오.", _
"플래시 무비 파일을 만들 때의 확장자는?" & vbCrLf & "--- *.□□□", _
"주파수, 진폭, 음색을 □□의 3요소라 한다.", _
"플래시에서 장면의 한 단위를 무엇이라고 하는가?" & vbCrLf & "--- □□□")


r() = Array(, "멀티미디어", , "swf", "소리", "프레임")


lbl_q = i & ". " & q(i)



End If

End Sub



Private Sub txt_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
  
  a(i) = txt
  
  If i = 2 Then
    If a(i) = "문자" Or a(i) = "그림" Or a(i) = "소리" Or a(i) = "애니메이션" Or a(i) = "동영상" Or a(i) = "시스템" Or a(i) = "통신망" Then
    Call sndplay("eff02.wav")
      i = i + 1
      txt = ""
      Call quiz
    Else
      i = 0
      Call quiz
    End If
  Else
    If a(i) = r(i) Then
    Call sndplay("eff02.wav")
      i = i + 1
      txt = ""
      Call quiz
    Else
      i = 0
      txt = ""
      Call quiz
    End If
  End If

End If

End Sub


Private Sub img_ht_Click()
Call sndplay("eff01.wav")
frm_5_hint.Show
ch = ch + 1
End Sub





