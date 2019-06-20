VERSION 5.00
Begin VB.Form frm_1_game 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "대진의 하루"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   Icon            =   "frm_1_game.frx":0000
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
      MouseIcon       =   "frm_1_game.frx":8732
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_1_game.frx":8A3C
      Top             =   120
      Width           =   720
   End
   Begin VB.Image img_a 
      Height          =   1920
      Index           =   4
      Left            =   5060
      MouseIcon       =   "frm_1_game.frx":AE36
      MousePointer    =   99  '사용자 정의
      Top             =   3465
      Width           =   2790
   End
   Begin VB.Image img_a 
      Height          =   1920
      Index           =   3
      Left            =   2385
      MouseIcon       =   "frm_1_game.frx":B140
      MousePointer    =   99  '사용자 정의
      Top             =   3465
      Width           =   2655
   End
   Begin VB.Image img_a 
      Height          =   1860
      Index           =   2
      Left            =   5060
      MouseIcon       =   "frm_1_game.frx":B44A
      MousePointer    =   99  '사용자 정의
      Top             =   1580
      Width           =   2780
   End
   Begin VB.Image img_a 
      Height          =   1845
      Index           =   1
      Left            =   2380
      MouseIcon       =   "frm_1_game.frx":B754
      MousePointer    =   99  '사용자 정의
      Top             =   1580
      Width           =   2655
   End
   Begin VB.Label lbl_q 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   20.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   880
      MouseIcon       =   "frm_1_game.frx":BA5E
      MousePointer    =   99  '사용자 정의
      TabIndex        =   0
      Top             =   3360
      Width           =   1180
   End
   Begin VB.Image img 
      Height          =   6000
      Left            =   0
      MouseIcon       =   "frm_1_game.frx":C328
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_1_game.frx":CBF2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8250
   End
End
Attribute VB_Name = "frm_1_game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim q() As Variant
Dim r As Variant
Dim a, a_chk(4) As Variant
Dim que As Variant
Dim i, j, k As Integer
Dim m, n(1 To 5) As Integer


Private Sub Form_Activate()
If ch > 0 Then
Else
img = LoadPicture(App.Path & "\img\img1_01_3.jpg")
i = 1
Call quiz
End If
End Sub


Sub quiz()

If i = 6 Then
res1 = 2
ch = 0
frm_1.Show
frm_1_game.Hide

ElseIf i = 0 Then
res1 = 1
ch = 0
frm_1.Show
frm_1_game.Hide

Else

q() = Array(, "AND", "OR", "NOT", "NAND", "NOR", "XOR", "XNOR")

Randomize

return_to1:
que = Int(Rnd * 7) + 1
For m = 1 To 5
If que = n(m) Then GoTo return_to1
Next
lbl_q = q(que)
n(i) = que
r = Int(Rnd * 4) + 1

For j = 1 To 4

If j = r Then
a = que
Call a_load

Else
return_to2:
a = Int(Rnd * 7) + 1
For k = 1 To 4
If a = que Or a = a_chk(k) Then
GoTo return_to2
End If
Next k
Call a_load

End If

a_chk(j) = a

Next j

End If

End Sub



Sub a_load()
Select Case a
Case 1
img_a(j) = LoadPicture(App.Path & "\img\img1_01_and.jpg")
Case 2
img_a(j) = LoadPicture(App.Path & "\img\img1_01_or.jpg")
Case 3
img_a(j) = LoadPicture(App.Path & "\img\img1_01_not.jpg")
Case 4
img_a(j) = LoadPicture(App.Path & "\img\img1_01_nand.jpg")
Case 5
img_a(j) = LoadPicture(App.Path & "\img\img1_01_nor.jpg")
Case 6
img_a(j) = LoadPicture(App.Path & "\img\img1_01_xor.jpg")
Case 7
img_a(j) = LoadPicture(App.Path & "\img\img1_01_xnor.jpg")
End Select

End Sub

Private Sub img_a_Click(Index As Integer)

If Index = r Then
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
frm_1_hint.Show
ch = ch + 1
End Sub




