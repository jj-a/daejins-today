VERSION 5.00
Begin VB.Form frm_2_game 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "대진의 하루"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   Icon            =   "frm_2_game.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Image img_bt 
      Height          =   450
      Left            =   4800
      MouseIcon       =   "frm_2_game.frx":8732
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_2_game.frx":8A3C
      Top             =   1400
      Width           =   1200
   End
   Begin VB.Image img3_bot 
      Height          =   2250
      Left            =   5600
      MouseIcon       =   "frm_2_game.frx":BC91
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_2_game.frx":BF9B
      Top             =   3740
      Width           =   2640
   End
   Begin VB.Image img3_top 
      Height          =   1950
      Left            =   5600
      MouseIcon       =   "frm_2_game.frx":12CE3
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_2_game.frx":12FED
      Top             =   1890
      Width           =   2640
   End
   Begin VB.Image img2_bot 
      Height          =   2265
      Left            =   2280
      MouseIcon       =   "frm_2_game.frx":1CE60
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_2_game.frx":1D16A
      Top             =   3720
      Width           =   3405
   End
   Begin VB.Image img2_top 
      Height          =   1950
      Left            =   2280
      MouseIcon       =   "frm_2_game.frx":246F4
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_2_game.frx":249FE
      Top             =   1890
      Width           =   3405
   End
   Begin VB.Image img1_bot 
      Height          =   2265
      Left            =   0
      MouseIcon       =   "frm_2_game.frx":2C79F
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_2_game.frx":2CAA9
      Top             =   3720
      Width           =   2400
   End
   Begin VB.Image img1_top 
      Height          =   1935
      Left            =   0
      MouseIcon       =   "frm_2_game.frx":34695
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_2_game.frx":3499F
      Top             =   1900
      Width           =   2400
   End
   Begin VB.Image img 
      Height          =   6000
      Left            =   0
      MouseIcon       =   "frm_2_game.frx":3E53C
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_2_game.frx":3EE06
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8250
   End
End
Attribute VB_Name = "frm_2_game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim a_top, a_bot, b_top, b_bot, c_top, c_bot


Private Sub Form_Activate()
img = LoadPicture(App.Path & "\img\img2_01_gm.jpg")
img1_top = LoadPicture(App.Path & "\img\img2_01_gm1_t1.jpg")
img1_bot = LoadPicture(App.Path & "\img\img2_01_gm1_b1.jpg")
img2_top = LoadPicture(App.Path & "\img\img2_01_gm2_t1.jpg")
img2_bot = LoadPicture(App.Path & "\img\img2_01_gm2_b1.jpg")
img3_top = LoadPicture(App.Path & "\img\img2_01_gm3_t1.jpg")
img3_bot = LoadPicture(App.Path & "\img\img2_01_gm3_b1.jpg")

End Sub



Private Sub img_bt_Click()
If a_top = 4 And a_bot = 4 And b_top = 4 And b_bot = 4 And c_top = 4 And c_bot = 4 Then
res2 = 2
Else
res2 = 1
End If
frm_2.Show
frm_2_game.Hide
End Sub



Private Sub img1_top_Click()


Select Case a_top
Case 0: img1_top = LoadPicture(App.Path & "\img\img2_01_gm1_t2.jpg"): a_top = 1
Case 1: img1_top = LoadPicture(App.Path & "\img\img2_01_gm1_t3.jpg"): a_top = 2
Case 2: img1_top = LoadPicture(App.Path & "\img\img2_01_gm1_t4.jpg"): a_top = 3
Case 3: img1_top = LoadPicture(App.Path & "\img\img2_01_gm1_t5.jpg"): a_top = 4
Case 4: img1_top = LoadPicture(App.Path & "\img\img2_01_gm1_t1.jpg"): a_top = 0
End Select

End Sub



Private Sub img1_bot_Click()

Select Case a_bot
Case 0: img1_bot = LoadPicture(App.Path & "\img\img2_01_gm1_b2.jpg"): a_bot = 1
Case 1: img1_bot = LoadPicture(App.Path & "\img\img2_01_gm1_b3.jpg"): a_bot = 2
Case 2: img1_bot = LoadPicture(App.Path & "\img\img2_01_gm1_b4.jpg"): a_bot = 3
Case 3: img1_bot = LoadPicture(App.Path & "\img\img2_01_gm1_b5.jpg"): a_bot = 4
Case 4: img1_bot = LoadPicture(App.Path & "\img\img2_01_gm1_b1.jpg"): a_bot = 0
End Select

End Sub








Private Sub img2_top_Click()

Select Case b_top
Case 0: img2_top = LoadPicture(App.Path & "\img\img2_01_gm2_t2.jpg"): b_top = 1
Case 1: img2_top = LoadPicture(App.Path & "\img\img2_01_gm2_t3.jpg"): b_top = 2
Case 2: img2_top = LoadPicture(App.Path & "\img\img2_01_gm2_t4.jpg"): b_top = 3
Case 3: img2_top = LoadPicture(App.Path & "\img\img2_01_gm2_t5.jpg"): b_top = 4
Case 4: img2_top = LoadPicture(App.Path & "\img\img2_01_gm2_t1.jpg"): b_top = 0
End Select

End Sub



Private Sub img2_bot_Click()

Select Case b_bot
Case 0: img2_bot = LoadPicture(App.Path & "\img\img2_01_gm2_b2.jpg"): b_bot = 1
Case 1: img2_bot = LoadPicture(App.Path & "\img\img2_01_gm2_b3.jpg"): b_bot = 2
Case 2: img2_bot = LoadPicture(App.Path & "\img\img2_01_gm2_b4.jpg"): b_bot = 3
Case 3: img2_bot = LoadPicture(App.Path & "\img\img2_01_gm2_b5.jpg"): b_bot = 4
Case 4: img2_bot = LoadPicture(App.Path & "\img\img2_01_gm2_b1.jpg"): b_bot = 0
End Select

End Sub







Private Sub img3_top_Click()

Select Case c_top
Case 0: img3_top = LoadPicture(App.Path & "\img\img2_01_gm3_t2.jpg"): c_top = 1
Case 1: img3_top = LoadPicture(App.Path & "\img\img2_01_gm3_t3.jpg"): c_top = 2
Case 2: img3_top = LoadPicture(App.Path & "\img\img2_01_gm3_t4.jpg"): c_top = 3
Case 3: img3_top = LoadPicture(App.Path & "\img\img2_01_gm3_t5.jpg"): c_top = 4
Case 4: img3_top = LoadPicture(App.Path & "\img\img2_01_gm3_t1.jpg"): c_top = 0
End Select

End Sub





Private Sub img3_bot_Click()

Select Case c_bot
Case 0: img3_bot = LoadPicture(App.Path & "\img\img2_01_gm3_b2.jpg"): c_bot = 1
Case 1: img3_bot = LoadPicture(App.Path & "\img\img2_01_gm3_b3.jpg"): c_bot = 2
Case 2: img3_bot = LoadPicture(App.Path & "\img\img2_01_gm3_b4.jpg"): c_bot = 3
Case 3: img3_bot = LoadPicture(App.Path & "\img\img2_01_gm3_b5.jpg"): c_bot = 4
Case 4: img3_bot = LoadPicture(App.Path & "\img\img2_01_gm3_b1.jpg"): c_bot = 0
End Select

End Sub

