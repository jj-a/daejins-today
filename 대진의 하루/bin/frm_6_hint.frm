VERSION 5.00
Begin VB.Form frm_6_hint 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "대진의 하루"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   Icon            =   "frm_6_hint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Image img_bk 
      Height          =   720
      Left            =   120
      MouseIcon       =   "frm_6_hint.frx":8732
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_6_hint.frx":8A3C
      Top             =   5160
      Width           =   720
   End
   Begin VB.Image img_ht 
      Height          =   6000
      Left            =   0
      MouseIcon       =   "frm_6_hint.frx":AE36
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_6_hint.frx":B140
      Top             =   0
      Width           =   8250
   End
End
Attribute VB_Name = "frm_6_hint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim pagenum


Private Sub Form_Activate()
img_ht = LoadPicture(App.Path & "\img\img_ht6_0.jpg")
img_bk.Visible = False
pagenum = 1
End Sub


Sub pagemove()
Select Case pagenum
Case 1: img_ht = LoadPicture(App.Path & "\img\img_ht6_1.jpg"): pagenum = 2
Case 2: img_ht = LoadPicture(App.Path & "\img\img_ht6_2.jpg"): pagenum = 3
Case 3: img_ht = LoadPicture(App.Path & "\img\img_ht6_3.jpg"): pagenum = 4
Case 4: img_ht = LoadPicture(App.Path & "\img\img_ht6_4.jpg"): pagenum = 5
Case 5: img_ht = LoadPicture(App.Path & "\img\img_ht6_5.jpg"): pagenum = 6
Case 6: img_ht = LoadPicture(App.Path & "\img\img_ht6_6.jpg"): pagenum = 7
Case 7: img_ht = LoadPicture(App.Path & "\img\img_ht6_7.jpg"): pagenum = 8
Case 8: img_ht = LoadPicture(App.Path & "\img\img_ht6_8.jpg"): pagenum = 9
: img_bk.Visible = True: img_ht.MouseIcon = LoadPicture(App.Path & "\img\mus_normal.cur")
End Select
End Sub


Private Sub img_ht_Click()
If pagenum >= 9 Then
Else
Call pagemove
End If
End Sub


Private Sub img_bk_Click()
Call sndplay("eff01.wav")
pagenum = 0
frm_6_hint.Hide
End Sub


