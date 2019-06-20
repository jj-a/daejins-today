VERSION 5.00
Begin VB.Form frm_ending 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "대진의 하루"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   Icon            =   "frm_ending.frx":0000
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
      Interval        =   1000
      Left            =   360
      Top             =   240
   End
   Begin VB.Image img 
      Height          =   6000
      Left            =   0
      MouseIcon       =   "frm_ending.frx":8732
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_ending.frx":8FFC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8250
   End
End
Attribute VB_Name = "frm_ending"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim pagenum As Integer

Private Sub Form_Activate()
img = LoadPicture(App.Path & "\img\img_end_0.jpg")
tmr.Enabled = True
End Sub

Private Sub img_Click()

End Sub

Private Sub tmr_Timer()
If pagenum = 12 Then
tmr.Enabled = False
Call move_to
Else
Select Case pagenum
Case 0: img = LoadPicture(App.Path & "\img\img_end_0.jpg"): pagenum = 1
Case 1: img = LoadPicture(App.Path & "\img\img_end_1.jpg"): pagenum = 2: tmr.Interval = 2000
Case 2: img = LoadPicture(App.Path & "\img\img_end_2.jpg"): pagenum = 3: tmr.Interval = 3000
Case 3: img = LoadPicture(App.Path & "\img\img_end_3.jpg"): pagenum = 4
Case 4: img = LoadPicture(App.Path & "\img\img_end_4.jpg"): pagenum = 5
Case 5: img = LoadPicture(App.Path & "\img\img_end_5.jpg"): pagenum = 6
Case 6: img = LoadPicture(App.Path & "\img\img_end_6.jpg"): pagenum = 7
Case 7: img = LoadPicture(App.Path & "\img\img_end_7.jpg"): pagenum = 8: tmr.Interval = 5000
Case 8: img = LoadPicture(App.Path & "\img\img_end_1.jpg"): pagenum = 9: tmr.Interval = 3000
Case 9: img = LoadPicture(App.Path & "\img\img_end_8.jpg"): pagenum = 10: tmr.Interval = 5000
Case 10: img = LoadPicture(App.Path & "\img\img_end_9.jpg"): pagenum = 11
Case 11: img = LoadPicture(App.Path & "\img\img_end_10.jpg"): pagenum = 12
End Select
End If
End Sub

Sub move_to()
frm_logo.Show
frm_ending.Hide
End Sub
