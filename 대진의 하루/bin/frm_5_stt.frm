VERSION 5.00
Begin VB.Form frm_5_stt 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "대진의 하루"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   Icon            =   "frm_5_stt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Image img 
      Height          =   6000
      Left            =   0
      MouseIcon       =   "frm_5_stt.frx":8732
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_5_stt.frx":8A3C
      Top             =   0
      Width           =   8250
   End
End
Attribute VB_Name = "frm_5_stt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim pagenum


Private Sub Form_Activate()
img = LoadPicture(App.Path & "\img\img_std.jpg")

pagenum = 1
End Sub


Sub pagemove()
Select Case pagenum
Case 1: img = LoadPicture(App.Path & "\img\img_std5_1.jpg"): pagenum = 2
Case 2: img = LoadPicture(App.Path & "\img\img_std5_2.jpg"): pagenum = 3
Case 3: img = LoadPicture(App.Path & "\img\img_std5_3.jpg"): pagenum = 4
End Select
End Sub

Private Sub img_Click()
If pagenum >= 4 Then
pagenum = 0
frm_5_tuto.Show
frm_5_stt.Hide
Else
Call pagemove
End If
End Sub










