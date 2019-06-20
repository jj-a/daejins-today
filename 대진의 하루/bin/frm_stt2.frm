VERSION 5.00
Begin VB.Form frm_stt2 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "대진의 하루"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   Icon            =   "frm_stt2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Image img_tip 
      Height          =   300
      Left            =   2640
      MouseIcon       =   "frm_stt2.frx":8732
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_stt2.frx":8FFC
      Top             =   200
      Width           =   3000
   End
   Begin VB.Image img_stt2 
      Height          =   6000
      Left            =   0
      MouseIcon       =   "frm_stt2.frx":C7B8
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_stt2.frx":CAC2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8250
   End
End
Attribute VB_Name = "frm_stt2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'frm_stt2->frm_1

Dim pagenum



Private Sub Form_Activate()
img_stt2 = LoadPicture(App.Path & "\img\empty.jpg")
End Sub


Private Sub img_stt2_Click()
If pagenum >= 5 Then
pagenum = 0
frm_1.Show
frm_stt2.Hide
Else
Call pagemove
End If
End Sub


Sub pagemove()

Select Case pagenum
Case 0: img_stt2 = LoadPicture(App.Path & "\img\img0_05_1.jpg"): pagenum = 1
Case 1: img_stt2 = LoadPicture(App.Path & "\img\img0_05_2.jpg"): pagenum = 2
Case 2: img_stt2 = LoadPicture(App.Path & "\img\img0_05_3.jpg"): pagenum = 3
Case 3: img_stt2 = LoadPicture(App.Path & "\img\img0_05_4.jpg"): pagenum = 4
Case 4: img_stt2 = LoadPicture(App.Path & "\img\empty.jpg"): pagenum = 5
End Select

End Sub
