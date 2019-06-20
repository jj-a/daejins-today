VERSION 5.00
Begin VB.Form frm_8_tuto 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "대진의 하루"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   Icon            =   "frm_8_tuto.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Image img_8 
      Height          =   6000
      Left            =   0
      MouseIcon       =   "frm_8_tuto.frx":8732
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_8_tuto.frx":8A3C
      Top             =   0
      Width           =   8250
   End
End
Attribute VB_Name = "frm_8_tuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub img_8_Click()
rt8 = 1
frm_8.Show
frm_8_tuto.Hide
End Sub
