VERSION 5.00
Begin VB.Form frm_exit 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  '고정 도구 창
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   3720
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frm_exit.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frm_exit.frx":8732
   MousePointer    =   99  '사용자 정의
   Moveable        =   0   'False
   Picture         =   "frm_exit.frx":8FFC
   ScaleHeight     =   2565
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Label lbl_no 
      BackStyle       =   0  '투명
      Height          =   495
      Left            =   1920
      MouseIcon       =   "frm_exit.frx":13265
      MousePointer    =   99  '사용자 정의
      TabIndex        =   1
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lbl_yes 
      BackStyle       =   0  '투명
      Height          =   495
      Left            =   720
      MouseIcon       =   "frm_exit.frx":1356F
      MousePointer    =   99  '사용자 정의
      TabIndex        =   0
      Top             =   1560
      Width           =   735
   End
End
Attribute VB_Name = "frm_exit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lbl_no_Click()
Call sndplay("eff01.wav")
frm_exit.Hide
End Sub

Private Sub lbl_yes_Click()
Call sndplay("eff01.wav")
End
End Sub
