VERSION 5.00
Begin VB.Form frm_help 
   BorderStyle     =   4  '고정 도구 창
   ClientHeight    =   2730
   ClientLeft      =   11220
   ClientTop       =   6540
   ClientWidth     =   3750
   Icon            =   "frm_help.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2820.569
   ScaleMode       =   0  '사용자
   ScaleWidth      =   4146.635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Image img_help 
      Height          =   2700
      Left            =   0
      MouseIcon       =   "frm_help.frx":8732
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_help.frx":8A3C
      Top             =   0
      Width           =   3750
   End
End
Attribute VB_Name = "frm_help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub img_help_Click()
Call sndplay("eff01.wav")
frm_help.Hide
End Sub
