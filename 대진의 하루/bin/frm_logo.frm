VERSION 5.00
Begin VB.Form frm_logo 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "대진의 하루"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   Icon            =   "frm_logo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frm_logo.frx":8732
   MousePointer    =   99  '사용자 정의
   Moveable        =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.Image img_logo 
      Height          =   6000
      Left            =   0
      MouseIcon       =   "frm_logo.frx":8FFC
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_logo.frx":98C6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8250
   End
End
Attribute VB_Name = "frm_logo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim pagenum

Private Sub Form_Activate()
Call sndplay("snd01.wav") '재생
res1 = 0: res2 = 0: res3 = 0: res4 = 0: res5 = 0: res6 = 0: res7 = 0: res8 = 0
img_logo = LoadPicture(App.Path & "\img\stt_0.jpg")
tmr.Enabled = True
End Sub


Private Sub tmr_Timer()
    Select Case pagenum
    Case 0: img_logo = LoadPicture(App.Path & "\img\stt_0.jpg"): pagenum = 1
    Case 1: img_logo = LoadPicture(App.Path & "\img\stt_1.jpg"): pagenum = 2
    Case 2: img_logo = LoadPicture(App.Path & "\img\stt_2.jpg"): pagenum = 3
    tmr.Interval = 4000
    Case 3: img_logo = LoadPicture(App.Path & "\img\stt_0.jpg"): pagenum = 4
    tmr.Interval = 1000
    Case 4: img_logo = LoadPicture(App.Path & "\img\stt_3.jpg"): pagenum = 5
    Case 5: img_logo = LoadPicture(App.Path & "\img\stt_4.jpg"): pagenum = 6
    tmr.Interval = 4000
    Case 6: img_logo = LoadPicture(App.Path & "\img\stt_0.jpg"): pagenum = 7
    tmr.Interval = 1000
    Case 7: img_logo = LoadPicture(App.Path & "\img\logo1.jpg"): pagenum = 8
    Case 8: img_logo = LoadPicture(App.Path & "\img\logo2.jpg"): pagenum = 9
    Case 9: img_logo = LoadPicture(App.Path & "\img\logo3.jpg"): pagenum = 10
    Case 10: img_logo = LoadPicture(App.Path & "\img\logo4.jpg"): pagenum = 11
    Case 11: img_logo = LoadPicture(App.Path & "\img\logo5.jpg"): pagenum = 12
    Case 12: img_logo = LoadPicture(App.Path & "\img\stt_0.jpg"): pagenum = 13
    Case 13: tmr.Enabled = False: frm_main.Show: frm_logo.Hide: pagenum = 0
    End Select
End Sub
