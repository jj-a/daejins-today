VERSION 5.00
Begin VB.Form frm_credit 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "대진의 하루"
   ClientHeight    =   6000
   ClientLeft      =   10335
   ClientTop       =   6240
   ClientWidth     =   8250
   Icon            =   "frm_credit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Image img_goback 
      Height          =   330
      Left            =   360
      MouseIcon       =   "frm_credit.frx":8732
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_credit.frx":8A3C
      Top             =   5280
      Width           =   810
   End
   Begin VB.Image img_n 
      Height          =   180
      Left            =   5320
      MouseIcon       =   "frm_credit.frx":BB88
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_credit.frx":BE92
      Top             =   4940
      Width           =   525
   End
   Begin VB.Image img_b 
      Height          =   195
      Left            =   1780
      MouseIcon       =   "frm_credit.frx":E41F
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_credit.frx":E729
      Top             =   4940
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image img_credit 
      Height          =   6000
      Left            =   0
      MouseIcon       =   "frm_credit.frx":10D23
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_credit.frx":115ED
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8250
   End
End
Attribute VB_Name = "frm_credit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim pagenum

Private Sub Form_Activate()
img_credit = LoadPicture(App.Path & "\img\credit1.jpg")
img_b.Visible = False
img_n.Visible = True
img_b.MouseIcon = LoadPicture(App.Path & "\img\mus_normal.cur")
img_b.MouseIcon = LoadPicture(App.Path & "\img\mus_click.cur")
End Sub

Sub pagemove()

Select Case pagenum
Case 0: img_credit = LoadPicture(App.Path & "\img\credit1.jpg")
: img_b.MouseIcon = LoadPicture(App.Path & "\img\mus_normal.cur")
: img_b.Visible = False
Case 1: img_credit = LoadPicture(App.Path & "\img\credit2.jpg")
: img_b.MouseIcon = LoadPicture(App.Path & "\img\mus_click.cur")
: img_b.Visible = True
Case 2: img_credit = LoadPicture(App.Path & "\img\credit3.jpg")
: img_n.MouseIcon = LoadPicture(App.Path & "\img\mus_click.cur")
: img_n.Visible = True
Case 3: img_credit = LoadPicture(App.Path & "\img\credit4.jpg")
: img_n.MouseIcon = LoadPicture(App.Path & "\img\mus_normal.cur")
: img_n.Visible = False
End Select

End Sub

Private Sub img_b_Click()
Call sndplay("eff01.wav")
pagenum = pagenum - 1
Call pagemove
End Sub

Private Sub img_credit_Click()

End Sub

Private Sub img_n_Click()
Call sndplay("eff01.wav")
pagenum = pagenum + 1
Call pagemove
End Sub

Private Sub img_goback_Click()
Call sndplay("eff01.wav")
frm_credit.Hide
frm_main.Show
pagenum = 0
End Sub






