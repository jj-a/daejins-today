VERSION 5.00
Begin VB.Form frm_stt 
   BackColor       =   &H000000FF&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "대진의 하루"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   Icon            =   "frm_stt.frx":0000
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
      MouseIcon       =   "frm_stt.frx":8732
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_stt.frx":8FFC
      Top             =   200
      Width           =   3000
   End
   Begin VB.Image img_bt 
      Height          =   615
      Left            =   2880
      MouseIcon       =   "frm_stt.frx":C7B8
      MousePointer    =   99  '사용자 정의
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Image img_stt 
      Height          =   6000
      Left            =   0
      MouseIcon       =   "frm_stt.frx":CAC2
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_stt.frx":CDCC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8250
   End
End
Attribute VB_Name = "frm_stt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'frm_stt->frm_stt2

Dim pag, num


Private Sub Form_Activate()
img_stt = LoadPicture(App.Path & "\img\img0_00_0.jpg")
End Sub

Sub pagemove()

num = num + 1

If pag = 0 Then
    Select Case num
    Case 1: img_stt = LoadPicture(App.Path & "\img\img0_00_1.jpg")
    Case 2: img_stt = LoadPicture(App.Path & "\img\img0_00_2.jpg")
    Case 3: img_stt = LoadPicture(App.Path & "\img\img0_00_3.jpg")
    Case 4: img_stt = LoadPicture(App.Path & "\img\img0_00_4.jpg")
    Case 5: img_stt = LoadPicture(App.Path & "\img\img0_00_5.jpg")
    Case 6: img_bt = LoadPicture(App.Path & "\img\img0_00_s1.jpg")
: img_stt.MouseIcon = LoadPicture(App.Path & "\img\mus_normal.cur")
    End Select
ElseIf pag = 1 Then
    Select Case num
    Case 1: img_stt = LoadPicture(App.Path & "\img\img0_01_1.jpg")
    Case 2: img_stt = LoadPicture(App.Path & "\img\img0_01_2.jpg")
    Case 3: img_stt = LoadPicture(App.Path & "\img\img0_01_3.jpg")
    Case 4: img_stt = LoadPicture(App.Path & "\img\img0_01_4.jpg")
    Case 5: img_stt = LoadPicture(App.Path & "\img\img0_01_5.jpg")
    Case 6: img_stt = LoadPicture(App.Path & "\img\img0_01_6.jpg"): pag = 2: num = 0
    End Select
ElseIf pag = 2 Then
    Select Case num
    Case 1: img_stt = LoadPicture(App.Path & "\img\img0_02_0.jpg")
    Case 2: img_stt = LoadPicture(App.Path & "\img\img0_02_1.jpg")
    Case 3: img_stt = LoadPicture(App.Path & "\img\img0_02_2.jpg"): pag = 3: num = 0
    End Select
ElseIf pag = 3 Then
    Select Case num
    Case 1: img_stt = LoadPicture(App.Path & "\img\img0_03_0.jpg")
    Case 2: img_stt = LoadPicture(App.Path & "\img\img0_03_1.jpg")
    Case 3: img_stt = LoadPicture(App.Path & "\img\img0_03_2.jpg")
    Case 4: img_stt = LoadPicture(App.Path & "\img\img0_03_3.jpg"): pag = 4: num = 0
    End Select
ElseIf pag = 4 Then
    Select Case num
    Case 1: img_stt = LoadPicture(App.Path & "\img\img0_04_0.jpg")
    Case 2: img_stt = LoadPicture(App.Path & "\img\img0_04_1.jpg")
    Case 3: img_stt = LoadPicture(App.Path & "\img\img0_04_2.jpg"): pag = 5: num = 0
    End Select
Else
    frm_stt2.Show
    frm_stt.Hide
    pag = 0
    num = 0
End If
    
End Sub

Private Sub img_stt_Click()
Call pagemove
End Sub


Private Sub img_bt_Click()
If pag = 0 And num >= 6 Then
Call sndplay("eff01.wav")
pag = 1
num = 0
img_bt.Visible = False
img_stt = LoadPicture(App.Path & "\img\img0_01_0.jpg")
img_stt.MouseIcon = LoadPicture(App.Path & "\img\mus_click.cur")
Else
Call pagemove
End If
End Sub

Private Sub img_bt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If pag = 0 And num >= 6 Then
img_bt = LoadPicture(App.Path & "\img\img0_00_s2.jpg")
End If
End Sub


Private Sub img_stt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If pag = 0 And num >= 6 Then
img_bt = LoadPicture(App.Path & "\img\img0_00_s1.jpg")
End If
End Sub
