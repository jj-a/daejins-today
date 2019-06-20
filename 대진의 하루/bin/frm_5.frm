VERSION 5.00
Begin VB.Form frm_5 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "대진의 하루"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   Icon            =   "frm_5.frx":0000
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
      Interval        =   200
      Left            =   7440
      Top             =   5160
   End
   Begin VB.Label lbl_re 
      BackStyle       =   0  '투명
      Height          =   495
      Left            =   120
      MouseIcon       =   "frm_5.frx":8732
      MousePointer    =   99  '사용자 정의
      TabIndex        =   1
      Top             =   5340
      Width           =   1815
   End
   Begin VB.Label lbl 
      BackStyle       =   0  '투명
      Height          =   735
      Left            =   2640
      MouseIcon       =   "frm_5.frx":8A3C
      MousePointer    =   99  '사용자 정의
      TabIndex        =   0
      Top             =   2520
      Width           =   3180
   End
   Begin VB.Image img_5 
      Height          =   6000
      Left            =   0
      MouseIcon       =   "frm_5.frx":8D46
      MousePointer    =   99  '사용자 정의
      Picture         =   "frm_5.frx":9050
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8250
   End
End
Attribute VB_Name = "frm_5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frm_5->frm_6
'5교시:멀티미디어


Dim pagenum, pag


Private Sub Form_Activate()
If res5 = 1 Then
    img_5 = LoadPicture(App.Path & "\img\img5_02_over.jpg")
ElseIf res5 = 2 Then
    img_5 = LoadPicture(App.Path & "\img\img5_02.jpg")
ElseIf rt5 = 1 Then
    pag = 1: pagenum = 2
    rt5 = 0
Else
    img_5 = LoadPicture(App.Path & "\img\img5_00.jpg")
    pag = 0
End If
End Sub


Sub pagemove()

If pagenum = 1 And pag = 0 Then
tmr.Enabled = True
: img_5.MouseIcon = LoadPicture(App.Path & "\img\mus_normal.cur")
: lbl.MouseIcon = LoadPicture(App.Path & "\img\mus_normal.cur")
: lbl_re.MouseIcon = LoadPicture(App.Path & "\img\mus_normal.cur")
End If

If res5 = 2 Then
    res5 = 0
   frm_6.Show
   frm_5.Hide
ElseIf res5 = 1 Then
Else
    If pag = 0 Then
        Select Case pagenum
        Case 0: img_5 = LoadPicture(App.Path & "\img\img5_01_0_00.jpg"): pagenum = 1
        End Select
    ElseIf pag = 1 Then
        Select Case pagenum
        Case 0: img_5 = LoadPicture(App.Path & "\img\img5_01_3.jpg"): pagenum = 1
        Case 1: img_5 = LoadPicture(App.Path & "\img\img5_01_1.jpg"): pagenum = 2
: frm_5_stt.Show: frm_5.Hide
        Case 2: img_5 = LoadPicture(App.Path & "\img\img5_01_4_0.jpg"): pagenum = 3
        Case 3: img_5 = LoadPicture(App.Path & "\img\img5_01_4_1.jpg"): pagenum = 4
        Case 4: img_5 = LoadPicture(App.Path & "\img\img5_01_4_2.jpg"): pag = 2: pagenum = 0
: img_5.MouseIcon = LoadPicture(App.Path & "\img\mus_normal.cur")
: lbl_re.MouseIcon = LoadPicture(App.Path & "\img\mus_normal.cur")
        End Select
    End If
End If


End Sub

Private Sub img_5_Click()
Call pagemove
End Sub



Private Sub lbl_Click()
If pag >= 2 Then
Call sndplay("eff01.wav")
pag = 0
pagenum = 0
img_5.MouseIcon = LoadPicture(App.Path & "\img\mus_click.cur")
lbl_re.MouseIcon = LoadPicture(App.Path & "\img\mus_click.cur")
frm_5_game.Show
frm_5.Hide
Else
Call pagemove
End If
End Sub

Private Sub lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If pag >= 2 Then
img_5 = LoadPicture(App.Path & "\img\img5_01_4_2_2.jpg")
End If
End Sub

Private Sub img_5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If pag >= 2 Then
img_5 = LoadPicture(App.Path & "\img\img5_01_4_2.jpg")
ElseIf res5 = 1 Then
img_5 = LoadPicture(App.Path & "\img\img5_02_over.jpg")
img_5.MouseIcon = LoadPicture(App.Path & "\img\mus_normal.cur")
lbl.MouseIcon = LoadPicture(App.Path & "\img\mus_normal.cur")
End If
End Sub


Private Sub lbl_re_Click()
If res5 = 1 Then
Call sndplay("eff01.wav")
res5 = 0
img_5.MouseIcon = LoadPicture(App.Path & "\img\mus_click.cur")
lbl.MouseIcon = LoadPicture(App.Path & "\img\mus_click.cur")
lbl_re.MouseIcon = LoadPicture(App.Path & "\img\mus_click.cur")
Call Form_Activate
Else
Call pagemove
End If
End Sub

Private Sub lbl_re_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If res5 = 1 Then
img_5 = LoadPicture(App.Path & "\img\img5_02_over_1.jpg")
End If
End Sub


Private Sub tmr_Timer()
    Select Case pagenum
    Case 1: img_5 = LoadPicture(App.Path & "\img\img5_01_0_01.jpg"): pagenum = 2
    Case 2: img_5 = LoadPicture(App.Path & "\img\img5_01_0_02.jpg"): pagenum = 3
    Case 3: img_5 = LoadPicture(App.Path & "\img\img5_01_0_03.jpg"): pagenum = 4
    Case 4: img_5 = LoadPicture(App.Path & "\img\img5_01_0_04.jpg"): pagenum = 5
    Case 5: img_5 = LoadPicture(App.Path & "\img\img5_01_0_05.jpg"): pagenum = 6
    Case 6: img_5 = LoadPicture(App.Path & "\img\img5_01_0_06.jpg"): pagenum = 7
    Case 7: img_5 = LoadPicture(App.Path & "\img\img5_01_0_07.jpg"): pagenum = 8
    Case 8: img_5 = LoadPicture(App.Path & "\img\img5_01_0_08.jpg"): pagenum = 9
    Case 9: img_5 = LoadPicture(App.Path & "\img\img5_01_0_09.jpg"): pagenum = 10
    Case 10: img_5 = LoadPicture(App.Path & "\img\img5_01_0_10.jpg"): pagenum = 11
    Case 11: img_5 = LoadPicture(App.Path & "\img\img5_01_0_11.jpg"): pagenum = 12
    Case 12: img_5 = LoadPicture(App.Path & "\img\img5_01_0_12.jpg"): pagenum = 13
    Case 13: img_5 = LoadPicture(App.Path & "\img\img5_01_0_13.jpg"): pagenum = 14
    Case 14: img_5 = LoadPicture(App.Path & "\img\img5_01_0_14.jpg"): pagenum = 15
    Case 15: img_5 = LoadPicture(App.Path & "\img\img5_01_0_15.jpg"): pagenum = 16
    Case 16: img_5 = LoadPicture(App.Path & "\img\img5_01_0_16.jpg"): pagenum = 17
    Case 17: img_5 = LoadPicture(App.Path & "\img\img5_01_0_17.jpg"): pagenum = 18
    Case 18: img_5 = LoadPicture(App.Path & "\img\img5_01_0_18.jpg"): pag = 1: pagenum = 0
: img_5.MouseIcon = LoadPicture(App.Path & "\img\mus_click.cur")
: lbl.MouseIcon = LoadPicture(App.Path & "\img\mus_click.cur")
: lbl_re.MouseIcon = LoadPicture(App.Path & "\img\mus_click.cur")
: tmr.Enabled = False
    End Select
End Sub





