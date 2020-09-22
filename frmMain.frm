VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10065
   ClientLeft      =   1905
   ClientTop       =   2175
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   ScaleHeight     =   671
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   812
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox SS 
      AutoRedraw      =   -1  'True
      Height          =   9015
      Left            =   8040
      ScaleHeight     =   597
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   797
      TabIndex        =   0
      Top             =   8880
      Visible         =   0   'False
      Width           =   12015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Finish As Boolean
Dim LastX As Single, LastY As Single
Dim InZ As Boolean

Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Const DegToRad As Single = 3.14159275180032 / 180
Private Const MouseSens As Single = 10

Public Sub Form_Load()
    Me.Show
    LastX = 100
    LastY = 100
    If MsgBox("Do you wish to use Lights? This will make the Game Load much slower. (2-5 Minutes)", vbYesNo, "Lighting...") = vbYes Then
        Init
        InitGame True
    Else
        Init
        InitGame False
    End If
    DoEvents
    Do While Finish = False
        DoEvents
        MLoop
    Loop
    End
End Sub

Public Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If InZ = False Then
            Zoom True
            InZ = True
        Else
            Zoom False
            InZ = False
        End If
    End If
    If Button = 1 Then
        Shoot
    End If
End Sub

Public Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RotateCamera (LastX - X) * -DegToRad / MouseSens, (LastY - Y) * -DegToRad / MouseSens
    If X <> LastX Or Y <> LastY Then
        SetCursorPos LastX, LastY
    End If
End Sub

Public Sub Form_Unload(Cancel As Integer)
    Finish = True
    DoEvents
    KillApp
End Sub
