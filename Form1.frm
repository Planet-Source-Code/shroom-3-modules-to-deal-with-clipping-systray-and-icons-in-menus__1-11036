VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4395
   ClientLeft      =   165
   ClientTop       =   750
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer RemClip 
      Interval        =   500
      Left            =   3600
      Top             =   1920
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   0
      Width           =   4815
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New..."
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuSysTray 
      Caption         =   "SysTray"
      Visible         =   0   'False
      Begin VB.Menu mnuSysTrayExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Not commented too well, seeing as how I keep a majority
'of the work in my head on small projects.
'Pretty self explanatory otherwise!

Private Sub Form_Load()
    AddToTray Me, "Example", Me.Icon
    SetClipVars 4095, 8295
    'I haven't figured out how to mask the icon
    'in the menu yet, hehe =)
    'SetMenuIcon (form).hwnd,Index of Main Menu, Index of Actual Menu Item, Flags, Unchecked Bitmap Handle, Checked Bitmap Handle
    SetMenuIcon Me.hwnd, 0, 2, 0, Picture1.Picture, Picture1.Picture
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim Message As Long
   On Error Resume Next
    Message = X / Screen.TwipsPerPixelX
    Select Case Message
        Case WM_RBUTTONUP
            'Something useful I just found out:
            ' You need to verify the height, otherwise
            ' it'll pop up the menu mid-form, if the
            ' form is big enough
            temp = GetY
            If temp > (Screen.Height / Screen.TwipsPerPixelY) - 30 Then
                PopupMenu mnuSysTray
            End If
    End Select
End Sub

Private Sub Form_Resize()
    ClipForForm Me, 4095, 8295
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RemoveClipping
    RemoveFromTray
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuSysTrayExit_Click()
    If MsgBox("Exit?", vbOKCancel, "System Tray Exit") = vbOK Then Unload Me
End Sub

Private Sub RemClip_Timer()
    'This Removes the automatic clipping that
    ' occurs when the form loads. Someone find a work-
    ' around and email me! "Lawrence Naegle" <hotrocket@uswet.net>
    RemoveClipping
    RemClip.Enabled = False
End Sub
