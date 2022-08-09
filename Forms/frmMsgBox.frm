VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmMsgBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Information"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5445
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1620
      Top             =   2310
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2130
      Top             =   2325
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4290
      TabIndex        =   1
      Top             =   2295
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   2280
      Left            =   0
      TabIndex        =   0
      Top             =   -60
      Width           =   5445
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   1860
         Left            =   150
         OleObjectBlob   =   "frmMsgBox.frx":0000
         TabIndex        =   2
         Top             =   270
         Width           =   5145
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   60
      OleObjectBlob   =   "frmMsgBox.frx":006C
      Top             =   2310
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim l_val As Integer
Private m_Trans As Class1

Dim intW As Integer
Dim intH As Integer
Dim BLFormStatus As Boolean

Private Sub Command1_Click()
Set m_Trans = Nothing
BLFormStatus = False
Timer2.Enabled = True
End Sub

Private Sub Form_Load()
Skin1.ApplySkin Me.hwnd
SkinLabel1.Caption = strMessage

BLFormStatus = True

Me.Top = FrmSplitter.Top + 3500
Me.Left = FrmSplitter.Left + 5500

Me.Height = 1000
Me.Width = 1000


Me.Refresh
End Sub

Private Sub Timer1_Timer()
'Timer for the Width
If BLFormStatus = True Then
    If intW >= 1000 Then
           intW = 0
           Me.Width = 5535
           Timer1.Enabled = False
           Timer2.Enabled = True
        Else
           Me.Width = Me.Width + intW
           Me.Left = Me.Left - (intW + 10)
    End If
    intW = intW + 100

ElseIf BLFormStatus = False Then
    If intW < -1000 Then
           intW = 0
           Timer1.Enabled = False
           Unload Me
        Else
           Me.Width = Me.Width + intW
           Me.Left = Me.Left + 500
    End If
    intW = intW - 100

End If
End Sub

Private Sub Timer2_Timer()
'Timer for the height
If BLFormStatus = True Then
    If intH >= 700 Then
           intH = 0
           Timer2.Enabled = False
           SetAlpha
           Me.Height = 3210
        Else
           Me.Height = Me.Height + intH
           Me.Top = Me.Top - (intH)
    End If
    intH = intH + 100
    
ElseIf BLFormStatus = False Then
    If intH < -700 Then
           intH = 0
           Timer1.Enabled = True
           Timer2.Enabled = False
        Else
           Me.Height = Me.Height + intH
           Me.Top = Me.Top + 500
    End If
    intH = intH - 100
End If
End Sub
Private Sub SetAlpha()
l_val = 220
Set m_Trans = New Class1
m_Trans.hwnd = Me.hwnd
m_Trans.Alpha = l_val

End Sub
