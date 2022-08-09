VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About FileSplitter"
   ClientHeight    =   3210
   ClientLeft      =   5160
   ClientTop       =   4515
   ClientWidth     =   5850
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2190
      Top             =   2775
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1680
      Top             =   2760
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "frmAbout.frx":0000
      Top             =   2760
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   390
      Left            =   4815
      TabIndex        =   7
      Top             =   2790
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   2790
      Left            =   0
      TabIndex        =   0
      Top             =   -60
      Width           =   5850
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   225
         Left            =   120
         OleObjectBlob   =   "frmAbout.frx":2A15F
         TabIndex        =   1
         Top             =   210
         Width           =   1470
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   225
         Left            =   120
         OleObjectBlob   =   "frmAbout.frx":2A1D8
         TabIndex        =   2
         Top             =   495
         Width           =   1470
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   225
         Left            =   120
         OleObjectBlob   =   "frmAbout.frx":2A245
         TabIndex        =   3
         Top             =   690
         Width           =   3330
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   225
         Left            =   120
         OleObjectBlob   =   "frmAbout.frx":2A2AC
         TabIndex        =   4
         Top             =   945
         Width           =   3330
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   225
         Left            =   135
         OleObjectBlob   =   "frmAbout.frx":2A33B
         TabIndex        =   5
         Top             =   1125
         Width           =   3330
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   885
         Left            =   120
         OleObjectBlob   =   "frmAbout.frx":2A3AA
         TabIndex        =   6
         Top             =   1845
         Width           =   5640
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   225
         Left            =   135
         OleObjectBlob   =   "frmAbout.frx":2A639
         TabIndex        =   8
         Top             =   1395
         Width           =   3330
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   225
         Left            =   120
         OleObjectBlob   =   "frmAbout.frx":2A6BC
         TabIndex        =   9
         Top             =   1590
         Width           =   4860
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_Trans As Class1
Dim l_val As Integer

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
           Me.Width = 5940
           Timer1.Enabled = False
           Timer2.Enabled = True
        Else
           Me.Width = Me.Width + intW
           Me.Left = Me.Left - (intW + 200)
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
               Me.Height = 3585
            Else
               Me.Height = Me.Height + intH
               Me.Top = Me.Top - (intH + 300)
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
