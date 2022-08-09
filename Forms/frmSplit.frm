VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmSplit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Split File"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
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
   ScaleHeight     =   3810
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2895
      Top             =   3330
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1575
      Top             =   3360
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2085
      Top             =   3375
   End
   Begin VB.CommandButton CmdSClose 
      Caption         =   "Close"
      Height          =   390
      Left            =   4830
      TabIndex        =   11
      Top             =   3390
      Width           =   945
   End
   Begin VB.CommandButton CmdSPlit 
      Caption         =   "Split"
      Enabled         =   0   'False
      Height          =   390
      Left            =   3960
      TabIndex        =   12
      Top             =   3390
      Width           =   885
   End
   Begin VB.Frame Frame1 
      Caption         =   "Source"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   5835
      Begin VB.CommandButton CmdSource 
         Caption         =   "..."
         Height          =   315
         Left            =   4815
         TabIndex        =   10
         Top             =   525
         Width           =   960
      End
      Begin VB.TextBox TxtSource 
         Height          =   315
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   525
         Width           =   4710
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   225
         Left            =   105
         OleObjectBlob   =   "frmSplit.frx":0000
         TabIndex        =   13
         Top             =   240
         Width           =   3435
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Destination"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   1020
      Visible         =   0   'False
      Width           =   5835
      Begin VB.TextBox TxtDestination 
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   480
         Width           =   4680
      End
      Begin VB.CommandButton CmdDestination 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   330
         Left            =   4800
         TabIndex        =   6
         Top             =   465
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   225
         Left            =   120
         OleObjectBlob   =   "frmSplit.frx":0099
         TabIndex        =   14
         Top             =   240
         Width           =   3435
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Split Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   1995
      Visible         =   0   'False
      Width           =   5835
      Begin VB.OptionButton OptDefault 
         Caption         =   "Default Split File Size (500 KB)"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   3015
      End
      Begin VB.OptionButton OptSpan 
         Caption         =   "Span Files to a Floppy Disk (1.36 MB)"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   3015
      End
      Begin VB.OptionButton OptCustom 
         Caption         =   "Custom Size"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox TxtSize 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   975
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   225
         Left            =   345
         OleObjectBlob   =   "frmSplit.frx":0128
         TabIndex        =   15
         Top             =   1005
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   225
         Left            =   2865
         OleObjectBlob   =   "frmSplit.frx":0195
         TabIndex        =   16
         Top             =   1005
         Width           =   975
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   15
      OleObjectBlob   =   "frmSplit.frx":01F6
      Top             =   3360
   End
End
Attribute VB_Name = "frmSplit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 '*******************************************************
 'This program is created by kshitij kumar
 'For any assistance mail me on krayknot@yahoo.com
 '*******************************************************
 
 Option Explicit
 Dim FSO As New FileSystemObject
      
Private m_Trans As Class1
Dim l_val As Integer

Dim intW As Integer
Dim intH As Integer
Dim BLFormStatus As Boolean

Private Sub Command1_Click()

End Sub

Private Sub CmdDestination_Click()
'Opens the selected file from the common dialog box
 frmFolderOpenDialog.Show vbModal
 TxtDestination.Text = strFolderOpenDialogFileName

End Sub


Private Sub CmdSClose_Click()
CmdSource.Enabled = False
CmdDestination.Enabled = False

Set m_Trans = Nothing
BLFormStatus = False
Timer2.Enabled = True
End Sub

Private Sub CmdSource_Click()
'Opens the selected file from the common dialog box
 frmFileOpenDialog.Show vbModal
 TxtSource.Text = strFileOpenDialogFileName
End Sub

Private Sub CmdSPlit_Click()

strCommand = "SPLIT"
frmBar.Show vbModal

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

Private Sub OptCustom_Click()
If OptCustom.Value = True Then
   TxtSize.Enabled = True
   TxtSize.BackColor = vbWhite
   TxtSize.SetFocus
Else
   TxtSize.Enabled = False
End If
End Sub

Private Sub OptDefault_Click()
TxtSize.Enabled = False
TxtSize.BackColor = &HE0E0E0
End Sub

Private Sub OptSpan_Click()
TxtSize.Enabled = False
TxtSize.BackColor = &HE0E0E0
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
           Me.Refresh
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
           Frame1.Visible = True
           Frame2.Visible = True
           Frame3.Visible = True
           Me.Refresh
           Timer2.Enabled = False
           SetAlpha
           Me.Height = 4185
        Else
           Me.Height = Me.Height + intH
           Me.Top = Me.Top - (intH + 300)
           Me.Refresh
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

Private Sub TxtDestination_Change()
If Trim(TxtSource.Text) <> "" And Trim(TxtDestination.Text) <> "" Then
    CmdSPlit.Enabled = True
End If
End Sub

Private Sub TxtSize_Change()
If Not IsNumeric(TxtSize.Text) Then
   TxtSize.Text = ""
   TxtSize.SetFocus
End If
End Sub
Private Sub SetAlpha()
l_val = 220
Set m_Trans = New Class1
m_Trans.hwnd = Me.hwnd
m_Trans.Alpha = l_val

End Sub

Private Sub TxtSource_Change()
CmdDestination.Enabled = True

If Trim(TxtSource.Text) <> "" And Trim(TxtDestination.Text) <> "" Then
    CmdSPlit.Enabled = True
End If
End Sub
