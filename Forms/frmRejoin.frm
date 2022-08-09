VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmRejoin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rejoin Splitted Files"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2445
      Top             =   3240
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2955
      Top             =   3255
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   210
      Top             =   3285
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4920
      TabIndex        =   8
      Top             =   3300
      Width           =   840
   End
   Begin VB.CommandButton CmdRejoin 
      Caption         =   "Rejoin"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4050
      TabIndex        =   9
      Top             =   3300
      Width           =   885
   End
   Begin VB.Frame Frame5 
      Caption         =   "Source of Splitted Header File"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5835
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   5160
      End
      Begin VB.CommandButton CmdJsource 
         Caption         =   "..."
         Height          =   330
         Left            =   5280
         TabIndex        =   6
         Top             =   465
         Width           =   375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   225
         Left            =   135
         OleObjectBlob   =   "frmRejoin.frx":0000
         TabIndex        =   10
         Top             =   240
         Width           =   3435
      End
   End
   Begin VB.Frame Frame6 
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
      Height          =   900
      Left            =   0
      TabIndex        =   2
      Top             =   900
      Width           =   5835
      Begin VB.CommandButton Command5 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   330
         Left            =   5265
         TabIndex        =   4
         Top             =   465
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   465
         Width           =   5160
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   225
         Left            =   120
         OleObjectBlob   =   "frmRejoin.frx":008F
         TabIndex        =   11
         Top             =   240
         Width           =   3435
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   1785
      Width           =   5835
      Begin VB.FileListBox File1 
         Height          =   675
         Left            =   4980
         TabIndex        =   1
         Top             =   180
         Visible         =   0   'False
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel LAbel7 
         Height          =   225
         Left            =   120
         OleObjectBlob   =   "frmRejoin.frx":010A
         TabIndex        =   12
         Top             =   285
         Width           =   4770
      End
      Begin ACTIVESKINLibCtl.SkinLabel Label8 
         Height          =   225
         Left            =   120
         OleObjectBlob   =   "frmRejoin.frx":0173
         TabIndex        =   13
         Top             =   555
         Width           =   4770
      End
      Begin ACTIVESKINLibCtl.SkinLabel LAbel9 
         Height          =   225
         Left            =   120
         OleObjectBlob   =   "frmRejoin.frx":01E6
         TabIndex        =   14
         Top             =   930
         Width           =   4770
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   885
      OleObjectBlob   =   "frmRejoin.frx":025D
      Top             =   3240
   End
End
Attribute VB_Name = "frmRejoin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim l_val As Integer
Private m_Trans As Class1

Dim intW As Integer
Dim intH As Integer
Dim BLFormStatus As Boolean

Dim FSO As New FileSystemObject

Private Sub CmdClose_Click()
CmdJsource.Enabled = True
Set m_Trans = Nothing
BLFormStatus = False
Timer2.Enabled = True
End Sub

Private Sub CmdJsource_Click()
'Opens the selected file from the common dialog box
 strFileOpenDialogFileExtension = "Adminmain.Split"
 frmFileOpenDialog.Show vbModal
 Text2.Text = strFileOpenDialogFileName
 Text3.Text = Replace(Text2.Text, Reverse(Mid$(Reverse(Text2.Text), 1, InStr(1, Reverse(Text2.Text), "\") - 1)), " ")
End Sub

Private Sub CmdRejoin_Click()
strCommand = "REJOIN"
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

Private Sub Text2_Change()
If Trim(Text2.Text) <> "" Then
    CmdRejoin.Enabled = True
End If
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
               Me.Height = 4095
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

