VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmSplitter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Splitter"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8400
   ControlBox      =   0   'False
   Icon            =   "FrmSplitter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1110
      Left            =   4740
      TabIndex        =   5
      Top             =   3135
      Visible         =   0   'False
      Width           =   2235
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   585
         Top             =   405
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   75
         Top             =   390
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   1200
         OleObjectBlob   =   "FrmSplitter.frx":1D2A
         Top             =   480
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4530
      Left            =   0
      TabIndex        =   6
      Top             =   -60
      Width           =   6855
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   135
         OleObjectBlob   =   "FrmSplitter.frx":2BE89
         TabIndex        =   7
         Top             =   300
         Width           =   3285
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   135
         OleObjectBlob   =   "FrmSplitter.frx":2BEF8
         TabIndex        =   8
         Top             =   555
         Width           =   3285
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmSplitter.frx":2BF5F
         TabIndex        =   9
         Top             =   825
         Width           =   3285
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   270
         Left            =   120
         OleObjectBlob   =   "FrmSplitter.frx":2BFD0
         TabIndex        =   10
         Top             =   1095
         Width           =   3285
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   135
         OleObjectBlob   =   "FrmSplitter.frx":2C059
         TabIndex        =   11
         Top             =   3945
         Width           =   3285
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   135
         OleObjectBlob   =   "FrmSplitter.frx":2C0E6
         TabIndex        =   12
         Top             =   4185
         Width           =   3285
      End
   End
   Begin VB.Frame Frame10 
      Height          =   4560
      Left            =   6855
      TabIndex        =   0
      Top             =   -75
      Width           =   1530
      Begin VB.CommandButton Command7 
         Caption         =   "Quit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   90
         TabIndex        =   4
         Top             =   3345
         Width           =   1365
      End
      Begin VB.CommandButton Command4 
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   90
         TabIndex        =   1
         Top             =   2295
         Width           =   1365
      End
      Begin VB.CommandButton Command3 
         Caption         =   "ReJoin File"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   90
         TabIndex        =   2
         Top             =   1245
         Width           =   1365
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Split File"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   90
         TabIndex        =   3
         Top             =   195
         Width           =   1365
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2760
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmSplitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim l_val As Integer
Private m_Trans As Class1

Dim intW As Integer
Dim intH As Integer
Dim BLFormStatus As Boolean
      
Private Sub Command2_Click()
strFileOpenDialogFileTitle = "Select the File to Split"
strFileOpenDialogFileExtension = "*.*"
frmSplit.Show vbModal

End Sub

Private Sub Command3_Click()
frmRejoin.Show vbModal
End Sub

Private Sub Command4_Click()
frmAbout.Show vbModal
End Sub

Private Sub Command7_Click()
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False

Set m_Trans = Nothing
BLFormStatus = False
Timer2.Enabled = True
End Sub

Private Sub Form_Load()
Skin1.ApplySkin Me.hwnd
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
           End
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
'               SetAlpha
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
