VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmFolderOpenDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4020
      OleObjectBlob   =   "frmFolderOpenDialog.frx":0000
      Top             =   180
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4995
      Top             =   210
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4515
      Top             =   210
   End
   Begin VB.Frame Frame1 
      Height          =   4395
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6090
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   105
         TabIndex        =   7
         Top             =   435
         Width           =   2940
      End
      Begin VB.FileListBox File1 
         Enabled         =   0   'False
         Height          =   3600
         Left            =   90
         TabIndex        =   5
         Top             =   750
         Width           =   2970
      End
      Begin VB.DirListBox Dir1 
         Height          =   3240
         Left            =   3045
         TabIndex        =   4
         Top             =   750
         Width           =   2985
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   3045
         TabIndex        =   3
         Top             =   4005
         Width           =   3000
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   225
         Left            =   135
         OleObjectBlob   =   "frmFolderOpenDialog.frx":2A15F
         TabIndex        =   6
         Top             =   210
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel Directory 
         Height          =   225
         Left            =   3105
         OleObjectBlob   =   "frmFolderOpenDialog.frx":2A1CC
         TabIndex        =   8
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   360
      Left            =   6135
      TabIndex        =   1
      Top             =   3420
      Width           =   1020
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   360
      Left            =   6135
      TabIndex        =   0
      Top             =   3885
      Width           =   1020
   End
End
Attribute VB_Name = "frmFolderOpenDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FSO As New FileSystemObject
Private m_Trans As Class1

Dim intW As Integer
Dim intH As Integer
Dim BLFormStatus As Boolean

Private Sub Command1_Click()
'Validation
'-----------------------------------------------------------
 If Trim(Text1.Text) = "" Then
     MsgBox "Error: There is an Error." & Chr(13) & _
            "Cause: No associate Folder has Selected" & Chr(13) & _
            "Resolution: Select the proper Folder or Click Cancel to cancel the Operation", vbCritical, "Connection: Error"
     Exit Sub
 End If
 
 strFolderOpenDialogFileName = Dir1.Path
' strFolderOpenDialogFileName = Replace(File1.Path & "\" & File1.FileName, "\\", "\")
 Unload Me
End Sub

Private Sub Command2_Click()
Set m_Trans = Nothing
BLFormStatus = False
Timer2.Enabled = True
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Dir1_Click()
Text1.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo ErrHandler
Dir1.Path = Drive1.Drive

Exit Sub
ErrHandler:
MsgBox "Error: There is an Error related to Drive specification" & Chr(13) & _
       "Error Number: " & Err.Number & Chr(13) & _
       "Description: " & Err.Description, vbCritical + vbExclamation, "Connection: Error"
End Sub

Private Sub File1_Click()
Text1.Text = File1.FileName
End Sub

Private Sub Form_Load()

Skin1.ApplySkin Me.hwnd

BLFormStatus = True

Me.Top = FrmSplitter.Top + 3500
Me.Left = FrmSplitter.Left + 5500

Me.Height = 1000
Me.Width = 1000

'Initialize the form
'--------------------------------------------------------------------
 Me.Caption = strFolderOpenDialogFileTitle
 File1.Pattern = strFolderOpenDialogFileExtension
 Me.Refresh
'--------------------------------------------------------------------
End Sub
Private Sub SetAlpha()
l_val = 220
Set m_Trans = New Class1
m_Trans.hwnd = Me.hwnd
m_Trans.Alpha = l_val

End Sub

Private Sub Timer1_Timer()
'Timer for the Width
If BLFormStatus = True Then
    If intW >= 1000 Then
           intW = 0
           Me.Width = 7335
           Timer1.Enabled = False
           Timer2.Enabled = True
        Else
           Me.Width = Me.Width + intW
           Me.Left = Me.Left - (intW + 50)
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
'               SetAlpha
               Me.Height = 4710
            Else
               Me.Height = Me.Height + intH
               Me.Top = Me.Top - (intH + 100)
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


