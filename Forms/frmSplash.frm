VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   2715
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   5220
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   0
      Top             =   810
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   405
   End
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   0
      Top             =   0
   End
   Begin VB.Image Image1 
      Height          =   2715
      Left            =   0
      Picture         =   "frmSplash.frx":000C
      Top             =   0
      Width           =   5220
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim l_val As Integer
Private m_Trans As Class1
Private Sub Command1_Click()
End
End Sub
Private Sub Form_Load()
l_val = 0
Set m_Trans = New Class1
m_Trans.hwnd = Me.hwnd
m_Trans.Alpha = l_val
End Sub

Private Sub Timer1_Timer()
If l_val >= 255 Then
    Timer1.Enabled = False
    Timer2.Enabled = True
    Exit Sub
Else
    m_Trans.Alpha = l_val
End If
l_val = l_val + 5

End Sub
Private Sub Timer2_Timer()
Static i As Integer
If i >= 2 Then
 FrmSplitter.Show vbModal
 'FrmSplitter.Show 0, frmSplash
 Unload Me
 'frmSplash.Show
 Timer2.Interval = 0
 Timer3.Enabled = True
End If
i = i + 1
End Sub
Private Sub Timer3_Timer()
If l_val <= 0 Then
    Timer3.Enabled = False
    Exit Sub
Else
    m_Trans.Alpha = l_val
End If
l_val = l_val - 5

End Sub


