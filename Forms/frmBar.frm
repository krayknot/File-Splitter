VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmBar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Processing Status"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   405
      Left            =   2655
      TabIndex        =   3
      Top             =   750
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   405
      Left            =   1440
      TabIndex        =   2
      Top             =   750
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   15
      Top             =   885
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   525
      Top             =   900
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4440
      OleObjectBlob   =   "frmBar.frx":0000
      Top             =   60
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   195
      Left            =   210
      OleObjectBlob   =   "frmBar.frx":2A15F
      TabIndex        =   1
      Top             =   210
      Width           =   3555
   End
   Begin File_Splitter.UserControl1 ProgressBAr 
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   510
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   6956042
   End
End
Attribute VB_Name = "frmBar"
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

Private Sub Command1_Click()
Dim fname As String, F_Size As String
Dim IntCnt As Long, i As Integer
Dim InputFile As Integer, OutputFile As Integer, MPointer As Integer
Dim IFIleName As String, OFileName As String, MFileName As String
Dim LPercent As Long
Dim StrMainFileName As String
Dim LFileSize As Long
  
Close 'Closes all open files

If frmSplit.OptDefault.Value = True Then
   LFileSize = 512000
ElseIf frmSplit.OptSpan.Value = True Then
       LFileSize = 1433600
ElseIf frmSplit.OptCustom.Value = True Then
       If frmSplit.TxtSize.Text = "" Then
          MsgBox "Please enter the split file size."
          frmSplit.TxtSize.SetFocus
          Exit Sub
       End If
       LFileSize = Val(frmSplit.TxtSize.Text)
End If

'Check that whether the txtsource and txtdestination are empty or not
 If frmSplit.TxtSource.Text = "" Or frmSplit.TxtDestination.Text = "" Then
    MsgBox "One of the Text Box [Source File Name] is empty."
    frmSplit.TxtSource.SetFocus
    Exit Sub
 End If
  
 If frmSplit.TxtDestination.Text = "" Then
    MsgBox "One of the Text Box [Destination Name] is empty."
    frmSplit.TxtDestination.SetFocus
    Exit Sub
 End If
  
 IntCnt = 1
 IFIleName = frmSplit.TxtSource.Text
 InputFile = FreeFile
 Open IFIleName For Binary As InputFile 'opens a file as a Source
 
 F_Size = String(LFileSize, " ") 'How much data has to be read
  
'Creates a different folder for each split project
 For i = 1 To 100
     If Not FSO.FolderExists(frmSplit.TxtDestination.Text & "\Admin_Split" & i) Then
        FSO.CreateFolder (frmSplit.TxtDestination.Text & "\Admin_Split" & i)
        Exit For
     End If
 Next i
   
 MFileName = frmSplit.TxtDestination.Text & "Admin_Split" & i & "\AdminMain" & ".Split"
 For IntCnt = 1 To (FileLen(frmSplit.TxtSource.Text) / LFileSize) + 1
     OutputFile = FreeFile 'FreeFile
     OFileName = frmSplit.TxtDestination.Text & "Admin_Split" & i & "\Admin" & IntCnt & ".Split"
     Open OFileName For Binary Access Write Lock Read As OutputFile 'opens a file for putting the data read from Source
        
     If FileLen(IFIleName) - Loc(OutputFile) < LFileSize Then F_Size = String(1024, " ")
     
     Get #InputFile, , F_Size
     Put #OutputFile, , F_Size
     Close #OutputFile
     
     LPercent = (LFileSize / FileLen(frmSplit.TxtSource.Text)) * 100
     If (ProgressBAr.Value + LPercent) > 100 Then
        ProgressBAr.Value = 100
     Else
        ProgressBAr.Value = frmBar.ProgressBAr.Value + LPercent
     End If
 
 Next
   
 'StrMainFileName = Reverse(Mid$(Reverse(frmSplit.CommonDialog1.FileName), 1, InStr(1, Reverse(frmSplit.CommonDialog1.FileName), "\") - 1))
 StrMainFileName = Reverse(Mid$(Reverse(strFileOpenDialogFileName), 1, InStr(1, Reverse(strFileOpenDialogFileName), "\") - 1))
  
 MPointer = FreeFile
 Open MFileName For Output As MPointer
 Print #MPointer, "<Filename> " & StrMainFileName & " <Split File Numbers> " & IntCnt - 1
 Close #MPointer
  
 ProgressBAr.Value = 0
 
 strMessage = "The File has splitted successfully" & Chr(13) & "Please Check your Destination Folder."
 BLFormStatus = False
 Timer2.Enabled = True
 'Set m_Trans = Nothing
End Sub

Private Sub Command2_Click()
Dim IntCnt As Integer, IntCount As Integer
Dim IFIleName As String, StrMyString As String, OFileName As String
Dim StrDFileName As String, BlDFileName As Boolean
Dim IntFileNum As Integer, BlFileNum As Boolean
Dim IntEndLocation As Integer
Dim strTemp As String
Dim F_Size As String
Dim LFileSize As Long

Close 'Closes all open files
'LFileSize = 512000

If Trim(frmRejoin.Text2.Text) = "" Then
   MsgBox "Error: Cannot Rejoin the File." & Chr(13) & _
          "Cause: No Header file Selected." & Chr(13) & _
          "Resolution: Select the Header First by the name Adminmain.Split", vbCritical, "File Selection Error"
    frmRejoin.Text2.SetFocus
    Exit Sub
End If

If Trim(frmRejoin.Text3.Text) = "" Then
   MsgBox "Error: Cannot Rejoin the File." & Chr(13) & _
          "Cause: No Destination folder path Selected." & Chr(13) & _
          "Resolution: Select the Destination Folder to make the Resultant file.", vbCritical, "File Selection Error"
   frmRejoin.Text3.SetFocus
   Exit Sub
End If


'Sets the path of the filelistbox
 frmRejoin.File1.Path = Replace(frmRejoin.Text2.Text, Reverse(Mid$(Reverse(frmRejoin.Text2.Text), 1, InStr(1, Reverse(frmRejoin.Text2.Text), "\") - 1)), " ")
 frmRejoin.File1.Refresh
 
 If Not FSO.FileExists(frmRejoin.File1.Path & "\" & "adminmain.split") Then
    MsgBox "Split information file does not exist"
    Exit Sub
 End If
 
 LFileSize = FileLen(frmRejoin.File1.Path & "\" & "admin1.split")
 
 IFIleName = frmRejoin.File1.Path & "\" & "AdminMain.Split"

'Reads the adminmain file
 Open IFIleName For Binary As 1
 BlDFileName = False
 BlFileNum = False
 
 Input #1, StrMyString   ' Read data into two variables.
 strTemp = Mid(StrMyString, 11)
 IntEndLocation = InStr(1, strTemp, "<", 1) - 1
 StrDFileName = Mid(strTemp, 1, IntEndLocation)
 
 strTemp = ""
 strTemp = InStr(IntEndLocation, StrMyString, ">", 1) + 1
 IntFileNum = Val(Mid(StrMyString, Val(strTemp), Trim(5)))
 Close #1
 
 frmRejoin.LAbel7.Caption = "FileName: " & StrDFileName
 frmRejoin.Label8.Caption = "No of files: " & IntFileNum
 
 For IntCnt = 1 To IntFileNum
     If Not FSO.FileExists(frmRejoin.File1.Path & "\Admin" & IntCnt & ".Split") Then
        MsgBox "Cannot Proceed. Admin" & IntCnt & ".Split not found"
        Exit Sub
     End If
 Next
 
 F_Size = String(LFileSize, " ") 'How much data has to be read
 IFIleName = ""
 OFileName = frmRejoin.File1.Path & "\" & StrDFileName
 
 Open OFileName For Binary Access Write Lock Read As 1 'opens a file for putting the data read from Source
 
 For IntCnt = 1 To IntFileNum
     IFIleName = frmRejoin.File1.Path & "\Admin" & IntCnt & ".Split"
     frmRejoin.LAbel9.Caption = "File in Process: " & IFIleName
     Open IFIleName For Binary As 2 'opens a file as a Source
     
     Get #2, , F_Size
     Put #1, , F_Size
     Close #2
     ProgressBAr.Value = (IntCnt / IntFileNum) * 100
     frmRejoin.Frame7.Refresh
 Next
 
 For IntCnt = 1 To IntFileNum
     FSO.DeleteFile (frmRejoin.File1.Path & "\Admin" & IntCnt & ".Split")
 Next
 FSO.DeleteFile (frmRejoin.File1.Path & "\AdminMain.Split")
 Close
 
 ProgressBAr.Value = 0
 strMessage = "The File has Joined successfully" & Chr(13) & "Please Check your Destination Folder."
 
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


End Sub

Private Sub Timer1_Timer()
'Timer for the Width
If BLFormStatus = True Then
    If intW >= 1000 Then
           intW = 0
           Me.Width = 5445
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
           frmMsgBox.Show vbModal
        Else
           Me.Width = Me.Width + intW
           Me.Left = Me.Left + 500
    End If
    intW = intW - 100

End If
End Sub

Private Sub Timer2_Timer()
'Timer for the height
On Error Resume Next
If BLFormStatus = True Then
    If intH >= 700 Then
           intH = 0
           Timer2.Enabled = False
           SetAlpha
           Me.Height = 1575
           Me.Refresh
           If strCommand = "SPLIT" Then
                Command1_Click
           ElseIf strCommand = "REJOIN" Then
                Command2_Click
           End If
        Else
           Me.Height = Me.Height + intH
           Me.Top = Me.Top - (intH + 10)
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
