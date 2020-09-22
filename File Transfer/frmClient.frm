VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtChat 
      Height          =   1575
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   6120
      Width           =   6855
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Top             =   7800
      Width           =   5535
   End
   Begin VB.CommandButton cmdSendChat 
      Caption         =   "Send Chat"
      Default         =   -1  'True
      Height          =   255
      Left            =   5760
      TabIndex        =   10
      Top             =   7800
      Width           =   1215
   End
   Begin VB.Timer tmrKBps 
      Interval        =   1000
      Left            =   0
      Top             =   8400
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send File"
      Height          =   495
      Left            =   5520
      TabIndex        =   6
      Top             =   3960
      Width           =   1215
   End
   Begin VB.DriveListBox lstDrive 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   3135
   End
   Begin VB.DirListBox lstDir 
      Height          =   1665
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   3135
   End
   Begin VB.FileListBox lstFiles 
      Height          =   1455
      Left            =   120
      System          =   -1  'True
      TabIndex        =   3
      Top             =   2400
      Width           =   3135
   End
   Begin VB.CommandButton cmdAddFile 
      Caption         =   "&Add File"
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   3960
      Width           =   1215
   End
   Begin VB.ListBox lstSend 
      Height          =   3375
      Left            =   3720
      TabIndex        =   1
      Top             =   360
      Width           =   3255
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove File"
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   3960
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   120
      Top             =   8160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   0
      Top             =   8760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox lstPath 
      Height          =   3375
      Left            =   3720
      TabIndex        =   7
      Top             =   360
      Width           =   3255
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   735
      Left            =   240
      TabIndex        =   13
      Top             =   4800
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   1296
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblKBps 
      Alignment       =   2  'Center
      Caption         =   "KBps:"
      Height          =   255
      Left            =   1200
      TabIndex        =   14
      Top             =   5640
      Width           =   4455
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6960
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line2 
      X1              =   3480
      X2              =   3480
      Y1              =   120
      Y2              =   4560
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Select Files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Send Files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFriend As String 'holds servers name
Dim strMyName As String 'holds your name
Dim strFileName As String 'holds the name of the file u are receiving
Dim strSize As String 'holds the size of the file
Dim strSoFar As String 'a var for calculating the KBps
Dim strBlock As String 'holds the data you are going to send
Dim strLOF As String 'holds the lenght of the file

Private Sub cmdAddFile_Click()
    If lstFiles.ListIndex = -1 Then 'if nothing is selected
        MsgBox "Please select a file, then click Add File", vbInformation, "Add File"
    Else
        lstSend.AddItem lstFiles.List(lstFiles.ListIndex)
        lstPath.AddItem lstDir.Path
    End If
    
End Sub

Private Sub cmdRemove_Click()
    If lstSend.ListIndex = -1 Then 'if nothing is selected
        MsgBox "Please select a file to remove, and then hit remove.", vbInformation, "Remove File"
    Else
        lstPath.RemoveItem lstSend.ListIndex
        lstSend.RemoveItem lstSend.ListIndex
    End If
    
End Sub

Private Sub cmdSendChat_Click()
    If Trim(txtSend.Text) = "" Then Exit Sub 'prevents someone trying to send nothing
    Winsock.SendData "Chat" & txtSend.Text 'sends the text to the chat
    txtChat.SelStart = Len(txtChat) 'put focus on the chat at the end so it is entered in the right place
    txtChat.SelText = strMyName & ":" & vbTab & txtSend.Text & vbCrLf 'puts the text in the chat
    txtSend.Text = "" 'clears the textbox u type in
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Winsock.Close 'closes winsock so program can end
    End 'closes program
End Sub

Private Sub lstDir_Change()
    lstFiles.Path = lstDir.Path 'links them together
    
End Sub

Private Sub lstDrive_Change()
On Error GoTo driveError 'if A: drive isnt ready (forexample)
    lstDir.Path = lstDrive.Drive
    Exit Sub
driveError:
    MsgBox "The current device is unavailable", vbCritical, "Error"
    lstDrive.ListIndex = 1 'goes to C:
    
End Sub

Private Sub tmrKBps_Timer()
On Error Resume Next 'prevents error
    lblKBps.Caption = "Transfering at: " & Format(strSoFar / 1000, "###0.0") & " / KBps" 'calculates the KBps
    strSoFar = 0 'resets it so it can be calculated again
End Sub

Private Sub Winsock_Connect()
    frmConnect.tmrClient.Enabled = False 'tell client to stop trying to connect cause it is connected =D
    Winsock.SendData "Nick" & frmConnect.txtName.Text 'sends ur name to server
    DoEvents 'i dunno why but i need it cause of winsock
    strMyName = frmConnect.txtName 'saves ur name into memory
    Me.Show 'shows frmclient
    Unload frmConnect 'obvious

End Sub

Private Sub cmdSend_Click()
On Error Resume Next 'prevents error
    strFileName = "" 'resets the filename
    strSize = "" 'resets the size

Dim intX As Integer
Dim strFile, strPath As String
        strFile = lstSend.List(0)
        strPath = lstPath.List(0)
        lstSend.RemoveItem 0
        lstPath.RemoveItem 0
        Open strPath & "\" & strFile For Binary As #1 'opens the file to be sent and reads it
        strLOF = LOF(1) 'gets the length of the file
        Winsock.SendData "Name" & strFile & ":" & strLOF 'sends the name of the first file and its length

End Sub
            
Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next 'prevents error
Dim strData As String 'holds data for select case
Dim strData2 As String 'holds data
    Call Winsock.GetData(strData, vbString) 'gets the data sent by the server
    strData2 = Mid(strData, 5) 'gets data
    strData = Left(strData, 4) 'gets data for select case
    Select Case strData 'goes to the right case depending on strData
        Case "File" 'a file transfer is in progress
            Put 1, , strData2 'puts data into file
            PBar.Value = PBar.Value + bytesTotal 'shows how much is done so far
            strSoFar = strSoFar + bytesTotal 'calculates KBps
            If Not LOF(1) >= PBar.Max Then
                Winsock.SendData "OKOK keep sending!" 'tells them ur done with the data and u want some more!
                DoEvents ' =D
            End If
        Case "Name" 'client has sent u the filename and is ready to begin transfer
            Dim intX As Integer 'holds position if :
            intX = InStr(1, strData2, ":", vbTextCompare) 'gets position of :
            strSize = Mid(strData2, intX + 1) 'holds the filesize
            PBar.Max = strSize 'sets up the progressbar
            strData = Mid(strData2, 1, intX - 1) 'holds filename
            strFileName = strData 'puts filename into memory
            Dim strResponse As String 'holds either a vbYEs or vbNo
            strResponse = MsgBox(strFriend & " wants to send you [" & strFileName & "].  Do you wish to receive this file?", vbYesNo, "File Exchange Requested") '<=- easy to understand
            If strResponse = vbYes Then 'if they said yes
                Dim strType As String 'holds the type of file
                strType = Right(strFileName, 3) 'gets the type of file
                CD.FileName = strFileName 'sets the filename into the commondialog box
                CD.Filter = "File Type (*." & strType & ")|*." & strType 'sets the filter to the filetype
                CD.Flags = cdlOFNOverwritePrompt 'asks u if u want to overwrite file
                CD.ShowSave 'shows the save commondialog box
                Open CD.FileName For Binary As #1 'opens a file with the name and path u want
                Winsock.SendData "OKOK i want the file" 'tell client u want the damn file
                Me.Enabled = False 'disables to form to PREVENT ERROR!!!!!!!!!!
            ElseIf strResponse = vbNo Then 'if they say no
                Winsock.SendData "Nope dont want it!" 'tell em u dont want their crap!
                DoEvents 'hmmm
            End If 'ok enough of that madness
        Case "Stop" 'the file exchange has ended
            Close #1 'closes the file
            'resets the progressbar
            PBar.Value = 0
            PBar.Max = 1
            '=====================
            Me.Enabled = True 'reenables the form!
            DoEvents
            Winsock.SendData "OKOKmore"
        Case "Nick" 'client has sent u their name
            strFriend = strData2 'saves their name into memory
        Case "Nope" 'tells u that they declined ur request to give em a file
            MsgBox strFriend & " declined your file transfer request.", vbInformation, "File Transfer Canceled!" '<=- easy to get again
            Close #1 'closes the file
            'stops the loops that was waiting for the boolean value to be true
            Do
            DoEvents
            Loop
            '==========================
        Case "OKOK" 'tells u they want more of the file
            If strData2 = "more" Then
                If lstSend.ListCount <> 0 Then
                    cmdSend_Click
                    Exit Sub
                Else
                    Exit Sub
                End If
            End If
            Me.Enabled = False 'keeps form disabled
            PBar.Max = strLOF 'sets progressbar max to filesize
            If Not EOF(1) Then 'does this if not the end of the file
                If strLOF - Loc(1) < 2040 Then 'if you are at the last chunk of data
                    strBlock = Space$(strLOF - Loc(1)) 'sets the block size to the size of the data (cause its less!)
                    Get 1, , strBlock 'gets data
                    Winsock.SendData "File" & strBlock 'sends data
                    DoEvents ' =/
                    PBar.Value = PBar.Value + Len(strBlock) 'sets progressbar
                    strSoFar = strSoFar + (strLOF - Loc(1)) 'sets KBps
                    Winsock.SendData "Stop the maddness!" 'tells client THE TRANSFER IS ENDED!
                    Close #1 'closes file
                    'resets the progressbar
                    PBar.Max = 1
                    PBar.Value = 0
                    '====================
                    Me.Enabled = True 'reenables the form
                Else 'if not the last chunk
                    strBlock = Space$(2040) 'sets block up to receive only 2040 bytes of data
                End If
                strSoFar = strSoFar + 2040 'calculates KBps
                Get 1, , strBlock 'gets data
                Winsock.SendData "File" & strBlock 'sends data
                DoEvents
                PBar.Value = PBar.Value + Len(strBlock) 'sets progressbar
            End If
        Case "Chat" 'if they are talking to ya
            txtChat.SelStart = Len(txtChat) 'sets cursor position in chatroom
            txtChat.SelText = strFriend & ":" & vbTab & strData2 & vbCrLf 'puts the chat into the room
    End Select
End Sub
