VERSION 5.00
Begin VB.Form frmConnect 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Connect"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   2415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Text            =   "Guest"
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   2415
      Begin VB.CommandButton cmdServer 
         Caption         =   "&Server"
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdClient 
         Caption         =   "C&lient"
         Height          =   375
         Left            =   1320
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      Default         =   -1  'True
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.Timer tmrClient 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1680
      Top             =   1920
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "© 2000 One"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   0
      TabIndex        =   5
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label lblIP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   2175
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!:::File Transfer:::!!!!!!!!!!!!!!!!!!!!!!!!!!!!'
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!By: Adam Spicer!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!'
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!'
'!!!!!!!!!!!!!!!!!!!!!!!!!!!EASIEST WAY TO TRANSFER!!!!!!!!!!!!!!!!!!!!!!!!!!'
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!FILES OVER A NETWORK!!!!!!!!!!!!!!!!!!!!!!!!!!!'
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!'
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!'
'!!!!!!!!!!!!!!!!!!!!!!!!!!copyright 2000; OneLink Software!!!!!!!!!!!!!!!!!!'
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!www.onelinksystems.com!!!!!!!!!!!!!!!!!!!!!!!!'
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!'
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!'

Option Explicit
Dim strConnect As String 'holds the persons name

Private Sub cmdClient_Click()
'disables and enabled controls
    cmdServer.Enabled = False
    cmdClient.Enabled = False
    txtIP.Enabled = True
    cmdConnect.Enabled = True
'============================
    txtIP.Text = "Server's IP" 'sets the text in txtIP to tell them they need to enter an IP adderss
    strConnect = "Client" 'lets the program know u are the client
End Sub

Private Sub cmdConnect_Click()
    Select Case strConnect 'will either use Server or Client case depending on strConnect
        Case "Server"
            frmServer.Caption = "Welcome [" & txtName & "]" 'sets form up
            frmServer.Winsock.Close 'prevents error
            frmServer.Winsock.LocalPort = CLng(4050) 'sets the local port to 4050
            frmServer.Winsock.Listen 'listens for anyone wanting to connect
        Case "Client"
            frmClient.Caption = "Welcome [" & txtName & "]" 'sets up form
            tmrClient.Enabled = True 'enabled timer that will try to connect
    End Select

End Sub

Private Sub cmdServer_Click()
'disables and enabled controls
    cmdServer.Enabled = False
    cmdClient.Enabled = False
    txtIP.Enabled = False
    cmdConnect.Enabled = True
    cmdConnect.Caption = "Listen"
'============================
    strConnect = "Server" 'lets the program know u are the server
End Sub

Private Sub Form_Load()
'disables and enabled controls
    txtIP.Enabled = False
    cmdConnect.Enabled = False
'=============================
    lblIP.Caption = "Your IP: " & frmServer.Winsock.LocalIP 'shows u your IP
    
End Sub


Private Sub lblIP_Click()
    txtIP.Text = frmServer.Winsock.LocalIP 'puts your IP in box IF YOU are wanting to connect to yourself
                                            'just easier than typing it in  =D
End Sub

Private Sub tmrClient_Timer()
    frmClient.Winsock.Close 'closes any previous connections
    frmClient.Winsock.Connect txtIP.Text, "4050" 'tries to connect to the IP of the server:4050 is the port

End Sub
