VERSION 5.00
Begin VB.Form frmconnect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Connect.."
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   Icon            =   "frmconnect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Connection Mode"
      Height          =   855
      Left            =   0
      TabIndex        =   8
      Top             =   3840
      Width           =   3375
      Begin VB.OptionButton optPassive 
         Caption         =   "Passive Connection"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   480
         Width           =   2295
      End
      Begin VB.OptionButton optActive 
         Caption         =   "Active Connection"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "File Transfer Mode"
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   3000
      Width           =   3375
      Begin VB.OptionButton optBin 
         Caption         =   "Binary File Transfer"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.OptionButton optAscii 
         Caption         =   "ASCII File Transfer"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   480
         Width           =   2535
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Connect"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox TxtPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox TxtUser 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox TxtPort 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Text            =   "21"
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox TxtServer 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Text            =   "ftp.geocities.com"
      ToolTipText     =   "Example ftp host"
      Top             =   360
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   3375
      Begin VB.Label Label6 
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Username:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Port:"
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "ftp://:"
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   360
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmconnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
DoEvents
DoEvents
frmconnect.Hide
DoEvents
DoEvents
frmftp.TxtConnectedTo.Text = frmconnect.TxtServer.Text
frmftp.Label1.Caption = frmconnect.TxtServer.Text
    If frmftp.mFTP.OpenConnection(TxtServer.Text, TxtPort.Text, TxtUser.Text, TxtPass.Text) Then
        
        If frmconnect.optActive = True Then
        frmftp.mFTP.SetModeActive
        Else
        frmftp.mFTP.SetModePassive
        End If
        
        If frmconnect.optBin = True Then
        frmftp.mFTP.SetTransferBinary
        Else
        frmftp.mFTP.SetTransferASCII
        End If
        
        frmftp.mFTP.SetFTPDirectory "/"
        frmftp.RefreshDirectoryListing
    End If
    Dim V As Integer

  V = TxtServer.Text = ""
  V = TxtUser.Text = ""
  V = TxtPass.Text = ""
  If V = True Then
   MsgBox "Ops!You must supply the FTP Host,User Name and Password before connecting.Try again!.."
   Else
  frmftp.RTBHeader.SelText = time & " > TRANSFERING DATA..." & vbCrLf
    frmftp.RTBHeader.SelText = time & " > OPENING FOLDER: " & Chr(34) & adr & Chr(34) & vbCrLf
frmftp.Show
End If

End Sub






