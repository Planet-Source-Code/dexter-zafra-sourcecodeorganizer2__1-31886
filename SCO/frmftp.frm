VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmftp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Extreme FTP Client"
   ClientHeight    =   7635
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9015
   Icon            =   "frmftp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3600
      ScaleHeight     =   375
      ScaleWidth      =   5295
      TabIndex        =   26
      Top             =   3000
      Width           =   5295
      Begin MSComctlLib.Toolbar TBsave 
         Height          =   360
         Left            =   3840
         TabIndex        =   27
         Top             =   0
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   635
         ButtonWidth     =   2117
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "Imdex"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Upload File"
               Object.ToolTipText     =   "Upload Queued Files"
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Get IP Host "
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
         BorderStyle     =   1
         OLEDropMode     =   1
      End
      Begin MSComctlLib.Toolbar TBopen 
         Height          =   360
         Left            =   2400
         TabIndex        =   28
         Top             =   0
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   635
         ButtonWidth     =   1984
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "Imdex"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Open  File"
               Object.ToolTipText     =   "Open Local File"
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Get IP Host "
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
         BorderStyle     =   1
         OLEDropMode     =   1
      End
   End
   Begin VB.PictureBox picTitle 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   8715
      TabIndex        =   20
      Top             =   480
      Width           =   8715
      Begin MSComctlLib.ProgressBar PB 
         Height          =   135
         Left            =   4200
         TabIndex        =   23
         Top             =   120
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   8280
         TabIndex        =   24
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   2280
         TabIndex        =   22
         Top             =   0
         Width           =   2535
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Remote Files Server:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Local File Queued Progress:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   5160
      TabIndex        =   15
      Top             =   5880
      Width           =   3735
      Begin VB.Label Label2 
         Caption         =   "Total Files Queued - 0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "Total Size - 0.00 KB"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label9 
         Caption         =   "Estimated Time: 00:00:00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label10 
         Caption         =   "Elapsed Time: 00:00:00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   2655
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   7200
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6840
      TabIndex        =   9
      Text            =   "0"
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1200
      Top             =   5520
   End
   Begin VB.TextBox TxtTotalBytesQueued 
      Height          =   285
      Left            =   4800
      TabIndex        =   8
      Text            =   "0"
      Top             =   5280
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   5400
   End
   Begin VB.TextBox TxtConnectedTo 
      Height          =   285
      Left            =   7440
      TabIndex        =   7
      Top             =   5400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox TxtRemotePath 
      Height          =   285
      Left            =   3360
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   3975
   End
   Begin MSComctlLib.ListView ListView3 
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   2990
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Local Path"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Remote Path"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Host Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "File Size "
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Command"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Height          =   330
      Left            =   6960
      Picture         =   "frmftp.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   3836
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File Name"
         Object.Width           =   6703
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "File Size"
         Object.Width           =   2471
      EndProperty
   End
   Begin RichTextLib.RichTextBox RTBHeader 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   5160
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   1296
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmftp.frx":0544
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   5760
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   22
      ImageHeight     =   20
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":05C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":0B6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":1110
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":16B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":1C58
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":21FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":27A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":2D44
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":32E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":388C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":3E30
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":43D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":4978
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":4F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":5030
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":5150
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   688
      ButtonWidth     =   767
      ButtonHeight    =   688
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      HotImageList    =   "imlToolbarHot"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Connect"
            Object.ToolTipText     =   "Connect"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Connect"
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Disconnect"
            Object.ToolTipText     =   "Disconnect"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Stop"
            Object.ToolTipText     =   "Stop operation"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh contents of current directory"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Transfer"
            Object.ToolTipText     =   "Transfer Queue"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Download"
            Object.ToolTipText     =   "Add files to queue for download"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Upload"
            Object.ToolTipText     =   "Add files to queue for upload"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CreateDirectory"
            Object.ToolTipText     =   "Create Directory..."
            ImageIndex      =   7
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete File"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Rename"
            Object.ToolTipText     =   "Rename File..."
            ImageIndex      =   9
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "View Large Icons"
            Object.ToolTipText     =   "View Large Icons"
            ImageIndex      =   10
            Style           =   2
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "View Small Icons"
            Object.ToolTipText     =   "View Small Icons"
            ImageIndex      =   11
            Style           =   2
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "View List"
            Object.ToolTipText     =   "View List"
            ImageIndex      =   12
            Style           =   2
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "View Details"
            Object.ToolTipText     =   "View Details"
            ImageIndex      =   13
            Style           =   2
            Value           =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarHot 
      Left            =   6600
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   22
      ImageHeight     =   20
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":5264
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":5808
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":5DAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":6350
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":68F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":6E98
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":743C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":79E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":7F84
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":8528
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":8ACC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":9070
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":9614
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":9BB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":9CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":9DEC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4800
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":9F00
            Key             =   "dir"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":A05A
            Key             =   "file"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":A4AC
            Key             =   "dll"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":A8FE
            Key             =   "web"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmftp.frx":D0B0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current File Progress:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   10
      Top             =   5880
      Width           =   4935
      Begin VB.Label Label6 
         Caption         =   "Extreme Design  FTP Client : V.1"
         BeginProperty Font 
            Name            =   "Chloreal"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         MouseIcon       =   "frmftp.frx":11D02
         MousePointer    =   99  'Custom
         TabIndex        =   25
         ToolTipText     =   "Developer's  website"
         Top             =   1320
         Width           =   4455
      End
      Begin VB.Label Label3 
         Caption         =   "Time Left: 00:00:00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Current File:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label11 
         Caption         =   "0 KB of  0 KB Transfered"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   3615
      End
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "Speed: 0 Kbps"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "Local Files Queued"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   -120
      TabIndex        =   4
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Menu menufile 
      Caption         =   "File"
      Begin VB.Menu menuconnect 
         Caption         =   "Connect"
      End
      Begin VB.Menu menuDisconnect 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu menuline2 
         Caption         =   "-"
      End
      Begin VB.Menu menuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu menucommands 
      Caption         =   "Commands"
      Begin VB.Menu menuremove 
         Caption         =   "Remove Item From Local  File list"
      End
      Begin VB.Menu menuremoveall 
         Caption         =   "Remove All From Local File list"
      End
      Begin VB.Menu menuline1 
         Caption         =   "-"
      End
      Begin VB.Menu menutransfer 
         Caption         =   "Upload Queued Files"
      End
      Begin VB.Menu menudownload 
         Caption         =   "Download Files"
      End
      Begin VB.Menu menuupload 
         Caption         =   "Open Local Files"
      End
      Begin VB.Menu menustop 
         Caption         =   "Stop"
      End
      Begin VB.Menu menuRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu spacer1 
         Caption         =   "-"
      End
      Begin VB.Menu menuMD 
         Caption         =   "Make New Directory"
      End
      Begin VB.Menu menuFtpdelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu menuRename 
         Caption         =   "Rename Files"
      End
      Begin VB.Menu mnuview 
         Caption         =   "View Local Files In Browser"
      End
      Begin VB.Menu mnuprop 
         Caption         =   "Properties"
      End
   End
End
Attribute VB_Name = "frmftp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents mFTP As cFTP
Attribute mFTP.VB_VarHelpID = -1
Private Header As Variant
Private BeginTransfer                   As Single
Private TransferRate                    As Single
Private Declare Function ClipCursor Lib "user32" _
    (lpRect As Any) As Long

Private FilePathName As String
Private Filename As String
Private FormName As String

Private Declare Function OSGetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationname As String, ByVal lpKeyname As String, ByVal nDefault As Long, ByVal lpfilename As String) As Long
Private Declare Function OSGetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpfilename As String) As Long
Private Declare Function OSGetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpfilename As String) As Long

Private Declare Function OSWritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function OSWritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As Any, ByVal lpfilename As String) As Long

Private Declare Function OSGetProfileInt Lib "kernel32" Alias "GetProfileIntA" (ByVal lpAppName As String, ByVal lpKeyname As String, ByVal nDefault As Long) As Long
Private Declare Function OSGetProfileSection Lib "kernel32" Alias "GetProfileSectionA" (ByVal lpAppName As String, ByVal lpreturnedstring As String, ByVal nSize As Long) As Long
Private Declare Function OSGetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyname As String, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long) As Long

Private Declare Function OSWriteProfileSection Lib "kernel32" Alias "WriteProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String) As Long
Private Declare Function OSWriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long

Private Const nBUFSIZEINI = 1024
Private Const nBUFSIZEINIALL = 4096
Private NewVersion As String
Private OldVersion As String
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Const SW_SHOWNORMAL = 1
Private Declare Function ShellExecuteEx Lib "shell32.dll" (ByRef s As SHELLEXECUTEINFO) As Long




Public Sub RefreshDirectoryListing()
    Dim Item As cDirItem
    Dim lstX As ListItem
    Dim sAttr As String
    mFTP_StateChanged (FTP_RETRIEVING_DIRECTORY_INFO)
    DoEvents
    mFTP.GetDirectoryListing "*.*"
    DoEvents
    ListView2.ListItems.Clear
    DoEvents
    Set lstX = ListView2.ListItems.Add(, , "..")
    lstX.SmallIcon = 1
    TxtRemotePath.Text = mFTP.GetFTPDirectory
    DoEvents
    DoEvents
    For Each Item In mFTP.Directory
         sAttr = ""
         With Item
              If .Archive Then sAttr = sAttr & " A " Else sAttr = sAttr & " - "
              If .Compressed Then sAttr = sAttr & " C " Else sAttr = sAttr & " - "
              If .Directory Then sAttr = sAttr & " D " Else sAttr = sAttr & " - "
              If .Hidden Then sAttr = sAttr & " H " Else sAttr = sAttr & " - "
              If .Normal Then sAttr = sAttr & " N " Else sAttr = sAttr & " - "
              If .Offline Then sAttr = sAttr & " O " Else sAttr = sAttr & " - "
              If .ReadOnly Then sAttr = sAttr & " R " Else sAttr = sAttr & " - "
              If .System Then sAttr = sAttr & " S " Else sAttr = sAttr & " - "
              If .Temporary Then sAttr = sAttr & " T " Else sAttr = sAttr & " - "
         End With
         
         Set lstX = ListView2.ListItems.Add(, , Item.Filename)
         DoEvents
         With lstX
            If Item.Directory Then
               .SmallIcon = 1
               .SubItems(1) = "< Directory >"
               DoEvents
            Else
               .SmallIcon = 2
               DoEvents
               .SubItems(1) = Item.FileSize
               DoEvents
            End If
            DoEvents
         End With
         DoEvents
    Next
    DoEvents
    TxtRemotePath.Text = mFTP.GetFTPDirectory
    mFTP_StateChanged (FTP_DIRECTORY_INFO_COMPLETED)
End Sub

Private Sub Command2_Click()
            mFTP.SetFTPDirectory ".."
            TxtRemotePath.Text = mFTP.GetFTPDirectory
            DoEvents
            DoEvents
            RefreshDirectoryListing
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
mFTP.CloseConnection
    ListView2.ListItems.Clear
    Timer2.Enabled = False
    
    Unload Me
   
End Sub

Private Sub Form_Unload(cancel As Integer)
mFTP.CloseConnection
    ListView2.ListItems.Clear
    Timer2.Enabled = False
    
    Unload Me
   
End Sub

Private Sub Label6_Click()
Call Shell("Start.exe " & "http://clik.to/ret", 0)
End Sub

Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
Me.PopupMenu menucommands
End Sub

Private Sub menuconnect_Click()
            frmconnect.Show vbModal, Me

End Sub

Private Sub menuDisconnect_Click()
    mFTP.CloseConnection
    ListView2.ListItems.Clear
    Timer2.Enabled = False
    Label1.Caption = "Disconnect"
End Sub

Private Sub menudownload_Click()
Call ListView2_DblClick
End Sub

Private Sub menuexit_Click()
mFTP.CloseConnection
    ListView2.ListItems.Clear
    Timer2.Enabled = False
    
    Unload Me
   
End Sub

Private Sub menuFtpdelete_Click()
   If Not (ListView2.SelectedItem Is Nothing) Then
      If MsgBox("Are you sure you want to delete " & ListView2.SelectedItem.Text & "?", vbQuestion + vbYesNo, "Delete?") = vbYes Then
         If mFTP.Directory(ListView2.SelectedItem.Text).Directory Then
            If mFTP.RemoveFTPDirectory(ListView2.SelectedItem.Text) Then
               MsgBox "Directory " & SName & " was successfully removed.", vbInformation
               RefreshDirectoryListing
            Else
               MsgBox mFTP.GetLastErrorMessage
            End If
         Else
            If mFTP.DeleteFTPFile(ListView2.SelectedItem.Text) Then
               MsgBox "The file " & SName & " was successfully deleted.", vbInformation
               RefreshDirectoryListing
            Else
               MsgBox mFTP.GetLastErrorMessage
            End If
         End If
      End If
   End If
End Sub

Private Sub menuMD_Click()
   Dim SName As String
   SName = Trim(InputBox("Please enter a name for this directory:"))
   If SName <> "" Then
      If mFTP.CreateFTPDirectory(SName) Then
         MsgBox "Directory " & SName & " was successfully created.", vbInformation
         RefreshDirectoryListing
      Else
         MsgBox mFTP.GetLastErrorMessage
      End If
   End If
End Sub

Private Sub menuRefresh_Click()
RefreshDirectoryListing
End Sub

Private Sub menuremove_Click()
        TxtTotalBytesQueued.Text = TxtTotalBytesQueued.Text - ListView3.SelectedItem.SubItems(4)
        DoEvents
        DoEvents

ListView3.ListItems.Remove ListView3.SelectedItem.Index
End Sub

Private Sub menuremoveall_Click()
ListView3.ListItems.Clear
TxtTotalBytesQueued.Text = "0"
End Sub

Private Sub menuRename_Click()
    If Not (ListView2.SelectedItem Is Nothing) Then
       If ListView2.SelectedItem.Text <> ".." Then
            If Not mFTP.Directory(ListView2.SelectedItem.Text).Directory Then
               Dim sNewName As String
               sNewName = Trim(InputBox("Please enter a new name for this file:"))
               If sNewName <> "" Then
                  If mFTP.RenameFTPFile(ListView2.SelectedItem.Text, sNewName) Then
                     MsgBox "File was successfuly renamed"
                     RefreshDirectoryListing
                  Else
                     MsgBox mFTP.GetLastErrorMessage
                  End If
               End If
            End If
         End If
      End If
End Sub

Private Sub menustop_Click()
On Error Resume Next
        Dim pstrmessage As String
        pstrmessage = MsgBox("This will stop your current transfer and disconnect you from the site, are you sure you want to continue?", vbYesNo)
        If pstrmessage = vbYes Then
            mFTP.CloseConnection
            ListView2.ListItems.Clear
            Timer2.Enabled = False
            MsgBox "Proccess Stopped"
        End If


End Sub

Private Sub menutransfer_Click()
    Dim lTimer As Long
    Dim strRemote As String
    Dim strLocal As String
    Dim y As String
    If ListView3.ListItems.Count = 0 Then
    MsgBox "Ops!The are no files queued."
    Exit Sub
    End If
    
    Do Until ListView3.ListItems.Count = 0
        If frmconnect.optActive = True Then
        frmftp.mFTP.SetModeActive
        DoEvents
        Else
        frmftp.mFTP.SetModePassive
        DoEvents
        End If
        DoEvents
        If frmconnect.optBin = True Then
        frmftp.mFTP.SetTransferBinary
        DoEvents
        Else
        frmftp.mFTP.SetTransferASCII
        DoEvents
        End If
               DoEvents
               DoEvents
   BeginTransfer = Timer
   Timer2.Enabled = True
   DoEvents
        Label7.Caption = "Current File: " & ListView3.SelectedItem.Text
        DoEvents
        y = TxtRemotePath.Text
        DoEvents
        DoEvents
        If y = ListView3.SelectedItem.SubItems(2) Then
        
        Else
        mFTP.SetFTPDirectory ListView3.SelectedItem.SubItems(2)
        DoEvents
        DoEvents
        RefreshDirectoryListing
        DoEvents
        DoEvents
        End If
         DoEvents
         DoEvents
         DoEvents
         DoEvents
         
         If ListView3.SelectedItem.SubItems(5) = "Download" Then
          strRemote = ListView3.SelectedItem.Text
          strLocal = ListView3.SelectedItem.SubItems(1) & "\" & ListView3.SelectedItem.Text

               If mFTP.FTPDownloadFile(strLocal, strRemote) Then
                
                Else
                MsgBox mFTP.GetLastErrorMessage & "Unable To Complete Request."
                Exit Sub
                End If
                DoEvents
                DoEvents
        End If
        
        If ListView3.SelectedItem.SubItems(5) = "Upload" Then
          strRemote = ListView3.SelectedItem.Text
          strLocal = ListView3.SelectedItem.SubItems(1)

               If mFTP.FTPUploadFile(strLocal, strRemote) Then
                
                Else
                MsgBox mFTP.GetLastErrorMessage & "Unable To Complete Request."
                Exit Sub
                End If
                DoEvents
                DoEvents
        End If
        TxtTotalBytesQueued.Text = TxtTotalBytesQueued.Text - ListView3.SelectedItem.SubItems(4)
        DoEvents
        DoEvents
        ListView3.ListItems.Remove 1
        DoEvents
  Loop
RefreshDirectoryListing
Timer2.Enabled = False
Label10.Caption = "Elapsed Time: 00:00:00"
Text1.Text = "0"
DoEvents
End Sub

Private Sub menuupload_Click()
    Dim vFiles As Variant
    Dim lFile As Long
    Dim y As Long
    With cd1
        .Filename = "" 'Clear the filename
        .CancelError = False 'Gives an error if cancel is pressed
        .DialogTitle = "Select File(s)...  (Multi Select)"
        .Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNHideReadOnly 'Falgs, allows Multi select, Explorer style and hide the Read only tag
        .Filter = "All files (*.*)|*.*"
        .MaxFileSize = 9999
        .ShowOpen
        vFiles = Split(.Filename, Chr(0)) 'Splits the filename up in segments
    If UBound(vFiles) = 0 Then ' If there is only 1 file then do this
    Open .Filename For Binary Access Read As #1
    Size = LOF(1)
    Close #1
    DoEvents
    DoEvents
    Set Item2 = ListView3.ListItems.Add(, , .FileTitle)
    Item2.SubItems(1) = .Filename
    DoEvents
    Item2.SubItems(2) = mFTP.GetFTPDirectory
    DoEvents
    Item2.SubItems(3) = frmftp.TxtConnectedTo.Text
    DoEvents
    Item2.SubItems(4) = Size
    y = Item2.SubItems(4)
    DoEvents
    DoEvents
    TxtTotalBytesQueued.Text = TxtTotalBytesQueued.Text + y
    Text1.Text = Text1.Text + y
    DoEvents
    DoEvents
    DoEvents
    Item2.SubItems(5) = "Upload"
    DoEvents
    Else
    For lFile = 1 To UBound(vFiles) ' More than 1 file then do this until there are no more files
    Open vFiles(0) + "\" & vFiles(lFile) For Binary Access Read As #1
    Size = LOF(1)
    Close #1
    DoEvents
    DoEvents
    Set Item2 = ListView3.ListItems.Add(, , vFiles(lFile))
    DoEvents
    Item2.SubItems(1) = vFiles(0) + "\" & vFiles(lFile)
    DoEvents
    Item2.SubItems(2) = mFTP.GetFTPDirectory
    DoEvents
    Item2.SubItems(3) = frmftp.TxtConnectedTo.Text
    DoEvents
    Item2.SubItems(4) = Size
    y = Item2.SubItems(4)
    DoEvents
    DoEvents
    TxtTotalBytesQueued.Text = TxtTotalBytesQueued.Text + y
    Text1.Text = Text1.Text + y
    DoEvents
    DoEvents
    DoEvents
    Item2.SubItems(5) = "Upload"
    DoEvents
    Next
    End If
    End With
End Sub

Public Sub mFTP_FileTransferProgress(lCurrentBytes As Long, lTotalBytes As Long)
On Error Resume Next
Dim j As Long
Dim j3 As Long
TransferRate = Format(Int(lCurrentBytes / (Timer - BeginTransfer)) / 1024, "####.00")
    PB.Max = lTotalBytes
    PB.Min = 0
  j = PB.Value

  DoEvents
        PB.Value = lCurrentBytes
         DoEvents
        PB.ToolTipText = PB.Value & " Bytes of " & PB.Max & " Bytes Transfered"
        DoEvents
        Label11.Caption = PB.Value \ 1024 & " KB of  " & PB.Max \ 1024 & " KB Transfered"
        DoEvents
        Label4.Caption = Format$(CLng((j / PB.Max) * 100)) + "%"
        DoEvents
        Label8.Caption = "Speed: " & Format(TransferRate, "##.#0#") & " Kbps"
        DoEvents
        Label3.Caption = "Time Left: " & ConvertTime(Int(((PB.Max - PB.Value) / 1024) / TransferRate))
        DoEvents
        Label9.Caption = "Estimated Time: " & ConvertTime(Int(((Text1.Text) / 1024) / TransferRate))
        DoEvents
        If PB.Value = PB.Max Then
        Label4.Caption = "100%"
        End If

End Sub

Private Sub Form_Load()
   Set mFTP = New cFTP
End Sub
Public Function ConvertTime(ByVal TheTime As Single) As String
    Dim NewTime                         As String
    Dim Sec                             As Single
    Dim Min                             As Single
    Dim H                               As Single
    If TheTime > 60 Then
        Sec = TheTime
        Min = Sec / 60
        Min = Int(Min)
        Sec = Sec - Min * 60
        H = Int(Min / 60)
        Min = Min - H * 60
        NewTime = H & ":" & Min & ":" & Sec
        If H < 0 Then H = 0
        If Min < 0 Then Min = 0
        If Sec < 0 Then Sec = 0
        NewTime = Format(NewTime, "HH:MM:SS")
        ConvertTime = NewTime
    End If
    If TheTime < 60 Then
        NewTime = "00:00:" & TheTime
        NewTime = Format(NewTime, "HH:MM:SS")
        ConvertTime = NewTime
    End If
End Function
Private Sub ListView2_DblClick()
On Error Resume Next
Dim y As Long
y = ListView2.SelectedItem.SubItems(1)

DoEvents
    If Not (ListView2.SelectedItem Is Nothing) Then
         If ListView2.SelectedItem.Text = ".." Then
            mFTP.SetFTPDirectory ListView2.SelectedItem.Text
            TxtRemotePath.Text = mFTP.GetFTPDirectory
DoEvents
         Else
            If mFTP.Directory(ListView2.SelectedItem.Text).Directory Then
               mFTP.SetFTPDirectory ListView2.SelectedItem.Text
               TxtRemotePath.Text = mFTP.GetFTPDirectory
               DoEvents
            End If
DoEvents
            
            If Not mFTP.Directory(ListView2.SelectedItem.Text).Directory Then
            DoEvents
            Set Item2 = ListView3.ListItems.Add(, , ListView2.SelectedItem.Text)
            Item2.SubItems(1) = App.Path & "\Downloads"
            DoEvents
            Item2.SubItems(2) = mFTP.GetFTPDirectory
            DoEvents
            Item2.SubItems(3) = frmftp.TxtConnectedTo.Text
            DoEvents
            Item2.SubItems(4) = y
            TxtTotalBytesQueued.Text = TxtTotalBytesQueued.Text + y
            Text1.Text = Text1.Text + y
            DoEvents
            Item2.SubItems(5) = "Download"
            Exit Sub
            End If
            End If
            End If
            DoEvents
            RefreshDirectoryListing
DoEvents
End Sub
Public Function DepleteChr(Chars As String, Optional ReplaceChr As String) As String
    Dim ChrCnt As Long
    ReplaceChr = Left(ReplaceChr, 1)
    If ReplaceChr = "" Then ReplaceChr = " "


    Do
        ChrCnt = InStr(1, Chars, ReplaceChr)
        If ChrCnt = 0 Then Exit Do
        Chars = Left(Chars, ChrCnt - 1) & Right(Chars, Len(Chars) - ChrCnt)
    Loop
    DepleteChr = Chars
End Function



Private Sub mFTP_StateChanged(State As FTP_CONNECTION_STATES)
    Select Case State
        Case FTP_CONNECTION_RESOLVING_HOST
            RTBHeader.SelText = "FTP_CONNECTION_RESOLVING_HOST" & vbNewLine
        Case FTP_CONNECTION_HOST_RESOLVED
            RTBHeader.SelText = "FTP_CONNECTION_HOST_RESOLVED" & vbNewLine
        Case FTP_CONNECTION_CONNECTED
            RTBHeader.SelText = "FTP_CONNECTION_CONNECTED" & vbNewLine
        Case FTP_CONNECTION_AUTHENTICATION
            RTBHeader.SelText = "FTP_CONNECTION_AUTHENTICATION" & vbNewLine
        Case FTP_USER_LOGGED
            RTBHeader.SelText = "FTP_USER_LOGGED" & vbNewLine
        Case FTP_ESTABLISHING_DATA_CONNECTION
            RTBHeader.SelText = "FTP_ESTABLISHING_DATA_CONNECTION" & vbNewLine
        Case FTP_DATA_CONNECTION_ESTABLISHED
            RTBHeader.SelText = "FTP_DATA_CONNECTION_ESTABLISHED" & vbNewLine
        Case FTP_RETRIEVING_DIRECTORY_INFO
            RTBHeader.SelText = "FTP_RETRIEVING_DIRECTORY_INFO" & vbNewLine
        Case FTP_DIRECTORY_INFO_COMPLETED
            RTBHeader.SelText = "FTP_DIRECTORY_INFO_COMPLETED" & vbNewLine
        Case FTP_TRANSFER_STARTING
            RTBHeader.SelText = "FTP_TRANSFER_STARTING" & vbNewLine
        Case FTP_TRANSFER_COMLETED
            RTBHeader.SelText = "FTP_TRANSFER_COMLETED" & vbNewLine
    End Select
End Sub





Private Sub mnuprop_Click()
Dim shInfo As SHELLEXECUTEINFO
If ListView3.SelectedItem Is Nothing Then
    MsgBox "Properties of what file?"
    Exit Sub
End If
Set lsItem = ListView3.SelectedItem
    With shInfo
        .cbSize = LenB(shInfo)
        .lpFile = strPath & lsItem.Text
        .nShow = SW_SHOW
        .fMask = SEE_MASK_INVOKEIDLIST
        .lpVerb = "properties"
    End With
    ShellExecuteEx shInfo
End Sub



Private Sub mnuview_Click()
If Not ListView3.SelectedItem Is Nothing Then
ShellExecute 0, vbNullString, strPath & ListView3.SelectedItem.Text, vbNullString, strPath, SW_SHOWNORMAL
Else: MsgBox "Ops!You must select a file..Nothing to open!", vbExclamation, App.Title
End If
End Sub

Private Sub mnuviewme_Click()
If Not ListView2.SelectedItem Is Nothing Then
ShellExecute 0, vbNullString, strPath & ListView2.SelectedItem.Text, vbNullString, strPath, SW_SHOWNORMAL
Else: MsgBox "Nothing to open!", vbExclamation, App.Title
End If
End Sub



Private Sub TBopen_ButtonClick(ByVal Button As MSComctlLib.Button)
Call menuupload_Click
End Sub

Private Sub TBsave_ButtonClick(ByVal Button As MSComctlLib.Button)
frmftp.RTBHeader.SelText = Time & " > TRANSFERING DATA..." & vbCrLf
    RTBHeader.SelText = Time & " > OPENING FOLDER: " & Chr(34) & adr & Chr(34) & vbCrLf

    Call menutransfer_Click
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
    Select Case Button.Key
        Case "Connect"
        Call menuconnect_Click
        Case "Disconnect"
        Call menuDisconnect_Click
        Case "Stop"
        Call menustop_Click
        Case "Refresh"
            RefreshDirectoryListing
        Case "Transfer"
        Call menutransfer_Click
        Case "Download"
        Call menudownload_Click
        Case "Upload"
        Call menuupload_Click
        Case "CreateDirectory"
          Call menuMD_Click
        Case "Delete"
            menuFtpdelete_Click
        Case "Rename"
           Call menuRename_Click
        Case "View Large Icons"
            ListView2.View = lvwIcon
        Case "View Small Icons"
            ListView2.View = lvwSmallIcon
        Case "View List"
            ListView2.View = lvwList
        Case "View Details"
            ListView2.View = lvwReport
    End Select
End Sub

Private Sub Timer1_Timer()
Label2.Caption = "Total Files Queued - " & ListView3.ListItems.Count
Label5.Caption = "Total Size - " & Format(TxtTotalBytesQueued.Text / 1024, "0.00") & " KB"
End Sub

Private Sub Timer2_Timer()
        If PB.Value = PB.Max Then
            Label10.Caption = "Elapsed Time: " & Format(Hr & ":" & Min & ":" & Sec, "HH:MM:SS")
        Else
            Sec = Sec + 1
            If Sec >= 60 Then
                Sec = 0
                Min = Min + 1
            ElseIf Min >= 60 Then
                Min = 0
                Hr = Hr + 1
            End If
           Label10.Caption = "Elapsed Time: " & Format(Hr & ":" & Min & ":" & Sec, "HH:MM:SS")

        End If

End Sub
