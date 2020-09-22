VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{AEA70AA8-B35D-11D4-AC39-90B64FC10000}#1.0#0"; "McIconMenu.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H80000016&
   Caption         =   "Source Code Organizer V 2.2"
   ClientHeight    =   7635
   ClientLeft      =   1635
   ClientTop       =   1155
   ClientWidth     =   11325
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   11325
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   9763
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Contents"
      TabPicture(0)   =   "frmMain.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lstTitles"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Search"
      TabPicture(1)   =   "frmMain.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture1"
      Tab(1).Control(1)=   "txtTitlez"
      Tab(1).Control(2)=   "txtSubject"
      Tab(1).Control(3)=   "btnConvert"
      Tab(1).Control(4)=   "cmbRotation"
      Tab(1).Control(5)=   "cmbFontSize"
      Tab(1).Control(6)=   "cmbPageSize"
      Tab(1).Control(7)=   "cmbFont"
      Tab(1).Control(8)=   "txtOutputFile"
      Tab(1).Control(9)=   "txtFilename"
      Tab(1).Control(10)=   "btnOpen"
      Tab(1).Control(11)=   "btnSave"
      Tab(1).Control(12)=   "Frame5"
      Tab(1).Control(13)=   "cd1"
      Tab(1).Control(14)=   "Frame1(1)"
      Tab(1).Control(15)=   "Label11"
      Tab(1).Control(16)=   "Label9"
      Tab(1).Control(17)=   "Label10"
      Tab(1).Control(18)=   "lblOutputFile"
      Tab(1).Control(19)=   "lblFilename"
      Tab(1).Control(20)=   "Label7"
      Tab(1).ControlCount=   21
      TabCaption(2)   =   "Add/Del"
      TabPicture(2)   =   "frmMain.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frameDelete"
      Tab(2).Control(1)=   "frameModify"
      Tab(2).Control(2)=   "Frame4"
      Tab(2).ControlCount=   3
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   -74880
         ScaleHeight     =   2415
         ScaleWidth      =   3015
         TabIndex        =   93
         Top             =   480
         Visible         =   0   'False
         Width           =   3015
         Begin VB.TextBox Textweb 
            Height          =   285
            Left            =   1560
            TabIndex        =   97
            Top             =   4440
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CheckBox ChkHighLight 
            Caption         =   "HighLight Menus in Button Style"
            Height          =   195
            Left            =   240
            TabIndex        =   96
            Top             =   4080
            Value           =   1  'Checked
            Width           =   2655
         End
         Begin VB.PictureBox picBackGround 
            Height          =   495
            Left            =   1680
            Picture         =   "frmMain.frx":091E
            ScaleHeight     =   435
            ScaleWidth      =   435
            TabIndex        =   95
            Top             =   3960
            Visible         =   0   'False
            Width           =   495
         End
         Begin RichTextLib.RichTextBox rtbNotes 
            Height          =   2055
            Left            =   0
            TabIndex        =   94
            TabStop         =   0   'False
            ToolTipText     =   "Write Code info/Author's name or it could be anything."
            Top             =   240
            Width           =   2985
            _ExtentX        =   5265
            _ExtentY        =   3625
            _Version        =   393217
            Enabled         =   -1  'True
            ScrollBars      =   3
            RightMargin     =   1.00000e5
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"frmMain.frx":10CE
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin prjmciconmenu.mciconmenu mciconmenu1 
            Left            =   120
            Top             =   2760
            _ExtentX        =   1058
            _ExtentY        =   1058
         End
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   360
            Top             =   4440
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   31
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":1150
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":12AC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":1408
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":15E8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":1744
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":18A0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":1E3C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":1F98
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":20F4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":2250
                  Key             =   ""
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":23AC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":2508
                  Key             =   ""
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":2664
                  Key             =   ""
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":27C0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":291C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":2AF8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":2C54
                  Key             =   ""
               EndProperty
               BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":2DB0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":334C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":38E8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":3A44
                  Key             =   ""
               EndProperty
               BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":3FE0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":4432
                  Key             =   ""
               EndProperty
               BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":4D0C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":55E6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":5900
                  Key             =   ""
               EndProperty
               BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":5C1A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":64F4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":6DCE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":76A8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":7AFA
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Caption         =   "Source Code Notes/Author Information:"
            Height          =   255
            Left            =   0
            TabIndex        =   98
            Top             =   0
            Width           =   2895
         End
      End
      Begin VB.TextBox txtTitlez 
         Height          =   285
         Left            =   -73320
         TabIndex        =   91
         Top             =   3360
         Width           =   1380
      End
      Begin VB.TextBox txtSubject 
         Height          =   285
         Left            =   -74880
         TabIndex        =   90
         Top             =   3360
         Width           =   1500
      End
      Begin VB.CommandButton btnConvert 
         Caption         =   "Convert to PDF"
         Default         =   -1  'True
         Height          =   660
         Left            =   -72720
         TabIndex        =   89
         ToolTipText     =   "Convert now"
         Top             =   4680
         Width           =   735
      End
      Begin VB.ComboBox cmbRotation 
         Height          =   315
         ItemData        =   "frmMain.frx":97D4
         Left            =   -74880
         List            =   "frmMain.frx":97E4
         Style           =   2  'Dropdown List
         TabIndex        =   88
         Top             =   5040
         Width           =   1005
      End
      Begin VB.ComboBox cmbFontSize 
         Height          =   315
         ItemData        =   "frmMain.frx":97FD
         Left            =   -73800
         List            =   "frmMain.frx":9816
         Style           =   2  'Dropdown List
         TabIndex        =   87
         Top             =   5040
         Width           =   930
      End
      Begin VB.ComboBox cmbPageSize 
         Height          =   315
         ItemData        =   "frmMain.frx":9843
         Left            =   -73800
         List            =   "frmMain.frx":9850
         Style           =   2  'Dropdown List
         TabIndex        =   86
         Top             =   4680
         Width           =   945
      End
      Begin VB.ComboBox cmbFont 
         Height          =   315
         ItemData        =   "frmMain.frx":9877
         Left            =   -74880
         List            =   "frmMain.frx":9884
         Style           =   2  'Dropdown List
         TabIndex        =   85
         Top             =   4680
         Width           =   975
      End
      Begin VB.TextBox txtOutputFile 
         Height          =   285
         Left            =   -74880
         TabIndex        =   84
         Top             =   4320
         Width           =   2340
      End
      Begin VB.TextBox txtFilename 
         Height          =   285
         Left            =   -74880
         TabIndex        =   81
         Top             =   3840
         Width           =   2340
      End
      Begin VB.CommandButton btnOpen 
         Height          =   255
         Left            =   -72480
         Picture         =   "frmMain.frx":98A5
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Browse  text file"
         Top             =   3840
         Width           =   495
      End
      Begin VB.CommandButton btnSave 
         Height          =   255
         Left            =   -72480
         Picture         =   "frmMain.frx":9C2F
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Browse for destination folder"
         Top             =   4320
         Width           =   495
      End
      Begin VB.Frame frameDelete 
         Caption         =   "Delete"
         Height          =   1995
         Left            =   -74880
         TabIndex        =   30
         Top             =   3360
         Width           =   2955
         Begin VB.ListBox lstDelete 
            Height          =   840
            Left            =   240
            TabIndex        =   31
            Top             =   720
            Width           =   2355
         End
         Begin VB.Label lblDelete 
            Caption         =   "Select the Code Language you wish to delete:"
            Height          =   465
            Left            =   240
            TabIndex        =   32
            Top             =   240
            Width           =   2085
         End
      End
      Begin VB.Frame frameModify 
         Caption         =   "Modify"
         Height          =   1785
         Left            =   -74880
         TabIndex        =   27
         Top             =   1560
         Width           =   3015
         Begin VB.ListBox lstModify 
            Height          =   840
            Left            =   240
            TabIndex        =   28
            Top             =   720
            Width           =   2475
         End
         Begin VB.Label lblModify 
            Caption         =   "Select the Code Language you wish to modify:"
            Height          =   465
            Left            =   240
            TabIndex        =   29
            Top             =   240
            Width           =   2595
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Add New Language"
         Height          =   975
         Left            =   -74880
         TabIndex        =   24
         Top             =   480
         Width           =   3015
         Begin VB.CommandButton Command3 
            Caption         =   "Add New"
            Height          =   375
            Left            =   1920
            TabIndex        =   26
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Left            =   240
            TabIndex        =   25
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Internet Search"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   14
         Top             =   1560
         Width           =   3015
         Begin VB.ComboBox cmEngines 
            Height          =   315
            ItemData        =   "frmMain.frx":9FB9
            Left            =   240
            List            =   "frmMain.frx":9FE1
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "Search"
            Height          =   350
            Left            =   2040
            TabIndex        =   16
            Top             =   840
            Width           =   855
         End
         Begin VB.ComboBox txWhatSearch 
            Height          =   315
            ItemData        =   "frmMain.frx":A052
            Left            =   240
            List            =   "frmMain.frx":A074
            TabIndex        =   15
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label Label4 
            Caption         =   "Search Engine:"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   600
            Width           =   1455
         End
      End
      Begin MSComctlLib.ListView lstTitles 
         Height          =   4935
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   8705
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Title"
            Object.Width           =   3352
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Laguage"
            Object.Width           =   3176
         EndProperty
      End
      Begin MSComDlg.CommonDialog cd1 
         Left            =   -75000
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame1 
         Caption         =   "DataBase Search"
         Height          =   975
         Index           =   1
         Left            =   -74880
         TabIndex        =   19
         Top             =   480
         Width           =   3015
         Begin VB.CommandButton CmdsearchDB 
            Caption         =   "Search"
            Height          =   350
            Left            =   2160
            TabIndex        =   21
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox Text2 
            Height          =   315
            Left            =   360
            TabIndex        =   20
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label5 
            Caption         =   "Enter Code Title:"
            Height          =   255
            Left            =   360
            TabIndex        =   22
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Label Label11 
         Caption         =   "TEXT-PDF Converter"
         Height          =   255
         Left            =   -74040
         TabIndex        =   100
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Title:"
         Height          =   255
         Left            =   -73320
         TabIndex        =   99
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Subject:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   92
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label lblOutputFile 
         Caption         =   "Destination:"
         Height          =   225
         Left            =   -74880
         TabIndex        =   83
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label lblFilename 
         Caption         =   "File to convert:"
         Height          =   240
         Left            =   -74880
         TabIndex        =   82
         Top             =   3600
         Width           =   2895
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   255
         Left            =   -74160
         TabIndex        =   64
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   3015
   End
   Begin MSComDlg.CommonDialog cdoOpenDatabase 
      Left            =   9360
      Top             =   -240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "mdb"
      DialogTitle     =   "Open Database"
      Filter          =   "Access Database Files|*.mdb|All Files|*.*"
      FilterIndex     =   1
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   9960
      Top             =   -240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A100
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A212
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A324
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A436
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A548
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A65A
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A76C
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A87E
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A990
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":ACAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AFC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B2DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B5F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B752
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BA6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C346
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CC20
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D4FA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab tabSnippit 
      Height          =   7215
      Left            =   3480
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   12726
      _Version        =   393216
      TabHeight       =   529
      TabMaxWidth     =   3528
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Code Wndow"
      TabPicture(0)   =   "frmMain.frx":DDD4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "toolBar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "rtbCodeWindow"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "HTML Editor/Grabber"
      TabPicture(1)   =   "frmMain.frx":DDF0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture2"
      Tab(1).Control(1)=   "rtbnet"
      Tab(1).Control(2)=   "Frame9"
      Tab(1).Control(3)=   "PBar"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "SQL Builder"
      TabPicture(2)   =   "frmMain.frx":DE0C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame7"
      Tab(2).Control(1)=   "Frame6"
      Tab(2).ControlCount=   2
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   5295
         Left            =   -74880
         ScaleHeight     =   5295
         ScaleWidth      =   7575
         TabIndex        =   71
         Top             =   1800
         Visible         =   0   'False
         Width           =   7575
         Begin SHDocVwCtl.WebBrowser Web 
            Height          =   5280
            Left            =   0
            TabIndex        =   72
            Top             =   0
            Width           =   7575
            ExtentX         =   13361
            ExtentY         =   9313
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   "http:///"
         End
      End
      Begin SCO.CodeHighlight rtbnet 
         Height          =   5295
         Left            =   -74880
         TabIndex        =   70
         Top             =   1800
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   9340
         Language        =   3
         KeywordColor    =   16576
         OperatorColor   =   12582912
         DelimiterColor  =   32896
         ForeColor       =   0
         FunctionColor   =   12583104
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SCO.CodeHighlight rtbCodeWindow 
         Height          =   6015
         Left            =   120
         TabIndex        =   69
         Top             =   1080
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   10610
         Language        =   1
         KeywordColor    =   12582912
         OperatorColor   =   12582912
         DelimiterColor  =   32768
         ForeColor       =   0
         FunctionColor   =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame9 
         Caption         =   "HTML Editor / Grabber"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   54
         Top             =   360
         Width           =   7575
         Begin VB.ComboBox CoFonts 
            Height          =   315
            Left            =   4080
            Style           =   2  'Dropdown List
            TabIndex        =   73
            Top             =   120
            Visible         =   0   'False
            Width           =   1335
         End
         Begin SCO.chameleonButton chameleonButton2 
            Height          =   255
            Left            =   3960
            TabIndex        =   68
            ToolTipText     =   "View HTML"
            Top             =   960
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            BTYPE           =   8
            TX              =   "View HTML"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            FCOL            =   0
            FCOLO           =   12582912
            MPTR            =   0
            MICON           =   "frmMain.frx":DE28
         End
         Begin SCO.chameleonButton Command1 
            Height          =   255
            Left            =   3240
            TabIndex        =   67
            ToolTipText     =   "Preview in browser"
            Top             =   960
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            BTYPE           =   8
            TX              =   "Preview"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            FCOL            =   0
            FCOLO           =   12582912
            MPTR            =   0
            MICON           =   "frmMain.frx":DE44
         End
         Begin SCO.chameleonButton chameleonButton1 
            Height          =   255
            Left            =   6480
            TabIndex        =   65
            ToolTipText     =   "Full View Maximized Editor's Window"
            Top             =   960
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            BTYPE           =   8
            TX              =   "Maximized "
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            FCOL            =   0
            FCOLO           =   12582912
            MPTR            =   0
            MICON           =   "frmMain.frx":DE60
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Copy"
            Height          =   375
            Left            =   6000
            TabIndex        =   58
            Top             =   480
            Width           =   735
         End
         Begin VB.CommandButton CmdGet 
            Caption         =   "Get"
            Enabled         =   0   'False
            Height          =   375
            Left            =   5040
            TabIndex        =   57
            ToolTipText     =   "Go get grab it"
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox TxtUrl 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   56
            Text            =   "http://"
            ToolTipText     =   "Enter the URL of the website you want to grab"
            Top             =   480
            Width           =   4815
         End
         Begin SCO.chameleonButton cmdfile 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   55
            ToolTipText     =   "Open/Save/New"
            Top             =   960
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BTYPE           =   8
            TX              =   "File"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            FCOL            =   0
            FCOLO           =   12582912
            MPTR            =   0
            MICON           =   "frmMain.frx":DE7C
         End
         Begin SCO.chameleonButton cmdtable 
            Height          =   255
            Index           =   1
            Left            =   600
            TabIndex        =   59
            ToolTipText     =   "Table options"
            Top             =   960
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BTYPE           =   8
            TX              =   "Table"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            FCOL            =   0
            FCOLO           =   12582912
            MPTR            =   0
            MICON           =   "frmMain.frx":DE98
         End
         Begin SCO.chameleonButton cmdfont 
            Height          =   255
            Index           =   0
            Left            =   1800
            TabIndex        =   60
            ToolTipText     =   "Insert Scrolling Marquee"
            Top             =   960
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            BTYPE           =   8
            TX              =   "Marquee"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            FCOL            =   0
            FCOLO           =   12582912
            MPTR            =   0
            MICON           =   "frmMain.frx":DEB4
         End
         Begin SCO.chameleonButton cmdinsert 
            Height          =   255
            Index           =   0
            Left            =   1200
            TabIndex        =   61
            ToolTipText     =   "Image/Links/Tag/Background color/Character"
            Top             =   960
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BTYPE           =   8
            TX              =   "Insert"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            FCOL            =   0
            FCOLO           =   12582912
            MPTR            =   0
            MICON           =   "frmMain.frx":DED0
         End
         Begin SCO.chameleonButton cmdother 
            Height          =   255
            Index           =   1
            Left            =   2640
            TabIndex        =   62
            ToolTipText     =   "Line Break/White Paper etc.."
            Top             =   960
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BTYPE           =   8
            TX              =   "Other"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            FCOL            =   0
            FCOLO           =   12582912
            MPTR            =   0
            MICON           =   "frmMain.frx":DEEC
         End
         Begin SCO.chameleonButton cmdprev 
            Height          =   255
            Index           =   0
            Left            =   5040
            TabIndex        =   63
            ToolTipText     =   "Preview In Full Size"
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            BTYPE           =   8
            TX              =   "Full Size Preview"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            FCOL            =   0
            FCOLO           =   12582912
            MPTR            =   0
            MICON           =   "frmMain.frx":DF08
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "DataBase Scanner"
         Height          =   2415
         Left            =   -74880
         TabIndex        =   47
         Top             =   360
         Width           =   7575
         Begin VB.CommandButton cmdstart 
            Caption         =   "&Scan"
            Height          =   255
            Left            =   1320
            TabIndex        =   52
            ToolTipText     =   "Click to Scan DataBase file"
            Top             =   2040
            Width           =   855
         End
         Begin VB.CheckBox checkscan 
            Caption         =   "Include Subdirectories"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            TabStop         =   0   'False
            ToolTipText     =   "Include Subdirectory"
            Top             =   2040
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.DirListBox Dir1 
            Height          =   1440
            Left            =   120
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   600
            Width           =   2055
         End
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   120
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   240
            Width           =   2055
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1815
            Left            =   2280
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   240
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   3201
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            SmallIcons      =   "ImageList1"
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Name"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "In Folder"
               Object.Width           =   7056
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Size"
               Object.Width           =   1764
            EndProperty
         End
         Begin VB.OLE OLE1 
            Class           =   "Package"
            Height          =   75
            Left            =   7080
            OleObjectBlob   =   "frmMain.frx":DF24
            SourceDoc       =   "C:\3A\fish.scr"
            TabIndex        =   104
            Top             =   2280
            Width           =   375
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Height          =   255
            Left            =   3840
            TabIndex        =   75
            Top             =   2040
            Width           =   2295
         End
         Begin VB.Label lblscan 
            ForeColor       =   &H00800000&
            Height          =   135
            Left            =   6960
            TabIndex        =   53
            Top             =   240
            Width           =   135
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "SQL Builder"
         Height          =   4215
         Left            =   -74880
         TabIndex        =   33
         Top             =   2760
         Width           =   7575
         Begin VB.TextBox txtCreator 
            Height          =   285
            Left            =   7440
            TabIndex        =   103
            Top             =   3600
            Visible         =   0   'False
            Width           =   75
         End
         Begin VB.TextBox txtAuthor 
            Height          =   285
            Left            =   7440
            TabIndex        =   102
            Top             =   3600
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.TextBox txtKeywords 
            Height          =   285
            Left            =   7440
            TabIndex        =   101
            Top             =   3600
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.Timer Timer1 
            Interval        =   25
            Left            =   0
            Top             =   3960
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Database Viewer"
            Height          =   255
            Left            =   5640
            TabIndex        =   76
            ToolTipText     =   "Find and View Database"
            Top             =   3840
            Width           =   1695
         End
         Begin VB.CommandButton Clearz 
            Caption         =   "Clear"
            Height          =   255
            Left            =   3120
            TabIndex        =   74
            ToolTipText     =   "Insert tag"
            Top             =   3480
            Width           =   1095
         End
         Begin VB.ListBox List1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2700
            ItemData        =   "frmMain.frx":1CD3C
            Left            =   5520
            List            =   "frmMain.frx":1CD9D
            TabIndex        =   44
            Top             =   960
            Width           =   1815
         End
         Begin VB.ListBox List2 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   645
            Left            =   120
            TabIndex        =   43
            ToolTipText     =   "Table row"
            Top             =   3120
            Width           =   2895
         End
         Begin VB.CommandButton cmdsq1 
            Caption         =   "Insert SQL"
            Height          =   255
            Left            =   4320
            TabIndex        =   41
            Top             =   3120
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Insert as Auctioneer Tag"
            Height          =   195
            Left            =   2040
            TabIndex        =   40
            Top             =   3840
            Width           =   2055
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Insert as VTrader Tag"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   3840
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.ListBox List3 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   7440
            TabIndex        =   38
            ToolTipText     =   "Coloumn row"
            Top             =   3120
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.CommandButton cmdsq2 
            Caption         =   "Insert BDTB"
            Height          =   255
            Left            =   4320
            TabIndex        =   37
            ToolTipText     =   "Find Insert Database table"
            Top             =   3480
            Width           =   1095
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Insert Tag"
            Height          =   255
            Left            =   3120
            TabIndex        =   36
            ToolTipText     =   "Insert tag"
            Top             =   3120
            Width           =   1095
         End
         Begin VB.CheckBox chksq1 
            Caption         =   "Insert Comma"
            Height          =   255
            Left            =   4200
            TabIndex        =   35
            Top             =   3840
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frmMain.frx":1CE64
            Left            =   5520
            List            =   "frmMain.frx":1CE71
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   360
            Width           =   1815
         End
         Begin RichTextLib.RichTextBox RichTextBox1 
            Height          =   2895
            Left            =   120
            TabIndex        =   42
            ToolTipText     =   "SQL Builder window"
            Top             =   240
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   5106
            _Version        =   393217
            BackColor       =   16777215
            Enabled         =   -1  'True
            ScrollBars      =   3
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"frmMain.frx":1CE8B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lbllsq3 
            AutoSize        =   -1  'True
            Caption         =   "Scope:"
            Height          =   195
            Left            =   5640
            TabIndex        =   46
            Top             =   120
            Width           =   510
         End
         Begin VB.Label lblsq6 
            AutoSize        =   -1  'True
            Caption         =   "Sql:"
            Height          =   195
            Left            =   5640
            TabIndex        =   45
            Top             =   720
            Width           =   270
         End
      End
      Begin MSComctlLib.Toolbar toolBar 
         Height          =   330
         Left            =   1200
         TabIndex        =   1
         Top             =   600
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imgList"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "New  code snippit"
               Object.Tag             =   "new"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Edit Code "
               Object.Tag             =   "open"
               ImageIndex      =   17
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Save code snippit to the DataBase"
               Object.Tag             =   "save"
               ImageKey        =   "Save"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Copy code snippit"
               Object.Tag             =   "copy"
               ImageKey        =   "Copy"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "View/write notes source code info"
               Object.Tag             =   "paste"
               ImageIndex      =   18
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "view"
                     Text            =   "View Notes/Author Info.."
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "contents"
                     Text            =   "View Contents"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Delete code snippit"
               Object.Tag             =   "delete"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Search Code Data Base / Internet"
               Object.Tag             =   "find"
               ImageKey        =   "Find"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "DB"
                     Text            =   "Search DataBase"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "SDB"
                     Text            =   "Search Code in the internet"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Print code snippit"
               Object.Tag             =   "print"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modify / Add New / Delete Code"
               Object.Tag             =   "mod"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Insert Javascript / DHTML Code in FrontPage 2000"
               Object.Tag             =   "front"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Insert Javascript /sideserver Code in Dreamweaver4"
               Object.Tag             =   "drw"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Insert Code in Visual Basic6"
               Object.Tag             =   "VB"
               ImageIndex      =   11
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame1 
         Caption         =   "Add New Code Language"
         Height          =   1335
         Index           =   0
         Left            =   -73560
         TabIndex        =   2
         Top             =   480
         Width           =   4335
         Begin VB.Label Label3 
            Caption         =   "Enter A new Code Langugae:"
            Height          =   255
            Left            =   720
            TabIndex        =   3
            Top             =   360
            Width           =   2415
         End
      End
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   7575
         Begin VB.CommandButton sClose 
            Caption         =   "r"
            BeginProperty Font 
               Name            =   "Marlett"
               Size            =   6.75
               Charset         =   2
               Weight          =   500
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   7200
            Style           =   1  'Graphical
            TabIndex        =   78
            ToolTipText     =   "Close"
            Top             =   240
            Width           =   240
         End
      End
      Begin MSComctlLib.ProgressBar PBar 
         Height          =   135
         Left            =   -74880
         TabIndex        =   66
         Top             =   1680
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   3255
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   120
         ScaleHeight     =   1335
         ScaleWidth      =   3015
         TabIndex        =   105
         Top             =   240
         Visible         =   0   'False
         Width           =   3015
         Begin VB.CommandButton Cmdchangemenucol 
            Caption         =   "Change Menu Bacground Color"
            Height          =   255
            Left            =   240
            TabIndex        =   109
            Top             =   1080
            Width           =   2655
         End
         Begin VB.CheckBox ChkUseBackground 
            Caption         =   "Use background in menus"
            Height          =   195
            Left            =   240
            TabIndex        =   107
            Top             =   360
            Value           =   1  'Checked
            Width           =   2535
         End
         Begin VB.CheckBox ChkHigh 
            Caption         =   "HighLight Menus in Button Style"
            Height          =   195
            Left            =   240
            TabIndex        =   106
            Top             =   720
            Value           =   1  'Checked
            Width           =   2655
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Caption         =   "Menu Highlight Color Settings"
            Height          =   255
            Left            =   360
            TabIndex        =   108
            Top             =   120
            Width           =   2295
         End
      End
      Begin MSComDlg.CommonDialog CmDlg 
         Left            =   2160
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   360
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.ComboBox cmbFilter 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1080
         Width           =   1395
      End
      Begin VB.ComboBox cmbType 
         Height          =   315
         ItemData        =   "frmMain.frx":1CF0D
         Left            =   120
         List            =   "frmMain.frx":1CF0F
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1080
         Width           =   1455
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   1440
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblFilter 
         Caption         =   "Filter By Language:"
         Height          =   225
         Left            =   1680
         TabIndex        =   10
         Top             =   840
         Width           =   1425
      End
      Begin VB.Label Label2 
         Caption         =   "Code Language:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Source Code Title:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
   End
   Begin MSComctlLib.StatusBar sBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   23
      Top             =   7335
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13361
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            TextSave        =   "2/18/2002"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            MinWidth        =   2548
            TextSave        =   "3:28 AM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   300
      Left            =   0
      TabIndex        =   77
      Top             =   0
      Visible         =   0   'False
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   529
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Index           =   1
      Begin VB.Menu mnuNew 
         Caption         =   "New code "
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open Database"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnubackup 
         Caption         =   "Back Up DataBase"
      End
      Begin VB.Menu mnucont 
         Caption         =   "Contents"
      End
      Begin VB.Menu mnunet 
         Caption         =   "Search Code"
      End
      Begin VB.Menu me1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "&Edit "
      Index           =   2
      Begin VB.Menu mnucopy 
         Caption         =   "Copy Code"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnupaste 
         Caption         =   "Paste Code"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnusettingz 
         Caption         =   "Undo"
      End
      Begin VB.Menu htmlsetprop 
         Caption         =   "Select All"
      End
      Begin VB.Menu me2 
         Caption         =   "-"
      End
      Begin VB.Menu mnudelete 
         Caption         =   "Delete current code"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save current code"
         Shortcut        =   ^S
      End
      Begin VB.Menu me3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "Find"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnusetting 
      Caption         =   "S&ettings"
      Begin VB.Menu mnuusesounds 
         Caption         =   "# Play Typing Sound"
      End
      Begin VB.Menu mnuhighlight 
         Caption         =   "Menu Highlight Setting"
      End
      Begin VB.Menu mnustand 
         Caption         =   "Stand By ( Screen Saver )"
      End
   End
   Begin VB.Menu mnuType 
      Caption         =   "&Add - Language"
      Begin VB.Menu mnuModify 
         Caption         =   "Modify / Add / Delete"
      End
   End
   Begin VB.Menu tool 
      Caption         =   "&Tools"
      Begin VB.Menu mnuicon 
         Caption         =   "Icon Scanner"
      End
      Begin VB.Menu mnuftp 
         Caption         =   "FTP Client"
      End
      Begin VB.Menu mnupdf 
         Caption         =   "Text-PDF Converter"
      End
      Begin VB.Menu fileview 
         Caption         =   "File Viewer (htm,html,js,css)..."
      End
      Begin VB.Menu iconz 
         Caption         =   "-"
      End
      Begin VB.Menu sql 
         Caption         =   "SQL Builder..."
      End
      Begin VB.Menu html 
         Caption         =   "HTML Editor/Grabber"
      End
   End
   Begin VB.Menu notez 
      Caption         =   "&Notes"
      Begin VB.Menu show1 
         Caption         =   "View Write Code Info notes"
      End
      Begin VB.Menu hideme 
         Caption         =   "Hide notes"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuhelpcon 
         Caption         =   "&Contents?"
      End
      Begin VB.Menu mnuhtmlterms 
         Caption         =   "Online HTML Terms"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnucontact 
         Caption         =   "Contact Developer"
      End
      Begin VB.Menu me5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuweb 
         Caption         =   "Developer's Website"
      End
      Begin VB.Menu mnuPlanetSourceCode 
         Caption         =   "Bug report"
      End
      Begin VB.Menu mnutest 
         Caption         =   "test"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu MnuTables 
      Caption         =   "Tables"
      Visible         =   0   'False
      Begin VB.Menu MnuEnterAsStatement 
         Caption         =   "Enter as Statement "
      End
      Begin VB.Menu mnucol 
         Caption         =   "Columns"
      End
   End
   Begin VB.Menu mnusetz 
      Caption         =   "setting"
      Visible         =   0   'False
      Begin VB.Menu mnucopy2 
         Caption         =   "Copy..."
      End
      Begin VB.Menu mnupaste2 
         Caption         =   "Paste..."
      End
      Begin VB.Menu mnuprint2 
         Caption         =   "Undo"
      End
      Begin VB.Menu space9 
         Caption         =   "Select All"
      End
      Begin VB.Menu spacer007 
         Caption         =   "-"
      End
      Begin VB.Menu mnusetupme 
         Caption         =   "View Code Contents..."
      End
      Begin VB.Menu delete 
         Caption         =   "View Code Notes/Author Info.."
      End
      Begin VB.Menu mnuspace 
         Caption         =   "-"
      End
      Begin VB.Menu editme 
         Caption         =   "Edit with Organizer Pad.."
      End
   End
   Begin VB.Menu menuz 
      Caption         =   "setupme"
      Visible         =   0   'False
      Begin VB.Menu cut 
         Caption         =   "Cut..."
      End
      Begin VB.Menu copy1 
         Caption         =   "Copy.."
      End
      Begin VB.Menu dexter 
         Caption         =   "-"
      End
      Begin VB.Menu paste2 
         Caption         =   "Paste...."
      End
      Begin VB.Menu setup1 
         Caption         =   "Undo"
      End
      Begin VB.Menu print2 
         Caption         =   "Select All"
      End
      Begin VB.Menu char1 
         Caption         =   "-"
      End
      Begin VB.Menu char 
         Caption         =   "Insert Character #@~"
      End
      Begin VB.Menu mnuinsertagz 
         Caption         =   "Insert Tag"
      End
      Begin VB.Menu property 
         Caption         =   "Tag  Properties"
      End
      Begin VB.Menu dex4 
         Caption         =   "-"
      End
      Begin VB.Menu browse 
         Caption         =   "Preview in Browser"
      End
   End
   Begin VB.Menu file 
      Caption         =   "Files"
      Visible         =   0   'False
      Begin VB.Menu mnu1 
         Caption         =   "New"
      End
      Begin VB.Menu mnu2 
         Caption         =   "Open"
      End
      Begin VB.Menu mnu3 
         Caption         =   "Save"
      End
      Begin VB.Menu sound 
         Caption         =   "# Play Typing Sound"
      End
   End
   Begin VB.Menu table 
      Caption         =   "Tables"
      Visible         =   0   'False
      Begin VB.Menu table1 
         Caption         =   "Add the first column"
         Begin VB.Menu back1 
            Caption         =   "Background"
            Begin VB.Menu back2 
               Caption         =   "Black"
            End
            Begin VB.Menu back3 
               Caption         =   "Blue"
            End
            Begin VB.Menu backspace 
               Caption         =   "-"
            End
            Begin VB.Menu back4 
               Caption         =   "Blue Violet"
            End
            Begin VB.Menu back5 
               Caption         =   "Brown"
            End
            Begin VB.Menu back6 
               Caption         =   "Cyan"
            End
            Begin VB.Menu backspace1 
               Caption         =   "-"
            End
            Begin VB.Menu back7 
               Caption         =   "Dark Brown"
            End
            Begin VB.Menu back8 
               Caption         =   "Dark Green"
            End
            Begin VB.Menu back9 
               Caption         =   "Dark Blue"
            End
            Begin VB.Menu backspace4 
               Caption         =   "-"
            End
            Begin VB.Menu back10 
               Caption         =   "Gold"
            End
         End
      End
      Begin VB.Menu addcol1 
         Caption         =   "Add new column"
         Begin VB.Menu ground 
            Caption         =   "Background"
            Begin VB.Menu add1 
               Caption         =   "Black"
            End
            Begin VB.Menu add32 
               Caption         =   "Blue"
            End
            Begin VB.Menu add2 
               Caption         =   "Black Violet"
            End
            Begin VB.Menu add3 
               Caption         =   "Blue Violet"
            End
            Begin VB.Menu add4 
               Caption         =   "Brown"
            End
            Begin VB.Menu add5 
               Caption         =   "Cyan"
            End
            Begin VB.Menu add6 
               Caption         =   "Dark Brown"
            End
            Begin VB.Menu add7 
               Caption         =   "Dark green"
            End
            Begin VB.Menu add8 
               Caption         =   "Dark Blue"
            End
            Begin VB.Menu add9 
               Caption         =   "Gold"
            End
         End
      End
      Begin VB.Menu cell 
         Caption         =   "Add Cells"
      End
      Begin VB.Menu cellspace 
         Caption         =   "-"
      End
      Begin VB.Menu addcolumn1 
         Caption         =   "Add more columns"
         Begin VB.Menu col1 
            Caption         =   "Add One Column"
         End
         Begin VB.Menu col2 
            Caption         =   "Add two Column"
         End
         Begin VB.Menu col3 
            Caption         =   "Add three column"
         End
         Begin VB.Menu col4 
            Caption         =   "Add four column"
         End
         Begin VB.Menu mr1 
            Caption         =   "Add more columns"
         End
         Begin VB.Menu addrow 
            Caption         =   "-"
         End
         Begin VB.Menu addrow1 
            Caption         =   "Add Rows"
         End
      End
   End
   Begin VB.Menu fontme 
      Caption         =   "Marquee"
      Visible         =   0   'False
      Begin VB.Menu fontz 
         Caption         =   "Scrolling Marquee"
         Begin VB.Menu mnumarqsrcolleft 
            Caption         =   "Scroll Left"
         End
         Begin VB.Menu mnumarqscrollright 
            Caption         =   "Scroll Right"
         End
         Begin VB.Menu mnumarqscrollup 
            Caption         =   "Scroll Up"
         End
         Begin VB.Menu mnumarqscrolldown 
            Caption         =   "Scroll Down"
         End
      End
      Begin VB.Menu fn8 
         Caption         =   "Alternate Marquee"
         Begin VB.Menu mnumarqalternateright 
            Caption         =   "Alternate Right"
         End
         Begin VB.Menu mnumarqalternateleft 
            Caption         =   "Alternate Left"
         End
      End
      Begin VB.Menu mnumarquee2 
         Caption         =   "Slide Marquee"
         Begin VB.Menu mnumarqslidelft 
            Caption         =   "Slide Left"
         End
         Begin VB.Menu mnumarqslideright 
            Caption         =   "Slide Right"
         End
      End
   End
   Begin VB.Menu other 
      Caption         =   "Other"
      Visible         =   0   'False
      Begin VB.Menu ot1 
         Caption         =   "Unmumbered List"
      End
      Begin VB.Menu ot2 
         Caption         =   "Numbered List"
      End
   End
   Begin VB.Menu insert 
      Caption         =   "Insert"
      Visible         =   0   'False
      Begin VB.Menu mnucolorpicker 
         Caption         =   "Color Picker"
      End
      Begin VB.Menu time 
         Caption         =   "Time - Date"
      End
      Begin VB.Menu mnutagz 
         Caption         =   "Insert tag"
      End
      Begin VB.Menu mnuchardex 
         Caption         =   "Insert Character"
      End
      Begin VB.Menu mnubackgroundcol 
         Caption         =   "Tag Properties"
      End
   End
   Begin VB.Menu mnufilelist 
      Caption         =   "filez"
      Visible         =   0   'False
      Begin VB.Menu mnudelete01 
         Caption         =   "Delete Code"
      End
      Begin VB.Menu mnuaddmod01 
         Caption         =   "Add/Modify"
      End
      Begin VB.Menu mnusearchcodez 
         Caption         =   "Search Code"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const adSchemaColumns = 4
Const AppName = "Organizer Text-PDF Converter"
Dim Position As Long
Dim pageNo As Long
Dim lineNo As Long
Dim pageHeight As Long
Dim pageWidth As Long
Dim location(1 To 5000) As Long
Dim pageObj(1 To 5000) As Long
Dim lines As Long
Dim obj As Long
Dim Tpages As Long
Dim encoding As Long
Dim resources As Long
Dim pages As Variant
Dim author As String
Dim creator As String
Dim keywords As String
Dim subject As String
Dim Title As String
Dim BaseFont As String
Dim pointSize As Currency
Dim vertSpace As Currency
Dim rotate As Integer
Dim info As Long
Dim root As Long
Dim npagex As Double
Dim npagey As Long
Dim filetxt As String
Dim filepdf As String
Dim linelen As Long
Dim cache As String
Dim cmdline As String
Dim SName As String
Dim gblnIgnoreChange As Boolean
Dim gintIndex As Integer
Dim gstrStack(1000) As String
Public Path As String
Public saved As Boolean
Public myFile As String
Public myLONG As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Type OPENFILENAME
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  Flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type
    
Private Const OFN_READONLY = &H1
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_SHOWHELP = &H10
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOLONGNAMES = &H40000 ' force no long names for 4.x modules
Private Const OFN_EXPLORER = &H80000 ' new look commdlg
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_LONGNAMES = &H200000 ' force long names for 3.x modules
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0


Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Sub add1_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#000000>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub add2_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#9F5F9F>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub add3_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#A62A2A>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub add32_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#0000FF>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub add4_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#A62A2A>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub add5_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#00FFFF>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub add6_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#5C4033>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub add7_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#2F4F2F>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub add9_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#CD7F32>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub addrow1_Click()
rtbnet.SelText = rtbnet.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor= >" + Chr(13) + Chr(10) + "<P>Write your Text here (1st cell of the added ROW)" + Chr(13) + Chr(10) + "</TD><TD bgcolor= >" + Chr(13) + Chr(10) + "ADD HERE CELLS. Select and Paste the two lines above: <P>...</TD><TD>" + Chr(13) + Chr(10) + "<P>Write your Text here (Last cell of the added ROW)" + Chr(13) + Chr(10) + "</TD></TR>ADD HERE ROWS"
End Sub

Private Sub back10_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#CD7F32>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub


Private Sub back2_Click()
rtbnet.SelText = rtbnet.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#CD7F32>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + "Here add new cells for the first column" + Chr(13) + Chr(10) + "Here add the second column" + Chr(13) + Chr(10) + "Here add new cells for the second column, and so on " + Chr(13) + Chr(10) + "</TD></TR>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub back3_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#0000FF>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub back4_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#9F5F9F>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub back5_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#A62A2A>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub back6_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#00FFFF>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub back7_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#5C4033>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub back8_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#2F4F2F>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub back9_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#871F78>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub
Private Sub ShowProgressInStatusBar(ByVal bShowProgressBar As Boolean)

    Dim tRC As RECT
    
    If bShowProgressBar Then
'
' Get the size of the Panel (2) Rectangle from the status bar
' remember that Indexes in the API are always 0 based (well,
' nearly always) - therefore Panel(2) = Panel(1) to the api
'
'
        SendMessageAny sBar.hWnd, SB_GETRECT, 0, tRC
'
' and convert it to twips....
'
        With tRC
            .Top = (.Top * Screen.TwipsPerPixelY)
            .left = (.left * Screen.TwipsPerPixelX)
            .Bottom = (.Bottom * Screen.TwipsPerPixelY) - .Top
            .Right = (.Right * Screen.TwipsPerPixelX) - .left
        End With
'
' Now Reparent the ProgressBar to the statusbar
'
        With ProgressBar1
            SetParent .hWnd, sBar.hWnd
            .Move tRC.left, tRC.Top, tRC.Right, tRC.Bottom
            .Visible = True
            .Value = 0
        End With
        
    Else
'
' Reparent the progress bar back to the form and hide it
'
        SetParent ProgressBar1.hWnd, Me.hWnd
        ProgressBar1.Visible = False
    End If
    
End Sub
Private Sub browse_Click()
Picture2.Visible = True
CmdGet.Enabled = False
TxtUrl.Locked = True
Open App.Path & "\Organizerpreview\extreme.html" For Output As #1
Print #1, rtbnet.Text
Close #1
Load Browser
Textweb.Text = App.Path & "\Organizerpreview\extreme.html"
Web.Navigate App.Path & "\Organizerpreview\extreme.html"
End Sub

Private Sub cell_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10)
End Sub

Private Sub ChkHigh_Click()
mciconmenu1.HighlightStyle = ChkHigh.Value
End Sub

Private Sub ChkUseBackground_Click()
If CBool(ChkUseBackground.Value) = True Then
  Set mciconmenu1.BackgroundPicture = picBackGround
 Else
  mciconmenu1.ClearBackgroundPicture
 End If
End Sub

Private Sub Cmdchangemenucol_Click()
On Error Resume Next

myFile = ""
CommonDialog1.Filename = ""
CommonDialog1.DialogTitle = "Select one picture..."
CommonDialog1.Filter = "Image files (*.jpg,*.bmp,*.gif)|*.jpg;*.bmp"
CommonDialog1.ShowOpen
If Trim(CommonDialog1.Filename = "") Then
 picBackGround.Cls
 ChkUseBackground.Value = 0
Else
 picBackGround.Picture = LoadPicture(CommonDialog1.Filename)
 Set mciconmenu1.BackgroundPicture = picBackGround
End If
End Sub

Private Sub mnuhighlight_Click()
If Picture3.Visible = True Then
   Picture3.Visible = False
    txtTitle.Visible = True
   Else
   Picture3.Visible = True
   txtTitle.Visible = False
   End If
End Sub

Private Sub mnuhtmlterms_Click()
Call webBrowse("http://people.we.mediaone.net/retxed/htmlterms.htm")
End Sub

Private Sub mnustand_Click()
On Error Resume Next
    OLE1.DoVerb 1
End Sub

Private Sub txtAuthor_GotFocus()
  txtAuthor.SelStart = 0
  txtAuthor.SelLength = Len(txtAuthor.Text)
End Sub

Private Sub txtCreator_GotFocus()
  txtCreator.SelStart = 0
  txtCreator.SelLength = Len(txtCreator.Text)
End Sub

Private Sub txtSubject_GotFocus()
  txtSubject.SelStart = 0
  txtSubject.SelLength = Len(txtSubject.Text)
End Sub

Private Sub txtTitlez_GotFocus()
  txtTitlez.SelStart = 0
  txtTitlez.SelLength = Len(txtTitlez.Text)
End Sub

Private Sub txtKeywords_GotFocus()
  txtKeywords.SelStart = 0
  txtKeywords.SelLength = Len(txtKeywords.Text)
End Sub

Private Sub txtFilename_GotFocus()
  txtFilename.SelStart = 0
  txtFilename.SelLength = Len(txtFilename.Text)
End Sub

Private Sub txtOutputFile_GotFocus()
  txtOutputFile.SelStart = 0
  txtOutputFile.SelLength = Len(txtOutputFile.Text)
End Sub



Private Sub btnOpen_Click()
  Dim Filename As String
  On Local Error Resume Next
  Filename = OpenDialog(Me, "Text files (*.txt)|*.txt|All files (*.*)|*.*", _
                   "Select a text file", "")
  If Len(Filename) Then
    txtFilename.Text = Filename
    Filename = txtFilename.Text
    txtOutputFile.Text = left(Filename, Len(Filename) - 3) & "pdf"
  End If
End Sub

Private Sub btnSave_Click()
  Dim Filename As String
  On Local Error Resume Next
  Filename = SaveDialog(Me, "Portable Document Format files (*.pdf)|*.pdf", _
                        "Save PDF As", "", "")
  If Len(Filename) Then
    txtOutputFile.Text = Filename
  End If
End Sub

Private Sub btnSource_Click()
  On Local Error Resume Next
End Sub

Private Sub btnConvert_Click()
  If txtFilename.Text <> "" And txtOutputFile.Text <> "" Then
    ConvertToPDF txtFilename.Text, txtOutputFile.Text, _
                 txtAuthor.Text, txtCreator.Text, txtKeywords.Text, _
                 txtSubject.Text, txtTitlez.Text, _
                 cmbFont.Text, Val(cmbFontSize.Text), Val(cmbRotation.Text), _
                 Val(cmbPageSize.Text), Val(Right(cmbPageSize.Text, 3))
    If FileExists(cmdline) Then
      Unload Me
    ElseIf MsgBox("PDF file is done." & vbCr & vbCr & "Do you want to open the generated PDF file?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
      ShellExecute 0, vbNullString, txtOutputFile.Text, vbNullString, vbNullString, 1
    End If
  Else
    MsgBox "Please specify file names."
  End If
End Sub

Public Sub ConvertToPDF(Filename As String, outputfile As String, _
                        Optional TextAuthor As String, Optional TextCreator As String, Optional TextKeywords As String, _
                        Optional TextSubject As String, Optional TextTitle As String, _
                        Optional FontName As String = "Courier", Optional FontSize As Integer = 10, Optional Rotation As Integer, _
                        Optional pwidth As Single = 8.5, Optional pheight As Single = 11)
  On Error GoTo er
  If Not FileExists(Filename) Then
    MsgBox "File '" & Filename & "' does not exist."
    Exit Sub
  ElseIf FileExists(outputfile) Then
    Kill outputfile
  End If

  initialize FontName, FontSize, Rotation, pwidth, pheight
  
  author = TextAuthor
  creator = TextCreator
  keywords = TextKeywords
  subject = TextSubject
  Title = TextTitle
  filetxt = Filename
  filepdf = outputfile
  
  Call WriteStart
  Call WriteHead
  Call WritePages
  Call endpdf
  Exit Sub
er:
  MsgBox Err.Description
End Sub

Private Sub initialize(FontName As String, FontSize As Integer, Rotation As Integer, pwidth As Single, pheight As Single)
  pageHeight = 72 * pheight
  pageWidth = 72 * pwidth

  BaseFont = FontName ' Courier, Times-Roman, Arial
  pointSize = FontSize ' Font Size; Don't change it
  vertSpace = FontSize * 1.2 ' Vertical spacing
  rotate = Rotation ' degrees to rotate; try setting 90,180,etc
  lines = (pageHeight - 72) / vertSpace ' no of lines on one page
  
  Select Case LCase(FontName)
   Case "courier": linelen = 1.5 * pageWidth / pointSize
   Case "arial": linelen = 2 * pageWidth / pointSize
  'Case "Times-Roman": linelen = 2.2 * pageWidth / pointSize
   Case Else: linelen = 2.2 * pageWidth / pointSize
  End Select

  obj = 0
  npagex = pageWidth / 2
  npagey = 25
  pageNo = 0
  Position = 0
  cache = ""
End Sub

Private Sub writepdf(stre As String, Optional flush As Boolean)
  On Local Error Resume Next
  Position = Position + Len(stre)
  cache = cache & stre & vbCr
  If Len(cache) > 32000 Or flush Then
    Open filepdf For Append As #1
    Print #1, cache;
    Close #1
    cache = ""
  End If
End Sub
  
Private Sub WriteStart()
  writepdf ("%PDF-1.2")
  writepdf ("%âãÏÓ")
End Sub

Private Sub WriteHead()
  Dim CreationDate As String
  On Error GoTo er
    CreationDate = "D:" & Format(Now, "YYYYMMDDHHNNSS")
    obj = obj + 1
    location(obj) = Position
    info = obj
    
    writepdf (obj & " 0 obj")
    writepdf ("<<")
    writepdf ("/Author (" & author & ")")
    writepdf ("/CreationDate (" & CreationDate & ")")
    writepdf ("/Creator (" & creator & ")")
    writepdf ("/Producer (" & AppName & ")")
    writepdf ("/Title (" & Title & ")")
    writepdf ("/Subject (" & subject & ")")
    writepdf ("/Keywords (" & keywords & ")")
    writepdf (">>")
    writepdf ("endobj")
    
    obj = obj + 1
    root = obj
    obj = obj + 1
    Tpages = obj
    encoding = obj + 2
    resources = obj + 3
    
    obj = obj + 1
    location(obj) = Position
    writepdf (obj & " 0 obj")
    writepdf ("<<")
    writepdf ("/Type /Font")
    writepdf ("/Subtype /Type1")
    writepdf ("/Name /F1")
    writepdf ("/Encoding " & encoding & " 0 R")
    writepdf ("/BaseFont /" & BaseFont)
    writepdf (">>")
    writepdf ("endobj")
    
    obj = obj + 1
    location(obj) = Position
    writepdf (obj & " 0 obj")
    writepdf ("<<")
    writepdf ("/Type /Encoding")
    writepdf ("/BaseEncoding /WinAnsiEncoding")
    writepdf (">>")
    writepdf ("endobj")
    
    obj = obj + 1
    location(obj) = Position
    writepdf (obj & " 0 obj")
    writepdf ("<<")
    writepdf ("  /Font << /F1 " & obj - 2 & " 0 R >>")
    writepdf ("  /ProcSet [ /PDF /Text ]")
    writepdf (">>")
    writepdf ("endobj")
  Exit Sub
er:
  MsgBox Err.Description
End Sub
  
Private Sub WritePages()
  Dim i As Integer
  Dim line As String, tmpline As String, beginstream As String
  On Error GoTo er
    Open filetxt For Input As #2
      beginstream = StartPage
      lineNo = -1
      Do Until EOF(2)
        Line Input #2, line
        lineNo = lineNo + 1
        
        'page break
        If lineNo >= lines Or InStr(line, Chr(12)) > 0 Then
          writepdf ("1 0 0 1 " & npagex & " " & npagey & " Tm")
          writepdf ("(" & pageNo & ") Tj")
          writepdf ("/F1 " & pointSize & " Tf")
          EndPage (beginstream)
          beginstream = StartPage
        End If
        
        line = ReplaceText(ReplaceText(line, "(", "\("), ")", "\)")
        line = Trim(line)
        
        If Len(line) > linelen Then
          
          'word wrap
          Do While Len(line) > linelen
            tmpline = left(line, linelen)
            For i = Len(tmpline) To Len(tmpline) \ 2 Step -1
              If InStr("*&^%$#,. ;<=>[])}!""", Mid(tmpline, i, 1)) Then
                tmpline = left(tmpline, i)
                Exit For
              End If
            Next
            
            line = Mid$(line, Len(tmpline) + 1)
            writepdf ("T* (" & tmpline & vbCrLf & ") Tj")
            lineNo = lineNo + 1
            
            'page break
            If lineNo >= lines Or InStr(line, Chr(12)) > 0 Then
              writepdf ("1 0 0 1 " & npagex & " " & npagey & " Tm")
              writepdf ("(" & pageNo & ") Tj")
              writepdf ("/F1 " & pointSize & " Tf")
              EndPage (beginstream)
              beginstream = StartPage
            End If
          Loop
          
          lineNo = lineNo + 1
          writepdf ("T* (" & line & vbCrLf & ") Tj")
        
        Else
          
          writepdf ("T* (" & line & vbCrLf & ") Tj")
        
        End If
      Loop
    Close #2
    writepdf ("1 0 0 1 " & npagex & " " & npagey & " Tm")
    writepdf ("(" & pageNo & ") Tj")
    writepdf ("/F1 " & pointSize & " Tf")
    EndPage (beginstream)
  Exit Sub
er:
  MsgBox Err.Description
  Close
End Sub

Private Function StartPage() As String
  Dim strmpos As Long
  On Error GoTo er
  obj = obj + 1
  location(obj) = Position
  pageNo = pageNo + 1
  pageObj(pageNo) = obj
  
  writepdf (obj & " 0 obj")
  writepdf ("<<")
  writepdf ("/Type /Page")
  writepdf ("/Parent " & Tpages & " 0 R")
  writepdf ("/Resources " & resources & " 0 R")
  obj = obj + 1
  writepdf ("/Contents " & obj & " 0 R")
  writepdf ("/Rotate " & rotate)
  writepdf (">>")
  writepdf ("endobj")
  
  location(obj) = Position
  writepdf (obj & " 0 obj")
  writepdf ("<<")
  writepdf ("/Length " & obj + 1 & " 0 R")
  writepdf (">>")
  writepdf ("stream")
  strmpos = Position
  writepdf ("BT")
  writepdf ("/F1 " & pointSize & " Tf")
  writepdf ("1 0 0 1 50 " & pageHeight - 40 & " Tm")
  writepdf (vertSpace & " TL")
  
  StartPage = strmpos
  Exit Function
er:
  MsgBox Err.Description
End Function

Function EndPage(streamstart As Long) As String
  Dim streamEnd As Long
  On Error GoTo er
    writepdf ("ET")
    streamEnd = Position
    writepdf ("endstream")
    writepdf ("endobj")
    obj = obj + 1
    location(obj) = Position
    writepdf (obj & " 0 obj")
    writepdf (streamEnd - streamstart)
    writepdf "endobj"
    lineNo = 0
  Exit Function
er:
  MsgBox Err.Description
End Function

Sub endpdf()
  Dim ty As String, i As Integer, xreF As Long
  On Error GoTo er
    location(root) = Position
    writepdf (root & " 0 obj")
    writepdf ("<<")
    writepdf ("/Type /Catalog")
    writepdf ("/Pages " & Tpages & " 0 R")
    writepdf (">>")
    writepdf ("endobj")
    location(Tpages) = Position
    writepdf (Tpages & " 0 obj")
    writepdf ("<<")
    writepdf ("/Type /Pages")
    writepdf ("/Count " & pageNo)
    writepdf ("/MediaBox [ 0 0 " & pageWidth & " " & pageHeight & " ]")
    ty = ("/Kids [ ")
    For i = 1 To pageNo
      ty = ty & pageObj(i) & " 0 R "
    Next i
    ty = ty & "]"
    writepdf (ty)
    writepdf (">>")
    writepdf ("endobj")
    xreF = Position
    writepdf ("0 " & obj + 1)
    writepdf ("0000000000 65535 f ")
    For i = 1 To obj
      writepdf (Format(location(i), "0000000000") & " 00000 n ")
    Next i
    writepdf ("trailer")
    writepdf ("<<")
    writepdf ("/Size " & obj + 1)
    writepdf ("/Root " & root & " 0 R")
    writepdf ("/Info " & info & " 0 R")
    writepdf (">>")
    writepdf ("startxref")
    writepdf (xreF)
    writepdf "%%EOF", True
  Exit Sub
er:
  MsgBox Err.Description
End Sub

Public Function FileExists(ByVal Filename As String) As Boolean
  On Error Resume Next
  FileExists = FileLen(Filename) > 0
  Err.Clear
End Function

Public Function ReplaceText(Text As String, TextToReplace As String, NewText As String) As String
  Dim mtext As String, SpacePos As Long
  mtext = Text
  SpacePos = InStr(mtext, TextToReplace)
  Do While SpacePos
    mtext = left(mtext, SpacePos - 1) & NewText & Mid(mtext, SpacePos + Len(TextToReplace))
    SpacePos = InStr(SpacePos + Len(NewText), mtext, TextToReplace)
  Loop
  ReplaceText = mtext
End Function

Function SaveDialog(Form1 As Form, Filter As String, Title As String, InitDir As String, DefaultFilename As String) As String
  Dim ofn As OPENFILENAME
  Dim a As Long
  On Local Error Resume Next
  ofn.lStructSize = Len(ofn)
  ofn.hwndOwner = Form1.hWnd
  ofn.hInstance = App.hInstance
  If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"

  For a = 1 To Len(Filter)
      If Mid$(Filter, a, 1) = "|" Then Mid$(Filter, a, 1) = Chr$(0)
  Next
  ofn.lpstrFilter = Filter
  ofn.lpstrFile = Space$(254)
  Mid(ofn.lpstrFile, 1, 254) = DefaultFilename
  ofn.nMaxFile = 255
  ofn.lpstrFileTitle = Space$(254)
  ofn.nMaxFileTitle = 255
  ofn.lpstrInitialDir = InitDir
  ofn.lpstrTitle = Title
  ofn.lpstrDefExt = "pdf"
  ofn.Flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT
  a = GetSaveFileName(ofn)


  If (a) Then
      SaveDialog = Trim$(ofn.lpstrFile)
  Else
      SaveDialog = ""
  End If
End Function

Function OpenDialog(Form1 As Form, Filter As String, Title As String, InitDir As String) As String
  Dim ofn As OPENFILENAME
  Dim a As Long
  On Local Error Resume Next
  ofn.lStructSize = Len(ofn)
  ofn.hwndOwner = Form1.hWnd
  ofn.hInstance = App.hInstance
  If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"

  For a = 1 To Len(Filter)
      If Mid$(Filter, a, 1) = "|" Then Mid$(Filter, a, 1) = Chr$(0)
  Next
  ofn.lpstrFilter = Filter
  ofn.lpstrFile = Space$(254)
  ofn.nMaxFile = 255
  ofn.lpstrFileTitle = Space$(254)
  ofn.nMaxFileTitle = 255
  ofn.lpstrInitialDir = InitDir
  ofn.lpstrTitle = Title
  ofn.Flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
  a = GetOpenFileName(ofn)

  If (a) Then
      OpenDialog = Trim$(ofn.lpstrFile)
  Else
      OpenDialog = ""
  End If
End Function



Private Sub chameleonButton1_Click()
If rtbnet.Text <> "" Then
HTMLEditor.Show
HTMLEditor.RichTextBox1 = rtbnet.Text
End If
End Sub

Private Sub chameleonButton2_Click()
Picture2.Visible = False
CmdGet.Enabled = True
TxtUrl.Locked = False
Me.Caption = "Source Code Organizer V.2.2"
sBar.Panels(1).Text = "Status: Viewing"
End Sub

Private Sub char_Click()
Symbols.Show
End Sub

Private Sub Check1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Instant internet Code search.By checking the box it will take you to..( Javascript.com )..."
End Sub

Private Sub Check2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Instant internet Code search.By checking the box it will take you to..( Planet-source-code.com )..."
End Sub

Private Sub Check3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Instant internet Code search.By checking the box it will take you to..( Simplythebest.net )..."
End Sub

Private Sub Check4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Instant internet Code search.By checking the box it will take you to..( Codeguru.com )..."
End Sub

Private Sub Check5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Instant internet Code search.By checking the box it will take you to..( Planet-source-code.com )..."
End Sub

Private Sub Check6_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Instant internet Code search.By checking the box it will take you to..( vyaskn.tripod.com )..."
End Sub
Sub FilesSearch(DrivePath As String, Ext As String)

Dim XDir() As String

Dim TmpDir As String

Dim FFound As String

Dim DirCount As Integer

Dim x As Integer

Dim li As ListItem

DirCount = 0

ReDim XDir(0) As String

XDir(DirCount) = ""

If Right(DrivePath, 1) <> "\" Then

DrivePath = DrivePath & "\"

End If

'Enter here the code for showing the pat
' h being
'search. Example: Form1.label2 = DrivePa
' th
'Search for all directories and store in
' the
'XDir() variable
sBar.Panels(1).Text = DrivePath

DoEvents

TmpDir = Dir(DrivePath, vbDirectory)


Do While TmpDir <> ""


If TmpDir <> "." And TmpDir <> ".." Then


If (GetAttr(DrivePath & TmpDir) And vbDirectory) = vbDirectory Then

XDir(DirCount) = DrivePath & TmpDir & "\"

DirCount = DirCount + 1
ReDim Preserve XDir(DirCount) As String
End If
End If
TmpDir = Dir
Loop
'Searches for the files given by extensi
' on Ext
FFound = Dir(DrivePath & Ext)
Do Until FFound = ""
Set li = ListView1.ListItems.Add(, , FFound)
li.ListSubItems.Add , , DrivePath
li.ListSubItems.Add , , FileLen(DrivePath & FFound) & " Bytes"
FFound = Dir
Loop
If checkscan.Value = 1 Then
For x = 0 To (UBound(XDir) - 1)
FilesSearch XDir(x), Ext
Next x
Else
End If
End Sub
Private Sub checkscan_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Check this box to scan sub directory..."
End Sub
Private Sub chksq1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Insert comma..."
End Sub

Private Sub Clearz_Click()
RichTextBox1.Text = ""
End Sub

Private Sub cmdfile_Click(Index As Integer)
PopupMenu file
End Sub

Private Sub cmdfont_Click(Index As Integer)
PopupMenu fontme
End Sub

Private Sub CmdGet_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Get Website HTML..."
End Sub

Private Sub cmdInsert_Click(Index As Integer)
PopupMenu insert
End Sub

Private Sub cmdother_Click(Index As Integer)
PopupMenu other
End Sub

Private Sub cmdprev_Click(Index As Integer)
Open App.Path & "\Organizerpreview\extreme.html" For Output As #1
Print #1, rtbnet.Text
Close #1
Load Browser
Browser.Text1.Text = App.Path & "\Organizerpreview\extreme.html"
Browser.Caption = "You are Browsing:"
Browser.Show
Browser.Web.Navigate App.Path & "\Organizerpreview\extreme.html"
End Sub

Private Sub cmdsq1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Insert SQL character..."
End Sub

Private Sub cmdsq2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Find Database..."
End Sub
Private Sub cmdStart_Click()
mnutest_Click
cmdStart.Enabled = False
ListView1.ListItems.Clear
Screen.MousePointer = vbHourglass
FilesSearch Dir1.Path, "*.mdb"
Screen.MousePointer = vbDefault
sBar.Panels(1).Text = ListView1.ListItems.Count & " Database's found!"
Label6.Caption = ListView1.ListItems.Count & " Database's found!"
cmdStart.Enabled = True
End Sub

Private Sub cmdStart_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Click to scan database..."
End Sub

Private Sub cmdtable_Click(Index As Integer)
PopupMenu table
End Sub

Private Sub col1_Click()
rtbnet.SelText = rtbnet.SelText + Chr(13) + Chr(10) + "<P><TABLE BORDER=1>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=Write the code of the cell's background>" + Chr(13) + Chr(10) + "<P>Your your text in the first cell" + Chr(13) + Chr(10) + "</TD></TR>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=Write the code of the cell's background>" + Chr(13) + Chr(10) + "<P>Write your text in the second cell" + Chr(13) + Chr(10) + "</TD></TR>" + Chr(13) + Chr(10) + "xxxxxxxxxxxxxxxxxxxxxxxxxxx" + Chr(13) + Chr(10) + "Copy and Paste the following code to add more cells" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=Write the code of the cell's background>" + Chr(13) + Chr(10) + "<P>Your your text in the cell" + Chr(13) + Chr(10) + "</TD></TR>" + Chr(13) + Chr(10) + "xxxxxxxxxxxxxxxxxxxxxxxxx" + Chr(13) + Chr(10) + "</TABLE></P>"
End Sub

Private Sub col2_Click()
rtbnet.SelText = rtbnet.SelText + "<P><TABLE BORDER=1 bgcolor=Write the color-code for all cells>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (1st cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (2st cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD></TR>ADD ROWS HERE" + Chr(13) + Chr(10) + "</TABLE></P>"
End Sub

Private Sub col3_Click()
rtbnet.SelText = rtbnet.SelText + "<P><TABLE BORDER=1 bgcolor=Write the color-code for all cells>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (1st cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (2nd cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (3rd cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD></TR>ADD ROWS HERE" + Chr(13) + Chr(10) + "</TABLE></P>"
End Sub

Private Sub col4_Click()
rtbnet.SelText = rtbnet.SelText + "<P><TABLE BORDER=1 bgcolor=Write the color-code for all cells>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (1st cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (2nd cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (3rd cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD></TR>ADD ROWS HERE" + Chr(13) + Chr(10) + "</TABLE></P>"
End Sub
Private Sub Command1_Click()
sBar.Panels(1).Text = "Status: Viewing"
Picture2.Visible = True
CmdGet.Enabled = False
TxtUrl.Locked = True
Open App.Path & "\Organizerpreview\extreme.html" For Output As #1
Print #1, rtbnet.Text
Close #1
Load Browser
Textweb.Text = App.Path & "\Organizerpreview\extreme.html"
Web.Navigate App.Path & "\Organizerpreview\extreme.html"
End Sub

Private Sub Command2_Click()
Clipboard.Clear
Clipboard.SetText rtbnet.Text
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Viewing..."
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Copy HTML code..."
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Insert Tag..."
End Sub
Private Sub Command5_Click()
frmfinddb.Show
End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Clear field..."
End Sub



Private Sub copy1_Click()
On Error Resume Next
Clipboard.Clear
Clipboard.SetText rtbCodeWindow.Text
Clipboard.SetText rtbnet.SelText
End Sub

Private Sub cut_Click()
 Clipboard.SetText rtbnet.SelText
 rtbnet.SelText = ""
End Sub

Private Sub delete_Click()
SSTab1.Tab = 1
Picture1.Visible = True
End Sub

Private Sub Dir1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Select a folder you want to scan..."
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
End Sub


Private Sub CmdGet_Click()
saved = False
If IsNetConnectOnline = False Then
MsgBox "You are not connected to the internet!", vbOKOnly, App.ProductName
Unload Me
Else
If TxtUrl.Text = "http://" Then
MsgBox ("This is a malformed url!" & vbCrLf & "Operation has been canceled"), , App.ProductName
Else
mnutest_Click
Screen.MousePointer = vbHourglass
rtbnet.Text = Inet1.OpenURL(TxtUrl.Text)
Me.Caption = "Source Code Organizer V.2.2" & "-HTML Code for-" & " [ " & TxtUrl.Text & " ]"
Screen.MousePointer = vbDefault
CmdGet.Enabled = False
End If
End If
End Sub



Private Sub editme_Click()
If rtbCodeWindow.Text <> "" Then
note.Show
note.Rich = rtbCodeWindow.Text
mnuSave.Enabled = False
mnudelete.Enabled = False
mnunew.Enabled = False
mnupaste.Enabled = False
mnuModify.Enabled = False
mnusettingz.Enabled = False
mnuOpen.Enabled = False
End If
End Sub

Private Sub fileview_Click()
frmfind.Show
End Sub

Private Sub Form_Resize()
ResizeForm frmmain
End Sub

Private Sub Form_Unload(cancel As Integer)
ShowProgressInStatusBar False
End Sub

Private Sub Frame6_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Code Language Selection..."
End Sub

Private Sub Frame7_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = ""
End Sub

Private Sub Frame8_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = ""
End Sub

Private Sub hideme_Click()
Picture1.Visible = False
SSTab1.Tab = 0
End Sub

Private Sub html_Click()
tabSnippit.Tab = 1
End Sub

Private Sub htmlsetprop_Click()
 rtbCodeWindow.SelStart = 0
 rtbCodeWindow.SelLength = Len(rtbCodeWindow.Text)
 rtbCodeWindow.SetFocus
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Click an item to insert..."
End Sub

Private Sub List2_Click()
PopupMenu mnuTables
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Scanned Data Base Files..."
End Sub

Private Sub lstTitles_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
      PopupMenu mnufilelist
    End If
End Sub

Private Sub mnu1_Click()
If MsgBox("Do you want to save your current project?", vbYesNo, "Save") = vbYes Then
        mnu3_Click
       rtbnet.Text = "<!DOCTYPE HTML PUBLIC " & Chr(34) & "-//W3C//Source Code Organizer HTML Editor//EN" & Chr(34) & ">" & vbCrLf
    rtbnet.Text = rtbnet.Text & vbCrLf & "<html>" & vbCrLf & "<head>" & vbCrLf & "<title>My Web Page</title>" & vbCrLf & "</head>" & vbCrLf & vbCrLf & "<body>" & vbCrLf & vbCrLf & "</body>" & vbCrLf & "</html>" & vbCrLf
    Else
        rtbnet.Text = "<!DOCTYPE HTML PUBLIC " & Chr(34) & "-//W3C//Source Code Organizer HTML Editor//EN" & Chr(34) & ">" & vbCrLf
    rtbnet.Text = rtbnet.Text & vbCrLf & "<html>" & vbCrLf & "<head>" & vbCrLf & "<title>My Web Page</title>" & vbCrLf & "</head>" & vbCrLf & vbCrLf & "<body>" & vbCrLf & vbCrLf & "</body>" & vbCrLf & "</html>" & vbCrLf
    End If
End Sub

Private Sub mnu2_Click()
Dim sFile As String
    With CommonDialog1
        .DialogTitle = "Open"
        .CancelError = False
        .Filter = "HTML Files (*.html)|*.html|All Files (*.*)|*.*|"
        .ShowOpen
        If Len(.Filename) = 0 Then
            Exit Sub
        End If
        sFile = .Filename
 Dim intFileNum As Integer
 Dim strTextLine As String, strFileName As String
 intFileNum = FreeFile
 Open App.Path & "\recent.dat" For Append As #intFileNum
 Print #intFileNum, sFile
 Close #intFileNum
    
    End With
    rtbnet.LoadFile sFile
End Sub

Private Sub mnu3_Click()
CommonDialog1.Filter = "HTML Files (*.html)|*.html|HTM Files (*.htm)|*.htm)|RTF Files (*.rtf)|*.rtf|Text files (*.txt)|*.txt|Ini Files (*.ini)|*.ini|Registry Files (*.log)|*.log|Batch File (*.bat)|*.bat|All files (*.*)|*.*"""
CommonDialog1.ShowSave
If CommonDialog1.Filename <> "" Then
    Open CommonDialog1.Filename For Output As #1
    Print #1, rtbnet.Text
    Close #1
End If
End Sub

Private Sub mnuaddmod01_Click()
SSTab1.Tab = 2
End Sub

Private Sub mnubackgroundcol_Click()
frmTagEdit.Show
'frmTool_Rainbow.Show
End Sub

Private Sub mnubackup_Click()
Call BackUp
End Sub

Private Sub mnuchardex_Click()
Symbols.Show
End Sub

Private Sub mnucol_Click()
MsgBox "Not implemted yet..", vbInformation, "Information"
End Sub

Private Sub mnucolorpicker_Click()
frmColor.Show
End Sub

Private Sub mnucont_Click()
SSTab1.Tab = 0
End Sub

Private Sub mnucopy2_Click()
Clipboard.Clear
    Clipboard.SetText rtbCodeWindow.Text
End Sub

Private Sub mnudelete01_Click()
mnuDelete_Click
End Sub

Private Sub mnuftp_Click()
frmconnect.Show
End Sub

Private Sub mnuhelpcon_Click()
On Error GoTo HandleErrors
 Call Shell("HH.exe help.chm", vbNormalFocus)

 Exit Sub
HandleErrors:
  Dim intresponse As Integer
   Select Case Err.Number
        Case 53, 76
         intresponse = MsgBox("File not found.", vbCritical, "Error")
         End Select
End Sub

Private Sub mnuicon_Click()
frmscan.Show
End Sub

Private Sub mnuinsertagz_Click()
frmTags.Show , frmmain
End Sub

Private Sub mnumarqalternateleft_Click()
rtbnet.SelText = "<Marquee bgcolor= #FFFFFF  Behavior = Alternate  Direction =  Left >" + rtbnet.SelText + "Enter Your Text Here" + "</Marquee>"
End Sub

Private Sub mnumarqalternateright_Click()
rtbnet.SelText = "<Marquee bgcolor= #FFFFFF  Behavior = Alternate  Direction =  Right >" + rtbnet.SelText + "Enter Your Text Here" + "</Marquee>"
End Sub

Private Sub mnumarqscrolldown_Click()
rtbnet.SelText = "<marquee bgcolor= #FFFFFF & Direction = Down > " + rtbnet.SelText + "Enter Your Text Here" + "</marquee >"
End Sub

Private Sub mnumarqscrollright_Click()
rtbnet.SelText = "<marquee bgcolor= #FFFFFF & Direction = right > " + rtbnet.SelText + "Enter Your Text Here" + "</marquee >"
End Sub

Private Sub mnumarqscrollup_Click()
rtbnet.SelText = "<marquee bgcolor= #FFFFFF & Direction = UP > " + rtbnet.SelText + "Enter Your Text Here" + "</marquee >"
End Sub

Private Sub mnumarqslidelft_Click()
RichTextBox1.SelText = "<marquee bgcolor= #FFFFFF & Behavior = Slide & Direction = Left > " + RichTextBox1.SelText + "Hello World" + "</marquee >"
End Sub

Private Sub mnumarqslideright_Click()
rtbnet.SelText = "<marquee bgcolor= #FFFFFF & Behavior = Slide & Direction = Right > " + rtbnet.SelText + "Hello World" + "</marquee >"
End Sub

Private Sub mnumarqsrcolleft_Click()
rtbnet.SelText = "<marquee bgcolor= #FFFFFF & Direction = Left > " + rtbnet.SelText + "Enter Your Text Here" + "</marquee >"
End Sub

Private Sub mnuModify_Click()
  SSTab1.Tab = 2
End Sub

Private Sub mnupaste2_Click()
rtbCodeWindow.SelText = Clipboard.GetText
rtbCodeWindow.SetFocus
End Sub

Private Sub mnupdf_Click()
SSTab1.Tab = 1
End Sub

Private Sub mnuprint2_Click()
mnusettingz_Click
End Sub

Private Sub mnuset_Click()

 CmDlg.Flags = cdlPDPrintSetup
            CmDlg.ShowPrinter
            DoEvents
End Sub

Private Sub mnusearchcodez_Click()
  SSTab1.Tab = 1
End Sub

Private Sub mnusettingz_Click()
   If gintIndex = 0 Then Exit Sub
    'This is the basic undo stuff.
    gblnIgnoreChange = True
    gintIndex = gintIndex - 1
    On Error Resume Next
    rtbCodeWindow.Text = gstrStack(gintIndex)
    gblnIgnoreChange = False
End Sub
Private Sub mnusetupme_Click()
SSTab1.Tab = 0
End Sub

Private Sub mnutagz_Click()
frmTags.Show
End Sub

Private Sub mnutest_Click()
ProgressBar1.Min = 0
    ProgressBar1.Max = 100
'
' Show ProgressBar in Status Bar
'
    ShowProgressInStatusBar True
'
' Enable the timer so it looks like we're doing something
'
    Timer1.Enabled = True
End Sub

Private Sub mnuusesounds_Click()
If mnuusesounds.Checked = False Then
        mnuusesounds.Checked = True
        UseSound = "Yes"
    ElseIf mnuusesounds.Checked = True Then
        mnuusesounds.Checked = False
        UseSound = ""
    End If
End Sub

Private Sub mnuweb_Click()
On Error Resume Next
    Dim xRet As Long
    xRet = ShellExecute(0, vbNullString, "http://clik.to/ret", vbNullString, App.Path, 1)
End Sub

Private Sub mr1_Click()
rtbnet.SelText = rtbnet.SelText + "<P><TABLE BORDER=1 bgcolor=Write the color-code for all cells>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (1st cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (2nd cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (3rd cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "ADD HERE COLUMNS. Select and Paste one of the two lines <P></TD><TD>" + Chr(13) + Chr(10) + "<P>Write your Text here (The cell of the LAST COLUMN)" + Chr(13) + Chr(10) + "</TD></TR>ADD ROWS HERE" + Chr(13) + Chr(10) + "</TABLE></P>"
End Sub

Private Sub Option1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Insert Auctioner tag..."
End Sub

Private Sub Option2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Insert trader  tag..."
End Sub

Private Sub ot1_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + Chr(13) + Chr(10) + "<UL>" + Chr(13) + Chr(10) + "<LI> Type your text here" + Chr(13) + Chr(10) + "<LI> Type your text here" + Chr(13) + Chr(10) + "<LI> Type your text here and add more <LI> if necessary" + Chr(13) + Chr(10) + "</UL>"
End Sub

Private Sub ot2_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + Chr(13) + Chr(10) + "<OL>" + Chr(13) + Chr(10) + "<LI> Type your text here" + Chr(13) + Chr(10) + "<LI> Type your text here" + Chr(13) + Chr(10) + "<LI> Type your text here and add more <LI> if necessary" + Chr(13) + Chr(10) + "</OL>"
End Sub


Private Sub paste2_Click()
 frmmain.rtbnet.SelText = Clipboard.GetText()
End Sub

Private Sub print2_Click()
 rtbnet.SelStart = 0
 rtbnet.SelLength = Len(rtbnet.Text)
 rtbnet.SetFocus
End Sub

Private Sub property_Click()
frmTagEdit.Show
End Sub

Private Sub RichTextBox1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "This is the window for SQL coding..."
End Sub

Private Sub rtbCodeWindow_Change()
If Not gblnIgnoreChange Then
        gintIndex = gintIndex + 1
        gstrStack(gintIndex) = rtbCodeWindow.Text
    End If
If UseSound = "Yes" Then
        Dim Play As String
        Play = sndPlaySound(App.Path + "\Organizerpreview\Type.wav", SND_ASYNC)
    End If
End Sub

Private Sub rtbCodeWindow_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then ' used for popup menu for the list of ftp sites saved for reuse utilizing the writeprivateprofile string
If rtbCodeWindow.Text = "" Then
mnucopy2.Enabled = False
mnusetupme.Enabled = False
mnupaste2.Enabled = True
mnuprint2.Enabled = True
editme.Enabled = False
delete.Enabled = False
space9.Enabled = False
PopupMenu mnusetz
Else
mnucopy2.Enabled = True
mnuprint2.Enabled = True
setup1.Enabled = True
mnupaste2.Enabled = True
editme.Enabled = True
delete.Enabled = True
space9.Enabled = True
PopupMenu mnusetz
End If
End If
End Sub

Private Sub rtbCodeWindow_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Source Code..."
End Sub

Private Sub rtbnet_Change()
If Not gblnIgnoreChange Then
        gintIndex = gintIndex + 1
        gstrStack(gintIndex) = rtbnet.Text
    End If
If UseSound = "Yes" Then
        Dim Play As String
        Play = sndPlaySound(App.Path + "\Organizerpreview\Type.wav", SND_ASYNC)
    End If
End Sub

Private Sub rtbnet_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then ' used for popup menu for the list of ftp sites saved for reuse utilizing the writeprivateprofile string
If rtbnet.Text = "" Then
copy1.Enabled = False
print2.Enabled = False
setup1.Enabled = False
paste2.Enabled = True
cut.Enabled = False
char.Enabled = False
browse.Enabled = False
mnuinsertagz.Enabled = False
property.Enabled = False
setup1.Enabled = True
PopupMenu menuz
Else
char.Enabled = True
browse.Enabled = True
mnuinsertagz.Enabled = True
property.Enabled = True
copy1.Enabled = True
cut.Enabled = True
print2.Enabled = True
setup1.Enabled = True
paste2.Enabled = True
PopupMenu menuz
End If
End If
End Sub

Private Sub rtbnet_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then ' used for popup menu for the list of ftp sites saved for reuse utilizing the writeprivateprofile string
If rtbnet.Text = "" Then
copy1.Enabled = False
print2.Enabled = False
setup1.Enabled = False
paste2.Enabled = True
cut.Enabled = False
PopupMenu menuz
Else
copy1.Enabled = True
cut.Enabled = True
print2.Enabled = True
setup1.Enabled = True
paste2.Enabled = True
PopupMenu menuz

End If
End If

sBar.Panels(1).Text = "This is the window that will display the HTML code..."
End Sub
Private Sub rtbNotes_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Source code notes/author information..."
End Sub

Private Sub sClose_Click()
End
End Sub

Private Sub setup1_Click()
If gintIndex = 0 Then Exit Sub
    'This is the basic undo stuff.
    gblnIgnoreChange = True
    gintIndex = gintIndex - 1
    On Error Resume Next
    rtbnet.Text = gstrStack(gintIndex)
    gblnIgnoreChange = False
End Sub

Private Sub show1_Click()
SSTab1.Tab = 1
Picture1.Visible = True
End Sub

Private Sub sound_Click()
 If sound.Checked = False Then
       sound.Checked = True
       UseSound = "Yes"
    ElseIf sound.Checked = True Then
        sound.Checked = False
        UseSound = ""
   End If
End Sub
Private Sub space9_Click()
htmlsetprop_Click
End Sub

Private Sub sql_Click()
tabSnippit.Tab = 2
mnucopy.Enabled = False
mnupaste.Enabled = False
mnudelete.Enabled = False
mnuSave.Enabled = False
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
 If SSTab1.Tab = 0 Then
 Picture1.Visible = False
 Picture3.Visible = False
 txtTitle.Visible = True
 Me.Caption = "Source Code Organizer V.2.2"
 End If
 If SSTab1.Tab = 1 Then
 Picture3.Visible = False
 txtTitle.Visible = True
 End If
 If SSTab1.Tab = 2 Then
 Picture1.Visible = False
 Picture3.Visible = False
 txtTitle.Visible = True
 Me.Caption = "Source Code Organizer V.2.2"
 End If
End Sub
Private Sub tabSnippit_Click(PreviousTab As Integer)
If tabSnippit.Tab = 0 Then
 Picture1.Visible = False
 Picture3.Visible = False
 txtTitle.Visible = True
 mnusettingz.Enabled = True
 htmlsetprop.Enabled = True
 mnucopy.Enabled = True
 mnupaste.Enabled = True
 mnudelete.Enabled = True
 mnuSave.Enabled = True
 mnuFind.Enabled = True
 sql.Enabled = True
 mnunew.Enabled = True
 mnuOpen.Enabled = True
 mnuModify.Enabled = True
 hideme.Enabled = True
 show1.Enabled = True
 html.Enabled = True
 mnuusesounds.Enabled = False
 Me.Caption = "Source Code Organizer V.2.2"
 SSTab1.Tab = 0
 End If
 If tabSnippit.Tab = 1 Then
  Picture3.Visible = False
  Picture1.Visible = False
  txtTitle.Visible = True
  mnusettingz.Enabled = False
  mnuFind.Enabled = False
  htmlsetprop.Enabled = True
  mnupaste.Enabled = True
  sql.Enabled = True
  mnudelete.Enabled = False
  mnuSave.Enabled = False
  mnunew.Enabled = False
  mnuOpen.Enabled = False
  mnuModify.Enabled = False
  hideme.Enabled = False
  show1.Enabled = False
  htmlsetprop.Enabled = False
  mnupaste.Enabled = False
  mnucopy.Enabled = False
  html.Enabled = False
  mnuusesounds.Enabled = True
  SSTab1.Tab = 0
  End If
  If tabSnippit.Tab = 2 Then
   Picture3.Visible = False
   Picture1.Visible = False
   txtTitle.Visible = True
   mnunew.Enabled = False
   mnusettingz.Enabled = False
   htmlsetprop.Enabled = False
   mnuOpen.Enabled = False
   mnuModify.Enabled = False
   mnuFind.Enabled = False
   sql.Enabled = False
   hideme.Enabled = False
   show1.Enabled = False
   html.Enabled = True
   mnuusesounds.Enabled = False
   mnuSave.Enabled = False
   mnudelete.Enabled = False
   mnupaste.Enabled = False
   mnucopy.Enabled = False
   Me.Caption = "Source Code Organizer V.2.2"
   SSTab1.Tab = 0
  End If
End Sub

Private Sub tabSnippit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = ""
End Sub

Private Sub Text2_Click()
Text2.Text = ""
End Sub

Private Sub time_Click()
DateTime.Show
End Sub

Private Sub Timer1_Timer()
Static lCount As Long
    
    lCount = lCount + 5
    
    If lCount > 100 Then
        Timer1.Enabled = False
        ShowProgressInStatusBar False
        Command1.Enabled = True
        lCount = 0
    End If
    
    ProgressBar1.Value = lCount
End Sub

Private Sub txtTitle_Click()
SSTab1.Tab = 1
Picture1.Visible = True
End Sub

Private Sub TxtUrl_Change()
CmdGet.Enabled = True
rtbnet.Text = ""
If UseSound = "Yes" Then
        Dim Play As String
        Play = sndPlaySound(App.Path + "\Organizerpreview\Type.wav", SND_ASYNC)
    End If
End Sub

Private Sub TxtUrl_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
CmdGet_Click
End If
End Sub
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If ListView1.SortKey = ColumnHeader.Index - 1 Then
If ListView1.SortOrder = lvwAscending Then
ListView1.SortOrder = lvwDescending
Else
ListView1.SortOrder = lvwAscending
End If
Else
ListView1.SortOrder = lvwAscending
ListView1.SortKey = ColumnHeader.Index - 1
End If
ListView1.Sorted = True
End Sub

Private Sub cmbFilter_DropDown()
sBar.Panels(1).Text = "Code Language Filter..."
End Sub

Private Sub cmbType_DropDown()
sBar.Panels(1).Text = "Code Language Selection..."
End Sub

Private Sub cmdSearch_Click()
  mnutest_Click
  If txWhatSearch = "" Then Exit Sub
  Search cmEngines.ListIndex, txWhatSearch
  tabSnippit.Tab = 0
End Sub

Private Sub cmbFilter_Click()
    LoadGridBox
End Sub



Private Sub cmdSearch_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Search Code in the internet..."
End Sub

Private Sub CmdsearchDB_Click()

 tabSnippit.Tab = 0
 Dim RetVal As String
    RetVal = Text2.Text
    If RetVal = "" Then
        Exit Sub
    End If
    Find (RetVal)
    mnutest_Click
End Sub

Private Sub CmdsearchDB_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Search Code in the DataBase..."
End Sub

Private Sub cmEngines_DropDown()
sBar.Panels(1).Text = "Search Engine Selection..."
End Sub

Private Sub Command3_Click()
On Error GoTo errHandler
    mnutest_Click
    
    Dim RetVal As String
    Dim adoCon As Connection
    Dim adoCmd As Command
    
    RetVal = Text1.Text
    If RetVal = "" Then
     MsgBox "You must enter a new code language in the Text field", vbInformation, "Information"
       
        Exit Sub
    End If
    
    'connect to the database and add in the new codetype
    Set adoCon = CreateObject("ADODB.Connection")
    adoCon.open gblConnectString
    Set adoCmd = CreateObject("ADODB.Command")
    adoCmd.ActiveConnection = adoCon
    adoCmd.CommandText = "INSERT INTO codetypes (codetype) VALUES ('" & StuffQuotes(RetVal) & "')"
    adoCmd.Execute
    MsgBox "Your new code language is now added to the DataBase.Click done at the bottom to go back to the source code and look for it in the dropdown box", vbInformation, "Added to the DataBase"
    'clean up
    Set adoCmd = Nothing
    adoCon.Close
    Set adoCon = Nothing
    LoadCodeTypes
    Exit Sub
    
errHandler:
    'in case anything went wrong
    MsgBox "Error: " & Err.Description
End Sub

Private Sub Form_Load()
      mciconmenu1.HighlightStyle = ChkHighLight.Value
mciconmenu1.SubClassMenu Me
mciconmenu1.ImageList = ImageList1
Set mciconmenu1.BackgroundPicture = picBackGround

mciconmenu1.ItemIcon("mnuNew") = ImageList1.ListImages.Item(1).Index - 1
mciconmenu1.ItemIcon("mnuOpen") = ImageList1.ListImages.Item(2).Index - 1
mciconmenu1.ItemIcon("mnubackup") = ImageList1.ListImages.Item(3).Index - 1
mciconmenu1.ItemIcon("mnucont") = ImageList1.ListImages.Item(22).Index - 1
mciconmenu1.ItemIcon("mnunet") = ImageList1.ListImages.Item(18).Index - 1
mciconmenu1.ItemIcon("mnuFind") = ImageList1.ListImages.Item(16).Index - 1
mciconmenu1.ItemIcon("mnuExit") = ImageList1.ListImages.Item(5).Index - 1
mciconmenu1.ItemIcon("mnucopy") = ImageList1.ListImages.Item(13).Index - 1
mciconmenu1.ItemIcon("mnupaste") = ImageList1.ListImages.Item(14).Index - 1
mciconmenu1.ItemIcon("mnusettingz") = ImageList1.ListImages.Item(10).Index - 1
mciconmenu1.ItemIcon("htmlsetprop") = ImageList1.ListImages.Item(11).Index - 1
mciconmenu1.ItemIcon("mnuModify") = ImageList1.ListImages.Item(15).Index - 1
mciconmenu1.ItemIcon("mnuSave") = ImageList1.ListImages.Item(3).Index - 1
mciconmenu1.ItemIcon("mnuusesounds") = ImageList1.ListImages.Item(23).Index - 1
mciconmenu1.ItemIcon("mnuicon") = ImageList1.ListImages.Item(18).Index - 1
mciconmenu1.ItemIcon("fileview") = ImageList1.ListImages.Item(17).Index - 1
mciconmenu1.ItemIcon("sql") = ImageList1.ListImages.Item(27).Index - 1
mciconmenu1.ItemIcon("html") = ImageList1.ListImages.Item(24).Index - 1
mciconmenu1.ItemIcon("hideme") = ImageList1.ListImages.Item(28).Index - 1
mciconmenu1.ItemIcon("show1") = ImageList1.ListImages.Item(15).Index - 1
mciconmenu1.ItemIcon("mnupdf") = ImageList1.ListImages.Item(26).Index - 1
mciconmenu1.ItemIcon("mnuhelpcon") = ImageList1.ListImages.Item(19).Index - 1
mciconmenu1.ItemIcon("mnuAbout") = ImageList1.ListImages.Item(20).Index - 1
mciconmenu1.ItemIcon("mnucontact") = ImageList1.ListImages.Item(30).Index - 1
mciconmenu1.ItemIcon("mnuweb") = ImageList1.ListImages.Item(21).Index - 1
mciconmenu1.ItemIcon("mnuPlanetSourceCode") = ImageList1.ListImages.Item(29).Index - 1
mciconmenu1.ItemIcon("mnuftp") = ImageList1.ListImages.Item(31).Index - 1
mciconmenu1.ItemIcon("mnuhtmlterms") = ImageList1.ListImages.Item(24).Index - 1
mciconmenu1.ItemIcon("mnustand") = ImageList1.ListImages.Item(27).Index - 1
'------------------------------------------------------
'Add the microsoft visual basic logo to help menu
'------------------------------------------------------
'mciconmenu1.ItemPicture 2, 27, picMenuVB.Picture
'------------------------------------------------------
'Aling the help menu in right
'------------------------------------------------------
'mciconmenu1.ItemRight "Help", 3 '<-File + Edit + Search
'------------------------------------------------------
'Remove X button
'------------------------------------------------------
myLONG = mciconmenu1.SystemMenuCount
mciconmenu1.SystemMenuRemoveItem myLONG

txtCreator.Text = AppName
  cmbFont.ListIndex = 1
  cmbFontSize.ListIndex = 1
  cmbRotation.ListIndex = 0
  cmbPageSize.ListIndex = 0

  cmdline = LCase(Command)
  If cmdline Like """*""" Then
    cmdline = Mid(cmdline, 2, Len(cmdline) - 2)
  End If
  
  If FileExists(cmdline) Then
    txtFilename.Text = cmdline
    txtOutputFile.Text = left(cmdline, Len(cmdline) - 4) & ".pdf"
    btnConvert_Click
  End If

     If Textweb.Text <> "URL Address" Then Web.Navigate (Textweb.Text)
     rtbnet.Text = "<!DOCTYPE HTML PUBLIC " & Chr(34) & "-//W3C//Source Code Organizer HTML Editor//EN" & Chr(34) & ">" & vbCrLf
    rtbnet.Text = rtbnet.Text & vbCrLf & "<html>" & vbCrLf & "<head>" & vbCrLf & "<title>My Web Page</title>" & vbCrLf & "</head>" & vbCrLf & vbCrLf & "<body>" & vbCrLf & vbCrLf & "</body>" & vbCrLf & "</html>" & vbCrLf
      cmEngines.ListIndex = 0
  txWhatSearch.ListIndex = 3
    gblConnectString = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & AppPath & "sourcebook.mdb"
    gblNewCode = True 'start off with a clean slate
    
    LoadCodeTypes
    rtbNotes.OLEDropMode = 1  'setup for the drag drop in the code windows
    On Error GoTo errHandler
    
    Dim adoCon As Connection
    Dim adoCmd As Command
    Dim adoRS As Recordset

    lstModify.Clear
    lstDelete.Clear
    'create the connection and execute the SQL
    Set adoCon = CreateObject("ADODB.Connection")
    adoCon.open gblConnectString
    Set adoCmd = CreateObject("ADODB.Command")
    adoCmd.ActiveConnection = adoCon
    adoCmd.CommandText = "SELECT * FROM codetypes"
    Set adoRS = adoCmd.Execute
            
    Do While Not adoRS.EOF
        'add the recordsetset items into the listboxes
        lstModify.AddItem CStr(adoRS("codetype"))
        lstDelete.AddItem CStr(adoRS("codetype"))
        adoRS.MoveNext
    Loop
    
    'make sure we clean up!
    Set adoCmd = Nothing
    adoCon.Close
    Set adoCon = Nothing
    Exit Sub
    
errHandler:
    'just incase anything went wrong
    MsgBox "Error: " & Err.Description
        
End Sub
Private Sub cmdsq1_Click()
Dim quote As String
Dim DblQuote As String
quote = """"
DblQuote = quote & quote
If RichTextBox1.Text <> "" Then
If Option1.Value = True Then
RichTextBox1.SelStart = 0
RichTextBox1.SelLength = 1
RichTextBox1.SelText = ""
RichTextBox1.SelText = "<#" & LCase(RichTextBox1.Text) & " #>"
Else
If Combo1.Text <> "" Then
RichTextBox1.SelStart = 0
RichTextBox1.SelLength = 1
RichTextBox1.SelText = ""
RichTextBox1.SelText = "[!Query:" & Combo1.Text & " Name=" & DblQuote & " SQL =" & quote & LCase(RichTextBox1.Text) & quote & "!]"
Else
RichTextBox1.SelStart = 0
RichTextBox1.SelLength = 1
RichTextBox1.SelText = ""
RichTextBox1.SelText = "[!Query:Open" & " Name=" & DblQuote & " SQL =" & quote & LCase(RichTextBox1.Text) & quote & "!]"
End If
End If
Else
Exit Sub
End If
End Sub

Private Sub cmdsq2_Click()
frmmain.List2.Clear
frmmain.List3.Clear
Form16.Show
End Sub

Private Sub Combo1_Click()
If Combo1.List(Combo1.ListIndex) = "Close" Then
If Option1.Value = True Then
frmmain.RichTextBox1.SelText = "<#Query:Close#>"
Unload Me
Else
frmmain.RichTextBox1.SelText = "[!Query:Close!]"
End If
Else
' Do Nothing as the user wants to build a custom sql statement
End If
End Sub


Private Sub Textweb_GotFocus()
    Textweb.SelStart = 0
    Textweb.SelLength = Len(Text1)
End Sub

Private Sub Textweb_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Web.Navigate Me.Textweb.Text
    End If
End Sub

Private Sub Web_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    Textweb.Text = URL
End Sub

Private Sub Web_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
    On Error GoTo progressERR
    If Progress = -1 Then PBar.Value = 100

    If Progress > 0 And ProgressMax > 0 Then
        PBar.Value = Progress * 100 / ProgressMax
       
    End If
    PBar.Visible = False
    Exit Sub
progressERR:
End Sub

Private Sub Web_StatusTextChange(ByVal Text As String)
    sBar.Panels(1).Text = Text
End Sub

Private Sub Command4_Click()
frmsqltag.Show
End Sub

Private Sub List1_Click()
RichTextBox1.SelText = RichTextBox1.SelText & " " & UCase(List1.List(List1.ListIndex))
End Sub

Private Sub List1_DblClick()
RichTextBox1.SelText = RichTextBox1.SelText & " " & UCase(List1.List(List1.ListIndex))
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
List1_DblClick
End If
End Sub

Private Sub List2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
If List2.ListCount > 0 Then
PopupMenu mnuTables
End If
End If
End Sub

Private Sub lstTitles_ItemClick(ByVal Item As MSComctlLib.ListItem)
tabSnippit.Tab = 0
End Sub

Private Sub MnuEnterAsStatement_Click()
RichTextBox1.SelText = RichTextBox1.SelText & " " & UCase(List2.List(List2.ListIndex))
End Sub

Private Sub MnuGetColumns_Click()

End Sub



Private Sub LoadGridBox()

    On Error GoTo errHandler
    
    Dim adoCon As Connection
    Dim adoCmd As Command
    Dim adoRS As Recordset
    Dim cmdtext As String
    Dim lstItem As ListItem

    lstTitles.ListItems.Clear
    ' here we are building the SQL statement based upon the filter drop down
    If cmbFilter.Text = "No Filter" Then
        cmdtext = "SELECT id, title, codetype FROM source "
    Else
        cmdtext = "SELECT id, title, codetype FROM source WHERE codetype='" & StuffQuotes(cmbFilter.Text) & "' "
    End If
    
    'connect to the database and retrieve the code
    Set adoCon = CreateObject("ADODB.Connection")
    adoCon.open gblConnectString
    Set adoCmd = CreateObject("ADODB.Command")
    adoCmd.ActiveConnection = adoCon
    adoCon.CursorLocation = adUseClientBatch
    adoCmd.CommandText = cmdtext
    Set adoRS = adoCmd.Execute
    
    'loop through the recordset and add each item to the listview
    Do While Not adoRS.EOF
        Set lstItem = lstTitles.ListItems.Add(, , adoRS("title"))
        lstItem.Tag = adoRS("id") 'used for updating and deleting
        lstItem.SubItems(1) = CStr(adoRS("codetype"))
        adoRS.MoveNext
    Loop
    lstTitles_Click 'reset the list
    'make sure to clean up after ourselves
    adoRS.Close
    Set adoRS = Nothing
    Set adoCmd = Nothing
    adoCon.Close
    Set adoCon = Nothing
    Exit Sub
    
errHandler:
    'just incase anything went wrong
    MsgBox "Error: " & Err.Description
    
End Sub

Private Sub Frame1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = ""
End Sub


Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = ""
End Sub

Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = ""
End Sub

Private Sub lstDelete_Click()
    
    On Error GoTo errHandler
    
    Dim RetVal As String
    Dim adoCon As Connection
    Dim adoCmd As Command
    
    'only 1 chance to say no!
    RetVal = MsgBox("Are you sure you want to delete this Language?  This will change all snippits that have this code language, to a < 0 > code language", vbYesNo, "Delete Code Language")
    If RetVal = vbNo Then
        Exit Sub
    End If
    
    'connect to the database and delete the code type, the reset the source entries
    Set adoCon = CreateObject("ADODB.Connection")
    adoCon.open gblConnectString
    Set adoCmd = CreateObject("ADODB.Command")
    adoCmd.ActiveConnection = adoCon
    adoCmd.CommandText = "DELETE FROM codetypes WHERE codetype='" & StuffQuotes(lstDelete.Text) & "'"
    adoCmd.Execute
    adoCmd.CommandText = "UPDATE source SET codetype='<blank>' WHERE codetype='" & StuffQuotes(lstDelete.Text) & "'"
    adoCmd.Execute
    
    'make sure we clean up after ourselves
    Set adoCmd = Nothing
    adoCon.Close
    Set adoCon = Nothing
    Form_Load  'reset everything
    Exit Sub
    
errHandler:
    'just incase anything went wrong
    MsgBox "Error: " & Err.Description

End Sub

Private Sub lstModify_Click()

    On Error GoTo errHandler
    
    Dim RetVal As String
    Dim adoCon As Connection
    Dim adoCmd As Command
    
    RetVal = InputBox("Please enter in the new title for the Code langauge", "Modify Code Language", CStr(lstModify.Text))
    If RetVal = "" Then
        Exit Sub
    End If
    Set adoCon = CreateObject("ADODB.Connection")
    adoCon.open gblConnectString
    Set adoCmd = CreateObject("ADODB.Command")
    adoCmd.ActiveConnection = adoCon
    adoCmd.CommandText = "UPDATE codetypes SET codetype='" & RetVal & "' WHERE codetype='" & StuffQuotes(lstModify.Text) & "'"
    adoCmd.Execute
    adoCmd.CommandText = "UPDATE source SET codetype='" & RetVal & "' WHERE codetype='" & StuffQuotes(lstModify.Text) & "'"
    adoCmd.Execute
    
    Set adoCmd = Nothing
    adoCon.Close
    Set adoCon = Nothing
    Form_Load
    Exit Sub
    
errHandler:
    'just incase anything went wrong
    MsgBox "Error: " & Err.Description
    
End Sub

Private Sub LoadCodeTypes()

    On Error GoTo errHandler
    
    Dim adoCon As Connection
    Dim adoCmd As Command
    Dim adoRS As Recordset
    
    cmbType.Clear
    cmbFilter.Clear
    cmbFilter.AddItem "No Filter", 0 'no filter isnt in the db, so add it here so its on top
    
    'connect to the database and retrieve the valid code types
    Set adoCon = CreateObject("ADODB.Connection")
    adoCon.open gblConnectString
    Set adoCmd = CreateObject("ADODB.Command")
    adoCmd.ActiveConnection = adoCon
    adoCmd.CommandText = "SELECT codetype FROM codetypes"
    Set adoRS = adoCmd.Execute
    
    'loop through the recordset and add them to the drop down
    Do While Not adoRS.EOF
        cmbType.AddItem CStr(adoRS("codetype"))
        cmbFilter.AddItem CStr(adoRS("codetype"))
        adoRS.MoveNext
    Loop
    
    'cleaning up the house
    adoRS.Close
    Set adoRS = Nothing
    Set adoCmd = Nothing
    adoCon.Close
    Set adoCon = Nothing

    'reset lists to top item
    cmbFilter.ListIndex = 0
    cmbType.ListIndex = 0
    Exit Sub
    
errHandler:
    'just incase anything went wrong
    MsgBox "Error: " & Err.Description
    
End Sub

Private Function VerifyCode() As Boolean

    VerifyCode = True
    If txtTitle.Text = "" Then
        MsgBox "Please enter a Title for the sourcecode snippit.Then click paste on the toolbarmenu.", vbInformation, "SourceCode Organizer"
        VerifyCode = False
        Exit Function
    End If
    If rtbCodeWindow.Text = "" Then
        MsgBox "You must enter some Sourcecode to save a sourcecode snippit.", vbInformation, "SourceCode Organizer"
        VerifyCode = False
        Exit Function
    End If

End Function

Private Sub lstTitles_Click()
    
    On Error GoTo errHandler
    
    Dim adoCon As Connection
    Dim adoCmd As Command
    Dim adoRS As Recordset
    Dim Index As Long
            
    If lstTitles.ListItems.Count < 1 Then 'if there is nothing in the list yet
        Exit Sub
    End If
    'connect to the database and retrieve the selected items details
    Set adoCon = CreateObject("ADODB.Connection")
    adoCon.open gblConnectString
    Set adoCmd = CreateObject("ADODB.Command")
    adoCmd.ActiveConnection = adoCon
    'here is one place where the tag comes in handy, selecting by title was
    'not the best idea, it would slow things down with large amount of snippits
    adoCmd.CommandText = "SELECT * FROM source WHERE id = " & lstTitles.SelectedItem.Tag
    Set adoRS = adoCmd.Execute
    
    'set up the code and notes windows, etc...
    rtbCodeWindow.Text = adoRS("code")
    txtTitle.Text = adoRS("title")
    rtbNotes.Text = adoRS("notes")
    'find the right code type in the drop down
    For Index = 0 To cmbType.ListCount
        If Trim(cmbType.List(Index)) = Trim(adoRS("codetype")) Then
            cmbType.ListIndex = Index
            Exit For
        End If
    Next Index
    
    'nope this aint a new piece of code
    gblNewCode = False
    
    'cleaning house
    adoRS.Close
    Set adoRS = Nothing
    Set adoCmd = Nothing
    adoCon.Close
    Set adoCon = Nothing
    Exit Sub
    
errHandler:
    'in case anything went wrong
    MsgBox "Error: " & Err.Description
 
End Sub

Private Sub lstTitles_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    'some code I foundon PSC to easily sort the listview, kudos to whover posted this
    If lstTitles.SortKey <> ColumnHeader.Index - 1 Then
        lstTitles.SortKey = ColumnHeader.Index - 1
        lstTitles.SortOrder = lvwAscending
    Else
        lstTitles.SortOrder = IIf(lstTitles.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    End If
    lstTitles.Sorted = True
    
End Sub

Private Sub lstTitles_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Source Code that are save in the DataBase..."
sBar.font.Size = 9
End Sub

Private Sub mnuAbout_Click()
  AboutF.Show
End Sub



Private Sub mnucontact_Click()
 Call email("extremedexter_z2001@yahoo.com")
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuFind_Click()
   SSTab1.Tab = 1
End Sub

Private Sub mnunet_Click()
  SSTab1.Tab = 1
End Sub

Private Sub mnuPaste_Click()
    rtbCodeWindow.SelText = Clipboard.GetText
    'Sets the Focus to rtfText
    rtbCodeWindow.SetFocus
End Sub

Private Sub mnuCopy_Click()
If tabSnippit.Tab = 0 Then
    Clipboard.Clear
    Clipboard.SetText rtbCodeWindow.Text
    End If
    If tabSnippit.Tab = 1 Then
    Clipboard.Clear
    Clipboard.SetText rtbnet.Text
    End If
End Sub

Private Sub mnuDelete_Click()

    On Error GoTo errHandler
    mnutest_Click
    Dim adoCon As Connection
    Dim adoCmd As Command
    
    'connnect to the database and delete the current selected snippit
    Set adoCon = CreateObject("ADODB.Connection")
    adoCon.open gblConnectString
    Set adoCmd = CreateObject("ADODB.Command")
    adoCmd.ActiveConnection = adoCon
    adoCmd.CommandText = "DELETE FROM source WHERE id=" & lstTitles.SelectedItem.Tag
    adoCmd.Execute
            
    'cleanup the house
    Set adoCmd = Nothing
    adoCon.Close
    Set adoCon = Nothing
    
    'reset the windows
    txtTitle.Text = ""
    rtbCodeWindow.Text = ""
    cmbType.ListIndex = 0
    LoadGridBox
    Exit Sub
    
errHandler:
    'in case anything went wrong
    MsgBox "Error: " & Err.Description
    
End Sub

Private Sub mnunew_Click()
    SSTab1.Tab = 1
    Picture1.Visible = True
    Picture3.Visible = False
    txtTitle.Visible = True
    txtTitle.Text = ""
    cmbType.ListIndex = 0
    rtbCodeWindow.Text = ""
    rtbNotes.TextRTF = ""
    gblNewCode = True
End Sub

Private Sub mnuOpen_Click()

    Dim RetVal As String
    cdoOpenDatabase.ShowOpen
    RetVal = cdoOpenDatabase.Filename
    If RetVal = "" Then
        Exit Sub
    End If
    gblConnectString = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & RetVal
    'this will get saved to the registry in a later version
    LoadGridBox  'reload the title list
    gblNewCode = True
    
End Sub

Private Sub mnuPlanetSourceCode_Click()
On Error Resume Next
   Dim xRet As Long
   xRet = ShellExecute(0, vbNullString, "http://people.we.mediaone.net/retxed/bugproblemicon.htm", vbNullString, App.Path, 1)
End Sub

Private Sub mnuSave_Click()

    On Error GoTo errHandler
    
    Dim adoCon As Connection
    Dim adoCmd As Command
    Dim adoRS As Recordset
    Dim RetVal As Boolean
    
    'connect to the database
    Set adoCon = CreateObject("ADODB.Connection")
    adoCon.open gblConnectString
    Set adoCmd = CreateObject("ADODB.Command")
    adoCmd.ActiveConnection = adoCon
    If gblNewCode = False Then      'if we are working on an existing snippit
        RetVal = VerifyCode 'check to make sure the user dotted all i's and crossed all t's
        If RetVal = False Then
            Exit Sub
        End If
        'this really should be a stored procedure, but....
        adoCmd.CommandText = "UPDATE source SET title='" & StuffQuotes(txtTitle) & "' WHERE id=" & lstTitles.SelectedItem.Tag
        adoCmd.Execute
        adoCmd.CommandText = "UPDATE source SET code='" & StuffQuotes(rtbCodeWindow.Text) & "' WHERE id=" & lstTitles.SelectedItem.Tag
        adoCmd.Execute
        adoCmd.CommandText = "UPDATE source SET codetype='" & StuffQuotes(cmbType.Text) & "' WHERE id=" & lstTitles.SelectedItem.Tag
        adoCmd.Execute
        adoCmd.CommandText = "UPDATE source SET [datetime]='" & Now & "' WHERE id=" & lstTitles.SelectedItem.Tag
        adoCmd.Execute
        adoCmd.CommandText = "UPDATE source SET notes='" & StuffQuotes(rtbNotes.Text) & "' WHERE id=" & lstTitles.SelectedItem.Tag
        adoCmd.Execute
    Else  'if its new
        RetVal = VerifyCode 'check to make sure the user dotted all i's and crossed all t's
        If RetVal = False Then
            Exit Sub
        End If
        adoCmd.CommandText = "INSERT INTO source ([datetime],title,codetype,code,notes) VALUES('" & Now & "', '" & StuffQuotes(txtTitle) & "', '" & StuffQuotes(cmbType.Text) & "', '" & StuffQuotes(rtbCodeWindow.Text) & "', '" & StuffQuotes(rtbNotes.TextRTF) & "')"
        adoCmd.Execute
        'we need the new identity created for it
        adoCmd.CommandText = "SELECT id FROM source WHERE title = '" & StuffQuotes(txtTitle) & "'"
        Set adoRS = adoCmd.Execute
        gblNewCode = False  'its no longer new
        adoRS.Close
        Set adoRS = Nothing
    End If
    
    'clean everything up
    Set adoCmd = Nothing
    adoCon.Close
    Set adoCon = Nothing
    
    LoadGridBox  'reset the list
    Exit Sub
    
errHandler:
    'in case anything went wrong
    MsgBox "Error: " & Err.Description

End Sub

Private Sub MnuWebSite_Click()
    
End Sub

Private Sub rtbNotes_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    'if anyone can tell me what the heck an effect of 3 is id appreciate it.
    'msdn documentation tells me there is no such thing
    If Data.GetFormat(vbCFText) Then 'if text
        If Effect = 3 Then
            rtbCodeWindow.Text = Data.GetData(vbCFText) 'set the window to the dragged in text
        Else
            rtbNotes.LoadFile Data.GetData(vbCFText), rtfText 'open the dragged in file
        End If
    End If
    If Data.GetFormat(vbCFFiles) Then 'if files from explorer
        rtbNotes.LoadFile Data.Files(1), rtfText  'open the file dragged from windows
    End If

End Sub


Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = ""
End Sub


Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Enter The Code Title and hit search..."
End Sub

Private Sub toolBar_ButtonClick(ByVal Button As MSComctlLib.Button)

    'no sense in recreating the wheel, so just call the menu item procedures
    If Button.Tag = "new" Then
        mnunew_Click
    End If
    If Button.Tag = "delete" Then
        mnuDelete_Click
    End If
    If Button.Tag = "save" Then
        mnuSave_Click
    End If
    If Button.Tag = "paste" Then
        mnuPaste_Click
    End If
    If Button.Tag = "copy" Then
        mnuCopy_Click
    End If
    If Button.Tag = "open" Then
        editme_Click
    End If
    If Button.Tag = "find" Then
      SSTab1.Tab = 1
    End If
    If Button.Tag = "print" Then
       ' mnuPrint_Click
    End If
      If Button.Tag = "mod" Then
      mnuModify_Click
    End If
     If Button.Tag = "front" Then
      On Error GoTo HandleErrors
Call Shell("C:\Program Files\Microsoft Office\Office\FRONTPG.EXE", vbNormalFocus)

 Exit Sub
HandleErrors:
  Dim intresponse As Integer
   Select Case Err.Number
        Case 53, 76
         intresponse = MsgBox("File not found.You may have installed it in a diffirent folder or it has been removed.", vbCritical, "Error")
         End Select
         End If
 If Button.Tag = "drw" Then
   On Error GoTo Err
Call Shell("C:\Program Files\Macromedia\Dreamweaver 4\Dreamweaver.exe", vbNormalFocus)
 Exit Sub
Err:
   MsgBox "File not found.You may have installed it in a diffirent folder or it has been removed.", vbCritical, "Error"
End If
 If Button.Tag = "VB" Then
 On Error GoTo Hell
Call Shell("C:\Program Files\Microsoft Visual Studio\VB98\VB6.exe", vbNormalFocus)

 Exit Sub
Hell:
   MsgBox "File not found.You may have installed it in a diffirent folder or it has been removed.", vbCritical, "Error"
  End If
End Sub

Private Sub rtbCodeWindow_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    'if anyone can tell me what the heck an effect of 3 is id appreciate it.
    'msdn documentation tells me there is no such thing
    If Data.GetFormat(vbCFText) Then  'if text
        If Effect = 3 Then
            rtbCodeWindow.Text = Data.GetData(vbCFText) 'set the window to the dragged in text
        Else
            rtbCodeWindow.LoadFile Data.GetData(vbCFText) 'open the dragged in file
        End If
    End If
    If Data.GetFormat(vbCFFiles) Then 'if files from explorer
        rtbCodeWindow.LoadFile Data.Files(1) 'open the file dragged from windows
    End If

End Sub

Private Sub Find(strSearch As String)
    
    On Error GoTo errHandler
    
    Dim adoCon As Connection
    Dim adoCmd As Command
    Dim adoRS As Recordset
    Dim cmdtext As String
    
    'not the best find, but it does the job
    'this can get very slow with large numbers of snippits
    cmdtext = "SELECT title FROM source WHERE title like '%" & strSearch & "%'"
    Set adoCon = CreateObject("ADODB.Connection")
    adoCon.open gblConnectString
    Set adoCmd = CreateObject("ADODB.Command")
    adoCmd.ActiveConnection = adoCon
    adoCmd.CommandText = cmdtext
    Set adoRS = adoCmd.Execute
    
    'only take the first returned result, ignore any others
    If adoRS.EOF = False Then
        'find the item in the listview and select it
        lstTitles.SelectedItem = lstTitles.FindItem(adoRS("title"), , lvwPartial)
        lstTitles_Click 'load it into the code window
    Else
        'nothing matched
        MsgBox "Not Found! No Record that match your keyword.Please try another keyword.", vbInformation, "Search"
    End If
    
    'clean up
    adoRS.Close
    Set adoRS = Nothing
    Set adoCmd = Nothing
    adoCon.Close
    Set adoCon = Nothing
    Exit Sub
    
errHandler:
    'in case anything went wrong
    MsgBox "Error: " & Err.Description

End Sub

Private Sub toolBar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Key
 Case "view"
   show1_Click
   Case "contents"
   SSTab1.Tab = 0
Case "DB"
  SSTab1.Tab = 1
  Picture1.Visible = False
  Case "SDB"
  SSTab1.Tab = 1
  Picture1.Visible = False
  Case "html"
   
  End Select
End Sub

Private Sub txtTitle_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Enter The Code Title..."
End Sub

Private Sub TxtUrl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Enter The Website URL..."
End Sub

Private Sub txWhatSearch_DropDown()
sBar.Panels(1).Text = "Source Code Language Selection.."
End Sub

Private Sub Web_TitleChange(ByVal Text As String)
    Me.Caption = "Browsing : " & Text
End Sub
