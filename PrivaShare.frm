VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "PrivaShare ver. 1.1b"
   ClientHeight    =   5325
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7350
   Icon            =   "PrivaShare.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   1164
      ButtonWidth     =   1508
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Connect"
            Key             =   "connect"
            ImageKey        =   "connect"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            Key             =   "search"
            ImageKey        =   "search"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Share"
            Key             =   "share"
            ImageKey        =   "share"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "No Share"
            Key             =   "noShare"
            ImageKey        =   "noShare"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Welcome"
            Key             =   "welcome"
            ImageKey        =   "welcome"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   4950
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Object.Width           =   4586
            MinWidth        =   4586
            Text            =   "PrivaShare (c) 2001 Gene Hamilton"
            TextSave        =   "PrivaShare (c) 2001 Gene Hamilton"
            Object.ToolTipText     =   "That's me!"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Object.ToolTipText     =   "Number of current connections open."
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "07/05/2021"
            Object.ToolTipText     =   "Todays date."
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Log:"
            TextSave        =   "Log:"
            Object.ToolTipText     =   "Logging Enabled/Disabled."
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   7435
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   12632256
      ForeColor       =   12582912
      TabCaption(0)   =   "Connections/Chat"
      TabPicture(0)   =   "PrivaShare.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "MMControl1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtSend"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdSend"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdDrop"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "tvwConnects"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtOutput"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdAddFavorites"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdSendSound"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Download"
      TabPicture(1)   =   "PrivaShare.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(2)=   "cmdRequestFile"
      Tab(1).Control(3)=   "Command1"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Share/Upload"
      TabPicture(2)   =   "PrivaShare.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(1)=   "cmdUpload"
      Tab(2).Control(2)=   "Frame5"
      Tab(2).Control(3)=   "Frame6"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Settings/Log"
      TabPicture(3)   =   "PrivaShare.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label6"
      Tab(3).Control(1)=   "Label2"
      Tab(3).Control(2)=   "Label4"
      Tab(3).Control(3)=   "txtName"
      Tab(3).Control(4)=   "Frame2"
      Tab(3).Control(5)=   "txtLocalIP"
      Tab(3).Control(6)=   "txtNetIP"
      Tab(3).Control(7)=   "Frame7"
      Tab(3).Control(8)=   "Frame8"
      Tab(3).ControlCount=   9
      Begin VB.CommandButton cmdSendSound 
         Caption         =   "Send .wav Recording to Selected"
         Height          =   255
         Left            =   4320
         TabIndex        =   52
         Top             =   3840
         Width           =   2655
      End
      Begin VB.CommandButton cmdAddFavorites 
         Caption         =   "&Add  Selected to Favorites"
         Height          =   495
         Left            =   5280
         TabIndex        =   50
         ToolTipText     =   "You can add selected connection above to you favorites list."
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Frame Frame8 
         Caption         =   "Preferences"
         Height          =   1215
         Left            =   -70680
         TabIndex        =   47
         Top             =   2760
         Width           =   2775
         Begin VB.CommandButton cmdLoadPreferences 
            Caption         =   "Load Preferences"
            Height          =   255
            Left            =   360
            TabIndex        =   49
            Top             =   360
            Width           =   2055
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save Current Preferences"
            Height          =   255
            Left            =   360
            TabIndex        =   48
            Top             =   720
            Width           =   2055
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Logging options"
         Height          =   1575
         Left            =   -70680
         TabIndex        =   43
         Top             =   1080
         Width           =   2775
         Begin VB.CommandButton Command6 
            Caption         =   "Start Logging"
            Height          =   255
            Left            =   360
            TabIndex        =   46
            Top             =   360
            Width           =   2055
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Stop Logging"
            Height          =   255
            Left            =   360
            TabIndex        =   45
            Top             =   720
            Width           =   2055
         End
         Begin VB.CommandButton cmdSaveLog 
            Caption         =   "Save && Clear Log"
            Height          =   255
            Left            =   360
            TabIndex        =   44
            Top             =   1200
            Width           =   2055
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "File Uploading"
         Height          =   975
         Left            =   -70080
         TabIndex        =   41
         Top             =   1800
         Width           =   1935
         Begin VB.CheckBox ChkUpload 
            Caption         =   "Permit uploads to this directory?"
            Height          =   375
            Left            =   120
            TabIndex        =   42
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "File Sharing"
         Height          =   1095
         Left            =   -70080
         TabIndex        =   38
         Top             =   480
         Width           =   1935
         Begin VB.OptionButton optShare 
            Caption         =   "Share files"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   40
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton optShare 
            Caption         =   "No file sharing"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   39
            Top             =   720
            Width           =   1575
         End
      End
      Begin VB.TextBox txtNetIP 
         Height          =   285
         Left            =   -71160
         TabIndex        =   36
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Show Hosts Files"
         Height          =   255
         Left            =   -70080
         TabIndex        =   35
         ToolTipText     =   "Look into connections file sharing directory."
         Top             =   600
         Width           =   2175
      End
      Begin VB.CommandButton cmdRequestFile 
         Caption         =   "Request Selected File"
         Height          =   255
         Left            =   -70080
         TabIndex        =   34
         ToolTipText     =   "Select file above, and press this button to download."
         Top             =   3600
         Width           =   2175
      End
      Begin VB.Frame Frame3 
         Caption         =   "Host download folder"
         Height          =   2535
         Left            =   -70200
         TabIndex        =   32
         Top             =   960
         Width           =   2415
         Begin VB.ListBox lstFiles 
            Height          =   2205
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.CommandButton cmdUpload 
         Caption         =   "Upload Selected"
         Height          =   495
         Left            =   -70080
         TabIndex        =   31
         ToolTipText     =   "If connection allows uploading, you can send them a file."
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Frame Frame4 
         Caption         =   "Folder to share Files from"
         Height          =   3375
         Left            =   -74880
         TabIndex        =   23
         Top             =   480
         Width           =   4575
         Begin VB.DriveListBox Drive2 
            Height          =   315
            Left            =   120
            TabIndex        =   26
            Top             =   2880
            Width           =   2055
         End
         Begin VB.DirListBox Dir2 
            Height          =   2340
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Width           =   2055
         End
         Begin VB.FileListBox File3 
            DragIcon        =   "PrivaShare.frx":04B2
            Height          =   2820
            Left            =   2280
            System          =   -1  'True
            TabIndex        =   24
            Top             =   360
            Width           =   2175
         End
      End
      Begin RichTextLib.RichTextBox txtOutput 
         Height          =   2175
         Left            =   240
         TabIndex        =   22
         Top             =   600
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   3836
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"PrivaShare.frx":08F4
      End
      Begin VB.TextBox txtLocalIP 
         Height          =   285
         Left            =   -72840
         TabIndex        =   19
         Top             =   720
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         Caption         =   "Logging"
         Height          =   2895
         Left            =   -74880
         TabIndex        =   14
         Top             =   1080
         Width           =   3975
         Begin VB.ListBox lstLogging 
            Height          =   2010
            Left            =   120
            TabIndex        =   18
            Top             =   720
            Width           =   3735
         End
         Begin VB.CheckBox ChkTransfers 
            Caption         =   "Log Transfers"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   1335
         End
         Begin VB.CheckBox ChkTime 
            Caption         =   "Log Time"
            Height          =   255
            Left            =   2760
            TabIndex        =   16
            Top             =   360
            Width           =   975
         End
         Begin VB.CheckBox ChkIPs 
            Caption         =   "Log Name/IP"
            Height          =   255
            Left            =   1440
            TabIndex        =   15
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   -74760
         TabIndex        =   13
         Text            =   "Newbie"
         Top             =   720
         Width           =   1695
      End
      Begin MSComctlLib.TreeView tvwConnects 
         Height          =   2175
         Left            =   3720
         TabIndex        =   12
         ToolTipText     =   "These are your connections."
         Top             =   600
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   3836
         _Version        =   393217
         HideSelection   =   0   'False
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         ImageList       =   "ImageList1"
         Appearance      =   1
      End
      Begin VB.CommandButton cmdDrop 
         Caption         =   "Drop Selected"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3720
         TabIndex        =   10
         ToolTipText     =   "This will drop connection to who's selected."
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         Caption         =   "Folder to Download Files into"
         Height          =   3375
         Left            =   -74880
         TabIndex        =   3
         Top             =   480
         Width           =   4575
         Begin VB.FileListBox File1 
            DragIcon        =   "PrivaShare.frx":0976
            Height          =   2820
            Left            =   2280
            System          =   -1  'True
            TabIndex        =   6
            Top             =   360
            Width           =   2175
         End
         Begin VB.DirListBox Dir1 
            Height          =   2340
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   2055
         End
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   120
            TabIndex        =   4
            Top             =   2880
            Width           =   2055
         End
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         Height          =   255
         Left            =   1080
         TabIndex        =   2
         ToolTipText     =   "Press to send the message you typ in above. Or just hit return."
         Top             =   3720
         Width           =   1695
      End
      Begin VB.TextBox txtSend 
         Height          =   525
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   3120
         Width           =   3375
      End
      Begin MCI.MMControl MMControl1 
         Height          =   375
         Left            =   6360
         TabIndex        =   51
         ToolTipText     =   "Press the circle to record 3 seconds of sound.  Play to hear it."
         Top             =   3360
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   661
         _Version        =   393216
         PlayEnabled     =   -1  'True
         RecordEnabled   =   -1  'True
         PrevVisible     =   0   'False
         NextVisible     =   0   'False
         PauseVisible    =   0   'False
         BackVisible     =   0   'False
         StepVisible     =   0   'False
         StopVisible     =   0   'False
         EjectVisible    =   0   'False
         DeviceType      =   ""
         FileName        =   ""
      End
      Begin VB.Label Label7 
         Caption         =   "Use your mic input on sound card to send 3 second wav file."
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   4080
         TabIndex        =   53
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Internet IP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71160
         TabIndex        =   37
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Local IP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72840
         TabIndex        =   21
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Your Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   20
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Enter your text here:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Chat window"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Currently connected to"
         Height          =   255
         Left            =   3720
         TabIndex        =   7
         Top             =   360
         Width           =   1815
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   0
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6480
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrivaShare.frx":0DB8
            Key             =   "shake"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrivaShare.frx":120C
            Key             =   "help1"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrivaShare.frx":1660
            Key             =   "help2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrivaShare.frx":1AB4
            Key             =   "view"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrivaShare.frx":1F08
            Key             =   "drop"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrivaShare.frx":235C
            Key             =   "lock"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrivaShare.frx":27B0
            Key             =   "secure"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrivaShare.frx":2C04
            Key             =   "fileClosed"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrivaShare.frx":3058
            Key             =   "fileOpen"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrivaShare.frx":34AC
            Key             =   "search"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrivaShare.frx":39F0
            Key             =   "dropAll"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrivaShare.frx":3E44
            Key             =   "connect"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrivaShare.frx":4298
            Key             =   "welcome"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrivaShare.frx":46EC
            Key             =   "x"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrivaShare.frx":4800
            Key             =   "share"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrivaShare.frx":4C54
            Key             =   "noShare"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   600
      Left            =   5880
      Top             =   1080
   End
   Begin VB.ListBox lstNodes 
      Height          =   255
      Left            =   4560
      TabIndex        =   28
      Top             =   1080
      Width           =   495
   End
   Begin VB.ListBox lstNodes_nodes 
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   29
      Top             =   1440
      Width           =   495
   End
   Begin VB.FileListBox File2 
      Height          =   480
      Left            =   3960
      TabIndex        =   30
      Top             =   1200
      Width           =   375
   End
   Begin VB.Timer Timer2 
      Interval        =   3000
      Left            =   2760
      Top             =   1080
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuPort 
         Caption         =   "Port"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuSearchNodes 
         Caption         =   "Search for file through your connections."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About PrivaShare"
      End
   End
   Begin VB.Menu mnuNodes 
      Caption         =   "Join"
      Visible         =   0   'False
      Begin VB.Menu JoinNode 
         Caption         =   "Join to this node"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddFavorites_Click()

On Error GoTo noOneSelected4

Dim i As Integer

For i = 0 To intNum_Connections - 1
        getSendersInfo i
        If tvwConnects.SelectedItem = strName Then
            'Put check routine here to see if already in list.
            '-------------------------------------------------
            If MsgBox("Add " & strName & " to favorite?", vbYesNo) = vbYes Then
                'Add to listboxes
                frmConnect.lstFavName.AddItem strName
                frmConnect.lstFavIP.AddItem strIP
                frmConnect.lstConnections.AddItem strName & vbTab & strIP
                saveFavorites
            End If
        End If
Next i
Exit Sub

noOneSelected4:
'No name was selected in tree view.  Siwwy wabbits.
MsgBox ("You must first select who you want to add.")
    
End Sub

Private Sub cmdDrop_Click()
    Dim i As Integer
    
On Error GoTo noOneSelected3
    
    'look through lstNodes starting at 0.
    For i = 0 To intNum_Connections - 1
        getSendersInfo i
        If tvwConnects.SelectedItem = strName Then
            If MsgBox("Disconnect from " & strName & "?", vbYesNo) = vbYes Then
                tvwConnects.Nodes.Remove strName & strIP
                Winsock1(i + 1).Close
                lstNodes.RemoveItem i * 4
                lstNodes.AddItem "", i * 4
                intNum_ConnectionsNow = intNum_ConnectionsNow - 1 'Decrease number of connections.
                'Connection was dropped.  That wasn't nice...
                
                'If no connections, disable buttons
                If intNum_ConnectionsNow = 0 Then
                    cmdSend.Enabled = False
                    cmdDrop.Enabled = False
                    txtName.Enabled = True
                End If
                
                txtOutput.Text = txtOutput.Text + vbCrLf + strName & " was dropped."
                txtOutput.SelStart = Len(txtOutput.Text)
                txtSend.SetFocus
                'Update log, IPs.
                If ChkIPs And blnLog Then
                    lstLogging.AddItem strName & " was dropped."
                End If
                'Update log, Date and time.
                If ChkTime And blnLog Then
                    lstLogging.AddItem Time
                End If
 
                Exit For
            End If
        End If
    Next i
    
    'Clear their connections.
    lstNodes_nodes(i + 1).Clear
    
    sendConnectionsToAll i + 1 'Update node list of all connections.
    
    Exit Sub
    
noOneSelected3:
'No name was selected in tree view.  Siwwy wabbits.
MsgBox ("You must first select who you want to drop.")

End Sub

Private Sub cmdGetConnections_Click()

On Error GoTo noOneSelected2
    Dim i As Integer

    For i = 0 To intNum_Connections - 1
        getSendersInfo i
        If tvwConnects.SelectedItem = strName Then
            i = Val(strIndex)
            Exit For
        End If
    Next i
    'send request, and null takes up space, not used.
    sendToOne i, "requestContacts,null"
    Exit Sub

noOneSelected2:
'No name was selected in tree view.  Siwwy wabbits.
MsgBox ("You must first select who you want to get directory from in Connections/Chat window.")

End Sub

Private Sub cmdRequestFile_Click()

Dim i As Integer

On Error GoTo fileError

    For i = 0 To intNum_Connections - 1
        getSendersInfo i
        If tvwConnects.SelectedItem = strName Then
            i = Val(strIndex)
            Exit For
        End If
    Next i

    'send request, and filename.
    sendToOne i, "requestFile," & lstFiles.Text
    'If logging file downloads, log it.
    If ChkTransfers And blnLog Then
        lstLogging.AddItem "File " & lstFiles.Text & "requested..."
    End If
    Exit Sub
    
fileError:
    MsgBox ("There was an error while trying to download file.")
    
End Sub

Private Sub cmdSaveLog_Click()

Dim strLog
Dim strMonth As String
Dim strDay As String
Dim strYear As String
Dim intCounter As Integer
Dim i As Integer
intCounter = 1

'Get current date.
strLog = Date

'Get the day and year for name of log file.
strMonth = Mid(strLog, 1, InStr(1, strLog, "/") - 1)
strLog = Mid(strLog, InStr(1, strLog, "/") + 1, Len(strLog))
strDay = Mid(strLog, 1, InStr(1, strLog, "/") - 1)
strYear = Mid(strLog, InStr(1, strLog, "/") + 1, Len(strLog))

'Build filename string.
strLog = "Log_" & strMonth & "_" & strDay & "_" & strYear & "."

On Error GoTo writeLog

    'See if this log already exsists.
    For i = 1 To 50
        Open appPath & strLog & intCounter For Input As #1
        Close #1
        intCounter = intCounter + 1
    Next i
    
writeLog:

    'Open the log file, all lines of listbox.
    Open appPath & strLog & intCounter For Output As #1
    For i = 0 To lstLogging.ListCount - 1
        lstLogging.ListIndex = i
        Write #1, lstLogging.Text
    Next i
    Close #1
    lstLogging.Clear


End Sub

Private Sub cmdSend_Click()

    'If send button pressed, send message to all then output textbox.
    sendToEveryone
    txtOutput.Text = txtOutput.Text + vbCrLf + "Me: " + txtSend.Text
    txtOutput.SelStart = Len(txtOutput.Text)
    txtSend.Text = ""
    txtSend.SetFocus
    
End Sub

Private Sub cmdSave_Click()

    Open appPath & "preferences.cfg" For Output As #1
        Write #1, txtName.Text
        Write #1, intPort
        Write #1, ChkTransfers
        Write #1, ChkIPs
        Write #1, ChkTime
        Write #1, optShare(0).Value
        Write #1, optShare(1).Value
        Write #1, ChkUpload.Value
        Write #1, frmWelcome.txtWelcome.Text
        Write #1, Dir1.Path
        Write #1, Dir2.Path
    Close #1
    
End Sub

Private Sub cmdLoadPreferences_Click()

On Error GoTo savePreferences

    Dim strTemp As String
    Open appPath & "preferences.cfg" For Input As #1
    
        'Load name.
        Input #1, strTemp
        txtName.Text = strTemp
        
        'Load port.
        Input #1, strTemp
        intPort = strTemp
                
        'Log File transfers?
        Input #1, strTemp
        ChkTransfers.Value = strTemp
        
        'Save Name/IP in log?
        Input #1, strTemp
        ChkIPs.Value = strTemp
        
        'Save time in log?
        Input #1, strTemp
        ChkTime.Value = strTemp
        
        'Share files?
        Input #1, strTemp
        optShare(0).Value = strTemp
        Input #1, strTemp
        optShare(1).Value = strTemp
        
        'Permit uploads?
        Input #1, strTemp
        ChkUpload.Value = strTemp
        
        'Get Welcome string.
        Input #1, strWelcome
        frmWelcome.txtWelcome.Text = strWelcome
        
        'Set the download folder.
        Input #1, strTemp
        Dir1.Path = strTemp
        Dir1.Refresh
        'Set the share folder.
        Input #1, strTemp
        Dir2.Path = strTemp
        Dir2.Refresh
        
    Close #1
    Exit Sub
    
savePreferences:
    cmdSave_Click

End Sub

Private Sub loadFavorites()

On Error GoTo saveFavoritesError

    
    Open appPath & "favorites.cfg" For Input As #1
    
    Do While Not EOF(1)
    
        'Load name.
        Input #1, strName
        frmConnect.lstFavName.AddItem strName
        'Load IP.
        Input #1, strIP
        frmConnect.lstFavIP.AddItem strIP
        frmConnect.lstConnections.AddItem strName & vbTab & strIP
        
    Loop
    
    Close #1
    Exit Sub
    
saveFavoritesError:
    'No favorites have been saved yet.
    saveFavorites
    
End Sub

Private Sub cmdSendSound_Click()
    
On Error GoTo noOneSelected

Dim i As Integer

    For i = 0 To intNum_Connections - 1
        getSendersInfo i
        If tvwConnects.SelectedItem = strName & strIP Then
            i = Val(strIndex)
            Exit For
        End If
    Next i
    
    'Save the .wav file and send it.
    MMControl1.Command = "Save"
    MMControl1.Command = "Close"
    'Send the sound.
    cmdSendSound.Enabled = False
    setupSendFile i, "PS_SoundFile.wav"
    Exit Sub
    
noOneSelected:
'No name was selected in tree view.  Siwwy wabbits.
MsgBox ("You must first select who you want to send sound to from Connections/Chat window.")
cmdSendSound.Enabled = True
MMControl1.Command = "Open"

End Sub

Private Sub cmdUpload_Click()
    
On Error GoTo noOneSelected2
    Dim i As Integer

    For i = 0 To intNum_Connections - 1
        getSendersInfo i
        If tvwConnects.SelectedItem = strName & strIP Then
            i = Val(strIndex)
            Exit For
        End If
    Next i
    'send request, and null takes up space, not used.
    sendToOne i, "uploadingFile," & File3.FileName
    Exit Sub

noOneSelected2:
'No name was selected in tree view.  Siwwy wabbits.
MsgBox ("You must first select who to upload to inConnections/Chat window.")


End Sub

Private Sub Command1_Click()
    
lstFiles.Clear

On Error GoTo noOneSelected
    Dim i As Integer

    For i = 0 To intNum_Connections - 1
        getSendersInfo i
        If tvwConnects.SelectedItem = strName Then
            i = Val(strIndex)
            Exit For
        End If
    Next i
    'send request, and null takes up space, not used.
    sendToOne i, "requestDir,null"
    Exit Sub

noOneSelected:
'No name was selected in tree view.  Siwwy wabbits.
MsgBox ("You must first select who you want to get directory from in Connections/Chat window.")

End Sub


Private Sub Command2_Click()
    'Close file when ready to send.
     MMControl1.Command = "Save"
    MMControl1.Command = "Close"
    
    
End Sub

Private Sub Command5_Click()

    'Stop sending text to logging listbox.
    blnLog = False
    lstLogging.AddItem "** Logging stopped: " & Date & " " & Time
    
    'Show log: off in status bar.
    StatusBar1.Panels(4).Text = "Log: Off"


End Sub

Private Sub Command6_Click()

    'Start sending text to logging listbox.
    blnLog = True
    lstLogging.AddItem "** Logging started: " & Date & " " & Time
    
    'Show log: on in status bar.
    StatusBar1.Panels(4).Text = "Log: On"
    
    'Enable save log button.
    cmdSaveLog.Enabled = True
 
End Sub

Private Sub Dir1_Change()

    'Chage directory looking to selected dir.
     File1.Path = Dir1.Path
     
End Sub

Private Sub Dir2_Change()
    'Change visible filelistbox used in requests.
    File3.Path = Dir2.Path
    'change hidden filelistbox used in requests.
    File2.Path = Dir2.Path
    
End Sub

Private Sub Drive1_Change()

    'Change directory window to match drive.
    Dir1.Path = Drive1.Drive
    
End Sub

Private Sub Drive2_Change()
    Dir2.Path = Drive2.Drive
End Sub


Public Sub Form_Load()
    
    'Turn off timer.
    Timer2.Enabled = False
    
    'Get the applications path.
    appPath = App.Path
    If Right$(appPath, 1) <> "\" Then appPath = appPath & "\"
    
    'Set up sound recording.
    MMControl1.DeviceType = "WaveAudio"
    MMControl1.FileName = appPath & "PS_SoundFile.wav"
    MMControl1.Command = "Open"
    'When file is downloaded blnWav is set to true if it was a sound to play.
    'Just in case... Works without it on win98.
    blnWav = False
    
    'Put local IP in text box
    txtLocalIP.Text = Winsock1(0).LocalIP
    
    'Load preferences.cfg if present.
    cmdLoadPreferences_Click
    
    'Load Favorite connections if present.
    loadFavorites
    
    'IF first time used, set port to default.
    If intPort = 0 Then
        intPort = 2001
    End If
    
    'Setup linening connection zero in winsock array.
    Winsock1(0).Close
    Winsock1(0).LocalPort = intPort
    Winsock1(0).Listen
    
    'turn off send button until conection established.
    cmdSend.Enabled = False
    
    'Call the statusbar update sub
    printConnections
    
    'Turn off timer until needed.
    Timer1.Enabled = False
    
    'Show log: off in status bar.
    StatusBar1.Panels(4).Text = "Log: Off"
    
    'Turn off save log button until log is started.
    cmdSaveLog.Enabled = False
    
End Sub

Public Sub cmdConnect_Click(connectionIP As String)

On Error GoTo errorhandler

    Dim intArrayNumber As Integer   'The array number to use.
    Dim i As Integer
    Dim blnConnected As Boolean
    
    'Check for errors in IP address
    If connectionIP = "" Then Exit Sub
    
    'For testing!
    alreadyConnected blnConnected, connectionIP
    
    If blnConnected = True Then
        MsgBox ("You are already connected to this node.")
        txtSend.SetFocus
    Else
    
        'Show connection text in output textbox.
        txtOutput.Text = txtOutput.Text + vbCrLf + "Connecting to IP " & connectionIP & "."
        txtOutput.SelStart = Len(txtOutput.Text)
        '-----------------------

        'Search if there's a used available Winsock control.
        For i = 0 To intNum_Connections
            'Is there an loaded unused index?
            If Winsock1(i).State = sckClosed Then
        
            intArrayNumber = i ' use a used closed spot.
                Exit For
            End If
        Next i 'Looking for used open number in array.
    
        'If none was found, create a new one.
        If intArrayNumber = 0 Then
    
            'Increment number of connections.
            intNum_Connections = intNum_Connections + 1
        
            'Load a new Winsock control for this connection. Only load new one after 2.
            Load Winsock1(intNum_Connections)
        
            'Make listbox array index to use for connection info.
            Load lstNodes_nodes(intNum_Connections)
        
            'Set the winsock index to new array number.
            intArrayNumber = intNum_Connections
        
            'Put in 4 blank spots in listbox lstNodes
            lstNodes.AddItem ""
            lstNodes.AddItem ""
            lstNodes.AddItem ""
            lstNodes.AddItem ""
        
        End If
  
        'connect if you can ---------------------
        Winsock1(intArrayNumber).Close
        Winsock1(intArrayNumber).LocalPort = 0
        Winsock1(intArrayNumber).Connect connectionIP, intPort
        'Increase number of current connections for statusbar.
        intNum_ConnectionsNow = intNum_ConnectionsNow + 1
    
        'Call the statusbar update sub
        printConnections
    
        'Turn on timer and set the connection integer to send myInfo to when .2 seconds have passed.
        intChannel = intArrayNumber
        Timer1.Enabled = True
           
        'Move to tab 1 to see connection.
        SSTab1.Tab = 0
    
    End If 'If blnConnected
    
Exit Sub

errorhandler:

    'Error connecting.
    txtOutput.Text = txtOutput.Text + vbCrLf + "Failed to connect."
    txtOutput.SelStart = Len(txtOutput.Text)
    Winsock1(1).Close

End Sub

Private Sub Form_Resize()

On Error GoTo noChange

    Me.Height = 6045
    Me.Width = 7470
    Exit Sub
    
noChange:

End Sub


Private Sub Form_Unload(Cancel As Integer)

    'Save and close wav file.
    MMControl1.Command = "Save"
    MMControl1.Command = "Close"
    
    'Unload all forms.
    Unload frmDefaultPort
    Unload frmWelcome
    
End Sub

Private Sub JoinNode_Click()

On Error GoTo alreadyConnected

Dim i As Integer
Dim Index As Integer
Dim strNodeName As String
Dim strNodeIP As String

    For i = 0 To intNum_Connections - 1
        getSendersInfo i
        If tvwConnects.SelectedItem.Parent = strName Then
            Index = Val(strIndex)
            Exit For
        End If
    Next i
   
    For i = 0 To lstNodes_nodes(Index).ListCount / 3 - 1 'ListCount is 1 based, so remove 1.
        'Get node name.
        lstNodes_nodes(Index).ListIndex = i * 3
        strNodeName = lstNodes_nodes(Index).Text
        
        If strNodeName = tvwConnects.SelectedItem Then
            'Get IP.
            lstNodes_nodes(Index).ListIndex = i * 3 + 1
            strNodeIP = lstNodes_nodes(Index).Text
            Exit For
        End If
    Next i
    
    If strNodeName <> "" Then
        If MsgBox("Connect to " & strNodeName & "?", vbYesNo) = vbYes Then
            'Try to connect to selected node.
            'frmConnect.txtConnection.Text = strNodeIP
            cmdConnect_Click strNodeIP
        
        End If ' yes/no
    End If ' node <> ""
    
    Exit Sub
    
alreadyConnected:
MsgBox ("You must choose a subnode to connect to.")

End Sub

Private Sub MMControl1_Done(NotifyCode As Integer)

    'When played. Go back to start.
    MMControl1.Command = "Prev"
    
End Sub


Private Sub MMControl1_RecordClick(Cancel As Integer)
    'IF play button is hit, set timer.
    'Timer2 is set for 3 sec.
    Timer2.Enabled = True
End Sub

Private Sub mnuAbout_Click()

    'Show about form.
    frmAbout.Show
    
End Sub

Private Sub mnuExit_Click()
Dim i As Integer

    'Close all open ports used.
    For i = 0 To intNum_Connections
        Winsock1(i).Close
    Next i
    
    'Unload forms.
    Unload frmConnect
    Unload frmDefaultPort
    Unload frmWelcome
    Unload Me
    End
    
End Sub


Private Sub mnuPort_Click()
    
    frmDefaultPort.Show
    
End Sub


Private Sub mnuSearchNodes_Click()

    frmConnect.Show
    frmConnect.txtSearch.SetFocus
    
End Sub

Private Sub Timer1_Timer()

    'If State <> 7 then failed to connect.
    If Winsock1(intChannel).State <> 7 Then
        intSelText = Len(txtOutput.Text)
        txtOutput.Text = txtOutput.Text + vbCrLf + "Connection " & intChannel & " Failed."
        'Failed, so decrease # of connections.
        intNum_ConnectionsNow = intNum_ConnectionsNow - 1
        'Call the statusbar update sub
        printConnections
        'Select new text for color change.
        txtOutput.SelStart = intSelText
        txtOutput.SelLength = Len(txtOutput.Text)
        txtOutput.SelColor = vbRed
        'Set select to end of text when done.
        txtOutput.SelStart = Len(txtOutput.Text)
        txtSend.SetFocus
        Winsock1(intChannel).Close
        Timer1.Enabled = False
        
    'Connection was successfull
    Else

        'SendInfo to selected connection(channel) when .2 seconds has passed on timer.
        sendMyInfo intChannel
        
        'Turn on Send button
        cmdSend.Enabled = True
        cmdDrop.Enabled = True
        txtName.Enabled = False
        
        'Turn timer off until needed again.
        Timer1.Enabled = False
        
    'Safe to turn on Send button.
    txtSend.SetFocus
    End If

End Sub

Private Sub Timer2_Timer()
    'When 3 seconds of recording is up. Save change.
    MMControl1.Command = "Save"
    MMControl1.Command = "Close"
    MMControl1.Command = "Open"
    Timer2.Enabled = False
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.key
        Case "connect"
            frmConnect.Show
            frmConnect.txtConnection.SetFocus
        
        Case "search"
            frmConnect.Show
            frmConnect.txtSearch.SetFocus
            
        Case "welcome"
            frmWelcome.Show
            
        Case "share"
            optShare(0) = True
                        
        Case "noShare"
            optShare(1) = True
                    
    End Select
End Sub


Private Sub Toolbar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'Display the menu when right mouse button is pressed
    If Button = vbRightButton Then
        PopupMenu mnuNodes
    End If
    
End Sub


Private Sub tvwConnects_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'Display the menu when right mouse button is pressed
    If Button = vbRightButton Then
        PopupMenu mnuNodes
    End If

End Sub

Private Sub txtName_Change()
    If InStr(1, txtName, ",") Then
        MsgBox ("Your name can not include a comma.")
        txtName = Mid(txtName, 1, InStr(1, txtName, ",") - 1)
        
    End If
        
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
    
    'If Enter key pressed, send text, clear text box.
    If KeyAscii = 13 Then
        cmdSend_Click
        txtSend.Text = ""
    End If
    
End Sub

Private Sub Winsock1_Close(Index As Integer)

    'Get information for disconnected node starting at 0
    getSendersInfo Index - 1
    
    'Connection was broke by other computer.
    txtOutput.Text = txtOutput.Text + vbCrLf + strName & " Disconnected."
    txtOutput.SelStart = Len(txtOutput.Text)
    txtSend.SetFocus
        
    'Close the connection just to be safe. I guess....
    Winsock1(Index).Close
    
On Error GoTo portScan

    'Remove node from treeview.
    tvwConnects.Nodes.Remove strName & strIP

    'Clear their connections.
    lstNodes_nodes(Index).Clear
    
    'Update log, IPs.
    If ChkIPs And blnLog Then
        lstLogging.AddItem strName & " disconnected."
    End If
    'Update log, Date and time.
    If ChkTime And blnLog Then
        lstLogging.AddItem Time
    End If
    
    txtSend.SetFocus

    'Connection has left, Open spot in array.
    lstNodes.RemoveItem (Index - 1) * 4
    lstNodes.AddItem "", (Index - 1) * 4

    'Update current connections.
    intNum_ConnectionsNow = intNum_ConnectionsNow - 1
    printConnections
    
    'If no connections, disable buttons
    If intNum_ConnectionsNow = 0 Then
        cmdSend.Enabled = False
        cmdDrop.Enabled = False
        txtName.Enabled = True
    End If
    
    SSTab1.Tab = 0 'Put view tab to connections to see that someone left.
       
    sendConnectionsToAll Index 'Update node list of all connections.
    
    Exit Sub
    
portScan:
    'Connection was broke by other computer.
    txtOutput.Text = txtOutput.Text + vbCrLf + "Someone might be scanning your connection port."
    txtOutput.SelStart = Len(txtOutput.Text)
    txtSend.SetFocus
    
End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
'A connection was requested from the server.

Dim i As Integer
Dim intArrayNumber As Integer

'only 0 in winsock array allowed for connecting.
If Index = 0 Then

'Search if there's a used available Winsock control.
For i = 0 To intNum_Connections
    If Winsock1(i).State = sckClosed Then
        intArrayNumber = i
        Exit For
    End If
Next i
    
    'If none was found, create a new one.
    If intArrayNumber = 0 Then
    
        'Increment number of connections.
        intNum_Connections = intNum_Connections + 1
        
        'Load a new Winsock control for this connection. Only load new one after 2.
        Load Winsock1(intNum_Connections)
        
        'Set the winsock index to new array number.
        intArrayNumber = intNum_Connections
        
        'Make listbox array index to use for connection info.
        Load lstNodes_nodes(intNum_Connections)
        
        'Add 4 blank spaces for new node connection.
        lstNodes.AddItem ""
        lstNodes.AddItem ""
        lstNodes.AddItem ""
        lstNodes.AddItem ""
        
    End If
    
    'Let system assign an open port to array spot.
    Winsock1(intArrayNumber).LocalPort = 0
    
    'Then accept connection on that port.
    Winsock1(intArrayNumber).Accept requestID
    
    'Enable the Send button, so you can talk back.
    cmdSend.Enabled = True
    
    'Post connection in window and set focus to send textbox.
    txtOutput.Text = txtOutput.Text + vbCrLf + "Connection with " & Winsock1(intArrayNumber).RemoteHostIP & " made."
    txtOutput.SelStart = Len(txtOutput.Text)
    txtSend.SetFocus
    
    'Increase number of current connections for statusbar.
    intNum_ConnectionsNow = intNum_ConnectionsNow + 1
    
    'Call the statusbar update sub
    printConnections
    
    'Turn on timer and set the connection integer to send myInfo to when .2 seconds have passed.
    intChannel = intArrayNumber

    'Enable clock to send myInfo in .2 seconds.
    Timer1.Enabled = True
    
End If ' index = 0

End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    
    'String for receiving data
    Dim Incoming As String
    
    'String - what to do.
    Dim strCut As String
    
    'temp index
    Dim i As Integer
    
    'Recieve incoming string over net.
    Winsock1(Index).GetData Incoming, vbString, bytesTotal
    
    'Cut off the instruction part.
    cutString strCut, Incoming
    
    'Find out what they want.
    Select Case strCut
        Case "print"
            getSendersInfo Index - 1
            txtOutput.Text = txtOutput.Text & vbCrLf & strName + ": " & Incoming
            txtOutput.SelStart = Len(txtOutput.Text)
            txtSend.SetFocus
            'Show first tab if someone sends message.  Anouying but needed.
            SSTab1.Tab = 0
        
        Case "myInfo"
            'Extract the name
            cutString strCut, Incoming
            
            'First part is name.
            strName = strCut
            
            'Next cut is the IP address.
            cutString strCut, Incoming
            strIP = Winsock1(Index).RemoteHostIP ' strCut
            
            'Share is the next part of string.
            cutString strCut, Incoming
            strShare = strCut
            
            'Index is index now.
            strIndex = Index
            
            'Remove blank spots in lstNodes listbox.
            lstNodes.RemoveItem 4 * (Index - 1)
            lstNodes.RemoveItem 4 * (Index - 1)
            lstNodes.RemoveItem 4 * (Index - 1)
            lstNodes.RemoveItem 4 * (Index - 1)
            
            'Add senders info to lstNodes listbox.
            lstNodes.AddItem strShare, 4 * (Index - 1)
            lstNodes.AddItem Index, 4 * (Index - 1)
            lstNodes.AddItem strIP, 4 * (Index - 1)
            lstNodes.AddItem strName, 4 * (Index - 1)
            
            'Add who connected to tree view.
            Set nodTemp = tvwConnects.Nodes.Add(, , strName & strIP, strName, "fileClosed", "fileOpen")
            tvwConnects.Nodes.Item(strName & strIP).Selected = True
            tvwConnects.Nodes.Item(strName & strIP).Expanded = True
    
            'Is there other connections at other end?
            If Incoming <> "0" Then
                'Get the other connections of new connection.
                getConnections Index, Incoming
               
            End If
                  
            'Update log, IPs.
            If ChkIPs And blnLog Then
                lstLogging.AddItem strName & "(" & strIP & ") Connected."
            End If
            'Update log, Date and time.
            If ChkTime And blnLog Then
                lstLogging.AddItem Time
            End If
            
            'Print name of who connected in output textbox.
            txtOutput.Text = txtOutput.Text & vbCrLf & strName & " is Connected"
            txtOutput.SelStart = Len(txtOutput.Text)
            txtSend.SetFocus
            
            'Turn on drop connection button, since there's at least 1 connection.
            cmdDrop.Enabled = True
            
            'Request Welcome message.
            sendToOne Index, "welcome,null"

        Case "requestContacts"
            'Refresh a connection to see who they're connected to.
            sendConnections Index
            
        Case "listOfContacts"
            'Receiving contact list from a connection.
            clearNodes Index
            cutString strCut, Incoming 'Take out first string to see if theres info coming.
            Incoming = strCut & "," & Incoming 'Then put it back in to be compatable with sub call.
            If strCut > 0 Then
                getConnections Index, Incoming
            End If
            
        Case "welcome"
            'Send welcome message to new connection.
            sendToOne Index, "printWelcome," & strWelcome
            
        Case "printWelcome"
            txtOutput.Text = txtOutput.Text + vbCrLf + Incoming
            txtOutput.SelStart = Len(txtOutput.Text)
            txtSend.SetFocus
            sendConnectionsToAll Index

        Case "requestDir"
            'send your file directory to who asked for it.
            sendDir Index
            
        Case "showFiles"
            'Fill box with files.
            fillDir Incoming
            
        Case "requestFile"
            'Someone requested a file.
            setupSendFile Index, Incoming
            
        Case "makeFile"
            cutString strCut, Incoming
            'Open a file for writing.
            makeFile Index, Val(strCut), Incoming
            
        Case "setupFile"
            'Store the other computers channel in your channel array.
            cutString strCut, Incoming
            intChannels(strCut) = Incoming
            sendFile Index, Val(strCut)
            
        Case "uploadingFile"
            'Someone wants to upload a file.
            If ChkUpload Then
                sendToOne Index, "requestFile," & Incoming
                getSendersInfo (Index - 1)
                
            'Log that it's a upload to you, not one of your downloads.
            If ChkTransfers Then
                lstLogging.AddItem Incoming & " is being uploaded from " & strName
            End If
            
            'IF uploading not permitted. Tell them.
            Else
                sendToOne Index, "print,Uploading not permited."
            End If
            
        Case "sendFile"
            'Send file block.
            'cutString strCut, Incoming
            sendFile Index, Val(Incoming)
            
        Case "moreFile"
            cutString strCut, Incoming
            moreFile Index, Val(strCut), Incoming
            
        Case "fileDone"
            'cutString strCut, Incoming
            'End of file reached, close channel, set array spot to 0(open).
            Close #Incoming
            intChannels(Incoming) = 0
            
            getSendersInfo Index - 1
            txtOutput.Text = txtOutput.Text + vbCrLf + "File successfully downloaded from " & strName & "."
            'If logging file downloads, log it.
            If ChkTransfers And blnLog Then
                lstLogging.AddItem "File successfully downloaded from " & strName & "."
            End If
            File1.Refresh
            File2.Refresh
            File3.Refresh
            
            'Was file a sound to play?
            If blnWav Then
                MMControl1.Command = "open"
                cmdSendSound.Enabled = True
                blnWav = False
                MMControl1.Command = "Play"
            End If
            
        Case "searchFor"
            cutString strCut, Incoming
            'Search for file locally.
            fileSearch Index, strCut, Incoming
            
        Case "fileFound"
            'frmConnect.lstSearch.AddItem "Working"
            fileFound Incoming
            
    End Select
        
    
End Sub

Private Sub sendToEveryone()

Dim i As Integer

For i = 1 To intNum_Connections
    If Winsock1(i).State <> sckClosed Then
        Winsock1(i).SendData ("print," & txtSend.Text)
    End If
Next i

       
End Sub

Private Sub sendToOne(Index As Integer, output As String)

    'Use this for sending nonchat information
    Winsock1(Index).SendData output
    
End Sub

Private Sub printConnections()

    'Update current connections in statusbar
    StatusBar1.Panels(2).Text = "Connections open: " & intNum_ConnectionsNow
         
End Sub


Private Sub cutString(strCut As String, Incoming As String)
On Error GoTo cutError

    'Seporate into 2 seporate strings with comma.
    'First get everything before the comma and put it in strControl.
    strCut = Mid(Incoming, 1, InStr(1, Incoming, ",") - 1)

On Error GoTo cutError2
    'Second get everything behind comma and put it in strData.
    Incoming = Mid(Incoming, InStr(1, Incoming, ",") + 1, Len(Incoming))
    Exit Sub
    
cutError:
    MsgBox ("cutString error #1")
    MsgBox (Incoming)
    Exit Sub
    
cutError2:
    MsgBox ("cutString error #2")
    MsgBox (strCut)
    MsgBox (Incoming)
    
End Sub

Private Sub sendMyInfo(Index As Integer)
    
Dim i As Integer
Dim output As String
Dim strContact As String

'Build up My Information string. 3 parts.
output = "myInfo," & txtName.Text & ","
output = output & txtLocalIP.Text & ","
'Does this node share files?
If optShare(0) Then
    output = output & "yes"
Else
    output = output & "no"
End If

'Add information of other connections
If intNum_ConnectionsNow > 1 Then 'IF there's other connections beside the one just made.
    getSendersInfo Index - 1 'Compensate for 0 based listbox.
    strContact = strName & strIP 'Get name of this connection.
    output = output & "," & intNum_ConnectionsNow - 1 'Number of other connections beside this one.
    For i = 0 To intNum_ConnectionsNow - 1
        getSendersInfo i
        If (strName & strIP) <> strContact Then 'Not this connection? Send info then.
            output = output & "," & strName
            output = output & "," & strIP
            output = output & "," & strShare
        End If
    Next i
    output = output & ",null"
Else
    output = output & ",0,null"
End If

'Send it using sendToOne subroutine
sendToOne Index, output

End Sub

Private Sub sendDir(Index As Integer)

Dim i As Integer
Dim intLength As Integer
Dim output As String

'Do I share files?
If optShare(1) Then
    'Do you share files?
    sendToOne Index, "print,Directory Search denied."

Else

    'Use second list box in case ftp tab is open when
    'request for file list is made. Otherwise, will
    'select each and slowly go down list as you watch...

    'Get number of files in filelistbox2.
    intLength = File2.ListCount

    'If nothing in directory, Nothing to send.
    If intLength = 0 Then
        Exit Sub
    End If

    output = "showFiles," & intLength

    For i = 0 To intLength - 1
        'Select filenames one at a time and append to string.
        File2.ListIndex = i
        output = output & "," & File2.FileName
    Next i

    'send the directory.
    sendToOne Index, output
 
End If
End Sub

Private Sub fillDir(Incoming As String)

Dim i As Integer
Dim strCut As String

    'Get the length of list in FTP Directory.
    cutString strCut, Incoming
    i = Val(strCut)
    
    'Loop if need to do more than once.
    If i > 1 Then
        For i = 1 To i - 1
             cutString strCut, Incoming
             lstFiles.AddItem strCut
        Next i
    End If
    
    'Add last one manually, nothing after last comma.
    'Would error if cutString used on last one.
    lstFiles.AddItem Incoming
    lstFiles.ListIndex = 0

End Sub

Private Sub setupSendFile(Index As Integer, Incoming As String)

Dim intLocalChannel As Integer
On Error GoTo FileNotFound

findChannel intLocalChannel
    
If Incoming = "PS_SoundFile.wav" Then 'Not a file, a sound.
    Open appPath & Incoming For Binary As #intLocalChannel

Else
    Open Dir1.Path & "\" & Incoming For Binary As #intLocalChannel

End If 'Sound file?

sendToOne Index, "makeFile," & intLocalChannel & "," & Incoming

Exit Sub

FileNotFound:

sendToOne Index, "print,File not found or currenty open. Refresh directory."

End Sub

Private Sub sendFile(Index As Integer, intLocalChannel As Integer)

    'All of file sent?
    If EOF(intLocalChannel) Then
        sendToOne Index, "fileDone," & intChannels(intLocalChannel)
        Close intLocalChannel
        intChannels(intLocalChannel) = 0
        'Log if logging on.
        If ChkTransfers And blnLog Then
            'Dim dToday As Date
            lstLogging.AddItem "File sent."
        End If
        
        'If sound was send, reenable sound.
        MMControl1.Command = "Open"
        cmdSendSound.Enabled = True
        
    'Send some more data.
    Else
        strFileString = "moreFile," & intChannels(intLocalChannel) & ","
        strFileString = strFileString & Input(4064, #intLocalChannel)
                  
        sendToOne Index, strFileString
    End If

End Sub

Private Sub makeFile(Index As Integer, intHostChannel As Integer, Incoming As String)

Dim intLocalChannel As Integer

On Error GoTo fileOpenError

    'Find a unused channel on this system.
    findChannel intLocalChannel

    'Save to channel used on other connection.
    'use it to other computer what file is being sent.
    intChannels(intLocalChannel) = intHostChannel
    
    'First check if sound to play.
    If Incoming = "PS_SoundFile.wav" Then 'Not a file, a sound.
        MMControl1.Command = "Close"
        cmdSendSound.Enabled = False
        'Set sound vaiable(not part of the control) so it will be played.
        blnWav = True
        'Save the sound file in app.path. Not a normal download.
        Open appPath & Incoming For Binary As #intLocalChannel
    Else
        'File transfer, so save in download directory.
        Open Dir1.Path & "\" & Incoming For Binary As #intLocalChannel
    End If
    
    'Tell host what channel was set up, and your local one.
    sendToOne Index, "setupFile," & intHostChannel & "," & intLocalChannel
    
    'Post that file is being downloaded.
    getSendersInfo Index - 1
    txtOutput.Text = txtOutput.Text & vbCrLf & "Getting " & Incoming & " from " & strName & "."
    'If logging file downloads, log it.
    If ChkTransfers And blnLog Then
        lstLogging.AddItem "Getting " & Incoming & " from " & strName & "."
    End If
    
    Exit Sub
    
fileOpenError:
    MsgBox ("Error opening file!")

End Sub

Private Sub moreFile(Index As Integer, intLocalChannel As Integer, Incoming As String)

Put intLocalChannel, , Incoming

sendToOne Index, "sendFile," & intChannels(intLocalChannel)

End Sub

Private Sub getSendersInfo(Index As Integer)

Dim i As Integer

    'i is the index to senders 4 info parts in listbox lstNodes.
    i = Index * 4
    
    'Fill the strings with senders info in listbox.
    lstNodes.ListIndex = i
    strName = lstNodes.Text
    lstNodes.ListIndex = i + 1
    strIP = lstNodes.Text
    lstNodes.ListIndex = i + 2
    strIndex = lstNodes.Text
    lstNodes.ListIndex = i + 3
    strShare = lstNodes.Text
    
End Sub

Private Sub findChannel(intLocalChannel As Integer)
Dim i As Integer

'Find unused channel to Write with/Put into.
For i = 2 To 202
    If intChannels(i) = 0 Then
        intLocalChannel = i
        Exit For
    End If
Next i

End Sub

Private Sub getConnections(Index As Integer, Incoming As String)
Dim strCut As String
Dim i As Integer
Dim key As String
Dim share As String
Dim nodeName As String 'The nodes name

    'Clear the listbox
    lstNodes_nodes(Index).Clear
    
    getSendersInfo Index - 1
    cutString strCut, Incoming
    For i = 0 To Val(strCut) - 1 'Compensate for 0 based listbox.
        cutString strCut, Incoming      'Add name
        lstNodes_nodes(Index).AddItem strCut
        nodeName = strCut
        cutString strCut, Incoming      'Add IP
        lstNodes_nodes(Index).AddItem strCut
        'Make a unique key with nodeName,nodeIP, and parents IP.
        key = nodeName & strCut & strIP
        cutString strCut, Incoming      'Add share?
        lstNodes_nodes(Index).AddItem strCut
        If strCut = "no" Then
            Set nodTemp = tvwConnects.Nodes.Add(strName & strIP, tvwChild, key, nodeName, "fileClosed", "fileClosed")
        Else
            Set nodTemp = tvwConnects.Nodes.Add(strName & strIP, tvwChild, key, nodeName, "fileOpen", "fileOpen")
        End If
        
    Next i
    
End Sub

Private Sub sendConnections(Index As Integer)

Dim strContact As String
Dim output As String
Dim i As Integer

If intNum_ConnectionsNow > 1 Then 'IF there's other connections beside the one just made.
    getSendersInfo Index - 1 'Compensate for 0 based listbox.
    strContact = strName & strIP 'Get name of this connection.
    output = "listOfContacts," & intNum_ConnectionsNow - 1 'Number of other connections beside this one.
    For i = 0 To intNum_ConnectionsNow - 1
        getSendersInfo i
        If (strName & strIP) <> strContact Then 'Not this connection? Send info then.
            output = output & "," & strName
            output = output & "," & strIP
            output = output & "," & strShare
        End If
    Next i
    
    output = output & ",null" 'Tack on an extra string for cutString sub to work right.
    
Else

    output = "listOfContacts,0,null" 'No other contacts.
    
End If

sendToOne Index, output

End Sub

Private Sub clearNodes(Index As Integer)

Dim i As Integer
Dim strCut As String

    'Are there nodes in there already?
    'Take out nodes.
    If lstNodes_nodes(Index).ListCount > 0 Then
        For i = 0 To lstNodes_nodes(Index).ListCount / 3 - 1 'ListCount is 1 based, so remove 1.
            lstNodes_nodes(Index).ListIndex = i * 3
            strCut = lstNodes_nodes(Index).Text
            'Add IP to the Key.
            lstNodes_nodes(Index).ListIndex = i * 3 + 1
            strCut = strCut & lstNodes_nodes(Index).Text
            
            getSendersInfo Index - 1
            strCut = strCut + strIP
            'Remove node.
            tvwConnects.Nodes.Remove strCut
        Next i
    End If
    
    'Clear the nodes listbox.
    lstNodes_nodes(Index).Clear

End Sub

Private Sub sendConnectionsToAll(Index As Integer)

Dim i As Integer

For i = 1 To intNum_Connections
    If Winsock1(i).State <> sckClosed Then
        If i <> Index Then
            sendConnections i
        End If
    End If
Next i

End Sub

Private Sub alreadyConnected(blnConnected As Boolean, strConnection As String)
Dim i As Integer
'Assume not connected.
blnConnected = False

'Trying to connect to yourself???
If strConnection = txtLocalIP.Text Then
    blnConnected = True
    Exit Sub
End If

'See if already connected to this IP before connecting.
If intNum_ConnectionsNow <> 0 Then
    For i = 1 To intNum_Connections
        If Winsock1(i).State <> sckClosed Then
            If strConnection = Winsock1(i).RemoteHostIP Then
                blnConnected = True
                Exit For
            End If
        End If
    Next i
End If
End Sub

Private Sub fileSearch(Index As Integer, strCut As String, Incoming As String)

    Dim i As Integer
    Dim strTemp As String
    Dim strNodeName As String
    Dim strNodeIP   As String
    Dim strFile As String
    Dim strFiles As String
    Dim strPassOn As String 'Build this string to pass on as you take incoming apart.
    Dim strFileFound As String 'Build this string and use if file found.
    
    'The text being looked for.
    strFile = strCut
    
    'Start building passOn string.
    strPassOn = "searchFor," & strFile & "," & txtName & "," & txtLocalIP
    strFileFound = txtName & "," & txtLocalIP
    
    'get previous sender in list.
    cutString strCut, Incoming
    strNodeName = strCut
    cutString strCut, Incoming
    
    'Internet IP fix. Done this way to be compatable with previous versions.
    strNodeIP = Winsock1(Index).RemoteHostIP ' This is used to get internet IPs.
    strPassOn = strPassOn & "," & strNodeName & "," & strNodeIP
            strFileFound = strFileFound & "," & strNodeName & "," & strNodeIP
    'End of internet fix.............
    
    If Incoming <> "null" Then
        'Look for end of incoming string.
        Do Until Incoming = "null"
    
            'get next sender in list.
            cutString strCut, Incoming
            strNodeName = strCut
            cutString strCut, Incoming
            strNodeIP = strCut
    
            If strNodeName & strNodeIP <> txtName & txtLocalIP Then
                'Sender not in list so far, keep going.
                strPassOn = strPassOn & "," & strNodeName & "," & strNodeIP
                strFileFound = strFileFound & "," & strNodeName & "," & strNodeIP
            Else
                Exit Sub 'Found self in string.  Kill search.
            End If
        
        Loop 'Keep adding everyone, one by one to list.
    End If ' Incoming = "null"
    
    strPassOn = strPassOn & ",null" 'null is end of string.
    strFileFound = strFileFound & ",null" 'Same here.
    
    'Send the search to all connections except where it came from.
    For i = 1 To intNum_Connections
        If Winsock1(i).State <> sckClosed Then
            If i <> Index Then
                sendToOne i, strPassOn
            End If
        End If
    Next i
    
    
    'Now look for file.
    strFiles = "Update your Privashare"
    For i = 0 To File2.ListCount - 1
        File2.ListIndex = i
        strTemp = InStr(1, File2.FileName, strFile, 1)
        If strTemp <> "0" Then
            'File found locally, Add to strFile.
            strFiles = strFiles & "*" & File2.FileName
        End If
        
    Next i
    
    'If files found, send message back to sender, thourgh path of connections.
    If strFiles <> "Update your Privashare" Then
        sendToOne Index, "fileFound," & strFiles & "," & strFileFound
    End If
    
End Sub

Private Sub saveFavorites()

Dim i As Integer
Dim strTemp As String

On Error GoTo saveFav

If frmConnect.lstFavName.ListCount <> 0 Then
 
    Open appPath & "favorites.cfg" For Output As #1

        For i = 0 To frmConnect.lstFavName.ListCount - 1
    
            frmConnect.lstFavName.ListIndex = i
            frmConnect.lstFavIP.ListIndex = i
    
            Write #1, frmConnect.lstFavName.Text
            Write #1, frmConnect.lstFavIP.Text
        
        Next i
        
        Close #1
        
End If

Exit Sub
    
saveFav:
    MsgBox ("saving favorites didn't work")


End Sub

Private Sub fileFound(Incoming As String)

Dim i As Integer
Dim strTemp As String
Dim strFileName As String
Dim strNodeName As String
Dim strNodeIP As String
Dim strFoundName As String
Dim strFoundIP  As String
Dim blnMoreFiles As Boolean

'Get the filename there.
cutString strTemp, Incoming
strFileName = strTemp

'Get the name and ip of computer with file.
cutString strTemp, Incoming
strFoundName = strTemp
cutString strTemp, Incoming
strFoundIP = strTemp

'Remove this computers info from string.
cutString strTemp, Incoming
cutString strTemp, Incoming



If Incoming = "null" Then ' You are the one who is looking for file.
    
    'Loop until there is no more file matches at host IP.
    blnMoreFiles = False
    Do Until blnMoreFiles = True
        
        If InStr(1, strFileName, "*", 1) <> "0" Then
            cutFiles strTemp, strFileName
            If strTemp <> "Update your Privashare" Then
                frmConnect.lstSearch.AddItem strTemp & vbTab & strFoundName & vbTab & strFoundIP
                frmConnect.lstSearchIP.AddItem strFoundIP
            End If
        Else
            frmConnect.lstSearch.AddItem strFileName & vbTab & strFoundName & vbTab & strFoundIP
            frmConnect.lstSearchIP.AddItem strFoundIP
            blnMoreFiles = True
        End If
    
    Loop
    
    frmConnect.lstSearch.ListIndex = frmConnect.lstSearch.ListCount - 1
    frmConnect.lstSearchIP.ListIndex = frmConnect.lstSearch.ListCount - 1
Else

    'Get computers info, next in list, to send "fileFound" back to.
    cutString strTemp, Incoming
    strNodeName = strTemp
    cutString strTemp, Incoming
    strNodeIP = strTemp

    'Find next connection to send back "fileFound" to.
    For i = 0 To intNum_Connections - 1
            getSendersInfo i
            If (strNodeName & strNodeIP) = (strName & strIP) Then
                i = Val(strIndex)
                Exit For
            End If
        Next i
    strTemp = "fileFound," & strFileName & "," & strFoundName & "," & strFoundIP & "," & strNodeName & "," & strNodeIP & "," & Incoming
    
    'Send string back one more computer.
    sendToOne i, strTemp
    
End If ' Incoming.

End Sub


Private Sub cutFiles(strCut As String, Incoming As String)
'On Error GoTo cutError

    'Seporate into 2 seporate strings with *.
    'First get everything before the comma and put it in strControl.
    strCut = Mid(Incoming, 1, InStr(1, Incoming, "*") - 1)

    'Second get everything behind comma and put it in strData.
    Incoming = Mid(Incoming, InStr(1, Incoming, "*") + 1, Len(Incoming))
    Exit Sub
    
End Sub
