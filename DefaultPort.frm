VERSION 5.00
Begin VB.Form frmDefaultPort 
   BackColor       =   &H00CCB7B9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connection Port"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3945
   Icon            =   "DefaultPort.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   3945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00CCB7B9&
      Caption         =   "Use Default"
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00CCB7B9&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00CCB7B9&
      Caption         =   "Set"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"DefaultPort.frx":0442
      Height          =   975
      Left            =   360
      TabIndex        =   5
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CCB7B9&
      Caption         =   "Connection Port"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   600
      TabIndex        =   3
      Top             =   1200
      Width           =   1530
   End
End
Attribute VB_Name = "frmDefaultPort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()

    Unload Me
    
End Sub
Private Sub cmdOK_Click()
'OK was clicked.

'Redefine the settings.
intPort = txtPort.Text

Unload Me
End Sub

Private Sub Command1_Click()

    'This is the Default port used for new connections.
    intPort = 2001
    txtPort.Text = intPort
    
End Sub

Private Sub Form_Load()

'Show the current setting in the text boxes.
txtPort.Text = intPort

End Sub

