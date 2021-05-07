VERSION 5.00
Begin VB.Form frmWelcome 
   BackColor       =   &H00CCB7B9&
   Caption         =   "New Welcome message"
   ClientHeight    =   2325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2325
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00CCB7B9&
      Caption         =   "Done"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00CCB7B9&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtWelcome 
      CausesValidation=   0   'False
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Enter Welcome message in here for visitors to read on connect."
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00CCB7B9&
      Caption         =   "Enter a welcome message for anyone who visits your node:"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    frmWelcome.Hide
    
End Sub

Private Sub Command2_Click()
    'Make a new Welcome message.
    strWelcome = txtWelcome.Text
    frmWelcome.Hide
    
End Sub

Private Sub Form_Load()
    'Put current welcome sting in textbox.
    txtWelcome = strWelcome
    
    
End Sub
