VERSION 5.00
Begin VB.Form frmConnect 
   BackColor       =   &H00CCB7B9&
   Caption         =   "Connect and Search"
   ClientHeight    =   3630
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   7515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00CCB7B9&
      Caption         =   "Connect to  file sorce"
      Height          =   255
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton CmdRemove 
      BackColor       =   &H00CCB7B9&
      Caption         =   "Remove Selected"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00CCB7B9&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00CCB7B9&
      Caption         =   "Search..."
      Height          =   255
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdConnectTo 
      BackColor       =   &H00CCB7B9&
      Caption         =   "Connect"
      Height          =   255
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtConnection 
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Top             =   480
      Width           =   3135
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   4920
      TabIndex        =   5
      Top             =   480
      Width           =   2415
   End
   Begin VB.ListBox lstConnections 
      Height          =   1620
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   3135
   End
   Begin VB.ListBox lstFavIP 
      Height          =   450
      Left            =   2760
      TabIndex        =   2
      Top             =   2280
      Width           =   495
   End
   Begin VB.ListBox lstSearch 
      Height          =   1620
      ItemData        =   "frmConnect.frx":0000
      Left            =   3600
      List            =   "frmConnect.frx":0002
      TabIndex        =   1
      Top             =   1440
      Width           =   3735
   End
   Begin VB.ListBox lstFavName 
      Height          =   450
      Left            =   1440
      TabIndex        =   0
      Top             =   2280
      Width           =   495
   End
   Begin VB.ListBox lstSearchIP 
      Height          =   255
      Left            =   3600
      TabIndex        =   15
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Matches found, Location name, and IP"
      Height          =   255
      Left            =   3600
      TabIndex        =   9
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter new IP address to connect to "
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Do a Search for:"
      Height          =   255
      Left            =   4920
      TabIndex        =   6
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name and IP of Favorite connections"
      ForeColor       =   &H80000017&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   2895
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdClose_Click()
    
    frmConnect.Hide
    frmMain.Show
    
End Sub

Private Sub cmdConnectTo_Click()

    'Pass the IP in the connection textbox
    frmMain.cmdConnect_Click txtConnection
    frmConnect.Hide
    
End Sub


Private Sub CmdRemove_Click()
    
    lstFavName.RemoveItem lstConnections.ListIndex
    lstFavIP.RemoveItem lstConnections.ListIndex
    lstConnections.RemoveItem lstConnections.ListIndex
    
End Sub

Private Sub cmdSearch_Click()

If Len(txtSearch) > 2 Then

    Dim i As Integer

    For i = 1 To intNum_Connections
        If frmMain.Winsock1(i).State <> sckClosed Then
            frmMain.Winsock1(i).SendData ("searchFor," & txtSearch.Text & "," & frmMain.txtName & "," & frmMain.txtLocalIP & ",null")
        End If
    Next i

Else

    MsgBox ("You must search for a string longer than 2 charactors.  Otherwise you could get back too many results.")

End If

End Sub

Private Sub Command1_Click()
    
    txtConnection = lstSearchIP
    cmdConnectTo_Click
End Sub

Private Sub lstConnections_Click()

    lstFavIP.ListIndex = lstConnections.ListIndex
    txtConnection.Text = lstFavIP.Text
    
End Sub

Private Sub lstSearch_Click()
    lstSearchIP.ListIndex = lstSearch.ListIndex
End Sub

Private Sub txtConnection_KeyPress(KeyAscii As Integer)
    'If Enter key pressed, send text, clear text box.
    If KeyAscii = 13 Then
        'Pass the IP in the connection textbox
        frmMain.cmdConnect_Click txtConnection
        frmConnect.Hide
    End If
End Sub


Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    'If Enter key pressed, send text, clear text box.
    If KeyAscii = 13 Then
        'Do that search thingy.
        cmdSearch_Click
    End If
End Sub
