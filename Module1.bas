Attribute VB_Name = "Module1"
Option Explicit

'Used by the timer routine to know who to send myInfo to.
Global intChannel As Integer

'Used to select text for color change.
Global intSelText As Integer

'Put app.path in a string

Global appPath As String

'******** Connection variables *********************

'Make temporary node to add to node tree view.
Global nodTemp As Node

'Number of connections in array.
Global intNum_Connections As Integer

'Number of connections in use.
Global intNum_ConnectionsNow As Integer

'Port address for default connectios.
Global intPort As Integer

'Interger array for multiple downloads.
'Each file has it's own channel #2-202 to save and read from.
'#1 is used for saving logs and preferences.
Global intChannels(2 To 202) As Integer

'Logging turned on?
Global blnLog As Boolean

'Welcome string shown when someone connects to you.
Global strWelcome As String

'Make string for file transfers.
Global strFileString As String

'Temp information strings
Global strName As String
Global strIP As String
Global strIndex As String
Global strShare As String

'Value is true if file being sent is a sound.
Global blnWav As Boolean

