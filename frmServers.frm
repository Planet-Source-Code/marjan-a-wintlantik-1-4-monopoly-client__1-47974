VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmServers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Servers: double-click server to connect"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
   Icon            =   "frmServers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnd 
      Caption         =   "Exit"
      Height          =   615
      Left            =   240
      Picture         =   "frmServers.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2880
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServers.frx":0B14
            Key             =   "server"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstServers 
      Height          =   4335
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7646
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Host"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Version"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Guests"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh..."
      Height          =   615
      Left            =   4920
      Picture         =   "frmServers.frx":0F6C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4440
      Width           =   1215
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4200
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AccessType      =   1
      Protocol        =   4
      URL             =   "http://"
   End
End
Attribute VB_Name = "frmServers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdEnd_Click()
Unload Me
'End
End Sub

Private Sub cmdRefresh_Click()
'if we provided command line parameter, use it
If command <> "" Then
    serverName = command
    serverPort = "1234"
    frmGames.Show
    Unload Me
    Exit Sub
    End If
Dim b As String
frmServers.Caption = "Retrieving servers..."
DoEvents
lstServers.ListItems.Clear
DoEvents
b = Inet1.OpenURL("http://gator.monopd.net/", icString)
DoEvents
DoEvents
ServersList b
frmServers.Caption = "Servers: double-click server to connect"
DoEvents
End Sub

Private Sub Form_Load()
Debug.Print "Servers LOAD"
centerForm Me
cmdRefresh_Click
End Sub


Private Sub ServersList(html As String)
Dim p As Integer
Dim q As Integer

Do
p = InStr(UCase(html), "<SERVER")
If p = 0 Then Exit Sub
q = InStr(p, html, "/>")

If q = 0 Then Exit Sub

FillList mID(html, p, q - p + 2)
html = mID(html, q + 1)
DoEvents
Loop

End Sub


Private Sub FillList(html As String)
Dim Host As String
Dim version As String
Dim users As String
Dim Port As String

Dim item As ListItem
'Debug.Print html
'host
Host = GetData("host=" & Chr(34), html)
'version
version = GetData("version=" & Chr(34), html)
'users
users = GetData("users=" & Chr(34), html)
'port
Port = GetData("port=" & Chr(34), html)

Set item = lstServers.ListItems.Add(, Host & ":" & Port, Host, , "server")
item.SubItems(1) = version
item.SubItems(2) = users
Set item = Nothing
End Sub



Private Sub Form_Unload(Cancel As Integer)
Debug.Print "Servers UNLOAD"
End Sub

Private Sub lstServers_DblClick()
'we selected server. Continue!
Dim p As Integer
serverName = lstServers.SelectedItem.Key
p = InStr(serverName, ":")
serverPort = mID(serverName, p + 1)
serverName = Left(serverName, p - 1)
frmGames.Show
Unload Me
End Sub
