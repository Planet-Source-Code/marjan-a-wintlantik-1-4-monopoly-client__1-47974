VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTrade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trade"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9960
   Icon            =   "frmtrade.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   9960
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4680
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtrade.frx":058A
            Key             =   "card"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtrade.frx":0B24
            Key             =   "bag"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add to trade"
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   9855
      Begin VB.ComboBox cmbFrom 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdAppend 
         Caption         =   "Append"
         Height          =   375
         Left            =   8520
         TabIndex        =   10
         Top             =   210
         Width           =   855
      End
      Begin VB.ComboBox cmbTo 
         Height          =   315
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtMoney 
         Height          =   285
         Left            =   3960
         TabIndex        =   6
         Text            =   "0"
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox cmbEstate 
         Height          =   315
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox cmbWhat 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "to"
         Height          =   195
         Left            =   6000
         TabIndex        =   8
         Top             =   300
         Width           =   135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "from"
         Height          =   195
         Left            =   1440
         TabIndex        =   7
         Top             =   300
         Width           =   300
      End
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Accept"
      Height          =   375
      Left            =   8760
      TabIndex        =   2
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdReject 
      Caption         =   "Reject"
      Height          =   375
      Left            =   7440
      TabIndex        =   1
      Top             =   2880
      Width           =   975
   End
   Begin MSComctlLib.ListView lsttrade 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   3201
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "From"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Item"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "To"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblAccepted 
      Alignment       =   1  'Right Justify
      Caption         =   "Label4"
      Height          =   255
      Left            =   480
      TabIndex        =   13
      Top             =   3000
      Width           =   135
   End
   Begin VB.Label lblText 
      Caption         =   "Label3"
      Height          =   255
      Left            =   720
      TabIndex        =   12
      Top             =   3000
      Width           =   5415
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu mnuDeleteItem 
         Caption         =   "Delete from list"
      End
   End
End
Attribute VB_Name = "frmtrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public tradeID As String
Public tradeRevision As String
Public canUnload As Boolean
Private lastButton As Integer

Private Sub cmbFrom_Click()
LoadEstates
End Sub

Private Sub cmbWhat_Click()
If cmbWhat.ListIndex = 0 Then
    cmbEstate.Visible = True
    txtMoney.Visible = False
    Else
    cmbEstate.Visible = False
    txtMoney.Visible = True
    End If
    

End Sub

Private Sub LoadEstates()
'fill combo box with estates
Dim colestate As cEstate
Dim colplayer As cPlayer

Dim pId As String

For Each colplayer In colplayers
If cmbFrom.text = colplayer.Name Then
    pId = colplayer.ID
    Exit For
    End If
Next

cmbEstate.Clear
For Each colestate In colestates
If colestate.Owner = pId Then cmbEstate.AddItem colestate.Name
Next

If cmbEstate.ListCount <> 0 Then cmbEstate.ListIndex = 0

End Sub


Private Sub cmdAccept_Click()
frmGames.kSend ".Ta" & tradeID & ":" & tradeRevision
End Sub

Private Sub cmdAppend_Click()
Dim n As Integer
Dim tCommand As String
Dim colestate As cEstate
'did we add estate already?
'For n = 1 To lsttrade.ListItems.Count
'If InStr(lsttrade.ListItems(n).SubItems(1), cmbEstate.text) > 0 Then Exit Sub
'Next



Select Case cmbWhat.ListIndex
Case 0
'estate
tCommand = ".Te" & tradeID & ":"
For Each colestate In colestates
If colestate.Name = cmbEstate.text Then
    tCommand = tCommand & colestate.ID & ":"
    Exit For
    End If
Next
tCommand = tCommand & GetPlayerId(cmbTo.text)

Case 1
'money
tCommand = ".Tm" & tradeID & ":"
tCommand = tCommand & GetPlayerId(cmbFrom.text) & ":" & GetPlayerId(cmbTo.text) & ":" & txtMoney.text
End Select

'Debug.Print "Append:" & tCommand
frmGames.kSend tCommand
End Sub

Private Sub cmdReject_Click()
frmGames.kSend ".Tr" & tradeID
End Sub


Private Sub Form_Load()
Debug.Print "Tarde LOAD"
canUnload = False
cmbWhat.AddItem "Estate"
cmbWhat.AddItem "Money"


Dim colplayer As cPlayer

For Each colplayer In colplayers
    cmbTo.AddItem colplayer.Name
    cmbFrom.AddItem colplayer.Name
Next
cmbWhat.ListIndex = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If canUnload = False Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Debug.Print "Trade UNLOAD"
End Sub

Private Sub lsttrade_ItemClick(ByVal item As MSComctlLib.ListItem)
If lastButton = 2 Then PopupMenu mnuPopup
End Sub

Private Sub lsttrade_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lastButton = Button
End Sub

Private Sub mnuDeleteItem_Click()
Dim tCommand As String
Dim colestate As cEstate
Dim item As ListItem

'we delete item from trade
Select Case InStr(lsttrade.SelectedItem.SubItems(1), "Money")
Case 0
'estate
tCommand = ".Te" & tradeID & ":"
For Each item In lsttrade.ListItems
    For Each colestate In colestates
    If InStr(item.SubItems(1), colestate.Name) > 0 Then
        tCommand = tCommand & colestate.ID & ":"
        Exit For
        End If
    Next
Next
tCommand = tCommand & "-1"

Case Else
'money
tCommand = ".Tm" & tradeID & ":"
tCommand = tCommand & GetPlayerId(cmbFrom.text) & ":" & GetPlayerId(cmbTo.text) & ":0"
End Select

Debug.Print "delete:" & tCommand
frmGames.kSend tCommand

End Sub
