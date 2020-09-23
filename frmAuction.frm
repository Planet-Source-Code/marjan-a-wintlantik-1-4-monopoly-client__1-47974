VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAuction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   Icon            =   "frmAuction.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   5160
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3360
      Top             =   3240
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
            Picture         =   "frmAuction.frx":058A
            Key             =   "user"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   4935
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.CommandButton cmdBid 
      Caption         =   "Bid"
      Height          =   405
      Left            =   1800
      TabIndex        =   2
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox txtBid 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   3300
      Width           =   1335
   End
   Begin MSComctlLib.ListView lstAuction 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Player"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Offer"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmAuction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'user cannot just close this form by clicking on 'x'
Public canClose As Boolean

Private Sub cmdBid_Click()
'we send our bid offer
frmGames.kSend ".ab" & auctId & ":" & txtBid.text
End Sub

Private Sub Form_Load()
canClose = False
Debug.Print "Auction LOAD"
Dim aPlayer As cPlayer
Dim item As ListItem
centerForm Me
canClose = False
'load players in list

For Each aPlayer In colplayers
Set item = lstAuction.ListItems.Add(, "p" & aPlayer.ID, aPlayer.Name, , "user")
item.SubItems(1) = "0"
Next


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'can form be unloaded?
If canClose = False Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Debug.Print "Auction UNLOAD"
End Sub
