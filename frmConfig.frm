VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuration"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPlayerName 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Apply"
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Player name:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   285
      Width           =   915
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this is form with config options
'currently only player name is implemented
Dim NameDirty As Boolean


Private Sub cmdOK_Click()
'if name is changed, change it on server, in game and config file
If NameDirty = True Then
    player.Name = txtPlayerName.text
    frmGames.kSend ".n" & txtPlayerName.text
    WriteConfig
    NameDirty = False
    Me.Hide
    End If
Unload Me
End Sub

Private Sub Form_Load()
Debug.Print "Config LOAD"
NameDirty = False
txtPlayerName.text = player.Name
End Sub



Private Sub Form_Unload(Cancel As Integer)
Debug.Print "Config UNLOAD"
End Sub

Private Sub txtPlayerName_KeyPress(KeyAscii As Integer)
'if player did some writing, apply it
NameDirty = True
End Sub
