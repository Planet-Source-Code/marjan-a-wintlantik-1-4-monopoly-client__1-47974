VERSION 5.00
Begin VB.Form frmDispEstate 
   Caption         =   "Estate data"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3150
   Icon            =   "frmDispEstate.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   3150
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmEstate 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   3135
      Begin VB.Label lblData 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   1200
      Picture         =   "frmDispEstate.frx":058A
      Top             =   2280
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   1200
      Picture         =   "frmDispEstate.frx":1454
      Top             =   2280
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image icoCrd 
      Height          =   240
      Left            =   3480
      Picture         =   "frmDispEstate.frx":231E
      Top             =   1800
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image icoMortgaged 
      Height          =   240
      Left            =   3960
      Picture         =   "frmDispEstate.frx":28A8
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image icoHotel 
      Height          =   480
      Left            =   3840
      Picture         =   "frmDispEstate.frx":2E32
      Top             =   1320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ico1House 
      Height          =   480
      Left            =   4560
      Picture         =   "frmDispEstate.frx":36FC
      Top             =   720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ico2House 
      Height          =   480
      Left            =   5040
      Picture         =   "frmDispEstate.frx":3FC6
      Top             =   720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ico3House 
      Height          =   480
      Left            =   4200
      Picture         =   "frmDispEstate.frx":4890
      Top             =   2400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ico4house 
      Height          =   480
      Left            =   4800
      Picture         =   "frmDispEstate.frx":515A
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblGroup 
      Alignment       =   2  'Center
      Caption         =   "Label1"
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
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "frmDispEstate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub DisplayEstate(estate As cEstate)
'display Estate
Dim eGroup As String
Dim eOwner As String
Dim ePlayer As cPlayer
Dim eCGroup As cEstateGroup

BackColor = estate.BgColor
'we don't want back color being black.... so we use estate's back color
If estate.Color = 0 Then
    lblName.BackColor = estate.BgColor
    lblGroup.BackColor = estate.BgColor
    Else
    lblName.BackColor = estate.Color
    lblGroup.BackColor = estate.Color
    End If
lblName.Caption = estate.Name

frmEstate.BackColor = estate.BgColor
lblData.BackColor = estate.BgColor

'is estate owned?
If estate.Owner <> "-1" Then
    For Each ePlayer In colplayers
    If ePlayer.ID = estate.Owner Then
        eOwner = ePlayer.Name
        Exit For
        End If
    Next
    Else
    eOwner = "None"
    End If
'does belong to any group
If estate.Group <> "-1" Then
    For Each eCGroup In colEstateGroups
    If eCGroup.ID = estate.Group Then
        eGroup = eCGroup.Name
        Exit For
        End If
    Next
    
    Else
    eGroup = "None"
    End If
    
lblGroup.Caption = eGroup
If estate.CanBeOwned = True Then
    lblData.Caption = "OWNER : " & eOwner & vbCrLf
    lblData.Caption = lblData.Caption & "GROUP : " & eGroup & vbCrLf
    lblData.Caption = lblData.Caption & "PRICE : " & estate.Price & vbCrLf
    lblData.Caption = lblData.Caption & "_________________________" & vbCrLf & vbCrLf
    lblData.Caption = lblData.Caption & "            RENT : " & estate.Rent0 & vbCrLf
    lblData.Caption = lblData.Caption & "  RENT + 1 HOUSE : " & estate.Rent1 & vbCrLf
    lblData.Caption = lblData.Caption & "RENT + 2 HOUSES : " & estate.Rent2 & vbCrLf
    lblData.Caption = lblData.Caption & "RENT + 3 HOUSES : " & estate.Rent3 & vbCrLf
    lblData.Caption = lblData.Caption & "RENT + 4 HOUSES : " & estate.Rent4 & vbCrLf
    lblData.Caption = lblData.Caption & "    RENT + HOTEL : " & estate.Rent5 & vbCrLf
    lblData.Caption = lblData.Caption & "_________________________" & vbCrLf & vbCrLf
    lblData.Caption = lblData.Caption & "     HOUSE PRICE : " & (estate.Houseprice * 2) & vbCrLf
    lblData.Caption = lblData.Caption & "     HOTEL PRICE : " & (estate.Houseprice * 2)
    End If
'diplay status icons...
If estate.CanBeOwned = True Then
    icoMortgaged.Visible = estate.Mortaged
    icoCrd.Visible = (estate.Owner = "-1")
    Select Case estate.Houses
    Case 1
    ico1House.Visible = True
    Case 2
    ico2House.Visible = True
    Case 3
    ico3House.Visible = True
    Case 4
    ico4house.Visible = True
    Case 5
    icoHotel.Visible = True
    End Select
    Else
    frmEstate.Visible = False
    lblGroup.Visible = False
    lblName.ForeColor = vbBlack
    End If

Show
End Sub

Private Sub Form_Load()
Debug.Print "Display Estate LOAD"
'initiate coordinates...
centerForm Me
ico1House.Left = 0
ico1House.Top = frmEstate.Top - ico1House.Height

ico2House.Left = 0
ico2House.Top = frmEstate.Top - ico1House.Height

ico3House.Left = 0
ico3House.Top = frmEstate.Top - ico1House.Height

ico3House.Left = 0
ico3House.Top = frmEstate.Top - ico1House.Height

ico4house.Left = 0
ico4house.Top = frmEstate.Top - ico1House.Height

icoHotel.Left = 0
icoHotel.Top = frmEstate.Top - ico1House.Height

icoCrd.Top = ico1House.Top
icoCrd.Left = ico1House.Left + ico1House.Width

icoMortgaged.Top = icoCrd.Top
icoMortgaged.Left = icoCrd.Left + icoCrd.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
Debug.Print "Display Estate UNLOAD"
End Sub
