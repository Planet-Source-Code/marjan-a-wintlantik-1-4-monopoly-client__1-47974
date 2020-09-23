VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmGame 
   AutoRedraw      =   -1  'True
   Caption         =   "w"
   ClientHeight    =   7380
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10305
   Icon            =   "frmGame.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   10305
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Games..."
      Height          =   615
      Left            =   3960
      Picture         =   "frmGame.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4200
      Width           =   1455
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5280
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGame.frx":0B14
            Key             =   "field"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGame.frx":10AE
            Key             =   "dice"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGame.frx":1648
            Key             =   "info"
         EndProperty
      EndProperty
   End
   Begin WinTlantik.ControlPlayer conPlayer 
      Height          =   975
      Index           =   0
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1720
   End
   Begin VB.Timer tPause 
      Enabled         =   0   'False
      Left            =   4920
      Top             =   4920
   End
   Begin VB.Frame frmEstate 
      Height          =   4095
      Left            =   3600
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   6495
      Begin MSComctlLib.ListView lstMsgs 
         Height          =   1575
         Left            =   360
         TabIndex        =   20
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   2778
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton cmdJailRoll 
         Caption         =   "Roll attempt"
         Height          =   375
         Left            =   2760
         TabIndex        =   17
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdJailPay 
         Caption         =   "Pay && leave"
         Height          =   375
         Left            =   1200
         TabIndex        =   16
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdJailUseCard 
         Caption         =   "Use card"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   2760
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdPayPercentage 
         Caption         =   "Command2"
         Height          =   375
         Left            =   4440
         TabIndex        =   13
         Top             =   3000
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdPayStatic 
         Caption         =   "Command1"
         Height          =   375
         Left            =   4440
         TabIndex        =   12
         Top             =   3360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdRoll 
         Caption         =   "Roll"
         Height          =   615
         Left            =   3360
         Picture         =   "frmGame.frx":1BE2
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdBuy 
         Caption         =   "Buy"
         Height          =   375
         Left            =   2280
         TabIndex        =   10
         Tag             =   ".eb"
         Top             =   3360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdEndturn 
         Caption         =   "End turn"
         Height          =   375
         Left            =   1200
         TabIndex        =   9
         Tag             =   ".E"
         Top             =   3360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdAuction 
         Caption         =   "Auction"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Tag             =   ".ea"
         Top             =   3360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Image ico4house 
         Height          =   480
         Left            =   4920
         Picture         =   "frmGame.frx":216C
         Top             =   2400
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image ico3House 
         Height          =   480
         Left            =   5640
         Picture         =   "frmGame.frx":2A36
         Top             =   2400
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image ico2House 
         Height          =   480
         Left            =   4200
         Picture         =   "frmGame.frx":3300
         Top             =   2280
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image ico1House 
         Height          =   480
         Left            =   5520
         Picture         =   "frmGame.frx":3BCA
         Top             =   1920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image icoHotel 
         Height          =   480
         Left            =   2400
         Picture         =   "frmGame.frx":4494
         Top             =   1680
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image icoMortgaged 
         Height          =   240
         Left            =   4200
         Picture         =   "frmGame.frx":4D5E
         Top             =   1680
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image icoCrd 
         Height          =   240
         Left            =   2400
         Picture         =   "frmGame.frx":52E8
         Top             =   1320
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblEstateGroup 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Left            =   4320
         TabIndex        =   21
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   780
         Left            =   2760
         Picture         =   "frmGame.frx":5872
         Top             =   360
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   855
         Left            =   3720
         Picture         =   "frmGame.frx":73B4
         Stretch         =   -1  'True
         Top             =   240
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.Label lblEstateName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   2400
         TabIndex        =   14
         Top             =   1320
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Start game"
      Height          =   615
      Left            =   7680
      Picture         =   "frmGame.frx":1480A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4200
      Width           =   1455
   End
   Begin MSComctlLib.StatusBar sBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   7005
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmConf 
      Caption         =   "Game configuration"
      Height          =   4095
      Left            =   3600
      TabIndex        =   2
      Top             =   0
      Width           =   6015
      Begin VB.CheckBox chkOpt 
         Caption         =   "Check1"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   5775
      End
   End
   Begin VB.TextBox txtChat 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   6600
      Width           =   3495
   End
   Begin VB.TextBox txtChatList 
      Height          =   4335
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2160
      Width           =   3495
   End
   Begin MSComctlLib.ListView lstPlayers 
      Height          =   2175
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   3836
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Host"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Image pictHotel 
      Height          =   480
      Index           =   0
      Left            =   5760
      Picture         =   "frmGame.frx":14D94
      Top             =   6480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image pict4house 
      Height          =   480
      Index           =   0
      Left            =   8520
      Picture         =   "frmGame.frx":1565E
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image pict3house 
      Height          =   480
      Index           =   0
      Left            =   7440
      Picture         =   "frmGame.frx":15F28
      Top             =   6120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image pict2House 
      Height          =   480
      Index           =   0
      Left            =   6720
      Picture         =   "frmGame.frx":167F2
      Top             =   6120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image pict1House 
      Height          =   480
      Index           =   0
      Left            =   6120
      Picture         =   "frmGame.frx":170BC
      Top             =   6120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image pictCard 
      Height          =   240
      Index           =   0
      Left            =   4800
      Picture         =   "frmGame.frx":17986
      Top             =   5880
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image pictMortaged 
      Height          =   240
      Index           =   0
      Left            =   5160
      Picture         =   "frmGame.frx":17F10
      Top             =   5880
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblPName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   6720
      TabIndex        =   7
      Top             =   4200
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image pIece 
      Height          =   720
      Index           =   0
      Left            =   6855
      Picture         =   "frmGame.frx":1849A
      Top             =   5040
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Menu mnuGame 
      Caption         =   "Game"
      Begin VB.Menu mnuConfig 
         Caption         =   "Configuration"
      End
      Begin VB.Menu mnugBankrupt 
         Caption         =   "Declare bancrupcy"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSend 
         Caption         =   "Send"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPDispData 
         Caption         =   "Display Data"
      End
      Begin VB.Menu mnuPRQTrade 
         Caption         =   "Request Trade with"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPMortgage 
         Caption         =   "Mortgage"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPUnmortgage 
         Caption         =   "Unmortgage"
         Visible         =   0   'False
      End
      Begin VB.Menu mnupSellTobank 
         Caption         =   "Sell estate to bank"
         Visible         =   0   'False
      End
      Begin VB.Menu mnupBuyHouse 
         Caption         =   "Buy House"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPSellHouse 
         Caption         =   "Sell House"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPBuyHotel 
         Caption         =   "Buy Hotel"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPSellHotel 
         Caption         =   "Sell Hotel"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnup2 
      Caption         =   "popup2"
      Visible         =   0   'False
      Begin VB.Menu mnup2rQTrade 
         Caption         =   "Trade with"
      End
   End
End
Attribute VB_Name = "frmgame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
 Dim mLastX As Single
 Dim mLastY As Single
 Dim mLastButton
 Dim doWait As Boolean

Private Sub chkOpt_Click(Index As Integer)
'option clicked. Can we send change? (where in game are we?)
If conIn = True Then Exit Sub

If chkOpt(Index).Value = 1 Then frmGames.kSend chkOpt(Index).tag & "1"
If chkOpt(Index).Value = 0 Then frmGames.kSend chkOpt(Index).tag & "0"
    

End Sub

Private Sub cmdAuction_Click()
'we auction estate
frmGames.kSend cmdAuction.tag
End Sub

Private Sub cmdback_Click()
'we quit game
mnuExit_Click
End Sub

Private Sub cmdBuy_Click()
'we try to buy estate (depends on money quantity)
frmGames.kSend cmdBuy.tag

End Sub

Private Sub cmdEndturn_Click()
'end turn command. my opinion: OBSOLETE
frmGames.kSend cmdEndturn.tag
End Sub

Private Sub cmdGo_Click()
'if I am master of the game, let's start it
frmGames.kSend ".gs"
End Sub

Private Sub cmdJailPay_Click()
'pay jail
frmGames.kSend ".jp"

End Sub

Private Sub cmdJailRoll_Click()
'roll doubles to get out
frmGames.kSend ".jr"
End Sub

Private Sub cmdJailUseCard_Click()
'use card
frmGames.kSend ".jc"
End Sub

Private Sub cmdPayPercentage_Click()
'pay tax percentage
frmGames.kSend cmdPayPercentage.tag
End Sub

Private Sub cmdPayStatic_Click()
'pay statis tax
frmGames.kSend cmdPayStatic.tag

End Sub

Private Sub cmdRoll_Click()
'roll dice
frmGames.kSend ".r"

End Sub




Private Sub conPlayer_RightClick(Index As Integer)
'we right-clicked on list of player's estates
If conPlayer(Index).PlayerID <> player.ID Then
    'if we didn't click ourself, enable 'request trade with' option
    mnup2rQTrade.Caption = "Request trade with " & GetPlayerName(conPlayer(Index).PlayerID)
    mnup2rQTrade.tag = conPlayer(Index).PlayerID
    PopupMenu mnup2
    End If

End Sub

Private Sub Form_Load()
Debug.Print "Game LOAD"
'canDraw: if we can draw board
canDraw = False
centerForm Me
sBar.Panels(1).AutoSize = sbrSpring
sBar.Panels(1).text = "Waiting for game to begin."
conIn = False
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'we clicked on board
Const clipBrd = 20

If frmEstate.Visible = False Then Exit Sub

Dim StartX As Long, startY As Long
Dim endX As Long, endY As Long

StartX = txtChat.Width + clipBrd + txtChat.Left
startY = clipBrd + lstPlayers.Top

endX = Me.Width - clipBrd - 130
endY = Me.Height - clipBrd - sBar.Height - 690

If X > endX Or X < StartX Or Y > endY Or Y < startY Then Exit Sub

Dim estate As cEstate
mLastX = X
mLastY = Y

If Button = 1 Then
    frmDispEstate.DisplayEstate colestates(GetEstateID(X, Y) + 1)
    End If
'if coordinates are within board range, we check mouse buttons
If Button = 2 Then
    'reset menu options
    mnuPMortgage.Visible = False
    mnuPUnmortgage.Visible = False
    mnupSellTobank.Visible = False
    mnuPBuyHotel.Visible = False
    mnuPSellHotel.Visible = False
    'and enable appropriate popup menu options
    Set estate = colestates(GetEstateID(X, Y) + 1)
    If estate.Owner <> player.ID And estate.Owner <> "-1" Then
        mnuPRQTrade.Caption = "Request trade with " & GetPlayerName(estate.Owner)
        mnuPRQTrade.Visible = True
        mnuPRQTrade.tag = estate.Owner
        Else
        mnuPRQTrade.Visible = False
        End If
        
    If estate.Owner = player.ID Then
        'what can we do with estate?
        'buy house?
        mnupSellTobank.Visible = canSell2Bank
        mnupBuyHouse.Visible = estate.CanBuyHouses
        mnuPSellHouse.Visible = estate.CanSellHouses
        
        If estate.Houses = 5 And mnuPSellHouse.Visible = True Then
            mnuPSellHotel.Visible = True
            mnuPSellHouse.Visible = False
            Else
            mnuPSellHotel.Visible = False
            End If
        
        
        If estate.Houses = 4 And mnupBuyHouse.Visible = True Then
            mnuPBuyHotel.Visible = True
            mnupBuyHouse.Visible = False
            Else
            mnuPBuyHotel.Visible = False
            End If
            
        mnuPMortgage.Visible = Not (estate.Mortaged)
        If estate.Houses <> 0 Then mnuPMortgage.Visible = False
        
        mnuPUnmortgage.Visible = estate.Mortaged
        End If
    
    PopupMenu mnuPopup
    End If
    
End Sub
Private Function GetEstateID(xc As Single, yc As Single) As Integer
'here we get estate ID from mouse coordinates
Dim boardWidth As Long, boardHeight As Long
Dim StartX As Long, startY As Long
Dim endX As Long, endY As Long
Const clipBrd = 20
Dim rowHeight As Long
Dim colWidth As Long
Dim n As Integer
Dim smallWidth As Long
Dim smallHeight As Long
Dim numInRow As Integer
Dim eID As Integer


'how many fields in one quadrant?
numInRow = colestates.Count / 4



StartX = txtChat.Width + clipBrd + txtChat.Left
startY = clipBrd + lstPlayers.Top

endX = Me.Width - clipBrd - 130
endY = Me.Height - clipBrd - sBar.Height - 690

boardHeight = endY - startY
boardWidth = endX - StartX


'we'll take that height is 1/6th of board height...
'and width of column is 1/6th of board width...

rowHeight = boardHeight / 8
colWidth = boardWidth / 8

'Debug.Print Time
smallHeight = (boardHeight - (2 * rowHeight)) / (numInRow - 1)
smallWidth = (boardWidth - (2 * colWidth)) / (numInRow - 1)

'bottom quadrant?
If xc >= (StartX + clipBrd + colWidth) And xc <= endX And yc >= (endY - rowHeight) And yc <= (endY) Then
    'Debug.Print "Bottom!"
    'we must now calculate X coord
    'corner square?
    If xc > (endX - colWidth) Then
        'YES
        eID = 0
        Else
        eID = (numInRow + 1) - ((xc - StartX) / smallWidth)
        End If
    End If

'left qudrant?
If xc >= (StartX + clipBrd) And xc <= (StartX + colWidth + clipBrd) And yc >= (startY + rowHeight) And yc <= (endY) Then
    'Debug.Print "Left!"
    'here we calculate Y
    If yc > (endY - rowHeight) Then
        'YES
        eID = numInRow
        Else
        eID = (numInRow - 1) + ((endY - yc) / smallHeight)
        End If
    End If
    
'top?
If xc >= StartX And xc <= (endX - colWidth) And yc >= startY And xc <= endX And yc <= (startY + rowHeight) Then
    'Debug.Print "Top"
    'we just check X
    'corner?
    If xc < (StartX + colWidth) Then
        eID = numInRow * 2
        Else
        eID = (numInRow * 2) + ((xc - StartX) / smallWidth) - 1
        End If
    End If
    
'right?
If xc >= (endX - colWidth) And xc <= endX And yc >= startY And yc <= (endY - rowHeight) Then
    'Debug.Print "Right!"
        'here we calculate Y
    If yc < (startY + rowHeight) Then
        'YES
        eID = numInRow * 3
        Else
        eID = (numInRow * 3) + ((yc - startY) / smallHeight) - 1
        End If
    End If
GetEstateID = eID
End Function
Private Sub Form_Resize()
'if we drawn board, we redraw it
If canDraw = False Then Exit Sub
'if window is minimised, we also don't draw board
If Me.WindowState = 1 Then Exit Sub
DrawBoard
End Sub

Private Sub Form_Unload(Cancel As Integer)
Debug.Print "Game UNLOAD"

prevPlistLine = ""
pcsLoaded = False
frmEstate.BackColor = Me.BackColor
gameStatus = onHold

frmGames.Show
Unload Me
End Sub



Private Sub lblPName_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'we must remember coordinates for click event
Form_MouseUp Button, Shift, X + lblPName(Index).Left, lblPName(Index).Top + Y
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuConfig_Click()
frmConfig.Show
End Sub

Private Sub mnuExit_Click()
frmGames.kSend ".gx"
Debug.Print "EXIT"
Unload Me
End Sub

Private Sub mnugBankrupt_Click()
frmGames.kSend ".D"
End Sub

Private Sub mnup2rQTrade_Click()
'we request trade by right-clicking on players list
TradeData.tEstateName = ""
TradeData.tInitiator = player.Name
TradeData.tSource = GetPlayerName(mnup2rQTrade.tag)
frmGames.kSend ".Tn" & mnup2rQTrade.tag
End Sub

Private Sub mnuPBuyHotel_Click()
frmGames.kSend ".hb" & CStr(GetEstateID(mLastX, mLastY))
End Sub

Private Sub mnupBuyHouse_Click()
frmGames.kSend ".hb" & CStr(GetEstateID(mLastX, mLastY))
End Sub

Private Sub mnuPDispData_Click()
frmDispEstate.DisplayEstate colestates(GetEstateID(mLastX, mLastY) + 1)
End Sub

Private Sub mnuPMortgage_Click()
frmGames.kSend ".em" & CStr(GetEstateID(mLastX, mLastY))
End Sub

Private Sub mnuPRQTrade_Click()
'we request trade by right-cliking on estate
TradeData.tInitiator = player.Name
TradeData.tMeStart = True
TradeData.tEstateName = colestates(GetEstateID(mLastX, mLastY) + 1).Name
TradeData.tSource = GetPlayerName(colestates(GetEstateID(mLastX, mLastY) + 1).Owner)
frmGames.kSend ".Tn" & mnuPRQTrade.tag
End Sub

Private Sub mnuPSellHotel_Click()
frmGames.kSend ".hs" & CStr(GetEstateID(mLastX, mLastY))
End Sub

Private Sub mnuPSellHouse_Click()
frmGames.kSend ".hs" & CStr(GetEstateID(mLastX, mLastY))
End Sub

Private Sub mnupSellTobank_Click()
frmGames.kSend ".es" & CStr(GetEstateID(mLastX, mLastY))
End Sub

Private Sub mnuPUnmortgage_Click()
frmGames.kSend ".em" & CStr(GetEstateID(mLastX, mLastY))
End Sub

Private Sub mnuSend_Click()
'this was used during development period
frmGames.kSend InputBox("String to send?")
End Sub

Private Sub pict1House_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'board must send click event even if any of pictures has ben clicked...
Form_MouseUp Button, Shift, pict1House(Index).Left + X, pict1House(Index).Top + Y
End Sub

Private Sub pict2House_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'board must send click event even if any of pictures has ben clicked...
Form_MouseUp Button, Shift, pict2House(Index).Left + X, pict2House(Index).Top + Y
End Sub

Private Sub pict3house_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'board must send click event even if any of pictures has ben clicked...
Form_MouseUp Button, Shift, pict3house(Index).Left + X, pict3house(Index).Top + Y
End Sub

Private Sub pict4house_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'board must send click event even if any of pictures has ben clicked...
Form_MouseUp Button, Shift, pict4house(Index).Left + X, pict4house(Index).Top + Y
End Sub

Private Sub pictCard_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'board must send click event even if any of pictures has ben clicked...
Form_MouseUp Button, Shift, pictCard(Index).Left + X, pictCard(Index).Top + Y
End Sub

Private Sub pictHotel_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'board must send click event even if any of pictures has ben clicked...
Form_MouseUp Button, Shift, pictHotel(Index).Left + X, pictHotel(Index).Top + Y
End Sub

Private Sub pictMortaged_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'board must send click event even if any of pictures has ben clicked...
Form_MouseUp Button, Shift, pictMortaged(Index).Left + X, pictMortaged(Index).Top + Y
End Sub

Private Sub pIece_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'board must send click event even if any of pictures has ben clicked...
Form_MouseUp Button, Shift, pIece(Index).Left + X, pIece(Index).Top + Y
End Sub

Private Sub tpause_Timer()
'for pausing after we move token
tPause.Enabled = False
doWait = False
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)
'send tekst on chat window
Dim sText As String

sText = txtChat.text

If KeyAscii = 13 Then
    sText = Replace(sText, "è", "Ä")
    sText = Replace(sText, "È", "ÄŒ")
    sText = Replace(sText, "š", "Å¡")
    sText = Replace(sText, "Š", "Å ")
    sText = Replace(sText, "ž", "Å¾")
    sText = Replace(sText, "Ž", "Å½")
    frmGames.kSend sText
    txtChat.text = ""
    End If
End Sub

Public Sub DrawBoard()
Debug.Print "DRAWBOARD"
Me.Cls
'in each line we draw 9 small boxes and one big in each quadrant
Dim boardWidth As Long, boardHeight As Long

Dim StartX As Long, startY As Long
Dim endX As Long, endY As Long
Const clipBrd = 20
Dim rowHeight As Long
Dim colWidth As Long
Dim n As Integer
Dim smallWidth As Long
Dim smallHeight As Long
Dim estate As cEstate
Dim numInRow As Integer

numInRow = colestates.Count / 4

StartX = txtChat.Width + clipBrd + txtChat.Left
startY = clipBrd + lstPlayers.Top

endX = Me.Width - clipBrd - 130
endY = Me.Height - clipBrd - sBar.Height - 690

boardHeight = endY - startY
boardWidth = endX - StartX


txtChat.Top = Height - sBar.Height - txtChat.Height - 700

txtChatList.Top = conPlayer(conPlayer.Count - 1).Top + conPlayer(conPlayer.Count - 1).Height + 100
txtChatList.Height = txtChat.Top - txtChatList.Top - 100

'we'll take that height is 1/6th of board height...
'and width of column is 1/6th of board width...

rowHeight = boardHeight / 8
colWidth = boardWidth / 8



If frmEstate.Visible = False Then frmEstate.Visible = True
frmEstate.Left = StartX + colWidth + clipBrd
frmEstate.Top = startY + rowHeight + clipBrd
frmEstate.Width = boardWidth - (2 * colWidth) - (2 * clipBrd)
frmEstate.Height = boardHeight - (2 * rowHeight) - (2 * clipBrd)

lblEstateName.Left = 1
lblEstateName.Top = 1
lblEstateName.Width = frmEstate.Width - 1
lblEstateGroup.Top = (lblEstateName.Top + lblEstateName.Height) - lblEstateGroup.Height
lblEstateGroup.Left = (lblEstateName.Left + lblEstateName.Width) - lblEstateGroup.Width

lstMsgs.Top = lblEstateName.Height + 2
lstMsgs.Width = frmEstate.Width * 0.9
lstMsgs.Left = frmEstate.Width / 2 - lstMsgs.Width / 2


lstMsgs.Height = frmEstate.Height - lblEstateName.Height - cmdAuction.Height - 300

'also we set coords for buttons and images

lblEstateName.Left = 1
lblEstateName.Top = 1
lblEstateName.Width = frmEstate.Width - 1
lblEstateName.Height = lstMsgs.Top - 2

icoCrd.Left = lblEstateName.Left
icoCrd.Top = lblEstateName.Top

icoMortgaged.Top = icoCrd.Top
icoMortgaged.Left = lblEstateName.Left + icoCrd.Width


ico1House.Left = lblEstateName.Left
ico1House.Top = (lblEstateName.Top + lblEstateName.Height) - ico1House.Height

ico2House.Left = lblEstateName.Left
ico2House.Top = (lblEstateName.Top + lblEstateName.Height) - ico1House.Height

ico3House.Left = lblEstateName.Left
ico3House.Top = (lblEstateName.Top + lblEstateName.Height) - ico1House.Height

ico4house.Left = lblEstateName.Left
ico4house.Top = (lblEstateName.Top + lblEstateName.Height) - ico1House.Height

icoHotel.Left = lblEstateName.Left
icoHotel.Top = (lblEstateName.Top + lblEstateName.Height) - ico1House.Height


cmdAuction.Top = frmEstate.Height - cmdAuction.Height - 50
cmdAuction.Left = 50

cmdBuy.Top = cmdAuction.Top
cmdBuy.Left = cmdAuction.Left + 50 + cmdAuction.Width

cmdEndturn.Top = cmdAuction.Top
cmdEndturn.Left = cmdBuy.Left + 50 + cmdBuy.Width

cmdRoll.Top = frmEstate.Height - cmdRoll.Height - 50
cmdRoll.Left = cmdEndturn.Left + cmdEndturn.Width + 200


cmdPayStatic.Top = cmdAuction.Top
cmdPayStatic.Left = 50

cmdPayPercentage.Top = cmdAuction.Top
cmdPayPercentage.Left = 50 + cmdPayStatic.Width + cmdPayStatic.Left

cmdJailPay.Left = 50
cmdJailPay.Top = cmdAuction.Top

cmdJailUseCard.Top = cmdAuction.Top
cmdJailUseCard.Left = cmdJailPay.Left + 50 + cmdJailPay.Width

cmdJailRoll.Top = cmdAuction.Top
cmdJailRoll.Left = cmdJailUseCard.Left + 50 + cmdJailUseCard.Width

'bottom qudrant
For n = 1 To numInRow

Select Case n
Case 1
Set estate = colestates(n)
DrawCornerBox endX - colWidth, endY - rowHeight, colWidth, rowHeight, estate.BgColor, estate.Mortaged
pictMortaged(n - 1).Top = endY - pictMortaged(0).Height
pictMortaged(n - 1).Left = endX - colWidth
pictMortaged(n - 1).Visible = estate.Mortaged

pictCard(n - 1).Top = pictMortaged(n - 1).Top - pictCard(n - 1).Height
pictCard(n - 1).Left = pictMortaged(n - 1).Left
pictCard(n - 1).Visible = (estate.Owner = "-1" And estate.CanBeOwned = True)

pictHotel(n - 1).Left = pictMortaged(n - 1).Left + pictMortaged(0).Width
pictHotel(n - 1).Top = endY - pictHotel(0).Height

pict1House(n - 1).Top = pictHotel(n - 1).Top
pict1House(n - 1).Left = pictHotel(n - 1).Left

pict2House(n - 1).Top = pictHotel(n - 1).Top
pict2House(n - 1).Left = pictHotel(n - 1).Left

pict3house(n - 1).Top = pictHotel(n - 1).Top
pict3house(n - 1).Left = pictHotel(n - 1).Left

pict4house(n - 1).Top = pictHotel(n - 1).Top
pict4house(n - 1).Left = pictHotel(n - 1).Left

'DrawBox startX, startY, colWidth, rowHeight
Case Else
'usual box
smallWidth = (boardWidth - (2 * colWidth)) / (numInRow - 1)
'we get colection of estates
Set estate = colestates(n)
If estate.Group <> "-1" Then
    DrawBox endX - colWidth - (smallWidth * (n - 1)), endY - rowHeight, smallWidth, rowHeight, oup, estate.BgColor, estate.Color, estate.Mortaged
    Else
    DrawBox endX - colWidth - (smallWidth * (n - 1)), endY - rowHeight, smallWidth, rowHeight, oup, estate.BgColor, , estate.Mortaged
    End If

pictMortaged(n - 1).Top = endY - pictMortaged(0).Height
pictMortaged(n - 1).Left = endX - colWidth - (smallWidth * (n - 1))
pictMortaged(n - 1).Visible = estate.Mortaged

pictCard(n - 1).Top = pictMortaged(n - 1).Top - pictCard(n - 1).Height
pictCard(n - 1).Left = pictMortaged(n - 1).Left
pictCard(n - 1).Visible = (estate.Owner = "-1" And estate.CanBeOwned = True)


pictHotel(n - 1).Left = pictMortaged(n - 1).Left + pictMortaged(0).Width
pictHotel(n - 1).Top = endY - pictHotel(0).Height

pict1House(n - 1).Top = pictHotel(n - 1).Top
pict1House(n - 1).Left = pictHotel(n - 1).Left

pict2House(n - 1).Top = pictHotel(n - 1).Top
pict2House(n - 1).Left = pictHotel(n - 1).Left

pict3house(n - 1).Top = pictHotel(n - 1).Top
pict3house(n - 1).Left = pictHotel(n - 1).Left

pict4house(n - 1).Top = pictHotel(n - 1).Top
pict4house(n - 1).Left = pictHotel(n - 1).Left

End Select
Next

'left quadrant
For n = 1 To numInRow
Select Case n

Case 1
Set estate = colestates(n + 10)
DrawCornerBox StartX, endY - rowHeight, colWidth, rowHeight, estate.BgColor, estate.Mortaged

pictMortaged(n + 9).Top = endY - pictMortaged(0).Height
pictMortaged(n + 9).Left = StartX
pictMortaged(n + 9).Visible = estate.Mortaged

pictCard(n + 9).Top = pictMortaged(n + 9).Top
pictCard(n + 9).Left = pictMortaged(n + 9).Left + pictMortaged(0).Width
pictCard(n + 9).Visible = (estate.Owner = "-1" And estate.CanBeOwned = True)

pictHotel(n + 9).Left = pictMortaged(n + 9).Left
pictHotel(n + 9).Top = pictMortaged(n + 9).Top + pictMortaged(0).Height

pict1House(n + 9).Top = pictHotel(n + 9).Top
pict1House(n + 9).Left = pictHotel(n + 9).Left

pict2House(n + 9).Top = pictHotel(n + 9).Top
pict2House(n + 9).Left = pictHotel(n + 9).Left

pict3house(n + 9).Top = pictHotel(n + 9).Top
pict3house(n + 9).Left = pictHotel(n + 9).Left

pict4house(n + 9).Top = pictHotel(n + 9).Top
pict4house(n + 9).Left = pictHotel(n + 9).Left


Case Else
'we get colection of estates

'usual box
'we get colection of estates
Set estate = colestates(n + 10)
smallHeight = (boardHeight - (2 * rowHeight)) / (numInRow - 1)
If estate.Group <> "-1" Then
    DrawBox StartX, endY - rowHeight - ((n - 1) * smallHeight), colWidth, smallHeight, oRight, estate.BgColor, estate.Color, estate.Mortaged
    Else
    DrawBox StartX, endY - rowHeight - ((n - 1) * smallHeight), colWidth, smallHeight, oRight, estate.BgColor, , estate.Mortaged
    End If

pictMortaged(n + 9).Top = endY - rowHeight - ((n - 1) * smallHeight)
pictMortaged(n + 9).Left = StartX
pictMortaged(n + 9).Visible = estate.Mortaged

pictCard(n + 9).Top = pictMortaged(n + 9).Top
pictCard(n + 9).Left = pictMortaged(n + 9).Left + pictMortaged(0).Width
pictCard(n + 9).Visible = (estate.Owner = "-1" And estate.CanBeOwned = True)

pictHotel(n + 9).Left = pictMortaged(n + 9).Left
pictHotel(n + 9).Top = pictMortaged(n + 9).Top + pictMortaged(0).Height

pict1House(n + 9).Top = pictHotel(n + 9).Top
pict1House(n + 9).Left = pictHotel(n + 9).Left

pict2House(n + 9).Top = pictHotel(n + 9).Top
pict2House(n + 9).Left = pictHotel(n + 9).Left

pict3house(n + 9).Top = pictHotel(n + 9).Top
pict3house(n + 9).Left = pictHotel(n + 9).Left

pict4house(n + 9).Top = pictHotel(n + 9).Top
pict4house(n + 9).Left = pictHotel(n + 9).Left

End Select
Next

'top quadrant
For n = 1 To numInRow
Select Case n

Case 1
Set estate = colestates(n + 20)
DrawCornerBox StartX, startY, colWidth, rowHeight, estate.BgColor, estate.Mortaged
pictMortaged(n + 19).Top = startY
pictMortaged(n + 19).Left = StartX
pictMortaged(n + 19).Visible = estate.Mortaged

pictCard(n + 19).Top = pictMortaged(n + 19).Top - pictCard(0).Height
pictCard(n + 19).Left = pictMortaged(n + 19).Left
pictCard(n + 19).Visible = (estate.Owner = "-1" And estate.CanBeOwned = True)

pictHotel(n + 19).Left = pictMortaged(n + 19).Left + pictMortaged(0).Width
pictHotel(n + 19).Top = pictMortaged(n + 19).Top

pict1House(n + 19).Top = pictHotel(n + 19).Top
pict1House(n + 19).Left = pictHotel(n + 19).Left

pict2House(n + 19).Top = pictHotel(n + 19).Top
pict2House(n + 19).Left = pictHotel(n + 19).Left

pict3house(n + 19).Top = pictHotel(n + 19).Top
pict3house(n + 19).Left = pictHotel(n + 19).Left

pict4house(n + 19).Top = pictHotel(n + 19).Top
pict4house(n + 19).Left = pictHotel(n + 19).Left

Case Else
'usual box
smallWidth = (boardWidth - (2 * colWidth)) / (numInRow - 1)
Set estate = colestates(n + 20)
If estate.Group <> "-1" Then
    DrawBox (StartX + colWidth) + (smallWidth * (n - 2)), startY, smallWidth, rowHeight, odown, estate.BgColor, estate.Color, estate.Mortaged
    Else
    DrawBox StartX + colWidth + (smallWidth * (n - 2)), startY, smallWidth, rowHeight, odown, estate.BgColor, , estate.Mortaged
    End If
pictMortaged(n + 19).Top = startY
pictMortaged(n + 19).Left = StartX + colWidth + (smallWidth * (n - 2))
pictMortaged(n + 19).Visible = estate.Mortaged

pictCard(n + 19).Top = pictMortaged(n + 19).Top + pictCard(0).Height
pictCard(n + 19).Left = pictMortaged(n + 19).Left
pictCard(n + 19).Visible = (estate.Owner = "-1" And estate.CanBeOwned = True)

pictHotel(n + 19).Left = pictMortaged(n + 19).Left + pictMortaged(0).Width
pictHotel(n + 19).Top = pictMortaged(n + 19).Top

pict1House(n + 19).Top = pictHotel(n + 19).Top
pict1House(n + 19).Left = pictHotel(n + 19).Left

pict2House(n + 19).Top = pictHotel(n + 19).Top
pict2House(n + 19).Left = pictHotel(n + 19).Left

pict3house(n + 19).Top = pictHotel(n + 19).Top
pict3house(n + 19).Left = pictHotel(n + 19).Left

pict4house(n + 19).Top = pictHotel(n + 19).Top
pict4house(n + 19).Left = pictHotel(n + 19).Left


End Select
Next

'right quadrant
For n = 1 To numInRow
Select Case n

Case 1
Set estate = colestates(n + 30)
DrawCornerBox endX - colWidth, startY, colWidth, rowHeight, estate.BgColor, estate.Mortaged
pictMortaged(n + 29).Top = startY
pictMortaged(n + 29).Left = endX - pictMortaged(0).Width
pictMortaged(n + 29).Visible = estate.Mortaged

pictCard(n + 29).Top = pictMortaged(n + 29).Top
pictCard(n + 29).Left = pictMortaged(n + 29).Left - pictMortaged(0).Width
pictCard(n + 29).Visible = (estate.Owner = "-1" And estate.CanBeOwned = True)

pictHotel(n + 29).Left = endX - pictHotel(0).Width
pictHotel(n + 29).Top = pictMortaged(n + 29).Top + pictMortaged(0).Height

pict1House(n + 29).Top = pictHotel(n + 29).Top
pict1House(n + 29).Left = pictHotel(n + 29).Left

pict2House(n + 29).Top = pictHotel(n + 29).Top
pict2House(n + 29).Left = pictHotel(n + 29).Left

pict3house(n + 29).Top = pictHotel(n + 29).Top
pict3house(n + 29).Left = pictHotel(n + 29).Left

pict4house(n + 29).Top = pictHotel(n + 29).Top
pict4house(n + 29).Left = pictHotel(n + 29).Left


Case Else
'usual box
smallHeight = (boardHeight - (2 * rowHeight)) / (numInRow - 1)
Set estate = colestates(n + 30)
If estate.Group <> "-1" Then
    DrawBox endX - colWidth, startY + rowHeight + ((n - 2) * smallHeight), colWidth, smallHeight, oLeft, estate.BgColor, estate.Color, estate.Mortaged
    Else
    DrawBox endX - colWidth, startY + rowHeight + ((n - 2) * smallHeight), colWidth, smallHeight, oLeft, estate.BgColor, , estate.Mortaged
    End If
pictMortaged(n + 29).Top = startY + rowHeight + ((n - 2) * smallHeight)
pictMortaged(n + 29).Left = endX - pictMortaged(0).Width
pictMortaged(n + 29).Visible = estate.Mortaged

pictCard(n + 29).Top = pictMortaged(n + 29).Top
pictCard(n + 29).Left = pictMortaged(n + 29).Left - pictMortaged(0).Width
pictCard(n + 29).Visible = (estate.Owner = "-1" And estate.CanBeOwned = True)

pictHotel(n + 29).Left = endX - pictHotel(0).Width
pictHotel(n + 29).Top = pictMortaged(n + 29).Top + pictMortaged(0).Height

pict1House(n + 29).Top = pictHotel(n + 29).Top
pict1House(n + 29).Left = pictHotel(n + 29).Left

pict2House(n + 29).Top = pictHotel(n + 29).Top
pict2House(n + 29).Left = pictHotel(n + 29).Left

pict3house(n + 29).Top = pictHotel(n + 29).Top
pict3house(n + 29).Left = pictHotel(n + 29).Left

pict4house(n + 29).Top = pictHotel(n + 29).Top
pict4house(n + 29).Left = pictHotel(n + 29).Left


End Select
Next

'and pieces
For n = 0 To pIece.Count - 1
If pIece(n).Visible = True Then putPiece n + 1, pcsPos(n)
Next

End Sub

Private Sub DrawBox(sX As Long, sY As Long, Width As Long, Height As Long, Orientation As tOrientation, Optional BackColor As Long = -1, Optional Color As Long = -1, Optional Mortgaged As Boolean = False)
'first we draw color label

'we take that label is 1/8th of width/height of field
Dim sW As Long
Dim sH As Long

'background
If BackColor <> -1 Then
    Me.Line (sX, sY)-Step(Width, Height), BackColor, BF
    End If




If Color <> -1 Then
Select Case Orientation
    Case oup
    sH = Height / 6
    Me.Line (sX, sY)-Step(Width, sH), Color, BF
    
    Case odown
    sH = Height / 6
    Me.Line (sX, (sY + Height) - sH)-Step(Width, sH), Color, BF
    
    Case oLeft
    sW = Width / 6
    Me.Line (sX, sY)-Step(sW, Height), Color, BF
    
    Case oRight
    sW = Width / 6
    Me.Line ((sX + Width) - sW, sY)-Step(sW, Height), Color, BF
    End Select
    End If


Me.Line (sX, sY)-Step(Width, Height), , B
End Sub

Private Sub DrawCornerBox(sX As Long, sY As Long, Width As Long, Height As Long, Optional BgColor As Long = -1, Optional Mortgaged As Boolean = False)
If BgColor <> -1 Then
    Me.Line (sX, sY)-Step(Width, Height), BgColor, BF
    End If

Me.Line (sX, sY)-Step(Width, Height), , B
End Sub

Public Sub MovePlayer(PlayerID As String, location As Integer, direct As Boolean)
'we move token and we confirm we moved it


'Debug.Print "Loc:" & location, "PID:" & PlayerID

'If location = -1 Then
'    tPause.Interval = 500
'    doWait = True
'    tPause.Enabled = True
'    Do
'    DoEvents
'    Loop While doWait = True
    
'    frmGames.kSend ".t-1"
'    DoEvents
'    Exit Sub
'    End If

Dim n As Integer
'Debug.Print "ID:" & playerID & ":" & location

For n = 1 To pIece.Count
DoEvents
If pIece(n - 1).tag = CStr(PlayerID) Then
    'If direct = True Then
        putPiece n, location
        pcsPos(n - 1) = location
        tPause.Interval = 500
        doWait = True
        tPause.Enabled = True
        Do
        DoEvents
        Loop While doWait = True
        If pIece(n - 1).Visible = False Then
            pIece(n - 1).Visible = True
            lblPName(n - 1).Visible = True
            frmGames.kSend ".t-1"
            Else
            frmGames.kSend ".t" & CStr(location)
            End If
'        frmGames.kSend ".t" & CStr(location)
        Exit For
        End If
Next

End Sub

Private Sub putPiece(ID As Integer, location As Integer)
'put piece on place..
Dim boardWidth As Long, boardHeight As Long
Dim StartX As Long, startY As Long
Dim endX As Long, endY As Long
Const clipBrd = 20
Dim rowHeight As Long
Dim colWidth As Long
Dim smallWidth As Long
Dim smallHeight As Long
Dim numInRow As Integer

numInRow = colestates.Count / 4

StartX = txtChat.Width + clipBrd + txtChat.Left
startY = clipBrd + lstPlayers.Top

endX = Me.Width - clipBrd - 130
endY = Me.Height - clipBrd - sBar.Height - 690

boardHeight = endY - startY
boardWidth = endX - StartX

'we'll take that height is 1/6th of board height...
'and width of column is 1/6th of board width...

rowHeight = boardHeight / 8
colWidth = boardWidth / 8

Select Case location
'bottom
Case 0
'endX - colWidth, endY - rowHeight, colWidth, rowHeight
pIece(ID - 1).Left = endX - (colWidth / 2) - (pIece(0).Width / 2)
pIece(ID - 1).Top = endY - pIece(0).Height


Case 1 To (numInRow - 1)
smallWidth = (boardWidth - (2 * colWidth)) / 9
'endX - colWidth - (smallWidth * (n - 1)), endY - rowHeight, smallWidth, rowHeight
pIece(ID - 1).Left = ((endX - colWidth - (smallWidth * location)) + colWidth / 2) - pIece(0).Width / 2
pIece(ID - 1).Top = endY - pIece(0).Height
DoEvents

Case numInRow
'left
'startX, endY - rowHeight, colWidth, rowHeight, estate.BgColor
pIece(ID - 1).Left = (StartX + colWidth / 2) - (pIece(0).Width / 2)
pIece(ID - 1).Top = endY - pIece(0).Height

Case (numInRow + 1) To numInRow + (numInRow - 1)
smallHeight = (boardHeight - (2 * rowHeight)) / 9
'startX , endY - rowHeight - ((n - 1) * smallHeight), colWidth, smallHeight
pIece(ID - 1).Left = (StartX + colWidth / 2) - pIece(0).Width / 2
pIece(ID - 1).Top = endY - ((smallHeight * (location - 10)) + smallHeight / 2) - pIece(0).Height


Case numInRow * 2
'up
'startX, startY, colWidth, rowHeight
pIece(ID - 1).Left = (StartX + colWidth / 2) - pIece(0).Width / 2
pIece(ID - 1).Top = (startY + rowHeight / 2) - pIece(0).Height / 2

'Case 21 To 29
Case ((numInRow * 2) + 1) To (numInRow * 2) + ((numInRow) - 1)
smallWidth = (boardWidth - (2 * colWidth)) / 9
'(startX + colWidth) + (smallWidth * (n - 2)), startY, smallWidth, rowHeight
pIece(ID - 1).Left = (StartX + colWidth + (smallWidth * (location - 21)) + colWidth / 2) - pIece(0).Width / 2
pIece(ID - 1).Top = (startY + rowHeight / 2) - pIece(0).Height / 2

'Case 30
Case (3 * numInRow)
'right
'endX - colWidth, startY, colWidth, rowHeight
pIece(ID - 1).Left = (endX - colWidth / 2) - pIece(0).Width / 2
pIece(ID - 1).Top = (startY + rowHeight / 2) - pIece(0).Height / 2

'Case 31 To 39
Case ((numInRow * 3) + 1) To (numInRow * 3) + (numInRow - 1)
smallHeight = (boardHeight - (2 * rowHeight)) / 9
'endX -colWidth, startY + rowHeight + ((n - 2) * smallHeight), colWidth, smallHeight
pIece(ID - 1).Left = (endX - colWidth / 2) - pIece(0).Width / 2
pIece(ID - 1).Top = (startY + rowHeight + (smallHeight * (location - 31))) + (smallHeight / 2 - pIece(0).Height / 2)


End Select

lblPName(ID - 1).Top = pIece(ID - 1).Top + pIece(0).Height - lblPName(ID - 1).Height
lblPName(ID - 1).Left = (pIece(ID - 1).Left + (pIece(0).Width / 2)) - (lblPName(ID - 1).Width / 2)


DoEvents
End Sub

Public Sub Mortgage(Mortgaged As Boolean, EstateID As Integer)
'if estate is mortaged, we must display/hide appropriate picture

If canDraw = False Then Exit Sub

Dim n As Integer
Dim estate As cEstate


Set estate = colestates(EstateID + 1)
pictMortaged(EstateID).Visible = estate.Mortaged

End Sub

Public Sub EstateBought(EstateID As Integer, Free As Boolean)
'if estate is being bought, we must hide appropriate image
If canDraw = False Then Exit Sub

Dim estate As cEstate

Set estate = colestates(EstateID + 1)
If estate.CanBeOwned = True Then pictCard(EstateID).Visible = Free

End Sub


Public Sub HouseBought(EstateID As Integer, numHouses As Integer)
'if house is being bought we must display appropiate number of houses
If canDraw = False Then Exit Sub

Dim estate As cEstate
pict1House(EstateID).Visible = False
pict2House(EstateID).Visible = False
pict3house(EstateID).Visible = False
pict4house(EstateID).Visible = False
pictHotel(EstateID).Visible = False

Select Case numHouses
Case 1
pict1House(EstateID).Visible = True

Case 2
pict2House(EstateID).Visible = True

Case 3
pict3house(EstateID).Visible = True

Case 4
pict4house(EstateID).Visible = True

Case 5
pictHotel(EstateID).Visible = True

End Select

End Sub
