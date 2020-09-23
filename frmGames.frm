VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmGames 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Games: double-click game to join/initiate"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   Icon            =   "frmGames.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tpause 
      Left            =   2760
      Top             =   2280
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Servers..."
      Height          =   615
      Left            =   120
      Picture         =   "frmGames.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Width           =   1335
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1680
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGames.frx":0B14
            Key             =   "game"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGames.frx":13EE
            Key             =   "user"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh..."
      Height          =   615
      Left            =   5280
      Picture         =   "frmGames.frx":1CC8
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   1335
   End
   Begin MSComctlLib.ListView lstGames 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   7223
      View            =   3
      Sorted          =   -1  'True
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
         Text            =   "Game"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Players"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Connecting"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   4440
      Width           =   1935
   End
End
Attribute VB_Name = "frmGames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim playerSet As Boolean
Dim gmList As Boolean
Dim doWait  As Boolean
Dim prevGListLine As String
Dim prevPData As String
Dim prevPdata2 As String
Dim prevMsg As Boolean
Dim prevGStatus As String
Dim dbg As Boolean
Dim msgQueue As String
Dim ToProcess As Boolean
Dim auctionStatus As tStatus
Dim isConnected As Boolean
Dim b As String
Public WithEvents mklient As CSocket
Attribute mklient.VB_VarHelpID = -1

Private Sub cmdback_Click()
'we go back to servers form

If isConnected = True Then kSend ".d"
isConnected = False

prevGListLine = ""
prevPlistLine = ""
prevPData = ""
prevPdata2 = ""

Unload Me
End Sub

Private Sub cmdRefresh_Click()
'reload games list
kSend ".gl"
End Sub

Private Sub Form_Load()
Debug.Print "Games LOAD"
dbg = False
Set mklient = New CSocket
mklient.RemoteHost = serverName
mklient.RemotePort = CInt(serverPort)
mklient.Protocol = sckTCPProtocol
auctionStatus = AuctNoAuction
mklient.Connect
playerSet = False
centerForm Me
level = 0

ReadConfig


gmList = False
ToProcess = False
prevMsg = False
End Sub

Public Sub kSend(hLine As String)
'here we send data on server
mklient.SendData hLine & vbLf
End Sub

Public Sub LoadGames(hLine As String)

Dim p As Integer, q As Integer
Dim game As String
Dim gType As String

'Debug.Print "GL:" & hLine

'we recieved commad for games list update
gType = GetData("type=" & Chr(34), hLine)

Select Case gType

Case "full"
lstGames.ListItems.Clear
Do
p = InStr(UCase(hLine), "<GAME")
If p = 0 Then Exit Sub

q = InStr(p + 5, hLine, "/>")
If q = 0 Then Exit Sub
q = q + 2

game = mID(hLine, p, q - p)
hLine = mID(hLine, q)

addGame game
Loop

Case "add"
If frmGames.Visible = True Then Beep
addGame hLine

Case "del"
DeleteGame hLine

Case "edit"
Editgame hLine

End Select

End Sub
Private Sub Editgame(hLine As String)
'change game properties
Dim indx As String, n As Integer
Dim p As String

indx = GetData("id=" & Chr(34), hLine)
For n = 1 To lstGames.ListItems.Count
p = InStr(lstGames.ListItems(n).Key, ":")



If indx = mID(lstGames.ListItems(n).Key, p + 1) Then
    lstGames.ListItems(n).SubItems(2) = GetData("players=" & Chr(34), hLine)
    Exit For
    End If

Next


End Sub

Private Sub DeleteGame(hLine As String)
'remove game from list
Dim indx As String, n As Integer
Dim p As String

indx = GetData("id=" & Chr(34), hLine)
For n = 1 To lstGames.ListItems.Count
p = InStr(lstGames.ListItems(n).Key, ":")

If indx = mID(lstGames.ListItems(n).Key, p + 1) Then
    lstGames.ListItems.Remove n
    Exit For
    End If

Next

End Sub

Private Sub addGame(hLine As String)
'add game on list
Dim p As Integer
Dim q As Integer
Dim ID As String
Dim Name As String
Dim gType As String
Dim comment As String
Dim item As ListItem
'ID
p = InStr(UCase(hLine), "ID=")
If p = 0 Then Exit Sub
p = p + 4

q = InStr(p + 1, hLine, Chr(34))
ID = mID(hLine, p, q - p)

'TYPE
p = InStr(UCase(hLine), "TYPE=")
If p = 0 Then Exit Sub
p = p + 6
q = InStr(p + 1, hLine, Chr(34))
gType = mID(hLine, p, q - p)

'Name
p = InStr(UCase(hLine), "NAME=")
If p = 0 Then Exit Sub
p = p + 6
q = InStr(p + 1, hLine, Chr(34))
Name = mID(hLine, p, q - p)

'Comment
p = InStr(UCase(hLine), "DESCRIPTION=")
If p = 0 Then Exit Sub
p = p + 13
q = InStr(p + 1, hLine, Chr(34))
comment = mID(hLine, p, q - p)

If ID = "-1" Then
    Set item = lstGames.ListItems.Add(, gType & "|" & lstGames.ListItems.Count & ":" & ID, "Create new " & Name, , "game")
    Else
    Set item = lstGames.ListItems.Add(, gType & "|" & lstGames.ListItems.Count & ":" & ID, "Join " & Name, , "user")
    End If
item.SubItems(1) = comment
item.SubItems(2) = GetData("players=" & Chr(34), hLine)

End Sub


Private Sub Form_Unload(Cancel As Integer)
Debug.Print "Games UNLOAD"
Set colplayers = Nothing
Set colEstateGroups = Nothing
Set colestates = Nothing
Set colTradeForms = Nothing

mklient.CloseSocket
Set mklient = Nothing

player.Cookie = ""
player.ID = ""

If command = "" Then frmServers.Show

End Sub

Private Sub lstGames_DblClick()
'user double-clicked game. let's join / create it
Dim p As Integer
Dim gid As String
Dim gType As String

p = InStr(lstGames.SelectedItem.Key, ":")

gid = mID(lstGames.SelectedItem.Key, p + 1)

p = InStr(lstGames.SelectedItem.Key, "|")
gType = Left(lstGames.SelectedItem.Key, p - 1)

If gid <> "-1" Then
    kSend ".gj" & gid
    'frmgame.cmdGo.Visible = False
    'frmgame.cmdback.Visible = True
    Else
    kSend ".gn" & gType
    'frmgame.cmdGo.Visible = True
    'frmgame.cmdback.Visible = True
    End If
    
End Sub

Private Sub mklient_OnConnect()
Debug.Print "Connected"
lblStatus = "Connected!"
Caption = "Games on " & serverName & ": double-click game to join/initiate"
DoEvents
isConnected = True
End Sub

Private Sub mklient_OnDataArrival(ByVal bytesTotal As Long)
Dim s As String
'here we recieve data from server
'it is implemented trough queue
'for their definitions go and see http://unixcode.org

mklient.GetData s, vbString, bytesTotal

msgQueue = msgQueue & s


If prevMsg = True Then
    ToProcess = True
    Exit Sub
    End If

processMessage msgQueue

End Sub

Private Sub processMessage(ByVal Mmsg As String)
'we process messages from server.
procagain:
'if there are messages in queue... process them
If ToProcess = True Then
    Mmsg = msgQueue
    msgQueue = ""
    ToProcess = False
    Else
    msgQueue = ""
    End If

prevMsg = True
'proces messages sent by klient control

Dim p As Long, q As Long, n As Long
Dim tag As String, untag As String
Dim pMsg As String




'Debug.Print "RCVD:" & Mmsg

p = InStr(Mmsg, "<monopd>")
If p = 0 Then GoTo finish

q = InStr(Mmsg, "</monopd>")
If q = 0 Then GoTo finish

Mmsg = Replace(Mmsg, "<monopd>", "")
Mmsg = Replace(Mmsg, "</monopd>", "")
Mmsg = Replace(Mmsg, vbLf, "")

'If dbg = True Then Debug.Print Mmsg

Do

Mmsg = Trim(Mmsg)

'Debug.Print "S:" & Mmsg


'let's find tag
p = InStr(Mmsg, ">")
If p = 0 Then Exit Do
tag = mID(Mmsg, 2, p - 2)

p = InStr(tag, " ")
If p > 0 Then tag = Left(tag, p - 1)


'Debug.Print "M:" & Mmsg


'Debug.Print tag

Select Case tag

Case "server"
untag = "/>"
p = InStr(Mmsg, untag)
If p = 0 Then Exit Do
pMsg = Left(Mmsg, p + Len(untag))
Mmsg = mID(Mmsg, p + Len(untag))
server.Name = GetData("name=" & Chr(34), pMsg)
server.version = GetData("version=" & Chr(34), pMsg)


Case "client"
untag = "/>"
p = InStr(Mmsg, untag)
If p = 0 Then Exit Do
pMsg = Left(Mmsg, p + Len(untag))
Mmsg = mID(Mmsg, p + Len(untag))
If player.Cookie = "" Then player.Cookie = GetData("cookie=" & Chr(34), pMsg)
If player.ID = "" Then player.ID = GetData("playerid=" & Chr(34), pMsg)

Case "playerupdate"
untag = "/>"
p = InStr(Mmsg, untag)
If p = 0 Then Exit Do
pMsg = Left(Mmsg, p + (Len(untag) - 1))
Mmsg = mID(Mmsg, p + Len(untag))
getPlayerData pMsg

Case "updateplayerlist"
untag = "</updateplayerlist>"
p = InStr(Mmsg, untag)
If p = 0 Then Exit Do
pMsg = Left(Mmsg, p + (Len(untag) - 1))
Mmsg = mID(Mmsg, p + Len(untag))
UpdatePlayerList pMsg

Case "configupdate"
untag = "</configupdate>"
p = InStr(Mmsg, untag)
If p = 0 Then Exit Do
pMsg = Left(Mmsg, p + (Len(untag) - 1))
Mmsg = mID(Mmsg, p + Len(untag))
ConfigUpdate pMsg

Case "deleteplayer"
untag = "/>"
p = InStr(Mmsg, untag)
If p = 0 Then Exit Do
pMsg = Left(Mmsg, p + (Len(untag) - 1))
Mmsg = mID(Mmsg, p + Len(untag))
DeletePlayer pMsg

Case "msg"
untag = "/>"
p = InStr(Mmsg, untag)
If p = 0 Then Exit Do
pMsg = Left(Mmsg, p + (Len(untag) - 1))
Mmsg = mID(Mmsg, p + Len(untag))
AddMessage pMsg


Case "updategamelist"
untag = "</updategamelist>"
p = InStr(Mmsg, untag)
If p = 0 Then Exit Do
pMsg = Left(Mmsg, p + (Len(untag) - 1))
Mmsg = mID(Mmsg, p + Len(untag))
LoadGames pMsg

Case "gameupdate"
untag = "/>"
p = InStr(Mmsg, untag)
If p = 0 Then Exit Do
pMsg = Left(Mmsg, p + (Len(untag) - 1))
Mmsg = mID(Mmsg, p + Len(untag))
gameUpdate pMsg

Case "estategroupupdate"
untag = "/>"
p = InStr(Mmsg, untag)
If p = 0 Then Exit Do
pMsg = Left(Mmsg, p + (Len(untag) - 1))
Mmsg = mID(Mmsg, p + Len(untag))
EstateGroupUpdate pMsg

Case "estateupdate"
untag = "/>"
p = InStr(Mmsg, untag)
If p = 0 Then Exit Do
pMsg = Left(Mmsg, p + (Len(untag) - 1))
Mmsg = mID(Mmsg, p + Len(untag))
EstateUpdate pMsg

Case "display"
untag = "</display>"
p = InStr(Mmsg, untag)
If p = 0 Then
    untag = "/>"
    p = InStr(Mmsg, untag)
    End If
pMsg = Left(Mmsg, p + (Len(untag) - 1))
Mmsg = mID(Mmsg, p + Len(untag))
Display pMsg

Case "cardupdate"
untag = "/>"
p = InStr(Mmsg, untag)
If p = 0 Then Exit Do
pMsg = Left(Mmsg, p + (Len(untag) - 1))
Mmsg = mID(Mmsg, p + Len(untag))


Case "auctionupdate"
untag = "/>"
p = InStr(Mmsg, untag)
If p = 0 Then Exit Do
pMsg = Left(Mmsg, p + (Len(untag) - 1))
Mmsg = mID(Mmsg, p + Len(untag))
AuctionUpdate pMsg

Case "tradeupdate"
'Debug.Print "mmsg:" & Mmsg
untag = "</tradeupdate>"
p = InStr(Mmsg, untag)
If p = 0 Then
    untag = "/>"
    p = InStr(Mmsg, untag)
    End If
pMsg = Left(Mmsg, p + (Len(untag) - 1))
Mmsg = mID(Mmsg, p + Len(untag))
TradeUpdate pMsg


Case Else
Debug.Print "Invalid tag?:" & Mmsg
Exit Do

End Select

If Mmsg = "" Then Exit Do
Loop

finish:
prevMsg = False
'Debug.Print "EPM"

If ToProcess = True Then
    GoTo procagain
    End If

End Sub

Private Sub TradeUpdate(hLine As String)
'update trade data
Dim tType As String
Dim tID As String
Dim n As Integer
Dim tfrm As frmtrade
Dim tRevision As String
Dim item As ListItem
Dim tAccept As String

Dim tMoney As String, tFrom As String, tTo As String, tEstate As String

'Debug.Print "TU:" & hLine


tType = GetData("type=" & Chr(34), hLine)
tID = GetData("tradeid=" & Chr(34), hLine)
tAccept = GetData("accept=" & Chr(34), hLine)

Select Case tType
Case "new"
'diplay trade window
Set tfrm = New frmtrade
tfrm.lblText.Caption = " from 2 players has accepted trade proposal."
tfrm.lblAccepted.Caption = "0"
tfrm.tradeID = tID
tfrm.Visible = True


If TradeData.tEstateName <> "" And TradeData.tMeStart = True Then
    tfrm.cmbFrom.text = TradeData.tSource
    tfrm.cmbEstate.text = TradeData.tEstateName
    tfrm.cmbTo.text = TradeData.tInitiator
    
    
        
    Else
    tfrm.cmbWhat.ListIndex = 1
    If TradeData.tInitiator <> "" Then tfrm.cmbFrom.text = TradeData.tInitiator
    If TradeData.tSource <> "" Then tfrm.cmbTo.text = TradeData.tSource
    End If

TradeData.tMeStart = False
TradeData.tEstateName = ""
TradeData.tInitiator = ""
TradeData.tSource = ""


colTradeForms.Add tfrm


Case "completed", "rejected", "accepted"
For n = 1 To colTradeForms.Count
If colTradeForms(n).tradeID = tID Then
    colTradeForms(n).canUnload = True
    Unload colTradeForms(n)
    colTradeForms.Remove n
    Exit For
    End If
Next

Case Else

tRevision = GetData("revision=" & Chr(34), hLine)
'update data
For Each tfrm In colTradeForms
If tfrm.tradeID = tID Then
    'this is window, whos data we change
    If tRevision <> "" Then tfrm.tradeRevision = tRevision
    If InStr(hLine, "<trademoney") > 0 Then
        tMoney = GetData("money=" & Chr(34), hLine)
        tFrom = GetData("playerfrom=" & Chr(34), hLine)
        tTo = GetData("playerto=" & Chr(34), hLine)
        'must we update sum?
        For Each item In tfrm.lsttrade.ListItems
            If GetPlayerId(item.text) = tFrom And GetPlayerId(item.SubItems(2)) = tTo And InStr(item.SubItems(1), "Money") > 0 Then
                'update sum!
                If tMoney = "0" Then
                    'delete proposal
                    tfrm.lsttrade.ListItems.Remove item.Index
                    Else
                    item.SubItems(1) = "Money:" & tMoney
                    End If
                GoTo skipItem
                End If
            Next
            
        Set item = tfrm.lsttrade.ListItems.Add(, , GetPlayerName(tFrom), , "bag")
        item.SubItems(1) = "Money:" & tMoney
        item.SubItems(2) = GetPlayerName(tTo)
skipItem:
        End If
        
    If InStr(hLine, "<tradeestate") > 0 Then
        tEstate = GetData("estateid=" & Chr(34), hLine)
        tTo = GetData("targetplayer=" & Chr(34), hLine)
        If tTo = "-1" Then
            'we must remove revision
            For n = 1 To tfrm.lsttrade.ListItems.Count
            If tfrm.lsttrade.ListItems(n).tag = tEstate Then
                tfrm.lsttrade.ListItems.Remove n
                Exit For
                End If
            Next
            Else
            'is trade already existant?
            For Each item In tfrm.lsttrade.ListItems
                If GetPlayerId(item.text) = colestates(CInt(tEstate) + 1).Owner And InStr(item.SubItems(1), colestates(CInt(tEstate) + 1).Name) > 1 And GetPlayerId(item.SubItems(2)) = tTo Then Exit Sub
                Next
                
            Set item = tfrm.lsttrade.ListItems.Add(, , GetPlayerName(colestates(CInt(tEstate) + 1).Owner), , "card")
            item.SubItems(1) = "Estate:" & colestates(CInt(tEstate) + 1).Name
            item.SubItems(2) = GetPlayerName(tTo)
            item.tag = tEstate
            End If
        End If
    If tAccept <> "" Then tfrm.lblAccepted.Caption = tAccept
    Exit For
    End If
Next




End Select


    

End Sub

Private Sub AuctionUpdate(hLine As String)

Dim aID As String
Dim aActor As String
Dim aEstateID As String
Dim aStatus As String
Dim aPlayer As cPlayer
Dim n As Integer
Dim aHighBid As String
Dim aHighBidder As String
Dim item As ListItem


'Debug.Print hline


aID = GetData("auctionid=" & Chr(34), hLine)
aActor = GetData("actor=" & Chr(34), hLine)
aEstateID = GetData("estateid=" & Chr(34), hLine)
aStatus = GetData("status=" & Chr(34), hLine)
aHighBid = GetData("highbid=" & Chr(34), hLine)
aHighBidder = GetData("highbidder=" & Chr(34), hLine)

If aActor <> "" Then
    For n = 1 To colplayers.Count
    Set aPlayer = colplayers(n)
    If aPlayer.ID = aActor Then Exit For
    Next
    End If
    
If auctionStatus = AuctNoAuction Then
    auctId = aID
    frmAuction.lblStatus.Caption = ""
    frmAuction.Caption = aPlayer.Name & " is auctioning " & colestates(CInt(aEstateID) + 1).Name
    frmAuction.txtBid.text = "1"
    frmAuction.Show
    auctionStatus = AuctStarted
    End If

If auctionStatus = AuctStarted Then
    If aHighBid <> "" Then
    For n = 1 To frmAuction.lstAuction.ListItems.Count
        Set item = frmAuction.lstAuction.ListItems(n)
        If mID(item.Key, 2) = aHighBidder Then
            item.SubItems(1) = aHighBid
            Exit For
            End If
        Next
        frmAuction.txtBid.text = CInt(aHighBid) + 1
        End If
    Select Case aStatus
    Case "0"
    frmAuction.lblStatus.Caption = ""
    Case "1"
    frmAuction.lblStatus.Caption = "Going once..."
    Case "2"
    frmAuction.lblStatus.Caption = "Going twice..."
    Case "3"
    frmAuction.lblStatus.Caption = "Sold!"
    auctionStatus = AuctNoAuction
    frmAuction.canClose = True
    Unload frmAuction
    End Select
    End If


End Sub

Private Sub Display(hLine As String)
'here we dsplay text and buttons
Dim eID As String
Dim text As String
Dim tClear As Boolean
Dim tButtonsClear As Boolean
Dim estate As cEstate
Dim Disp As String
Dim p As Integer
Dim n As Integer
Dim numHouses As Integer
'Debug.Print "Display:" & hLine

hLine = Replace(hLine, vbLf, "")

Do

hLine = Trim(hLine)

p = InStr(hLine, "/>")
If p = 0 Then Exit Do

Disp = Left(hLine, p + 1)
hLine = mID(hLine, p + 2)


'Debug.Print "D:" & Disp

text = GetData("text=" & Chr(34), Disp)

If InStr(Disp, "button ") > 0 Then
    DisplayButton Disp
    GoTo dLoop
    End If


If (GetData("clearbuttons=" & Chr(34), Disp) = "1") Then
    frmgame.cmdAuction.Visible = False
    frmgame.cmdBuy.Visible = False
    frmgame.cmdEndturn.Visible = False
    frmgame.cmdGo.Visible = False
    frmgame.cmdback.Visible = False
    frmgame.cmdJailPay.Visible = False
    frmgame.cmdJailRoll.Visible = False
    frmgame.cmdJailUseCard.Visible = False
    frmgame.cmdPayPercentage.Visible = False
    frmgame.cmdPayStatic.Visible = False
    frmgame.cmdRoll.Visible = False
    End If
    

If InStr(Disp, "estateid=") > 0 Then
    eID = GetData("estateid=" & Chr(34), Disp)
    tClear = (GetData("cleartext=" & Chr(34), Disp) = "1")
    If eID <> "" Then
        If CLng(eID) > colestates.Count Then
            Debug.Print "Èez!!"
            GoTo dLoop
            End If
        
        If eID <> "-1" Then
            Set estate = colestates(CInt(eID) + 1)
            frmgame.frmEstate.BackColor = estate.BgColor
            If estate.Color = 0 Then
                frmgame.lblEstateName.BackColor = estate.BgColor
                Else
                frmgame.lblEstateName.BackColor = estate.Color
                End If
            frmgame.lblEstateGroup.BackColor = frmgame.lblEstateName.BackColor
            frmgame.lblEstateGroup.Caption = UCase(GetEstateGroupName(estate.Group))
            frmgame.lblEstateName.Caption = estate.Name
            frmgame.icoMortgaged.Visible = estate.Mortaged
            frmgame.icoCrd.Visible = (estate.Owner = "-1" And estate.CanBeOwned = True)
            frmgame.ico1House.Visible = False
            frmgame.ico2House.Visible = False
            frmgame.ico3House.Visible = False
            frmgame.ico4house.Visible = False
            frmgame.icoHotel.Visible = False
            
            numHouses = estate.Houses
            If InStr(text, "buys") > 0 Then numHouses = numHouses + 1
            If InStr(text, "sells") > 0 Then numHouses = numHouses - 1
            
            
            Select Case numHouses
            Case 1
            frmgame.ico1House.Visible = True
            Case 2
            frmgame.ico2House.Visible = True
            Case 3
            frmgame.ico3House.Visible = True
            Case 4
            frmgame.ico4house.Visible = True
            Case 5
            frmgame.icoHotel.Visible = True
            End Select
            Else
            Debug.Print "Pod"
            End If
        
        End If
        
    If tClear = True Then frmgame.lstMsgs.ListItems.Clear
    
    GoTo dLoop
    End If
    
If InStr(Disp, "playerupdate") > 0 Then
    getPlayerData hLine
    GoTo dLoop
    End If
    
    
Debug.Print "Not implemented:"; Disp

dLoop:

If text <> "" Then
    text = Precode(text)
    If InStr(text, "roll") > 0 Then
        frmgame.lstMsgs.ListItems.Add , , text, , "dice"
        frmgame.cmdRoll.Visible = False
        GoTo dloops
        End If
    If InStr(text, "lands") > 0 Then
        frmgame.lstMsgs.ListItems.Add , , text, , "field"
        If estate.CanBeOwned = True Then
            frmgame.lstMsgs.ListItems.Add , , "Price:" & estate.Price, , "info"
            If GetPlayerName(estate.Owner) = "" Then
                frmgame.lstMsgs.ListItems.Add , , "Owner:None", , "info"
                Else
                frmgame.lstMsgs.ListItems.Add , , "Owner:" & GetPlayerName(estate.Owner), , "info"
                End If
            End If
            
        GoTo dloops
        End If
    If InStr(text, "goes to") > 0 Then
        frmgame.lstMsgs.ListItems.Add , , text, , "field"
        GoTo dloops
        End If
        
    frmgame.lstMsgs.ListItems.Add , , text, , "info"
    GoTo dloops
    End If
dloops:
DoEvents
Loop

End Sub

Private Sub EstateUpdate(hLine As String)

Dim eName As String
Dim eHouses As String 'Long
Dim eMoney As String 'Long
Dim ePassmoney As String 'Long
Dim eMortaged As String 'Boolean
Dim eID As String
Dim eColor As String 'Long
Dim eBgColor As String 'Long
Dim eOwner As String
Dim eHouseprice As String 'Long
Dim eSellHousePrice As String 'Long
Dim eMortagePrice As String 'Long
Dim eUnmortageprice As String 'Long
Dim eGroup As String
Dim eCanBeOwned As String 'Boolean
Dim eCanToggleMortage As String 'Boolean
Dim eCanBuyHouses As String 'Boolean
Dim eCanSellHouses As String 'Boolean
Dim ePrice As String 'Long
Dim eRent0 As String 'Long
Dim eRent1 As String 'Long
Dim eRent2 As String 'Long
Dim eRent3 As String 'Long
Dim eRent4 As String 'Long
Dim eRent5 As String 'Long
Dim pControl As Integer
Dim n As Integer

'Debug.Print "EUPDATE:" & hLine

Dim estate As cEstate
eName = Precode(GetData("name=" & Chr(34), hLine))
eHouses = GetData(" houses=" & Chr(34), hLine)
eMoney = GetData("money=" & Chr(34), hLine)
ePassmoney = GetData("passmoney=" & Chr(34), hLine)
eMortaged = GetData("mortgaged=" & Chr(34), hLine)
eID = GetData("estateid=" & Chr(34), hLine)
eColor = GetData("color=" & Chr(34), hLine)
eBgColor = GetData("bgcolor=" & Chr(34), hLine)
eOwner = GetData("owner=" & Chr(34), hLine)
eHouseprice = GetData("houseprice=" & Chr(34), hLine)
eSellHousePrice = GetData("sellhouseprice=" & Chr(34), hLine)
eMortagePrice = GetData("mortgageprice=" & Chr(34), hLine)
eUnmortageprice = GetData("unmortgageprice=" & Chr(34), hLine)
eGroup = GetData("group=" & Chr(34), hLine)
eCanBeOwned = GetData("can_be_owned=" & Chr(34), hLine)
eCanToggleMortage = GetData("can_toggle_mortgage=" & Chr(34), hLine)
eCanBuyHouses = GetData("can_buy_houses=" & Chr(34), hLine)
eCanSellHouses = GetData("can_sell_houses=" & Chr(34), hLine)
ePrice = GetData("price=" & Chr(34), hLine)
eRent0 = GetData("rent0=" & Chr(34), hLine)
eRent1 = GetData("rent1=" & Chr(34), hLine)
eRent2 = GetData("rent2=" & Chr(34), hLine)
eRent3 = GetData("rent3=" & Chr(34), hLine)
eRent4 = GetData("rent4=" & Chr(34), hLine)
eRent5 = GetData("rent5=" & Chr(34), hLine)


If colestates.Count < (CInt(eID) + 1) Then
    Set estate = New cEstate
    Else
    Set estate = colestates(CInt(eID) + 1)
    End If
    
    
If eID <> "" Then estate.ID = eID
If eColor <> "" Then estate.Color = GetColor(eColor)
If eBgColor <> "" Then estate.BgColor = GetColor(eBgColor)
If eOwner <> "" Then estate.Owner = eOwner
If eHouseprice <> "" Then estate.Houseprice = CLng(eHouseprice)
If eSellHousePrice <> "" Then estate.SellHousePrice = CLng(eSellHousePrice)
If eMortagePrice <> "" Then estate.MortagePrice = CLng(eMortagePrice)
If eUnmortageprice <> "" Then estate.Unmortageprice = CLng(eUnmortageprice)
If eGroup <> "" Then estate.Group = eGroup
If eCanBeOwned <> "" Then estate.CanBeOwned = (eCanBeOwned = "1")
If eCanToggleMortage <> "" Then estate.CanToggleMortage = (eCanToggleMortage = "1")
If eCanBuyHouses <> "" Then estate.CanBuyHouses = (eCanBuyHouses = "1")
If eCanSellHouses <> "" Then estate.CanSellHouses = (eCanSellHouses = "1")
If ePrice <> "" Then estate.Price = CLng(ePrice)
If eRent0 <> "" Then estate.Rent0 = CLng(eRent0)
If eRent1 <> "" Then estate.Rent1 = CLng(eRent1)
If eRent2 <> "" Then estate.Rent2 = CLng(eRent2)
If eRent3 <> "" Then estate.Rent3 = CLng(eRent3)
If eRent4 <> "" Then estate.Rent4 = CLng(eRent4)
If eRent5 <> "" Then estate.Rent5 = CLng(eRent5)
If eName <> "" Then estate.Name = eName
If eHouses <> "" Then estate.Houses = CInt(eHouses)
If eMoney <> "" Then estate.Money = CLng(eMoney)
If ePassmoney <> "" Then estate.Passmoney = CLng(ePassmoney)
If eMortagePrice <> "" Then estate.MortagePrice = CLng(eMortagePrice)
If eUnmortageprice <> "" Then estate.Unmortageprice = CLng(eUnmortageprice)
If eSellHousePrice <> "" Then estate.SellHousePrice = CLng(eSellHousePrice)
If eMortaged <> "" Then
    estate.Mortaged = (eMortaged = "1")
    If canDraw = True Then frmgame.Mortgage estate.Mortaged, CInt(estate.ID)
    End If
    

If colestates.Count < (eID + 1) Then colestates.Add estate




If estate.Owner <> "-1" And estate.Owner <> "" Then
    'estate is bought/selled
    'previous owner?
    If estate.PrevOwner <> "-1" Then
            'erase it from previous player
            For n = 0 To frmgame.conPlayer.Count - 1
            If frmgame.conPlayer(n).PlayerID = estate.PrevOwner Then
                frmgame.conPlayer(n).DrawEstate CInt(estate.ID), False
                Exit For
                End If
            Next
            End If
    'draw it to new owner
    For n = 0 To frmgame.conPlayer.Count - 1
    If frmgame.conPlayer(n).PlayerID = estate.Owner Then
        frmgame.conPlayer(n).DrawEstate CInt(estate.ID), True
        Exit For
        End If
    Next
    estate.PrevOwner = estate.Owner
    frmgame.EstateBought estate.ID, False
    Else
    If estate.PrevOwner <> "-1" Then
            'erase it from previous player
            For n = 0 To frmgame.conPlayer.Count - 1
            If frmgame.conPlayer(n).PlayerID = estate.PrevOwner Then
                frmgame.conPlayer(n).DrawEstate CInt(estate.ID), False
                Exit For
                End If
            Next
            End If
    estate.PrevOwner = "-1"
    frmgame.EstateBought estate.ID, True
    End If



If eHouses <> "" Then
    frmgame.HouseBought estate.ID, CInt(GetData("houses=" & Chr(34), hLine))
    End If


Set estate = Nothing
If CInt(eID) = 39 Then dbg = True
End Sub

Private Function GetColor(hLine As String) As Long
'this function transforms text color code into long type

If hLine = "" Then Exit Function
hLine = Replace(hLine, "#", "")

Dim n As Integer
Dim cVal As Integer
Dim Char As String




GetColor = RGB(Val("&H" & Left(hLine, 2)), Val("&H" & mID(hLine, 3, 2)), Val("&H" & Right(hLine, 2)))





End Function

Private Sub EstateGroupUpdate(hLine As String)
Dim estateGroup As New cEstateGroup

estateGroup.ID = GetData("groupid=" & Chr(34), hLine)
estateGroup.Name = GetData("name=" & Chr(34), hLine)

colEstateGroups.Add estateGroup

Set estateGroup = Nothing
End Sub

Private Sub gameUpdate(hLine As String)
'here game status is being updatet and appropiate actions taken
Dim gStatus As String
Dim n As Integer

'If prevGStatus = hLine Then Exit Sub
'prevGStatus = hLine

'Debug.Print hline

gStatus = GetData("status=" & Chr(34), hLine)

Select Case gStatus

Case "config"
Set colEstateGroups = Nothing
Set colestates = Nothing
Set colplayers = Nothing

Debug.Print "CONFIG"
pcsLoaded = False
Me.Hide
frmgame.Show
frmgame.Caption = lstGames.SelectedItem.text & ":" & lstGames.SelectedItem.SubItems(1) & " on " & serverName
DoEvents

Case "init"
Debug.Print "INIT"
If pcsLoaded = False Then
    canDraw = False
    frmgame.sBar.Panels(1).text = "Retrieving full game data..."
    DoEvents
    pcsLoaded = True
    End If
frmgame.pIece(0).tag = mID(frmgame.lstPlayers.ListItems(1).Key, 2)
frmgame.lblPName(0).Caption = frmgame.lstPlayers.ListItems(1).text

For n = 2 To frmgame.lstPlayers.ListItems.Count
Load frmgame.pIece(n - 1)
Load frmgame.lblPName(n - 1)
frmgame.pIece(n - 1).tag = mID(frmgame.lstPlayers.ListItems(n).Key, 2)
frmgame.lblPName(n - 1).Caption = frmgame.lstPlayers.ListItems(n).text
frmgame.lblPName(n - 1).ZOrder 0
Next


For n = 0 To 9
pcsPos(n) = -1
Next


Case "run"
Debug.Print "RUN"
frmgame.Caption = "Wintlantik: Game is running"
DoEvents
For n = 1 To colestates.Count - 1
Load frmgame.pictMortaged(n)
Load frmgame.pictCard(n)
Load frmgame.pict1House(n)
Load frmgame.pict2House(n)
Load frmgame.pict3house(n)
Load frmgame.pict4house(n)
Load frmgame.pictHotel(n)
Next

frmgame.conPlayer(0).PlayerName = colplayers(1).Name
Set frmgame.conPlayer(0).Estates = colestates
Set frmgame.conPlayer(0).EstateGroups = colEstateGroups
frmgame.conPlayer(0).PlayerID = colplayers(1).ID
frmgame.conPlayer(0).Visible = True
frmgame.conPlayer(0).InitEstateList

For n = 2 To colplayers.Count
Load frmgame.conPlayer(n - 1)
frmgame.conPlayer(n - 1).PlayerName = colplayers(n).Name
frmgame.conPlayer(n - 1).Left = frmgame.conPlayer(0).Left
frmgame.conPlayer(n - 1).Top = frmgame.conPlayer(n - 2).Top + frmgame.conPlayer(0).Height
Set frmgame.conPlayer(n - 1).Estates = colestates
Set frmgame.conPlayer(n - 1).EstateGroups = colEstateGroups
frmgame.conPlayer(n - 1).PlayerID = colplayers(n).ID
frmgame.conPlayer(n - 1).Visible = True
frmgame.conPlayer(n - 1).InitEstateList
Next

frmgame.frmConf.Visible = False
frmgame.lstPlayers.Visible = False


frmgame.sBar.Panels(1).text = "Game in progress..."
DoEvents
canDraw = True
frmgame.cmdback.Visible = False

gameStatus = Running
frmgame.DrawBoard

Case "end"
Debug.Print "END"

End Select

End Sub

Private Sub getPlayerData(hLine As String)
'here we update player's data (money, name, location...)
Dim pname As String
Dim colplayer As cPlayer
Dim pId As String
Dim pCookie As String

pId = GetData("playerid=" & Chr(34), hLine)
pCookie = GetData("cookie=" & Chr(34), hLine)


'If player.ID = "" And pID <> "" Then player.ID = pID
'If player.Cookie = "" And pCookie <> "" Then player.Cookie = pCookie



If playerSet = False Then
    kSend ".n" & player.Name
    playerSet = True
    End If
If gmList = False Then
    kSend ".gl"
    gmList = True
    End If
    
If colplayers Is Nothing Then Exit Sub

Dim pMoney As String
Dim pDoublecount As String
Dim pJailCount As String
Dim pBankrupt As String
Dim pJailed As String
Dim pHasturn As String
Dim pSpectator As String
Dim pCanroll As String
Dim pCanRollAgain As String
Dim pCanBuyEstate As String
Dim pCanAuction As String
Dim pHasdebt As String
Dim pCanUseCard As String
Dim pLocation As String
Dim pDirectMove As String
Dim m As Integer
Dim item As ListItem

'If prevPdata2 = hline Then Exit Sub
'prevPdata2 = hline

pMoney = GetData("money=" & Chr(34), hLine)

'Debug.Print "PL:" & hLine

pname = GetData("name=" & Chr(34), hLine)
pDoublecount = GetData("doublecount=" & Chr(34), hLine)
pJailCount = GetData("jailcount=" & Chr(34), hLine)
pBankrupt = GetData("bankrupt=" & Chr(34), hLine)
pJailed = GetData("jailed=" & Chr(34), hLine)
pHasturn = GetData("hasturn=" & Chr(34), hLine)
pSpectator = GetData("spectator=" & Chr(34), hLine)
pCanroll = GetData("can_roll=" & Chr(34), hLine)
pCanRollAgain = GetData("can_rollagain=" & Chr(34), hLine)
pCanBuyEstate = GetData("can_buyestate=" & Chr(34), hLine)
pCanAuction = GetData("canauction=" & Chr(34), hLine)
pHasdebt = GetData("hasdebt=" & Chr(34), hLine)
pCanUseCard = GetData("canusecard=" & Chr(34), hLine)
pLocation = GetData("location=" & Chr(34), hLine)
pDirectMove = GetData("directmove=" & Chr(34), hLine)
pHasturn = GetData("hasturn=" & Chr(34), hLine)


For Each colplayer In colplayers
    DoEvents
    If colplayer.ID = pId Then
        If pLocation <> "" And pDirectMove <> "" Then
            'player has moved!
            colplayer.location = CInt(pLocation)
            frmgame.MovePlayer colplayer.ID, colplayer.location, (pDirectMove = "1")
            End If

        If pMoney <> "" Then colplayer.Money = CLng(pMoney)
        If pDoublecount <> "" Then colplayer.Doublecount = (pDoublecount = "1")
        If pJailCount <> "" Then colplayer.Doublecount = CInt(pJailCount)
        If pBankrupt <> "" Then colplayer.Bankrupt = (pBankrupt = "1")
        If pJailed <> "" Then colplayer.Jailed = (pJailed = "1")
        If pHasturn <> "" Then colplayer.Hasturn = (pHasturn = "1")
        If pSpectator <> "" Then colplayer.Spectator = (pSpectator = "1")
        If pCanroll <> "" Then colplayer.Canroll = (pCanroll = "1")
        If pCanRollAgain <> "" Then colplayer.CanRollAgain = (pCanRollAgain = "1")
        If pCanBuyEstate <> "" Then colplayer.CanBuyEstate = (pCanBuyEstate = "1")
        If pCanAuction <> "" Then colplayer.CanAuction = (pCanAuction = "1")
        If pHasdebt <> "" Then colplayer.Hasdebt = (pHasdebt = "1")
        If pCanUseCard <> "" Then colplayer.CanUseCard = (pCanUseCard = "1")
        If pMoney <> "" Then
            For m = 1 To frmgame.conPlayer.Count
            If frmgame.conPlayer(m - 1).PlayerID = pId Then
                frmgame.conPlayer(m - 1).PlayerMoney = pMoney
                Exit For
                End If
            Next
            End If
        If (colplayer.Canroll = True And pCanroll <> "") Or (colplayer.CanRollAgain = True And pCanRollAgain <> "") Then
            If colplayer.Jailed = False Or colplayer.Hasdebt = True Then
                frmgame.cmdRoll.Visible = True
                frmgame.mnugBankrupt.Visible = True
                End If
            End If
        'Exit For
        If pname <> "" Then
            colplayer.Name = pname
            If frmgame.frmEstate.Visible = True Then
                'change player's name
                For m = 0 To frmgame.conPlayer.Count - 1
                    If frmgame.conPlayer(m).PlayerID = colplayer.ID Then
                        frmgame.conPlayer(m).PlayerName = pname
                        Exit For
                        End If
                    Next
                Else
                For m = 1 To frmgame.lstPlayers.ListItems.Count
                If mID(frmgame.lstPlayers.ListItems(m).Key, 2) = colplayer.ID Then
                    frmgame.lstPlayers.ListItems(m).text = pname
                    Exit For
                    End If
                Next
                End If
            End If
        
        End If
    Next
    
    If pHasturn <> "" Then
        For m = 0 To frmgame.conPlayer.Count - 1
        If frmgame.conPlayer(m).PlayerID = pId Then
            frmgame.conPlayer(m).HiglightPlayer (pHasturn = "1")
            Exit For
            End If
         Next
         End If
    
End Sub


Private Sub UpdatePlayerList(hLine As String)
'update player list. How?
Dim gType As String


'type?
gType = GetData("type=" & Chr(34), hLine)

Select Case UCase(gType)
Case "FULL"
updatePlistFull (hLine)

Case "ADD"
addPlayer (hLine)

Case "DEL"
DeletePlayer hLine

Case "EDIT"
'editPlayer hline
End Select

End Sub

Private Sub updatePlistFull(ByVal hLine As String)
'we replace all players in list and collection

If prevPlistLine = hLine Then Exit Sub
prevPlistLine = hLine


Dim pl As String, n As Integer
Dim p As Integer, q As Integer
Dim item As ListItem
Dim cpl As cPlayer

Dim pId As String, pname As String, phost As String, pmaster As Boolean



'Debug.Print hline
'erase players collection
frmgame.lstPlayers.ListItems.Clear
DoEvents

If colplayers Is Nothing Then
    Set colplayers = New Collection
    Else
    Set colplayers = Nothing
    Set colplayers = New Collection
    End If
    
'now we go player by player
Do
p = InStr(UCase(hLine), "<PLAYER ")
If p = 0 Then GoTo send
p = p + 8

q = InStr(p, hLine, "/>")
If q = 0 Then GoTo send

pl = mID(hLine, p, q - p)

'let's add player
pId = GetData("playerid=" & Chr(34), pl)

pname = GetData("name=" & Chr(34), pl)
phost = GetData("host=" & Chr(34), pl)
pmaster = (GetData("master=" & Chr(34), pl) = "1")


Set item = frmgame.lstPlayers.ListItems.Add(, "p" & pId, pname)
item.SubItems(1) = phost


Set cpl = New cPlayer
cpl.Host = phost
cpl.ID = pId
cpl.IsMaster = pmaster
cpl.Name = pname

colplayers.Add cpl

hLine = mID(hLine, q + 2)
sloop:
Loop
    
send:

End Sub

Private Sub AddMessage(hLine As String)
'we display messages sent by server
Dim mType As String
Dim pname As String, text As String

mType = GetData("type=" & Chr(34), hLine)

Select Case UCase(mType)

Case "CHAT"
pname = GetData("author=" & Chr(34), hLine)
text = GetData("value=" & Chr(34), hLine)
text = Precode(text)
frmgame.txtChatList.text = frmgame.txtChatList.text & pname & ":" & text & vbCrLf
frmgame.txtChatList.SelStart = Len(frmgame.txtChatList.text)


Case "INFO"
text = GetData("value=" & Chr(34), hLine)
text = Precode(text)
frmgame.txtChatList.text = frmgame.txtChatList.text & "INFO:" & text & vbCrLf
frmgame.txtChatList.SelStart = Len(frmgame.txtChatList.text)

Case "ERROR"
text = GetData("value=" & Chr(34), hLine)
text = Precode(text)
frmgame.txtChatList.text = frmgame.txtChatList.text & "ERROR:" & text & vbCrLf
frmgame.txtChatList.SelStart = Len(frmgame.txtChatList.text)

Case "STANDBY"
text = GetData("value=" & Chr(34), hLine)
text = Precode(text)
frmgame.txtChatList.text = frmgame.txtChatList.text & "STANDBY:" & text & vbCrLf
frmgame.txtChatList.SelStart = Len(frmgame.txtChatList.text)


Case "STARTGAME"
text = GetData("value=" & Chr(34), hLine)
text = Precode(text)
frmgame.txtChatList.text = frmgame.txtChatList.text & "STARTGAME:" & text & vbCrLf
frmgame.txtChatList.SelStart = Len(frmgame.txtChatList.text)

End Select
End Sub

Public Function Precode(hLine As String) As String
'slovenian WIN characters from ISO code
hLine = Replace(hLine, "Ä", "è")
hLine = Replace(hLine, "ÄŒ", "È")
hLine = Replace(hLine, "Å¡", "š")
hLine = Replace(hLine, "Å ", "Š")
hLine = Replace(hLine, "Å¾", "ž")
hLine = Replace(hLine, "Å½", "Ž")
hLine = Replace(hLine, "©", "Š")
hLine = Replace(hLine, "¹", "š")
hLine = Replace(hLine, "®", "Ž")
hLine = Replace(hLine, "¾", "ž")

Precode = hLine
End Function


Private Sub addPlayer(hLine As String)
If colplayers Is Nothing Then Exit Sub
'We add player, if it does not exist
'first we check, if player exists by player id

Dim pId As String, pname As String, phost As String, pmaster As Boolean
Dim n As Integer, item As ListItem, cpl As cPlayer

pId = GetData("playerid=" & Chr(34), hLine)

For n = 1 To colplayers.Count
If colplayers(n).ID = pId Then Exit Sub 'player exists
Next

'player does not exists yet
pname = GetData("name=" & Chr(34), hLine)
phost = GetData("host=" & Chr(34), hLine)
pmaster = (GetData("master=" & Chr(34), hLine) = "1")

Set item = frmgame.lstPlayers.ListItems.Add(, "p" & pId, pname)
item.SubItems(1) = phost

Set cpl = New cPlayer
cpl.Host = phost
cpl.ID = pId
cpl.IsMaster = pmaster
cpl.Name = pname

colplayers.Add cpl
Beep

'Debug.Print "add player:" & colPlayers.Count
End Sub


Private Sub DeletePlayer(hLine As String)
'delete player, if exists
If colplayers Is Nothing Then Exit Sub
Dim n As Integer
Dim plID As String

plID = GetData("playerid=" & Chr(34), hLine)
For n = 1 To colplayers.Count
If colplayers(n).ID = plID Then
    colplayers.Remove (n)
    Exit For
    End If
Next

For n = 1 To frmgame.lstPlayers.ListItems.Count
If frmgame.lstPlayers.ListItems(n).Key = "p" & plID Then
    frmgame.lstPlayers.ListItems.Remove n
    Beep
    Exit For
    End If

Next
End Sub


Private Sub ConfigUpdate(hLine As String)
'we display config options

conIn = True
Dim p As Integer
Dim q As Integer
Dim opt As String
Dim n As Integer
Dim indx As Integer

Dim oTitle As String
Dim oCommand As String
Dim oValue As Boolean
Dim oEdit As Boolean

'how would you unload dynamic options?

For n = frmgame.chkOpt.Count To 2 Step -1
    Unload frmgame.chkOpt(n - 1)
    Next

indx = 0
'we begin empty. if index=0 then we use exsistant control

Do
p = InStr(hLine, "option")
If p = 0 Then GoTo send
p = p + 6


'untag?
q = InStr(hLine, "/>")
q = q + 2

opt = mID(hLine, p, q - p)
hLine = mID(hLine, q + 1)

oTitle = GetData("title=" & Chr(34), opt)
oCommand = GetData("command=" & Chr(34), opt)
oValue = (GetData("value=" & Chr(34), opt) = "1")
oEdit = (GetData("edit=" & Chr(34), opt) = "1")

frmgame.cmdback.Visible = True
'we add control
If indx = 0 Then
    'we can use existing one
    frmgame.chkOpt(indx).Caption = oTitle
    If oValue = True Then
        frmgame.chkOpt(indx).Value = 1
        Else
        frmgame.chkOpt(indx).Value = 0
        End If
    frmgame.chkOpt(indx).Enabled = oEdit
    frmgame.chkOpt(indx).tag = oCommand
    'if first option is enabled, then I am master of the game
    If oEdit = True Then
        frmgame.cmdGo.Visible = True
        End If
    frmgame.chkOpt(indx).Visible = True
    indx = 1
    Else
    Load frmgame.chkOpt(indx)
    frmgame.chkOpt(indx).Caption = oTitle
    If oValue = True Then
        frmgame.chkOpt(indx).Value = 1
        Else
        frmgame.chkOpt(indx).Value = 0
        End If
    frmgame.chkOpt(indx).Enabled = oEdit
    frmgame.chkOpt(indx).tag = oCommand
    frmgame.chkOpt(indx).Left = frmgame.chkOpt(0).Left
    frmgame.chkOpt(indx).Top = frmgame.chkOpt(indx - 1).Top + frmgame.chkOpt(0).Height + 100
    frmgame.chkOpt(indx).Visible = True
    indx = indx + 1
    End If
    
Loop

send:

For n = 0 To frmgame.chkOpt.Count - 1
If frmgame.chkOpt(n).tag = ".geS" Then
    canSell2Bank = (frmgame.chkOpt(n).Value = 1)
    Exit For
    End If
Next


conIn = False
End Sub


Private Sub mklient_Connect()
Debug.Print "connected"
End Sub



Private Sub DisplayButton(hLine As String)
'display appropiate button
Dim bCommand As String
Dim bCaption As String
Dim bEnabled As Boolean

bCommand = GetData("command=" & Chr(34), hLine)
bCaption = GetData("caption=" & Chr(34), hLine)
bEnabled = (GetData("enabled=" & Chr(34), hLine) = "1")

'Debug.Print hLine

Select Case bCommand
Case ".ea"
frmgame.cmdAuction.Visible = bEnabled

Case ".E"
frmgame.cmdEndturn.Visible = bEnabled

Case ".eb"
frmgame.cmdBuy.Visible = bEnabled

Case ".T$"
If bEnabled = True Then
    frmgame.cmdPayStatic.Caption = bCaption
    frmgame.cmdPayStatic.tag = bCommand
    frmgame.cmdPayStatic.Visible = True
    Else
    frmgame.cmdPayStatic.Visible = False
    End If

Case ".T%"
If bEnabled = True Then
    frmgame.cmdPayPercentage.Caption = bCaption
    frmgame.cmdPayPercentage.tag = bCommand
    frmgame.cmdPayPercentage.Visible = True
    Else
    frmgame.cmdPayPercentage.Visible = False
    End If

Case ".jc"
If bEnabled = True Then frmgame.cmdJailUseCard.Visible = True
frmgame.cmdRoll.Visible = False
Case ".jp"
If bEnabled = True Then frmgame.cmdJailPay.Visible = True
frmgame.cmdRoll.Visible = False

Case ".jr"
If bEnabled = True Then frmgame.cmdJailRoll.Visible = True
frmgame.cmdRoll.Visible = False

End Select
End Sub

Private Sub tpause_Timer()
'if we pause... OBSOLETE!
doWait = False
End Sub
