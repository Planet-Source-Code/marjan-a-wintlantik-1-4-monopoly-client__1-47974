Attribute VB_Name = "Globali"
Option Explicit
Global level As Integer
Global serverName As String
Global serverPort As String
Global canGet As Boolean
Global player As tClient
Global server As tServer
Global colEstateGroups As New Collection
Global colestates As New Collection
Global prevPlistLine As String
Global colplayers As Collection
Global canDraw As Boolean
Global conIn As Boolean
Global Const plName = "MMR"
Global pcsLoaded As Boolean
Global pcsPos(10) As Integer
Global auctId As String
Global gameStatus As tStatus
Global colTradeForms As New Collection
Global canSell2Bank As Boolean
Global TradeData As tTrade

'game status
Enum tStatus
onHold = 0
Config
init
BeforeRun
Running
AuctNoAuction
AuctStarted
End Enum

'server struct
Type tServer
Name As String
version As String
End Type

'client struct
Type tClient
Name As String
ID As String
Cookie As String
End Type

'board orientation enum
Enum tOrientation
oup = 0
odown
oLeft
oRight
End Enum

'trade struct
Type tTrade
tInitiator As String
tSource As String
tEstateName As String
tMeStart As Boolean
End Type


Function GetData(hOrient As String, hLine As String) As String
'here we collect data from string
'data is gathered from last character in hOrient (wich is ")
'and until next " character

Dim p As Integer, q As Integer

p = InStr(hLine, hOrient)
If p = 0 Then Exit Function
p = p + Len(hOrient)

'second "
q = InStr(p, hLine, Chr(34))

GetData = mID(hLine, p, q - p)

End Function

Sub centerForm(frm As Form)
'center form
frm.Left = Screen.Width / 2 - frm.Width / 2
frm.Top = Screen.Height / 2 - frm.Height / 2
End Sub


Function GetPlayerId(Name As String) As String
'we get player's name from his ID

Dim colplayer As cPlayer


For Each colplayer In colplayers
If colplayer.Name = Name Then
    GetPlayerId = colplayer.ID
    Exit Function
    End If
Next

End Function

Function GetPlayerName(PlayerID As String) As String
'we get player's ID from his name (potentialy dangerous)
Dim colplayer As cPlayer


For Each colplayer In colplayers
If colplayer.ID = PlayerID Then
    GetPlayerName = colplayer.Name
    Exit Function
    End If
Next

End Function

Function GetEstateGroupName(EstateGroupID As String) As String
'we get estate group name...
Dim estateGroup As cEstateGroup

For Each estateGroup In colEstateGroups
If estateGroup.ID = EstateGroupID Then
    GetEstateGroupName = estateGroup.Name & " "
    Exit For
    End If
Next
End Function

Sub WriteConfig()
'write configuration file if changes occures
Open App.Path & "\wintlantik.ini" For Output As #1
Print #1, player.Name
Close
End Sub

Sub ReadConfig()
'read configuration file
Open App.Path & "\wintlantik.ini" For Input As #1
Line Input #1, player.Name
Close

End Sub
