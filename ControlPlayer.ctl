VERSION 5.00
Begin VB.UserControl ControlPlayer 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3120
   ScaleHeight     =   795
   ScaleWidth      =   3120
   Begin VB.Label lblMoney 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label lblName 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "ControlPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mName As String
Dim mMoney As String
Dim mEstates As New Collection
Dim mNumGroups As Integer
Dim GPos() As Integer
Dim mID As String
Const clipBrd = 30
Const clipCrd = 150
Const slip = 30
Const Deck = 30
Const CardHeight = 250
Const cGray = 191
Const CardWidth = 180

Public Event RightClick()

Public Sub HiglightPlayer(Active As Boolean)
If Active = True Then
    lblMoney.BackColor = RGB(58, 110, 165)
    lblName.BackColor = RGB(58, 110, 165)
    Else
    lblMoney.BackColor = vbBlack
    lblName.BackColor = vbBlack
    End If
End Sub


Public Property Let PlayerID(vdata As String)
mID = vdata
End Property
Public Property Get PlayerID() As String
PlayerID = mID
End Property



Private Sub lblMoney_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then RaiseEvent RightClick
End Sub

Private Sub lblName_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then RaiseEvent RightClick
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'user har right-clicked control
If Button = 2 Then RaiseEvent RightClick
End Sub

Private Sub UserControl_Resize()
'if control is resized...
lblName.Width = Width / 2
lblMoney.Left = Width / 2
lblMoney.Width = Width / 2
End Sub


Public Property Let PlayerName(vdata As String)
'set player's name
mName = vdata
lblName.Caption = " " & vdata
End Property


Public Property Let PlayerMoney(vdata As String)
'set player's money
mMoney = vdata
lblMoney.Caption = vdata & " "
End Property

Public Property Set Estates(vdata As Collection)
'load estates collection, but only vital properties
Dim n As Integer
Dim estate As cConCard

For n = 1 To vdata.Count
If vdata(n).CanBeOwned = True Then
    Set estate = New cConCard
    estate.Color = vdata(n).Color
    estate.GroupID = vdata(n).Group
    estate.ID = vdata(n).ID
    mEstates.Add estate
    End If
Next


End Property

Public Property Set EstateGroups(vdata As Collection)
'how many estate groups? (to calculate shifts when drawing estate cards)
mNumGroups = vdata.Count
ReDim GPos(mNumGroups) As Integer
End Property

Public Sub InitEstateList()
'We draw estate cards on control
'I used following principe:
'Each card is object, and here I set it's coords.
'If I need to higlight it, it is then faster to obtain coords

Dim StartX As Long, endX As Long
Dim startY As Long, endY As Long
Dim boardWidth As Long, boardHeight As Long
Dim n As Integer
Dim estate As cConCard
Dim gid As Integer
Dim cX As Long, cY As Long


'let's draw estates and set their coords
StartX = clipBrd
endX = Width - clipBrd

startY = lblMoney.Height + clipBrd
endY = Height - clipBrd

boardWidth = endX - StartX
boardHeight = endY - startY


For n = 1 To mEstates.Count
'one by one
Set estate = mEstates(n)
'groupID?
gid = CInt(estate.GroupID)

'start coords
cX = StartX
cX = cX + (gid * (CardWidth + clipCrd)) + (slip * GPos(gid))

cY = startY
cY = cY + (slip * 3 * GPos(gid))

estate.X = cX
estate.Y = cY
'first clear space
Line (cX, cY)-Step(CardWidth, CardHeight), vbWhite, BF
'draw border
Line (cX, cY)-Step(CardWidth, CardHeight), RGB(cGray, cGray, cGray), B
'draw colored bar (gray here, since there are no owners yet)
Line (cX, cY)-Step(CardWidth, Deck), RGB(cGray, cGray, cGray), BF


GPos(gid) = GPos(gid) + 1 'where to place next control
Next


End Sub

Public Sub DrawEstate(EstateID As Integer, Optional DrawActive As Boolean = False)
'here we highlight/dehiglight estate cards
Dim estate As cConCard

For Each estate In mEstates
If estate.ID = EstateID Then estate.Owned = DrawActive
If (estate.ID = EstateID And DrawActive = True) Or estate.Owned = True Then
    'we draw colored bar and some lines...
    Line (estate.X, estate.Y)-Step(CardWidth, CardHeight), vbWhite, BF
    Line (estate.X, estate.Y)-Step(CardWidth, CardHeight), RGB(cGray, cGray, cGray), B
    Line (estate.X, estate.Y)-Step(CardWidth, Deck), estate.Color, BF
    Line (estate.X + 30, estate.Y + 30 + Deck)-Step(CardWidth - 60, 0), RGB(cGray, cGray, cGray)
    Line (estate.X + 30, estate.Y + 60 + Deck)-Step(CardWidth - 60, 0), RGB(cGray, cGray, cGray)
    Line (estate.X + 30, estate.Y + 90 + Deck)-Step(CardWidth - 60, 0), RGB(cGray, cGray, cGray)
    Line (estate.X + 30, estate.Y + 120 + Deck)-Step(CardWidth - 100, 0), RGB(cGray, cGray, cGray)
    Line (estate.X + 30, estate.Y + 150 + Deck)-Step(CardWidth - 100, 0), RGB(cGray, cGray, cGray)
    Line (estate.X + 30, estate.Y + 180 + Deck)-Step(CardWidth - 60, 0), RGB(cGray, cGray, cGray)
    estate.Owned = True
    Else
    'if there is no owner... draw in grey
    Line (estate.X, estate.Y)-Step(CardWidth, CardHeight), vbWhite, BF
    Line (estate.X, estate.Y)-Step(CardWidth, CardHeight), RGB(cGray, cGray, cGray), B
    Line (estate.X, estate.Y)-Step(CardWidth, Deck), RGB(cGray, cGray, cGray), BF
    End If
'remember for drawing in future
If estate.ID = EstateID Then estate.Owned = DrawActive
Next


End Sub
