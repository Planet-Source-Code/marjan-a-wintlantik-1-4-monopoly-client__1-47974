VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'player object
Public ID As String
Public Name As String
Public Host As String
Public IsMaster As Boolean
Public Cookie As String
Public Money As Long
Public Doublecount As Boolean
Public JailCount As Integer
Public Bankrupt As Boolean
Public Jailed As Boolean
Public Hasturn As Boolean
Public Spectator As Boolean
Public Canroll As Boolean
Public CanRollAgain As Boolean
Public CanBuyEstate As Boolean
Public CanAuction As Boolean
Public Hasdebt As Boolean
Public CanUseCard As Boolean
Public location As Integer
