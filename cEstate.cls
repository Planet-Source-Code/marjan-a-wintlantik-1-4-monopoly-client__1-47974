VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEstate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'estate object
Public Name As String
Public Houses As Long
Public Money As Long
Public Passmoney As Long
Public MortagePrice As Long
Public Unmortageprice As Long
Public SellHousePrice As Long
Public Mortaged As Boolean
Public ID As String
Public Color As Long
Public BgColor As Long
Public Owner As String
Public Houseprice As Long
Public Group As String
Public CanBeOwned As Boolean
Public CanToggleMortage As Boolean
Public CanBuyHouses As Boolean
Public CanSellHouses As Boolean
Public Price As Long
Public Rent0 As Long
Public Rent1 As Long
Public Rent2 As Long
Public Rent3 As Long
Public Rent4 As Long
Public Rent5 As Long
Public PrevOwner As String

Private Sub Class_Initialize()
PrevOwner = "-1"
End Sub
