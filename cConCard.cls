VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cConCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'card in owner's control
Public Owned As Boolean
Public ID As String
Public GroupID As String
Public Color As Long
Public X As Single
Public Y As Single


Private Sub Class_Initialize()
'set default values here for each object...
Owned = False
ID = "-1"
GroupID = "-1"
Color = -1
X = 0
Y = 0

End Sub
