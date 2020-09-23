VERSION 5.00
Begin VB.Form frmEstate 
   Caption         =   "Estate data"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3195
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   3195
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmEstate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub DisplayEstate(estate As cEstate)
BackColor = estate.BgColor

If estate.Color = 0 Then
    lblName.BackColor = estate.BgColor
    Else
    lblName.BackColor = estate.Color
    End If
lblName.Caption = estate.Name

End Sub

Private Sub Form_Load()

End Sub
