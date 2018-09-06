VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} monthuserform 
   Caption         =   "New month macros"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6615
   OleObjectBlob   =   "monthuserform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "monthuserform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton5_Click()
Call firstdayofthemonth_samequarter
Hide
End Sub

Private Sub CommandButton6_Click()
Call firstdayofthemonth_newquarter
Hide
End Sub

Private Sub CommandButton7_Click()
Call erase_data
End Sub

Private Sub CommandButton8_Click()
Hide
End Sub
