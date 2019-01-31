VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} monthuserform 
   Caption         =   "New month settings"
   ClientHeight    =   5400
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
Call NewMonth_Main(SameQuarter)
Hide
End Sub

Private Sub CommandButton6_Click()
Call NewMonth_Main(NewQuarter)
Hide
End Sub

Private Sub CommandButton7_Click()
Call EraseCopiedData
Hide
End Sub

Private Sub CommandButton8_Click()
Hide
End Sub

Private Sub Label14_Click()

End Sub

Private Sub Label17_Click()

End Sub

Private Sub Label19_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub Label7_Click()

End Sub

Private Sub Label8_Click()

End Sub

Private Sub UserForm_Click()

End Sub
