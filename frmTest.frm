VERSION 5.00
Begin VB.Form frmTest 
   AutoRedraw      =   -1  'True
   Caption         =   "Sizeable ColorFade Test Form"
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6090
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   6090
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()

    If FadeCycles = 0 Then
        Fade.PaintObj Me, FadeDirection
    Else
        Fade.PaintObj2 Me, FadeDirection, FadeCycles
    End If

End Sub
