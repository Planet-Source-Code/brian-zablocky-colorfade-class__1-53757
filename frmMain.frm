VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ColorFade Class Tester"
   ClientHeight    =   4305
   ClientLeft      =   2355
   ClientTop       =   2385
   ClientWidth     =   8160
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   8160
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdGO 
      Caption         =   "FADE"
      Default         =   -1  'True
      Height          =   465
      Left            =   5265
      TabIndex        =   2
      Top             =   1080
      Width           =   1545
   End
   Begin VB.Frame frameColor 
      Caption         =   "End Color"
      Height          =   735
      Index           =   1
      Left            =   6120
      TabIndex        =   23
      Top             =   180
      Width           =   1905
      Begin VB.PictureBox picColor 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   90
         ScaleHeight     =   315
         ScaleWidth      =   360
         TabIndex        =   24
         Top             =   270
         Width           =   420
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "Pick Color"
         Height          =   375
         Index           =   1
         Left            =   585
         TabIndex        =   1
         Top             =   270
         Width           =   1185
      End
   End
   Begin VB.Frame frameColor 
      Caption         =   "Start Color"
      Height          =   735
      Index           =   0
      Left            =   4050
      TabIndex        =   21
      Top             =   180
      Width           =   1905
      Begin VB.CommandButton cmdColor 
         Caption         =   "Pick Color"
         Height          =   375
         Index           =   0
         Left            =   585
         TabIndex        =   0
         Top             =   270
         Width           =   1185
      End
      Begin VB.PictureBox picColor 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   90
         ScaleHeight     =   315
         ScaleWidth      =   360
         TabIndex        =   22
         Top             =   270
         Width           =   420
      End
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   2475
      Top             =   4275
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Gradient Options"
      Height          =   2490
      Left            =   4005
      TabIndex        =   18
      Top             =   1710
      Width           =   4020
      Begin MSComCtl2.UpDown spnCycles 
         Height          =   285
         Left            =   3105
         TabIndex        =   13
         Top             =   1710
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "txtCycles"
         BuddyDispid     =   196614
         OrigLeft        =   3285
         OrigTop         =   1710
         OrigRight       =   3480
         OrigBottom      =   1995
         Max             =   12
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtCycles 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2790
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "0"
         Top             =   1710
         Width           =   315
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   960
         Left            =   2205
         ScaleHeight     =   960
         ScaleWidth      =   1590
         TabIndex        =   20
         Top             =   360
         Width           =   1590
         Begin VB.OptionButton optApply 
            Caption         =   "New Form"
            Height          =   195
            Index           =   2
            Left            =   90
            TabIndex        =   12
            Top             =   585
            Width           =   1320
         End
         Begin VB.OptionButton optApply 
            Caption         =   "Tip Box"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   11
            Top             =   315
            Width           =   1320
         End
         Begin VB.OptionButton optApply 
            Caption         =   "This Form"
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   10
            Top             =   45
            Value           =   -1  'True
            Width           =   1320
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1050
         Left            =   180
         ScaleHeight     =   1050
         ScaleWidth      =   2085
         TabIndex        =   19
         Top             =   1305
         Width           =   2085
         Begin VB.OptionButton optCorner 
            Caption         =   "LowerRight"
            Enabled         =   0   'False
            Height          =   195
            Index           =   3
            Left            =   540
            TabIndex        =   9
            Top             =   810
            Width           =   1815
         End
         Begin VB.OptionButton optCorner 
            Caption         =   "LowerLeft"
            Enabled         =   0   'False
            Height          =   195
            Index           =   2
            Left            =   540
            TabIndex        =   8
            Top             =   540
            Width           =   1815
         End
         Begin VB.OptionButton optCorner 
            Caption         =   "UpperRight"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   540
            TabIndex        =   7
            Top             =   270
            Width           =   1815
         End
         Begin VB.OptionButton optCorner 
            Caption         =   "UpperLeft"
            Enabled         =   0   'False
            Height          =   195
            Index           =   0
            Left            =   540
            TabIndex        =   6
            Top             =   0
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin VB.OptionButton optType 
         Caption         =   "Diagonal"
         Height          =   240
         Index           =   2
         Left            =   315
         TabIndex        =   5
         Top             =   990
         Width           =   1500
      End
      Begin VB.OptionButton optType 
         Caption         =   "Vertical"
         Height          =   240
         Index           =   1
         Left            =   315
         TabIndex        =   4
         Top             =   675
         Width           =   1500
      End
      Begin VB.OptionButton optType 
         Caption         =   "Horizontal"
         Height          =   240
         Index           =   0
         Left            =   315
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1635
      End
      Begin VB.Label lblJunk 
         Alignment       =   2  'Center
         Caption         =   "Cycles"
         Height          =   240
         Left            =   2610
         TabIndex        =   26
         Top             =   1485
         Width           =   780
      End
   End
   Begin VB.PictureBox picTip 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4065
      Left            =   120
      Picture         =   "frmMain.frx":0442
      ScaleHeight     =   4005
      ScaleWidth      =   3675
      TabIndex        =   15
      Top             =   120
      Width           =   3735
      Begin VB.CommandButton cmdNextTip 
         Caption         =   "&Next Tip"
         Height          =   375
         Left            =   2385
         TabIndex        =   14
         Top             =   90
         Width           =   1215
      End
      Begin VB.Label lblTipCount 
         BackStyle       =   0  'Transparent
         Caption         =   "0 Different Tips"
         Height          =   240
         Left            =   135
         TabIndex        =   27
         Top             =   3690
         Width           =   2490
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Did you know..."
         Height          =   255
         Left            =   540
         TabIndex        =   17
         Top             =   180
         Width           =   1755
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Height          =   2625
         Left            =   180
         TabIndex        =   16
         Top             =   810
         Width           =   3345
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The in-memory database of tips.
Dim Tips As New Collection

' Name of tips file
Const TIP_FILE = "TIPOFDAY.TXT"

' Index in collection of tip currently being displayed.
Dim CurrentTip As Long


Private Sub DoNextTip()

    Dim OldTip As Long
        
    CurrentTip = CurrentTip + 1
    If Tips.Count < CurrentTip Then
        CurrentTip = 1
    End If
    
    ' Show it.
    frmMain.DisplayCurrentTip
    
End Sub

Function LoadTips(sFile As String) As Boolean
    Dim NextTip As String   ' Each tip read in from file.
    Dim InFile As Integer   ' Descriptor for file.
    Dim TipCount As Long
    
    ' Obtain the next free file descriptor.
    InFile = FreeFile
    
    ' Make sure a file is specified.
    If sFile = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Make sure the file exists before trying to open it.
    If Dir(sFile) = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Read the collection from a text file.
    Open sFile For Input As InFile
    While Not EOF(InFile)
        Line Input #InFile, NextTip
        Tips.Add NextTip
        TipCount = TipCount + 1
    Wend
    Close InFile

    ' Display a tip at random.
    DoNextTip
    
    LoadTips = True
    lblTipCount.Caption = Trim(Str(TipCount)) & " Different Tips"
End Function



Private Sub cmdColor_Click(Index As Integer)
    
    If Index = 0 Then
        cdlg.Color = Fade.FadeStartColor
        cdlg.ShowColor
        Fade.FadeStartColor = cdlg.Color
        picColor(0).BackColor = cdlg.Color
    Else
        cdlg.Color = Fade.FadeEndColor
        cdlg.ShowColor
        Fade.FadeEndColor = cdlg.Color
        picColor(1).BackColor = cdlg.Color
    End If
    
End Sub








Private Sub cmdGO_Click()

    If optApply(0).Value Then
        If FadeCycles = 0 Then
            Fade.PaintObj Me, FadeDirection
        Else
            Fade.PaintObj2 Me, FadeDirection, FadeCycles
        End If
    End If
    
    
    If optApply(1).Value Then
        If FadeCycles = 0 Then
            Fade.PaintObj Me.picTip, FadeDirection
        Else
            Fade.PaintObj2 Me.picTip, FadeDirection, FadeCycles
        End If
    End If
    
    
    If optApply(2).Value Then
        frmTest.Show 1
    End If

End Sub






Private Sub cmdNextTip_Click()
    DoNextTip
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    ' Seed Rnd
    Randomize
    
    ' Read in the tips file and display a tip at random.
    If LoadTips(App.Path & "\" & TIP_FILE) = False Then
        lblTipText.Caption = "That the " & TIP_FILE & " file was not found? " & vbCrLf & vbCrLf & _
           "Create a text file named " & TIP_FILE & " using NotePad with 1 tip per line. " & _
           "Then place it in the same directory as the application. "
    End If

    
End Sub

Public Sub DisplayCurrentTip()
    If Tips.Count > 0 Then
        lblTipText.Caption = Tips.Item(CurrentTip)
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    End
End Sub



Private Sub optCorner_Click(Index As Integer)
    
    If optCorner(0).Value Then FadeDirection = DiagUpperLeftGradient
    If optCorner(1).Value Then FadeDirection = DiagUpperRightGradient
    If optCorner(2).Value Then FadeDirection = DiagLowerLeftGradient
    If optCorner(3).Value Then FadeDirection = DiagLowerRightGradient

End Sub

Private Sub optType_Click(Index As Integer)
    
    Dim i As Integer
    For i = 0 To 3
        optCorner(i).Enabled = optType(2).Value
    Next i
    
    Select Case Index
        Case 0
            FadeDirection = HorizontalGradient
        Case 1
            FadeDirection = VerticalGradient
        Case 2
            If optCorner(0).Value Then FadeDirection = DiagUpperLeftGradient
            If optCorner(1).Value Then FadeDirection = DiagUpperRightGradient
            If optCorner(2).Value Then FadeDirection = DiagLowerLeftGradient
            If optCorner(3).Value Then FadeDirection = DiagLowerRightGradient
    End Select
                
End Sub

Private Sub spnCycles_Change()
    FadeCycles = CLng(spnCycles.Value)
End Sub
