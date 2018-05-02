VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBar 
   Caption         =   "Progress"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5640
   OleObjectBlob   =   "ProgressBar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim progress As Double, maxProgress As Double, maxWidth As Long, startTime As Double
Public Sub Initialize(title As String, Optional max As Long = 100)
'Initialize and shor progress bar
    Me.Caption = title
    Me.ltext.Caption = ""
    maxProgress = max:  maxWidth = lBar.Width:    lBar.Width = 0
    lProgress.Caption = "0%"
    Me.Show False
End Sub
Public Sub AddProgress(Optional inc As Long = 1)
'Increase progress by an increment
    progress = progress + inc
    If progress > maxProgress Then progress = maxProgress
    lBar.Width = CLng(CDbl(progress) / maxProgress * maxWidth)
    lProgress.Caption = "" & CLng(CDbl(progress) / maxProgress * 100) & "%"
    If progress = maxProgress Then Me.Hide
    DoEvents
End Sub

Public Sub AddCaption(text As String)
    ltext.Caption = text
    DoEvents
End Sub
