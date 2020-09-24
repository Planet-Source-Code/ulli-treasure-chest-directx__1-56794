VERSION 5.00
Begin VB.Form fTreasureChest 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "A TREASURE CHEST"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   7500
   StartUpPosition =   2  'Bildschirmmitte
End
Attribute VB_Name = "fTreasureChest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private FPSLastCheck    As Long
Private FPSCount        As Long


Private Sub Form_Load()

  Dim Move As Single

    Busy = Initialise()
    Me.Show
    Do While Busy
        Move = Move + Pi / 36
        If Move > 10000 Then
            Move = 0
        End If
        If GetTickCount() - FPSLastCheck >= 1000 Then
            FPSCurrent = FPSCount
            FPSCount = 0
            FPSLastCheck = GetTickCount()
        End If
        FPSCount = FPSCount + 1
        Render Move
        DoEvents
    Loop

    Set D3DDevice = Nothing
    Set D3D = Nothing
    Set Dx = Nothing

    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Cancel = Busy
    Busy = False

End Sub

':) Ulli's VB Code Formatter V2.17.4 (2004-Okt-17 20:05)  Decl: 2  Code: 50  Total: 52 Lines
':) CommentOnly: 0 (0%)  Commented: 0 (0%)  Empty: 18 (34,6%)  Max Logic Depth: 3
