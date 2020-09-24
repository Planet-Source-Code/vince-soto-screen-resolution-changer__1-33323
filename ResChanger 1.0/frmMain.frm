VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Resolution Test"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   1680
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdResSet 
      Caption         =   "Set Screen Resolution"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Current Resolution:"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdResSet_Click()
   frmAdjustRes.Show
   Label1.Caption = "Current Resolution:  " & CurrentRes
   Me.Refresh
   Me.WindowState = vbMaximized
End Sub

Private Sub Form_Load()
   Label1.Caption = "Current Resolution:  " & CurrentRes
End Sub
