VERSION 5.00
Begin VB.Form frmAdjustRes 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Resolution"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   281
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   266
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Reset Original Resolution"
      Height          =   495
      Left            =   1320
      TabIndex        =   6
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Left            =   2520
      Top             =   2760
   End
   Begin VB.TextBox txtCurrentRes 
      Height          =   285
      Left            =   2280
      TabIndex        =   4
      Top             =   2280
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Left            =   2280
      Picture         =   "frmAdjustRes.frx":0000
      ScaleHeight     =   1755
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdSet 
      Appearance      =   0  'Flat
      Caption         =   "Set Resolution"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   1095
   End
   Begin VB.ListBox lstRes 
      Height          =   3375
      ItemData        =   "frmAdjustRes.frx":0F4F
      Left            =   120
      List            =   "frmAdjustRes.frx":0F51
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      Caption         =   "Cancel / Exit"
      Height          =   495
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label lblCurrent 
      Caption         =   "Current Resolution:"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   2040
      Width           =   1695
   End
End
Attribute VB_Name = "frmAdjustRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdSet_Click()
   ResChange (lstRes.Text)
   Unload Me
End Sub

Private Sub Command1_Click()
   ResChange (oRES)
   Unload Me
End Sub

Private Sub Form_Activate()
   'adds resolutions to listbox
   Dim X As Integer
   For X = 1 To NumModes
      lstRes.AddItem (Res(X))
   Next X
End Sub

Private Sub lstRes_DblClick()
   ResChange (lstRes.Text)
   Unload Me
End Sub

