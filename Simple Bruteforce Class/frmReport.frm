VERSION 5.00
Begin VB.Form frmReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test Results"
   ClientHeight    =   3465
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2047.236
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Report"
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.TextBox txtReport 
         Appearance      =   0  'Flat
         Height          =   2415
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCopy_Click()
Clipboard.SetText txtReport
End Sub
