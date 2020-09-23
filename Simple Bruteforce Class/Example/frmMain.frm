VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Character Generation Speed Test"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Statistics"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   4455
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Combination Length:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   20
         Top             =   600
         Width           =   1485
      End
      Begin VB.Label lblComboLen 
         Height          =   255
         Left            =   2640
         TabIndex        =   19
         Top             =   600
         Width           =   1695
      End
      Begin VB.Line Line1 
         X1              =   360
         X2              =   4200
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label lblTotalCombo 
         Height          =   255
         Left            =   2640
         TabIndex        =   11
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label lblComboPS 
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblCurrentCOmbo 
         Height          =   255
         Left            =   2640
         TabIndex        =   9
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Current Combination:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   1545
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Total Combinations:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   1320
         Width           =   1425
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Combinations Per Second:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   1080
         Width           =   1875
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtStartCombo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   18
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtEndLen 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   16
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txtStartLen 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start"
         Height          =   375
         Left            =   3120
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtTarget 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "target"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtCharacterset 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Starting Combination"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2520
         TabIndex        =   17
         Top             =   1560
         Width           =   1500
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "End Length:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1320
         TabIndex        =   15
         Top             =   1560
         Width           =   870
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Start Length:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1560
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Target Combination:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Character Set:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuBenchmark 
      Caption         =   "Benchmark"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private WithEvents cBrute As clsBruteforce
Attribute cBrute.VB_VarHelpID = -1
Dim bStop As Boolean
Dim intBenchmarkSec As Integer


Private Sub cBrute_Statistics(lngTotalCombos As Long, lngCombinationsPerSec As Long, strCurrentPassword As String)
Static lngSeconds As Long

    lblComboPS = Format(lngCombinationsPerSec, "#,###,##0")
    lblCurrentCOmbo = strCurrentPassword
    lblComboLen = Len(strCurrentPassword)
    lblTotalCombo = Format(lngTotalCombos, "###,###,###,##0")
    
    If intBenchmarkSec > 0 And lngSeconds >= intBenchmarkSec Then
        bStop = True
        lngSeconds = 0
    Else
        lngSeconds = lngSeconds + 1
    End If
End Sub

Private Sub cmdStart_Click()
txtCharacterset.Enabled = Not txtCharacterset.Enabled
txtStartLen.Enabled = Not txtStartLen.Enabled
txtEndLen.Enabled = Not txtEndLen.Enabled
txtStartCombo.Enabled = Not txtStartCombo.Enabled
txtTarget.Enabled = Not txtTarget.Enabled

If cmdStart.Caption = "Start" Then
    cmdStart.Caption = "Stop"
    
    With cBrute
        .CharacterSet = txtCharacterset
        .StartLength = Val(txtStartLen)
        .EndLength = Val(txtEndLen)
        If Len(txtStartCombo) > 0 Then .StartWord = txtStartCombo
    End With
    
    bStop = False
    intBenchmarkSec = 0
    DummyLoop
Else
    cmdStart.Caption = "Start"
    bStop = True
End If
End Sub

Private Sub Form_Load()
Set cBrute = New clsBruteforce

With cBrute
    txtCharacterset = .CharacterSet
    txtStartLen = .StartLength
    txtEndLen = .EndLength
    txtStartCombo = .StartWord
End With

End Sub

Sub DummyLoop()
Dim strTarget As String

strTarget = txtTarget

Do While bStop <> True
    If cBrute.BruteForce = strTarget Then
        cBrute.ForceCombinationCalc
        MsgBox "Combination Reached", vbApplicationModal + vbInformation, Me.Caption
        cmdStart_Click
    End If
Loop

cBrute.ResetStats

End Sub

Private Sub mnuBenchmark_Click()
Dim strNormal As String, strLongCharSet As String, strLongCombo As String, strMessage As String, strAverage As String

intBenchmarkSec = 10
strAverage = 0
strMessage = "This Benchmark Test will take approx. " & intBenchmarkSec * 3 & " Seconds."

MsgBox strMessage, vbApplicationModal + vbInformation, Me.Caption
frmSplash.Show
'=====================================================================================================================
' Normal Settings Test
'=====================================================================================================================

bStop = False

With cBrute
    .CharacterSet = "abcdefghijklmnopqrstuvwxyz"
    .StartLength = 1
    .EndLength = 10
End With

Do While bStop <> True
    Call cBrute.BruteForce
Loop

strAverage = Val(strAverage) + Val(cBrute.sTotalCombo)
strNormal = Format(cBrute.sTotalCombo, "###,###,##0")
cBrute.ResetStats

'=====================================================================================================================
' Long Characterset Test
'=====================================================================================================================

bStop = False

With cBrute
    .CharacterSet = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!@#$%^&*()_+-=[];,./{}:"
    .StartLength = 1
    .EndLength = 10
End With

Do While bStop <> True
    Call cBrute.BruteForce
Loop

strAverage = Val(strAverage) + Val(cBrute.sTotalCombo)
strLongCharSet = Format(cBrute.sTotalCombo, "###,###,##0")
cBrute.ResetStats

'=====================================================================================================================
' Long Combination Test
'=====================================================================================================================

bStop = False

With cBrute
    .StartLength = 100
    .EndLength = 110
End With

Do While bStop <> True
    Call cBrute.BruteForce
Loop

strAverage = Val(strAverage) + Val(cBrute.sTotalCombo)
strLongCombo = Format(cBrute.sTotalCombo, "###,###,##0")
cBrute.ResetStats

Unload frmSplash

strMessage = "CPU: " & GetProcessor() & vbCrLf & vbCrLf
strMessage = strMessage & "Results over " & intBenchmarkSec & "sec" & vbCrLf & vbCrLf
strMessage = strMessage & "Normal Test: " & strNormal & vbCrLf
strMessage = strMessage & "Long Character set Test: " & strLongCharSet & vbCrLf
strMessage = strMessage & "Long Combination Test: " & strLongCombo & vbCrLf & vbCrLf
strMessage = strMessage & "Your Average speed was: " & Format(Val(Val(strAverage / 3) / 10), "###,###,##0")

frmReport.Show
frmReport.txtReport = strMessage
End Sub
