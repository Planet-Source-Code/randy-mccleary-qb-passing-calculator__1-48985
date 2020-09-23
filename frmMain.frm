VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Football QB Passing Calculator"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6900
   DrawStyle       =   5  'Transparent
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":08CA
   ScaleHeight     =   250
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   460
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Calculated Stats:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   2655
      Left            =   3480
      TabIndex        =   11
      Top             =   240
      Width           =   3180
      Begin VB.Label lblPassEff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1920
         TabIndex        =   19
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000014&
         Caption         =   "Passing Efficiency:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2070
         Width           =   1695
      End
      Begin VB.Label Label13 
         BackColor       =   &H80000014&
         Caption         =   "Yards / Completion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1470
         Width           =   1695
      End
      Begin VB.Label lblYardsPerComp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1920
         TabIndex        =   16
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblYardsPerAtt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1920
         TabIndex        =   15
         Top             =   825
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000014&
         Caption         =   "Yards per Attempt:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   885
         Width           =   1695
      End
      Begin VB.Label lblCompletionPct 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1920
         TabIndex        =   13
         Top             =   315
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000014&
         Caption         =   "Completion %:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Input QB Passing Stats:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3180
      Begin VB.TextBox txtINTs 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   10
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtTDs 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   9
         Top             =   1710
         Width           =   975
      End
      Begin VB.TextBox txtYards 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   8
         Top             =   1260
         Width           =   975
      End
      Begin VB.TextBox txtCompletions 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   2
         Top             =   810
         Width           =   975
      End
      Begin VB.TextBox txtAttempts 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000014&
         Caption         =   "Interceptions:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000014&
         Caption         =   "Passing TD's:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1725
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000014&
         Caption         =   "Total Yards:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1290
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000014&
         Caption         =   "Attempts:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000014&
         Caption         =   "Completions:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   855
         Width           =   1335
      End
   End
   Begin VB.Image imgExit 
      Height          =   570
      Left            =   4320
      Picture         =   "frmMain.frx":11FE
      Top             =   3000
      Width           =   1650
   End
   Begin VB.Image imgClear 
      Height          =   570
      Left            =   2640
      Picture         =   "frmMain.frx":21EF
      Top             =   3000
      Width           =   1650
   End
   Begin VB.Image imgCalculate 
      Height          =   570
      Left            =   960
      Picture         =   "frmMain.frx":32DB
      Top             =   3000
      Width           =   1650
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub imgCalculate_Click()
   Dim lngAttempts As Long
   Dim lngCompletions As Long
   Dim lngYards As Long
   Dim lngTDs As Long
   Dim lngINTs As Long
   Dim dblCompletion As Double
   Dim dblYardsPerAtt As Double
   Dim dblYardsPerComp As Double
   Dim dblPassEff As Double
   
   '** Get the value from the input textboxes ***
   lngAttempts = Val(txtAttempts.Text)
   lngCompletions = Val(txtCompletions.Text)
   lngYards = Val(txtYards.Text)
   lngTDs = Val(txtTDs.Text)
   lngINTs = Val(txtINTs.Text)
   
   '** If Attempts or Completions is = to 0 then exit sub ***
   If lngAttempts <= 0 Or lngCompletions <= 0 Then
      Exit Sub
   End If
   
   '** Calculate the stats ***
   dblCompletion = (lngCompletions / lngAttempts) * 100
   dblYardsPerAtt = lngYards / lngAttempts
   dblYardsPerComp = lngYards / lngCompletions
   dblPassEff = (dblCompletion) + (8.4 * dblYardsPerAtt) + (330 * (lngTDs / lngAttempts)) + (-200 * (lngINTs / lngAttempts))

   '** Display the passing stats ***
   lblCompletionPct.Caption = FormatNumber(dblCompletion, 1) & "  "
   lblYardsPerAtt.Caption = FormatNumber(dblYardsPerAtt, 1) & "  "
   lblYardsPerComp.Caption = FormatNumber(dblYardsPerComp, 1) & "  "
   lblPassEff.Caption = FormatNumber(dblPassEff, 1) & "  "
End Sub

Private Sub imgClear_Click()
   txtAttempts.Text = ""
   txtCompletions.Text = ""
   txtYards.Text = ""
   txtTDs.Text = ""
   txtINTs.Text = ""
   
   lblCompletionPct.Caption = ""
   lblYardsPerAtt.Caption = ""
   lblYardsPerComp.Caption = ""
   lblPassEff.Caption = ""
End Sub

Private Sub imgExit_Click()
   Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Set frmMain = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload Me
End Sub

Private Sub txtAttempts_KeyPress(KeyAscii As Integer)
   If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub txtCompletions_KeyPress(KeyAscii As Integer)
   If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub txtYards_KeyPress(KeyAscii As Integer)
   If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub txtTDs_KeyPress(KeyAscii As Integer)
   If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub txtINTs_KeyPress(KeyAscii As Integer)
   If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then KeyAscii = 0
End Sub

