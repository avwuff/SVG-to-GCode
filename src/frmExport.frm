VERSION 5.00
Begin VB.Form frmExport 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Export GCode"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export Now"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   6300
      TabIndex        =   19
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Loop Cut while raising table"
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
      Left            =   4020
      TabIndex        =   8
      Top             =   600
      Width           =   4095
      Begin VB.TextBox txtMM 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   17
         Text            =   "1"
         Top             =   1620
         Width           =   975
      End
      Begin VB.TextBox txtLoops 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   15
         Text            =   "6"
         Top             =   840
         Width           =   975
      End
      Begin VB.CheckBox chkLoop 
         Caption         =   "Perform job multiple times"
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
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "mm"
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
         TabIndex        =   18
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Raise the bed this much after each loop:"
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
         Left            =   240
         TabIndex        =   16
         Top             =   1320
         Width           =   2910
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "This many loops:"
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
         Left            =   240
         TabIndex        =   14
         Top             =   900
         Width           =   1200
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Make it easy to cut through heavy plastics."
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
         Left            =   240
         TabIndex        =   12
         Top             =   300
         Width           =   3105
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Export Options"
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
      Left            =   180
      TabIndex        =   3
      Top             =   600
      Width           =   3675
      Begin VB.TextBox txtPPI 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1140
         TabIndex        =   11
         Text            =   "111111"
         Top             =   1260
         Width           =   1335
      End
      Begin VB.CheckBox chkPPI 
         Caption         =   "PPI Mode"
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
         Left            =   180
         TabIndex        =   9
         Top             =   1020
         Width           =   2115
      End
      Begin VB.TextBox txtFeedRate 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1140
         TabIndex        =   6
         Text            =   "20"
         Top             =   300
         Width           =   1335
      End
      Begin VB.CheckBox chkZPlunge 
         Caption         =   "Z-plunge (for engraver)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "PPI:"
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
         Left            =   540
         TabIndex        =   10
         Top             =   1320
         Width           =   300
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "in/min"
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
         Left            =   2580
         TabIndex        =   7
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Feed Rate:"
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
         Left            =   180
         TabIndex        =   5
         Top             =   360
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdChoosePath 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7500
      TabIndex        =   2
      Top             =   120
      Width           =   675
   End
   Begin VB.TextBox txtPath 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1260
      TabIndex        =   1
      Top             =   120
      Width           =   6195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Export Path:"
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
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   915
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdChoosePath_Click()

  
    
    With frmInterface.COMDLG
        
        .FileName = getFolderNameFromPath(.FileName) & "\" & getFileNameNoExten(getFileNameFromPath(.FileName)) & ".ngc"
        
        .Filter = "GCODE Files (*.ngc)|*.ngc"
        
        .DialogTitle = "Export GCODE"
        .ShowSave
        .CancelError = False
        txtPath.Text = .FileName
    End With
    
    
    
End Sub

Private Sub cmdExport_Click()
    If txtPath = "" Then
        MsgBox "Please specify an export path.", vbInformation
        Exit Sub
    End If
    
    If myDir(txtPath) <> "" Then
        If MsgBox("The file already exists.  Overwrite?", vbYesNo Or vbQuestion) <> vbYes Then Exit Sub
    End If

    ' Save everything.
    SetSet "FeedRate", txtFeedRate.Text
    SetSet "ZPlunge", IIf(FromCheck(chkZPlunge), "Y", "N")
    SetSet "PPI", IIf(FromCheck(chkPPI), "Y", "N")
    SetSet "PPI Rate", txtPPI.Text
    SetSet "Loop", IIf(FromCheck(chkLoop), "Y", "N")
    SetSet "Loops", txtLoops.Text
    SetSet "RaiseDist", txtMM.Text
    LastExportPath = txtPath.Text
    

    exportGCODE txtPath.Text, Val(txtFeedRate.Text), FromCheck(chkZPlunge), _
        FromCheck(chkPPI), Val(txtPPI.Text), _
        FromCheck(chkLoop), Val(txtLoops.Text), Val(txtMM.Text)
        
    MsgBox "Export complete!", vbInformation
    

End Sub

Private Function SetSet(Sett As String, Value As String)
    SaveSetting "Av's SVG to GCode", "Export", Sett, Value
End Function
Private Function GetSet(Sett As String, Optional DefaultValue As String) As String
    GetSet = GetSetting("Av's SVG to GCode", "Export", Sett, DefaultValue)
End Function

Private Sub Form_Load()

    txtFeedRate.Text = GetSet("FeedRate", "20")
    If GetSet("ZPlunge") = "Y" Then chkZPlunge.Value = vbChecked
    If GetSet("PPI") = "Y" Then chkPPI.Value = vbChecked
    If GetSet("Loop") = "Y" Then chkLoop.Value = vbChecked
    txtPPI.Text = GetSet("PPI", "111111")
    txtLoops.Text = GetSet("Loops", "6")
    txtMM.Text = GetSet("RaiseDist", "0.5")
    txtPath.Text = LastExportPath
    
    
    If txtPath = "" Then
        txtPath = getFolderNameFromPath(CurrentFile) & "\" & getFileNameNoExten(getFileNameFromPath(CurrentFile)) & ".ngc"

    End If

End Sub
