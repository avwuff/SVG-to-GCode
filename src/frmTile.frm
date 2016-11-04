VERSION 5.00
Begin VB.Form frmTile 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tile"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   2565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   375
      Left            =   1440
      TabIndex        =   12
      Top             =   2340
      Width           =   1035
   End
   Begin VB.TextBox txtInput 
      Height          =   315
      Index           =   0
      Left            =   1440
      TabIndex        =   11
      Text            =   "5"
      Top             =   120
      Width           =   1035
   End
   Begin VB.TextBox txtInput 
      Height          =   315
      Index           =   5
      Left            =   1440
      TabIndex        =   10
      Text            =   "0"
      Top             =   1920
      Width           =   1035
   End
   Begin VB.TextBox txtInput 
      Height          =   315
      Index           =   4
      Left            =   1440
      TabIndex        =   8
      Text            =   "0"
      Top             =   1560
      Width           =   1035
   End
   Begin VB.TextBox txtInput 
      Height          =   315
      Index           =   3
      Left            =   1440
      TabIndex        =   6
      Text            =   "0.1"
      Top             =   1200
      Width           =   1035
   End
   Begin VB.TextBox txtInput 
      Height          =   315
      Index           =   2
      Left            =   1440
      TabIndex        =   4
      Text            =   "0.1"
      Top             =   840
      Width           =   1035
   End
   Begin VB.TextBox txtInput 
      Height          =   315
      Index           =   1
      Left            =   1440
      TabIndex        =   2
      Text            =   "5"
      Top             =   480
      Width           =   1035
   End
   Begin VB.Label Label6 
      Caption         =   "Column Offset"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1980
      Width           =   1155
   End
   Begin VB.Label Label5 
      Caption         =   "Row Offset"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1620
      Width           =   1155
   End
   Begin VB.Label Label4 
      Caption         =   "Height Gap"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1260
      Width           =   1155
   End
   Begin VB.Label Label3 
      Caption         =   "Width Gap"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   900
      Width           =   1155
   End
   Begin VB.Label Label2 
      Caption         =   "Columns (Width)"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   540
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Rows (Height)"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1155
   End
End
Attribute VB_Name = "frmTile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGo_Click()

    ' Save
    Dim i As Long
    For i = 0 To txtInput.ubound
        SaveSetting "SVG to GCODE", "Tile", i, txtInput(i)
    Next
    
    frmInterface.goTile Val(txtInput(0)), Val(txtInput(1)), _
            Val(txtInput(2)), Val(txtInput(3)), _
            Val(txtInput(4)), Val(txtInput(5))

        
End Sub

Private Sub Form_Load()
    ' Load last values
    Dim i As Long
    Dim AA As String
    For i = 0 To txtInput.ubound
        AA = GetSetting("SVG to GCODE", "Tile", i, "")
        If AA <> "" Then
            txtInput(i) = AA
        End If
    Next
End Sub
