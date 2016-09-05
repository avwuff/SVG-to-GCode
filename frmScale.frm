VERSION 5.00
Begin VB.Form frmScale 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scale"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   3855
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtScale 
      Height          =   315
      Left            =   1260
      TabIndex        =   7
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   360
      TabIndex        =   6
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2100
      TabIndex        =   5
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CheckBox chkAspect 
      Caption         =   "Keep Aspect Ratio"
      Height          =   255
      Left            =   1260
      TabIndex        =   4
      Top             =   1500
      Value           =   1  'Checked
      Width           =   1875
   End
   Begin VB.TextBox txtHeight 
      Height          =   315
      Left            =   1260
      TabIndex        =   3
      Top             =   660
      Width           =   1335
   End
   Begin VB.TextBox txtWidth 
      Height          =   315
      Left            =   1260
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Scale:"
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
      TabIndex        =   8
      Top             =   1140
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Height:"
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
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Width:"
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
      TabIndex        =   0
      Top             =   300
      Width           =   975
   End
End
Attribute VB_Name = "frmScale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public updatingValue As Boolean
Public originalAspect As Double

Public setW As Double
Public setH As Double
Public originalW As Double
Public originalH As Double


Private Sub cmdApply_Click()
    setW = Val(txtWidth)
    setH = Val(txtHeight)
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
     Unload Me
     
     
End Sub

Private Sub txtHeight_Change()
    If updatingValue Then Exit Sub
    
    ' Calculate the new wodtj
    If chkAspect.Value = vbChecked Then
        updatingValue = True
        txtWidth = Round(Val(txtHeight) * originalAspect, 5)
        txtScale = Round(Val(txtWidth) / originalW * 100, 2)
        updatingValue = False
        
        
    End If

End Sub

Private Sub txtWidth_Change()
    If updatingValue Then Exit Sub
    
    ' Calculate the new height
    If chkAspect.Value = vbChecked Then
        updatingValue = True
        txtHeight = Round(Val(txtWidth) / originalAspect, 5)
        txtScale = Round(Val(txtWidth) / originalW * 100, 2)
        updatingValue = False
        
        
    End If

End Sub


Private Sub txtScale_Change()
    If updatingValue Then Exit Sub
    
    ' Calculate the new height
        updatingValue = True
        txtHeight = Round(originalH * (Val(txtScale) / 100), 5)
        txtWidth = Round(originalW * (Val(txtScale) / 100), 5)
        updatingValue = False
        
    

End Sub
