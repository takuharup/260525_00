VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormNodeSelect
   Caption         =   "節点番号選択"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4215
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.Label lblInstruction
      Caption         =   "転記する節点番号を選択してください（複数選択可）"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
   End
   Begin MSForms.ListBox lstNodes
      Height          =   3615
      Left            =   120
      MultiSelect     =   1
      TabIndex        =   0
      Top             =   600
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   6376
   End
   Begin MSForms.CommandButton btnOK
      Caption         =   "OK"
      Height          =   435
      Left            =   840
      TabIndex        =   1
      Top             =   4800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   767
   End
   Begin MSForms.CommandButton btnCancel
      Caption         =   "キャンセル"
      Height          =   435
      Left            =   2400
      TabIndex        =   2
      Top             =   4800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   767
   End
End
Attribute VB_Name = "FormNodeSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnOK_Click()
    Me.Tag = "OK"
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    Me.Tag = "Cancel"
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Me.Tag = "Cancel"
        Me.Hide
        Cancel = True
    End If
End Sub
