VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Triple Click Text Example..."
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   315
      Left            =   3270
      TabIndex        =   2
      Top             =   1740
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   270
      TabIndex        =   0
      Text            =   "Just triple click me to see the action!"
      Top             =   570
      Width           =   3975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Triple click TextBox to select whole string..."
      Height          =   195
      Left            =   300
      TabIndex        =   1
      Top             =   270
      Width           =   3060
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private WithEvents NewObject As clsTripleClick
Attribute NewObject.VB_VarHelpID = -1

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set NewObject = New clsTripleClick
    Set NewObject.vObject = Text1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set NewObject = Nothing
End Sub

Private Sub NewObject_TripleClick()
    'Select whole text when triple clicked
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
End Sub


