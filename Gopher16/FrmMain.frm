VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gopher16 Hasher"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Clear form"
      Height          =   255
      Left            =   2880
      TabIndex        =   9
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hash text"
      Height          =   255
      Left            =   4200
      TabIndex        =   8
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Result"
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   5295
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   5055
      End
      Begin VB.Label Label3 
         Caption         =   "Returned hash"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   5055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Input"
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FrmMain.frx":0000
         Left            =   120
         List            =   "FrmMain.frx":000A
         TabIndex        =   2
         Text            =   "True (Longer Time)"
         Top             =   3360
         Width           =   5055
      End
      Begin VB.TextBox Text1 
         Height          =   2415
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   600
         Width           =   5055
      End
      Begin VB.Label Label2 
         Caption         =   "Return Mixed-case hash:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   3120
         Width           =   5055
      End
      Begin VB.Label Label1 
         Caption         =   "Text to be hashed by Gopher16():"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   3855
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim TTime As String

    Command1.Enabled = False
    Command1.Caption = "Working..."
    Text1.Enabled = False
    
    If Left(Combo1.Text, 4) = "True" Then
        Text2.Text = Gopher16(Text1.Text, True, TTime)
        Me.Caption = "Gopher16 Hasher - Hashed in: " & TTime
    Else
        Text2.Text = Gopher16(Text1.Text, , TTime)
        Me.Caption = "Gopher16 Hasher - Hashed in: " & TTime
    End If
    
    Text1.Enabled = True
    Command1.Caption = "Hash Text!"
    Command1.Enabled = True
    Beep
End Sub

Private Sub Command2_Click()
    Text1.Text = ""
    Text2.Text = ""
    Combo1.Text = "True"
End Sub
