VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Made by Roni Tovi - Any questions are welcome at root@mutluhost.com"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   10320
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Please note:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   4440
      Width           =   9975
      Begin VB.Label Label2 
         Caption         =   $"Form1.frx":0000
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   9495
      End
      Begin VB.Label Label1 
         Caption         =   $"Form1.frx":00C1
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   9495
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Maximize the form"
      Height          =   615
      Left            =   6960
      TabIndex        =   3
      Top             =   3720
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Bind to Microsoft Access Database"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Bind to MS SQL Database"
      Height          =   615
      Left            =   3840
      TabIndex        =   1
      Top             =   3720
      Width           =   3015
   End
   Begin Project1.UserControl1 uc 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   6376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
uc.ConnString = "Driver={SQL Native Client};Server=localhost;Database=medi;Trusted_Connection=yes;"
If InStr(1, uc.ConnString, "medi") > 0 Then
    MsgBox "You should edit the ConnectionString relative to your MS SQL database first!", vbInformation
    Exit Sub
End If
uc.BindToSQL "SELECT * FROM departmanlar"
End Sub

Private Sub Command2_Click()
uc.ConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dtn.mdb;Persist Security Info=False"
uc.BindToSQL "SELECT * FROM tab1"
End Sub

Private Sub Command3_Click()
    If WindowState = vbMaximized Then
        WindowState = vbNormal
        Command3.Caption = "Maximize the form"
    Else
        WindowState = vbMaximized
        Command3.Caption = "Go back to normal mode"
    End If

End Sub

Private Sub Form_Resize()
On Error Resume Next
uc.Move 0, 0, Me.ScaleWidth
Frame1.Move Frame1.Left, Frame1.Top, Me.ScaleWidth - (Frame1.Left * 2)
Label1.Move Label1.Left, Label1.Top, Frame1.Width - (Label1.Left * 2)
Label2.Move Label2.Left, Label2.Top, Frame1.Width - (Label2.Left * 2)
End Sub
