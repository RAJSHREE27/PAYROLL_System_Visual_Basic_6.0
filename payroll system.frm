VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H8000000D&
   Caption         =   "Form2"
   ClientHeight    =   4530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6210
   LinkTopic       =   "Form2"
   ScaleHeight     =   4530
   ScaleWidth      =   6210
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "ADM LOGIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      MaskColor       =   &H000080FF&
      TabIndex        =   7
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   6
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "EMP LOGIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      MaskColor       =   &H000080FF&
      TabIndex        =   5
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000A&
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2520
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000C&
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   120
      Picture         =   "payroll system.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FFFF&
      Caption         =   "PASSWORD :"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "USER ID :"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "RKLS PRODUCTIONS PVT. LTD."
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   600
      Width           =   4095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Command1_Click()
DataEnvironment1.Search Text1.Text
With DataEnvironment1.rsSearch

If DataEnvironment1.rsSearch.EOF Then
MsgBox "Record Of Employee Not Found"
Else
If (Text1.Text = DataEnvironment1.rsSearch.Fields("UserID")) And (Text2.Text = DataEnvironment1.rsSearch.Fields("Password")) Then
Form3.Show
Form3.Text16.Text = Form2.Text1.Text


Unload Me
Else
MsgBox "Invalid Credentials"
End If
End If
.Close
End With

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
If (Text1.Text = "160000") And (Text2.Text = "admin123") Then
Form1.Show
Unload Me
Else
MsgBox "Invalid Credentials"
End If

End Sub




