VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Base Converter"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Convert"
      Height          =   615
      Left            =   1920
      TabIndex        =   3
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   4815
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   1200
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Text = BaseConvert(Trim(Text1.Text), CInt(Trim(Text2.Text)), CInt(Trim(Text3.Text)))
End Sub

