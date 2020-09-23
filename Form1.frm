VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Encode Decode base 64"
   ClientHeight    =   2685
   ClientLeft      =   3915
   ClientTop       =   1365
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   4110
   Begin VB.CommandButton Command4 
      Caption         =   "Clear"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Decode"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encode"
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2040
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Base64 encoded text below. Press decode to view plain text above."
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Type or paste text below. Press encode to view Base64 encoding."
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Dim strToEncode As String
    
    strToEncode = Text1
    Text2 = EncodeStr64(strToEncode)
    
End Sub

Private Sub Command2_Click()
    Text1 = ""
End Sub

Private Sub Command3_Click()
    Dim strToDecode As String
    
    strToDecode = Text2
    Text1 = DecodeStr64(strToDecode)
    
End Sub

Private Sub Command4_Click()
    Text2 = ""
End Sub
