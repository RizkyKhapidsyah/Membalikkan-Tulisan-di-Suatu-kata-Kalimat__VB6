VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Membalikkan Tulisan di Suatu kata/Kalimat"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6555
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   840
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   2160
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Created by Rizky Khapidsyah
'Source Code Program Dimulai Dari Sini

Private Sub Command1_Click()
  Text1.Text = BalikkanString(Text1.Text)
End Sub

Function BalikkanString(strKalimat As String) As String
Dim i As Integer, Panjang As Integer
Dim strTampung As String
  Panjang = Len(strKalimat)
   For i = Panjang To 1 Step -1
      strTampung = strTampung & Mid(strKalimat, i, 1)
   Next i
   BalikkanString = strTampung
End Function

Private Sub Form_Load()
Text1.Text = "Rizky Khapidsyah"
End Sub
