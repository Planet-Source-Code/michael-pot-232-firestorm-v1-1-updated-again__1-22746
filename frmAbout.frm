VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   3495
   ClientLeft      =   3060
   ClientTop       =   2640
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   233
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   435
      Left            =   1785
      TabIndex        =   0
      Top             =   2895
      Width           =   1200
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   $"frmAbout.frx":0000
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   4635
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "FIRESTORM V1.0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   30
      TabIndex        =   1
      Top             =   180
      Width           =   4635
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Endd As Boolean, I As Long
Private Sub Command1_Click()
frmAbout.Hide
frmMain.Command3_Click
End Sub

Public Sub Doeffect()
Show 1, frmMain
End Sub

