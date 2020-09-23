VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Firestorm"
   ClientHeight    =   5865
   ClientLeft      =   1845
   ClientTop       =   1845
   ClientWidth     =   6900
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   391
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   460
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command5 
      Caption         =   "Add a preset"
      Height          =   435
      Left            =   270
      TabIndex        =   28
      Top             =   5385
      Width           =   1275
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1725
      TabIndex        =   27
      Text            =   "Presets"
      Top             =   4455
      Width           =   1470
   End
   Begin VB.CommandButton Command4 
      Caption         =   "About"
      Height          =   435
      Left            =   270
      TabIndex        =   26
      Top             =   4905
      Width           =   1275
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   6045
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   25
      Top             =   5055
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Pause"
      Height          =   435
      Left            =   270
      TabIndex        =   24
      Top             =   4440
      Width           =   1275
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   5850
      Top             =   4890
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Random Shift"
      Height          =   240
      Left            =   5010
      TabIndex        =   22
      Top             =   5415
      Width           =   1545
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   3780
      TabIndex        =   20
      Text            =   "1"
      Top             =   5430
      Width           =   960
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   4965
      TabIndex        =   18
      Text            =   "5"
      Top             =   4785
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   3795
      TabIndex        =   16
      Text            =   "100"
      Top             =   4785
      Width           =   960
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   495
      Left            =   3690
      TabIndex        =   3
      Top             =   225
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   873
      _Version        =   393216
      Max             =   60
      SelStart        =   2
      TickFrequency   =   3
      Value           =   2
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load Picture"
      Height          =   600
      Left            =   1710
      TabIndex        =   2
      Top             =   3300
      Width           =   1425
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   600
      Left            =   285
      TabIndex        =   1
      Top             =   3300
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3000
      Left            =   270
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   0
      Top             =   180
      Width           =   3000
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   495
      Left            =   3690
      TabIndex        =   4
      Top             =   1050
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   873
      _Version        =   393216
      Min             =   -60
      Max             =   60
      SelStart        =   4
      TickFrequency   =   3
      Value           =   4
   End
   Begin MSComctlLib.Slider Slider3 
      Height          =   495
      Left            =   3690
      TabIndex        =   8
      Top             =   1860
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   873
      _Version        =   393216
      Max             =   10000
      SelStart        =   1000
      TickFrequency   =   500
      Value           =   1000
   End
   Begin MSComctlLib.Slider Slider4 
      Height          =   495
      Left            =   3690
      TabIndex        =   10
      Top             =   2610
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   873
      _Version        =   393216
      Max             =   200
      SelStart        =   4
      TickFrequency   =   10
      Value           =   4
   End
   Begin MSComctlLib.Slider Slider5 
      Height          =   495
      Left            =   3690
      TabIndex        =   12
      Top             =   3360
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   873
      _Version        =   393216
      Max             =   255
      SelStart        =   102
      TickFrequency   =   8
      Value           =   102
   End
   Begin MSComctlLib.Slider Slider6 
      Height          =   495
      Left            =   3690
      TabIndex        =   14
      Top             =   4080
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   873
      _Version        =   393216
      Min             =   -1000
      Max             =   1000
      SelStart        =   1
      TickFrequency   =   50
      Value           =   1
   End
   Begin VB.Label lblFPS 
      Caption         =   "FPS:"
      Height          =   240
      Left            =   285
      TabIndex        =   23
      Top             =   4185
      Width           =   2850
   End
   Begin VB.Label Label8 
      Caption         =   "Blur Iterations:"
      Height          =   180
      Left            =   3780
      TabIndex        =   21
      Top             =   5205
      Width           =   1155
   End
   Begin VB.Label Label10 
      Caption         =   "Y:"
      Height          =   180
      Left            =   4965
      TabIndex        =   19
      Top             =   4560
      Width           =   990
   End
   Begin VB.Label Label9 
      Caption         =   "X:"
      Height          =   180
      Left            =   3795
      TabIndex        =   17
      Top             =   4560
      Width           =   990
   End
   Begin VB.Label Label6 
      Caption         =   "Gravity:"
      Height          =   390
      Left            =   3735
      TabIndex        =   15
      Top             =   3885
      Width           =   2625
   End
   Begin VB.Label Label5 
      Caption         =   "Length:"
      Height          =   390
      Left            =   3690
      TabIndex        =   13
      Top             =   3165
      Width           =   2625
   End
   Begin VB.Label Label4 
      Caption         =   "Base width:"
      Height          =   390
      Left            =   3690
      TabIndex        =   11
      Top             =   2415
      Width           =   2625
   End
   Begin VB.Label Label3 
      Caption         =   "Density:"
      Height          =   390
      Left            =   3690
      TabIndex        =   9
      Top             =   1665
      Width           =   2625
   End
   Begin VB.Label lblBits 
      Caption         =   "Bits per pixel:"
      Height          =   240
      Left            =   285
      TabIndex        =   7
      Top             =   3945
      Width           =   2850
   End
   Begin VB.Label Label2 
      Caption         =   "Force:"
      Height          =   390
      Left            =   3690
      TabIndex        =   6
      Top             =   870
      Width           =   2625
   End
   Begin VB.Label Label1 
      Caption         =   "Deviation:"
      Height          =   390
      Left            =   3690
      TabIndex        =   5
      Top             =   60
      Width           =   2625
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private BPF As Long, Hgt As Long, BPP As Long, C As Single
Private B(0 To 40400) As Byte, B2(0 To 40400) As Byte, X As Long, Y As Long, D As Long, Ending As Boolean, Pos As Long, I As Long, J As Long, CC As Long, II As Long
Private Dist(-20 To 20, -20 To 20) As Single
Private Type Particle
X As Single
Y As Single
V As Single
SV As Single
Life As Single
End Type

Private Type Preset
Name As String
PS(1 To 10) As Long
End Type

Private Type PresetList
PreCount As Long
Pre() As Preset
End Type

Private Pause As Boolean
Private Col As Long, VV As Single, SVV As Single, Svv2 As Single, S3 As Long, S4 As Long, Sze As Long, Heat As Single, Grav As Single, sx As Long, sy As Long, Bl As Long, T As Long, Tim
Private P(0 To 10000) As Particle
Private Pr As PresetList

Private Sub Combo1_Click()
ApplyPreset (Combo1.ListIndex + 1)
End Sub

Private Sub Command1_Click()
Form_Unload 0
End Sub

Private Sub Command2_Click()
Pause = True
CD.Filter = "Bitmaps|*.bmp"
CD.ShowOpen
If CD.FileName = "" Then Pause = False: Exit Sub
Picture2.Picture = LoadPicture(CD.FileName)

Dim PI As BITMAP
GetObject Picture2.Picture, Len(PI), PI
If PI.bmBitsPixel <> 8 Then MsgBox "8 Bit pictures only!": GoTo Resu
If PI.bmWidth <> 200 Then MsgBox "Width is not 200 pixels!": GoTo Resu
If PI.bmHeight <> 200 Then MsgBox "Height is not 200 pixels!": GoTo Resu

Picture1.Picture = LoadPicture(CD.FileName)

Resu:
Pause = False
End Sub

Public Sub Command3_Click()
Pause = (Pause = False)
If Pause Then Command3.Caption = "Resume" Else Command3.Caption = "Pause"
End Sub

Sub SetVars()
VV = Slider2.Value / 8
SVV = Slider1.Value / 8
Svv2 = Slider1.Value / 16
S3 = Slider4.Value
S4 = Slider4.Value / 2
Heat = Slider5.Value / 200
Grav = Slider6.Value / 1000
sx = Val(Text1.Text)
sy = Val(Text2.Text)

Label1.Caption = "Deviation: " & Slider1.Value
Label2.Caption = "Force: " & Slider2.Value
Label3.Caption = "Density: " & Slider3.Value
Label4.Caption = "Base width: " & Slider4.Value
Label5.Caption = "Length: " & Slider5.Value
Label6.Caption = "Gravity: " & Slider6.Value
End Sub

Private Sub Command4_Click()
Pause = False
Command3_Click
frmAbout.Doeffect
If frmAbout.Endd = True Then Command3_Click
End Sub

Private Sub Command5_Click()
AddPreset
End Sub

Private Sub Form_Load()
'This code is intended for show only, and was not expected to teach
'anyone anything, thus the lack of comments...

'To use: 1. Compile into an EXE otherwise it goes too slow.
'        2. Click on Load picture and click on one of the bitmaps in the directory.
'        3. Tweak the settings until you're happy.
'        4. Sit back and enjoy
'        5. Improve it!

Show
On Error Resume Next

Bl = 1

BPP = GetBPP(Picture1)
lblBits.Caption = "Bits per pixel: " & BPP
BPP = BPP / 8
If BPP = 0 Then BPP = 1
BPF = Picture1.Width * BPP
Hgt = Picture1.Height

LoadPresets

For I = -5 To 5 'This is not used...
For J = -5 To 5
Dist(I, J) = 1 - (Distance(I, J, 0, 0) / 3)
Next
Next

SetVars

Tim = Timer

Do 'BEGIN MAIN LOOP
DoEvents
    For Y = Hgt - 1 To Hgt + 2 'Clear top and bottom pixels.
    For X = 1 To BPF
    B((BPF * Y) + X) = 0
    Next
    Next
    
If Pause = False Then
    
S3 = Slider4.Value + ((Rnd * 10) - 5)
S4 = (Slider4.Value / 2) + ((Rnd * 10) - 5)

For I = 0 To Slider3.Value 'BEGIN PARTICLE ENGINE
With P(I)
    
    If .Life <= 0 Then
    .X = sx + Int(Rnd * S3) - S4
    .SV = (Rnd * SVV) - Svv2
    .Y = sy
    .V = Rnd * VV
    .Life = (Rnd * 155 + 100) * Heat
    If .Life < 0 Then .Life = 0
    End If

    .X = .X + .SV
    .Y = .Y + .V
    .V = .V - Grav
    If .X <= 0 Or .X >= 200 Then .Life = 1
    If .Y <= 1 Or .Y >= 198 Then .Y = 0: .Life = 1
    
    
    Pos = ((200 - Int(.Y)) * BPF) + Int(.X)
    Col = B(Pos) * 2 + (.Life * 2)
    If Col <= 10 Then Col = 0
    If Col > 255 Then Col = 255
    B(Pos) = Col
    
    .Life = .Life - 1
    
End With
Next 'END PARTICLE ENGINE
    
    If Check1.Value = 0 Then 'BEGIN BLUR FILTERING
    For I = 1 To Bl
    For X = 0 To BPF Step BPP
    C = -Int(Rnd * 3 - 2)
    For Y = 1 To Hgt
    Pos = (Y * BPF) + X
    B(Pos - BPF) = (CInt(B(Pos - BPF)) + B(Pos + BPF) + B(Pos - BPP) + B(Pos + BPP) + B(Pos + BPP + BPF) + B(Pos - BPP + BPF) + B(Pos + BPP - BPF) + B(Pos - BPP - BPF)) / 8.51
    Next
    Next
    Next
    
Else

    For I = 1 To Bl
    For X = 0 To BPF Step BPP
    C = -Int(Rnd * 3 - 2)
    For Y = 1 To Hgt
    Pos = (Y * BPF) + X
    B(Pos - BPF * C) = (CInt(B(Pos - BPF)) + B(Pos + BPF) + B(Pos - BPP) + B(Pos + BPP) + B(Pos + BPP + BPF) + B(Pos - BPP + BPF) + B(Pos + BPP - BPF) + B(Pos - BPP - BPF)) / 8.51
    Next
    Next
    Next


End If 'END BLUR FILTERING
    
    
    
    SetBitmapBits Picture1.Picture, UBound(B), B(1)
    Picture1.Refresh

End If 'end of Pause block

    T = T + 1 'FPS COUNTER
    If Timer >= Tim + 1 Then
    lblFPS = "FPS: " & T
    Tim = Timer
    T = 0
    End If

Loop Until Ending = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
SavePresets
Ending = True
Unload Me
End
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
CC = 255
For J = -2 To 2
For I = -2 To 2
If Y + J > 200 Then GoTo Skip
If Y - J < 0 Then GoTo Skip
If X + I > 200 Then GoTo Skip
If X - I < 0 Then GoTo Skip

'Paint pixels to picture under mouse cursor

B(((Y + J) * BPF) + (X + I)) = CC
B(((Y - J) * BPF) + (X - I)) = CC
Skip:
Next
Next

End Sub

Private Sub Slider1_Scroll()
SetVars
End Sub

Private Sub Slider2_Scroll()
SetVars

End Sub

Private Sub Slider3_Scroll()
SetVars

End Sub

Private Sub Slider4_Scroll()
SetVars

End Sub
Private Sub Slider5_Scroll()
SetVars

End Sub

Private Sub Slider6_Scroll()
SetVars

End Sub

Private Sub Slider1_Change()
SetVars
End Sub
Private Sub Slider2_Change()
SetVars
End Sub
Private Sub Slider3_Change()
SetVars
End Sub
Private Sub Slider4_Change()
SetVars
End Sub
Private Sub Slider5_Change()
SetVars
End Sub
Private Sub Slider6_Change()
SetVars
End Sub
Private Sub Slider7_Change()
SetVars
End Sub


Private Sub Slider7_Scroll()
SetVars

End Sub

Private Sub Text1_Change()
SetVars
End Sub

Private Sub Text2_Change()
SetVars
End Sub

Private Sub Text3_Change()
Bl = Val(Text3.Text)
End Sub

Sub LoadPresets()
Dim Xw As Long
If Dir(App.Path & "\Presets.dat") = "" Then Exit Sub

Open App.Path & "\Presets.dat" For Binary As #1
Get #1, , Pr
Close #1

For Xw = 1 To Pr.PreCount
Combo1.AddItem Pr.Pre(Xw).Name
Next


End Sub

Sub ApplyPreset(Ind As Long)
With Pr.Pre(Ind)
Slider1.Value = .PS(1)
Slider2.Value = .PS(2)
Slider3.Value = .PS(3)
Slider4.Value = .PS(4)
Slider5.Value = .PS(5)
Slider6.Value = .PS(6)
Text1.Text = CStr(.PS(7))
Text2.Text = CStr(.PS(8))
Text3.Text = CStr(.PS(9))
Check1.Value = .PS(10)
End With
End Sub

Sub AddPreset()
Pr.PreCount = Pr.PreCount + 1
ReDim Preserve Pr.Pre(0 To Pr.PreCount) As Preset
With Pr.Pre(Pr.PreCount)
.PS(1) = Slider1.Value
.PS(2) = Slider2.Value
.PS(3) = Slider3.Value
.PS(4) = Slider4.Value
.PS(5) = Slider5.Value
.PS(6) = Slider6.Value
.PS(7) = Val(Text1.Text)
.PS(8) = Val(Text2.Text)
.PS(9) = Val(Text3.Text)
.PS(10) = Check1.Value
.Name = InputBox("Type in a name for this preset:")
Combo1.AddItem .Name
End With
End Sub

Sub SavePresets()
Open App.Path & "\Presets.dat" For Binary As #1
Put #1, , Pr
Close #1
End Sub
