VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Screen Splitter"
   ClientHeight    =   1905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4575
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1905
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDump 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Top             =   960
      Width           =   2580
   End
   Begin VB.TextBox txtY 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Text            =   "4"
      Top             =   600
      Width           =   775
   End
   Begin VB.TextBox txtX 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Text            =   "4"
      Top             =   240
      Width           =   775
   End
   Begin VB.CommandButton cmdSplitScreen 
      Caption         =   "SPLIT AND SAVE"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   4335
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4560
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label lblDump 
      Caption         =   "Image Dump Directory:   [                                                         ]"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   4455
   End
   Begin VB.Label lblVertical 
      Caption         =   "Vertical Partitions            [                 ]"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label lblHorizontal 
      Caption         =   "Horizontal Partitions        [                 ]"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ScreenX As Single, ScreenY As Single

Private Sub cmdSplitScreen_Click()
txtDump_LostFocus
If Dir(txtDump.Text, vbDirectory) = "" Then MsgBox "The specified directory for your image dump is invalid.", vbExclamation + vbOKOnly, "Error": Exit Sub

Me.Hide
Split_N_Save Val(txtX.Text), Val(txtY.Text), txtDump.Text
Me.Show

End Sub

Private Sub Form_Load()
txtDump.Text = App.Path
ScreenX = Screen.Width / Screen.TwipsPerPixelX
ScreenY = Screen.Height / Screen.TwipsPerPixelY

End Sub

Function Split_N_Save(X As Single, Y As Single, Dump As String)
Dim xCount As Single, yCount As Single

For i = 0 To (X - 1)
For j = 0 To (Y - 1)
currentslice = CStr(i) & ", " & CStr(j)
SavePicture CaptureWindow(GetDesktopWindow(), True, (ScreenX / X) * i, (ScreenY / Y) * j, (ScreenX / X), (ScreenY / Y)), Dump & currentslice & ".bmp"
Next j
Next i
End Function

Private Sub txtDump_LostFocus()
If Right(txtDump.Text, 1) <> "\" Then txtDump.Text = txtDump.Text & "\"
End Sub

Private Sub txtX_LostFocus()
txtX.Text = Val(txtX.Text)
If Val(txtX.Text) < 1 Then txtX.Text = "1"
If Val(txtX.Text) > 16 Then txtX.Text = "16"
End Sub

Private Sub txtY_LostFocus()
txtY.Text = Val(txtY.Text)
If Val(txtY.Text) < 1 Then txtY.Text = "1"
If Val(txtY.Text) > 16 Then txtY.Text = "16"
End Sub
