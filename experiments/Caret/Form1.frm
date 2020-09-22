VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   312
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   473
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox RichTextBox2 
      Height          =   2055
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3625
      _Version        =   393217
      TextRTF         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   2566
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":0086
   End
   Begin VB.Label Hdn 
      AutoSize        =   -1  'True
      Caption         =   "Hdn"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label Label1 
      Caption         =   "INS"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const VK_INSERT = &H2D

Dim RT1Ovr As Boolean
Dim RT2Ovr As Boolean

Private Sub Label2_Click()

End Sub


'NOTES:
' The form's ScaleMode is 3 - Pixel
' The Hdn label is not visibile, and AutoSize is true.  It is used to calculate the size of the caret.

Private Sub RichTextBox1_KeyDown(KeyCode As Integer, Shift As Integer)
    If GetKeyState(VK_INSERT) < 0 Then 'Insert toggled
        RT1Ovr = Not RT1Ovr
    End If
End Sub

Private Sub RichTextBox1_KeyUp(KeyCode As Integer, Shift As Integer)
    If RT1Ovr Then
        BlockCursor RichTextBox1
        Label1.Caption = "OVR"
    Else
        LineCursor RichTextBox1
        Label1.Caption = "INS"
    End If
End Sub

Private Sub RichTextBox2_KeyUp(KeyCode As Integer, Shift As Integer)
    If RT2Ovr Then
        BlockCursor RichTextBox2
        Label1.Caption = "OVR"
    Else
        Label1.Caption = "INS"
    End If
End Sub

Private Sub RichTextBox2_KeyDown(KeyCode As Integer, Shift As Integer)
    If GetKeyState(VK_INSERT) < 0 Then 'Insert toggled
        RT2Ovr = Not RT2Ovr
    End If
End Sub

'Private Sub RichTextBox1_GotFocus()
'    If CursorRight(RichTextBox1) Then
'        Label1.Caption = "INS"
'    Else
'        Label1.Caption = "OVR"
'    End If
'
'End Sub

Public Sub BlockCursor(rtb As RichTextBox)
    
    'Create a measurement of this character
    Hdn.Caption = Mid(rtb.Text, rtb.SelStart + 1, 1)
    Set Hdn.Font = rtb.Font
    
    'Create the cursor
    CreateCaret rtb.hWnd, 0, Hdn.Width, Hdn.Height
    ShowCaret rtb.hWnd

End Sub

Public Sub LineCursor(rtb As RichTextBox)
    
    'Create a measurement of this character
    Hdn.Caption = "X"
    Set Hdn.Font = rtb.Font
    
    'Create the cursor
    CreateCaret rtb.hWnd, 0, 0, Hdn.Height
    ShowCaret rtb.hWnd

End Sub

