VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "5 Language Translator"
   ClientHeight    =   6930
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear All"
      Height          =   495
      Left            =   2160
      TabIndex        =   9
      Top             =   5520
      Width           =   855
   End
   Begin RichTextLib.RichTextBox Text3 
      Height          =   1455
      Left            =   240
      TabIndex        =   6
      Top             =   3960
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   2566
      _Version        =   393217
      TextRTF         =   $"Trans.frx":0000
   End
   Begin RichTextLib.RichTextBox Text2 
      Bindings        =   "Trans.frx":0082
      Height          =   1575
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   2778
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"Trans.frx":008D
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   1095
      Left            =   6360
      TabIndex        =   4
      Top             =   2160
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1931
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"Trans.frx":010F
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Trans.frx":0191
      Left            =   3120
      List            =   "Trans.frx":01A4
      TabIndex        =   3
      Text            =   "Target Language"
      Top             =   5520
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Trans.frx":01D3
      Left            =   240
      List            =   "Trans.frx":01E6
      TabIndex        =   2
      Text            =   "Source Language"
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Request"
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Traslate"
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   5880
      Width           =   855
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6360
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   120
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   6480
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Source Text Below,  Translated Text Above"
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   3720
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   2160
      Left            =   960
      Picture         =   "Trans.frx":0215
      Top             =   0
      Width           =   3240
   End
   Begin VB.Menu FileMenu 
      Caption         =   "File"
      Begin VB.Menu ExitPro 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu TransOptions 
      Caption         =   "Translation Options"
      Begin VB.Menu ClearTrans 
         Caption         =   "Clear &Translation"
         Shortcut        =   ^T
      End
      Begin VB.Menu ClearSource 
         Caption         =   "Clear So&urce"
         Shortcut        =   ^U
      End
      Begin VB.Menu ClearAll 
         Caption         =   "Clear &All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu HelpMenu 
      Caption         =   "Help"
      Begin VB.Menu HowTo 
         Caption         =   "&How to use this program"
         Shortcut        =   ^H
      End
      Begin VB.Menu AboutItem 
         Caption         =   "A&bout"
         Shortcut        =   ^B
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo2_Click()
Select Case Combo2.Text
    Case "French": Combo2.Tag = "FR"
    Case "German": Combo2.Tag = "DE"
    Case "Spanish": Combo2.Tag = "ES"
    Case "English": Combo2.Tag = "EN"
    Case "Italian": Combo2.Tag = "IT"
End Select
End Sub

Private Sub Combo1_Click()
Select Case Combo1.Text
    Case "French": Combo1.Tag = "FR"
    Case "German": Combo1.Tag = "DE"
    Case "Spanish": Combo1.Tag = "ES"
    Case "English": Combo1.Tag = "EN"
    Case "Italian": Combo1.Tag = "IT"
End Select
End Sub

Private Sub Command1_Click()
Dim Start As Integer
Dim Done As Integer
Start = InStr(1, Text1.Text, "<textarea name=""wl_result""")
Start = Start + 114
Done = InStr(Start, Text1.Text, "</textarea>")
Text2.Text = Mid(Text1.Text, Start, Done - Start)
Label2.Caption = "Enter your text and hit the request button."
Command1.Enabled = False
End Sub

Private Sub Command2_Click()
Label2.Caption = "Please Wait..."
Dim ReturnStr As String
Text1.Text = Inet1.OpenURL("http://www.worldlingo.com/wl/Translate?wl_text=" & Text3.Text & "&wl_gloss=1&wl_srclang=" & Combo1.Tag & "&wl_trglang=" & Combo2.Tag, icString)
ReturnStr = Inet1.GetChunk(2048, icString)
Command1.Enabled = True
Label2.Caption = "Done! Hit the translate button."
Do While Len(ReturnStr) <> 0


    DoEvents

        Command1.Enabled = False
        Label2.Caption = "Pease Wait..."
        Text1.Text = Text1.Text & ReturnStr
        ReturnStr = Inet1.GetChunk(2048, icString)
        Command1.Enabled = True
        Label2.Caption = "Done! Hit the translate button."
    Loop
End Sub



Private Sub ExitPro_Click()
Unload Me
End Sub

Private Sub Form_Load()
Command1.Enabled = False
Label2.Caption = "Enter your text and hit the request button."
End Sub
Private Sub ClearAll_Click()
Text2.Text = ""
Text3.Text = ""
End Sub

Private Sub ClearSource_Click()
Text3.Text = ""
End Sub

Private Sub ClearTrans_Click()
Text2.Text = ""
End Sub

Private Sub Command3_Click()
Text2.Text = ""
Text3.Text = ""
End Sub
