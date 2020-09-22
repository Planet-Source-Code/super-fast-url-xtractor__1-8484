VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   10770
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C79345&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1470
      ItemData        =   "Form1.frx":0000
      Left            =   720
      List            =   "Form1.frx":0002
      TabIndex        =   3
      Top             =   720
      Width           =   8655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go !"
      Height          =   285
      Left            =   8760
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H0059BD51&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   7935
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4455
      Left            =   720
      TabIndex        =   0
      Top             =   3120
      Width           =   8655
      ExtentX         =   15266
      ExtentY         =   7858
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "res://C:\WINDOWS\SYSTEM\SHDOCLC.DLL/dnserror.htm#http:///"
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C79345&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   480
      Width           =   8655
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00634BB4&
      Caption         =   "Label4"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   720
      TabIndex        =   6
      Top             =   2280
      Width           =   8655
   End
   Begin VB.Label Label2 
      Caption         =   "Links"
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "URL"
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
List1.Clear
Label5.Caption = ""
WebBrowser1.Navigate (Text1.Text)
End Sub

Private Sub Form_Load()
WebBrowser1.Navigate ("www.infoseek.com")
End Sub

Private Sub List1_DblClick()
Label5.Caption = ""
WebBrowser1.Navigate (List1.Text)
List1.Clear
End Sub

Private Sub WebBrowser1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
List1.Clear
Label5.Caption = ""
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
If (pDisp Is WebBrowser1.Object) Then
List1.Clear
Label5.Caption = "That page contains " & WebBrowser1.Document.links.length & " links"
For i = 0 To WebBrowser1.Document.links.length - 1
    If Left$(LCase(WebBrowser1.Document.links.Item(i)), 4) = "http" Then List1.AddItem (WebBrowser1.Document.links.Item(i))
Next i
Label5.Caption = "That page contains " & WebBrowser1.Document.links.length & " links, " & List1.ListCount & " usable"
End If
End Sub


Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
    Label4.Caption = Text
End Sub
