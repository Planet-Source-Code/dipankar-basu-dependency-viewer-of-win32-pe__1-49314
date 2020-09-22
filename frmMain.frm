VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dependencies"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   6120
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   480
      TabIndex        =   7
      Top             =   3120
      Width           =   3015
   End
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   3960
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5400
      TabIndex        =   6
      ToolTipText     =   "Browse File"
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox txtTemp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Text            =   "Temporary Variable Storage Buffer"
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      ToolTipText     =   "Close"
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About ..."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      ToolTipText     =   "Contact Author"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdDep 
      Caption         =   "&Dependencies"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      ToolTipText     =   "Get File Info"
      Top             =   840
      Width           =   1575
   End
   Begin VB.ListBox LstDep 
      BackColor       =   &H00E0E0E0&
      Height          =   1740
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   3855
   End
   Begin VB.TextBox txtpath 
      BackColor       =   &H00E0E0E0&
      Height          =   405
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "File Name"
      Top             =   240
      Width           =   4935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Dependencies Viewer for Win32 PE Files.
' Written by Dipankar Basu on August 13, 2003.
' Modified on October 28, 2003 by Dipankar Basu.
' web : http://www.geocities.com/basudip_in/
' Copyright 2003 by Dipankar Basu
' Credits: Roger Gilchrist for reporting bugs, Thanks a lot.
Option Explicit
Private Sub cmdAbout_Click()
Dim msgStr As String
msgStr = "Dependency Viewer is developed by Dipankar Basu," & vbCrLf
msgStr = msgStr + "visit my web at http://www.geocities.com/basudip_in/." & vbCrLf
msgStr = msgStr + "Please feel free to drop in your sugessions and comments," & vbCrLf
msgStr = msgStr + "Your effort in this direction will be gladly appreciated."
Call MsgBox(msgStr, vbOKOnly, "Dependency Viewer  " & App.Major & "." & App.Minor)
End Sub
Private Sub cmdClose_Click()
Unload Me: End
End Sub
Private Sub cmdDep_Click()
Me.Caption = "[Searching] Please wait . . ."
LstDep.Clear
GetDepFiles
DoEvents
RemoveDupEntry
UpdateList
Me.Caption = "Dependencies"
End Sub
Private Sub cmdOpen_Click()
On Error GoTo eh:
With cDialog
.DialogTitle = "Open the file to view its Dependencies"
.Filter = "Windows Portable Executables|*.exe;*.ocx;*.dll|All Files|*.*"
.Flags = cdlOFNFileMustExist Or cdlOFNReadOnly
.ShowOpen
txtpath.Text = .FileName
End With
Exit Sub
eh:
MsgBox Err.Source & " reports " & Err.Description, , "Error " & Err.Number
End Sub
Private Sub txtpath_Change()
If txtpath.Text = vbNullString Then
cmdDep.Enabled = False
Else
cmdDep.Enabled = True
End If
End Sub
Private Sub GetDepFiles()
Dim Contents As String
On Local Error Resume Next
List1.Clear
Open txtpath.Text For Binary As #1
Contents = Space$(LOF(1))
Get #1, , Contents: Close #1: Contents = UCase$(Contents)
ReadContentsFor Contents, ".DLL"
DoEvents
ReadContentsFor Contents, ".OCX"
End Sub
Private Sub ReadContentsFor(ByVal strContents As String, ByVal strFind As String)
Dim lngExtLocation As Long, lngBlankLocation As Long
lngExtLocation = InStr(1, strContents, strFind)
lngBlankLocation = lngExtLocation
Do
    Do
    lngBlankLocation = lngBlankLocation - 1
    txtTemp.Text = Mid$(strContents, lngBlankLocation, 1)
    If Len(txtTemp.Text) = 0 Or Trim$(txtTemp.Text) = "\" Then Exit Do
    Loop
txtTemp.Text = Mid$(strContents, lngBlankLocation + 1, _
 lngExtLocation - lngBlankLocation + Len(strFind) - 1)
List1.AddItem Trim$(txtTemp.Text)
lngExtLocation = InStr(lngExtLocation + 4, strContents, strFind)
lngBlankLocation = lngExtLocation: DoEvents
Loop While lngExtLocation > 0
End Sub
Private Sub RemoveDupEntry()
Dim str As String, i As Long, d As Long
On Local Error Resume Next
Do While i < List1.ListCount
str = UCase$(List1.List(i))
If LenB(str) Then
    d = 1 + i
    With List1
    Do While d < .ListCount
    If str = UCase(.List(d)) Then
    .RemoveItem (d)
    d = d - 1
    End If
    d = d + 1
    Loop
    End With
End If
i = i + 1
Loop
DoEvents
End Sub
Private Sub UpdateList()
Dim i As Long
With List1
For i = 0 To .ListCount
    If .List(i) <> vbNullString And _
    Left$(Trim$(Right$(.List(i), 4)), 1) = "." Then _
LstDep.AddItem .List(i)
Next i
End With
End Sub
