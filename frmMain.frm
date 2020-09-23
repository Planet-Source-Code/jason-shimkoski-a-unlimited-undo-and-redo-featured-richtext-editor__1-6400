VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Stuff"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox rtfText 
      Height          =   3255
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   5741
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":000C
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Select &All"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   3975
   End
   Begin VB.CommandButton cmdPaste 
      Caption         =   "&Paste"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdCut 
      Caption         =   "Cu&t"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdRedo 
      Caption         =   "&Redo"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   "&Undo"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                     This Code was created by Jason Shimkoski               '
'                                 Copyright 2000                             '
'                                                                            '
'        You can use this code in your apps, as long as you mention          '
'        me a in your About Box.                                             '
'                                                                            '
'        Please note that sections of this sample app were revisions of code '
'        found at www.visualstatement.com/vb                                 '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        There is only one bug that I know of in this app that has to        '
'        do with the redo, but other than that, it is pretty much safe.      '
'        If you figure out how to fix it, send the revised code to           '
'        basspler@aol.com.  Thanks                                           '
'                                                                            '
'        Oh, one more thing, within the next few weeks,                      '
'        I should have a new version of this that will be                    '
'        using the windows API instead.                                      '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
'These are the variables for Undo and Redo
Dim gblnIgnoreChange As Boolean
Dim gintIndex As Integer
Dim gstrStack(1000) As String

'Here I'm just setting up the form.
Private Sub Form_Load()
    Form_Resize
End Sub

'Still just setting up the form
Private Sub Form_Resize()
    Me.Height = 4830
    Me.Width = 4305
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' For any retards out there, rtfText is the RichTextBox Control.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private Sub cmdRedo_Click()
    'This is the basic redo stuff.
    gblnIgnoreChange = True
    gintIndex = gintIndex + 1
    On Error Resume Next
    rtfText.TextRTF = gstrStack(gintIndex)
    gblnIgnoreChange = False
End Sub

Private Sub cmdUndo_Click()
    'This says that if the Index is = to 0, then It shouldn't undo anymore
    If gintIndex = 0 Then Exit Sub
    
    'This is the basic undo stuff.
    gblnIgnoreChange = True
    gintIndex = gintIndex - 1
    On Error Resume Next
    rtfText.TextRTF = gstrStack(gintIndex)
    gblnIgnoreChange = False
End Sub

Private Sub rtfText_Change()
    'Basically this updates the Undo and Redo
    If Not gblnIgnoreChange Then
        gintIndex = gintIndex + 1
        gstrStack(gintIndex) = rtfText.TextRTF
    End If
End Sub

Private Sub cmdCut_Click()
    'Clears the Clipboard to put text on it
    Clipboard.Clear
    'Sets the Text from rtfText onto the Clipboard
    Clipboard.SetText rtfText.SelText
    'Deletes the Selected Text on rtfText
    rtfText.SelText = ""
    'Sets the Focus to rtfText
    rtfText.SetFocus
End Sub

Private Sub cmdCopy_Click()
    'Clears the Clipboard to put text on it
    Clipboard.Clear
    'Sets the Text from rtfText onto the Clipboard
    Clipboard.SetText rtfText.SelText
    'Sets the Focus to rtfText
    rtfText.SetFocus
End Sub

Private Sub cmdPaste_Click()
    'Puts the Text from the clipboard into rtfText
    rtfText.SelText = Clipboard.GetText
    'Sets the Focus to rtfText
    rtfText.SetFocus
End Sub

Private Sub cmdSelectAll_Click()
    'Sets the cursors position to zero
    rtfText.SelStart = 0
    'Selects the full length of rtfText
    rtfText.SelLength = Len(rtfText.Text)
    'Sets the Focus to rtfText
    rtfText.SetFocus
End Sub
