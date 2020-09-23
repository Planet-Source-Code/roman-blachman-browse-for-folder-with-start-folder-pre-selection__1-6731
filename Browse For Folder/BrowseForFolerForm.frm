VERSION 5.00
Begin VB.Form BrowseForFolerForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Browse for folders with directory pre-selection, Roman Blachman"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSelected 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Text            =   "Selected path"
      Top             =   1560
      Width           =   7095
   End
   Begin VB.CommandButton BrowseForFolders 
      Caption         =   "&Browse with selection"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox txtStart 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "Start path to browse from"
      Top             =   480
      Width           =   4815
   End
   Begin VB.Label Label2 
      Caption         =   "Selected path:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Write a path to start browsing from:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "BrowseForFolerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Function BrowseForFolder(selectedPath As String) As String
Dim Browse_for_folder As BROWSEINFOTYPE
Dim itemID As Long
Dim selectedPathPointer As Long
Dim tmpPath As String * 256
With Browse_for_folder
    .hOwner = Me.hWnd ' Window Handle
    .lpszTitle = "Browse for folders with directory pre-selection, Roman Blachman" ' Dialog Title
    .lpfn = FunctionPointer(AddressOf BrowseCallbackProcStr) ' Dialog callback function that preselectes the folder specified
    selectedPathPointer = LocalAlloc(LPTR, Len(selectedPath) + 1) ' Allocate a string
    CopyMemory ByVal selectedPathPointer, ByVal selectedPath, Len(selectedPath) + 1 ' Copy the path to the string
    .lParam = selectedPathPointer ' The folder to preselect
End With
itemID = SHBrowseForFolder(Browse_for_folder) ' Execute the BrowseForFolder API
If itemID Then
    If SHGetPathFromIDList(itemID, tmpPath) Then ' Get the path for the selected folder in the dialog
        BrowseForFolder = Left$(tmpPath, InStr(tmpPath, vbNullChar) - 1) ' Take only the path without the nulls
    End If
    Call CoTaskMemFree(itemID) ' Free the itemID
End If
Call LocalFree(selectedPathPointer) ' Free the string from the memory
End Function

Private Sub BrowseForFolders_Click()
Dim tmpPath As String
tmpPath = txtStart ' Take the selected path from txtStart
If Len(tmpPath) > 0 Then
    If Not Right$(tmpPath, 1) <> "\" Then tmpPath = Left$(tmpPath, Len(tmpPath) - 1) ' Remove "\" if the user added
End If
txtStart = tmpPath
tmpPath = BrowseForFolder(tmpPath) ' Browse for folder
If tmpPath = "" Then
    txtSelected = "No folder selected !" ' If the user pressed cancel
Else
    txtSelected = "Folder Selected: " & tmpPath ' If the user selected a folder
End If
End Sub

Private Sub Form_Load()
MsgBox "Browse for folders with directory pre-selection" & vbCrLf & "created by Roman Blachman, eMail: romaz@inter.net.il"
End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox "Browse for folders with directory pre-selection" & vbCrLf & "created by Roman Blachman, eMail: romaz@inter.net.il"
End Sub
