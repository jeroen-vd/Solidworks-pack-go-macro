VERSION 5.00
Begin VB.Form PackAndRenameForm 
   Caption         =   "Pack & Rename"
   ClientHeight    =   3900
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnRun
      Caption         =   "Uitvoeren"
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   3360
      Width           =   1200
   End
   Begin VB.CommandButton btnCancel
      Caption         =   "Annuleren"
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   3360
      Width           =   1200
   End
   Begin VB.CommandButton btnBrowse
      Caption         =   "Bladeren..."
      Height          =   315
      Left            =   5160
      TabIndex        =   5
      Top             =   2040
      Width           =   1200
   End
   Begin VB.CheckBox chkIncludeDrawings
      Caption         =   "Tekeningen meenemen"
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Value           =   1  'Checked
      Width           =   2400
   End
   Begin VB.TextBox txtOutputFolder
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   4920
   End
   Begin VB.TextBox txtPrefix
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1920
   End
   Begin VB.Label lblHelp
      Caption         =   ""
      Height          =   840
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   6240
   End
   Begin VB.Label Label2
      Caption         =   "Exportmap"
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   1200
   End
   Begin VB.Label Label1
      Caption         =   "Prefix"
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   720
   End
End
Attribute VB_Name = "PackAndRenameForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnBrowse_Click()
    Dim folder As String
    folder = BrowseForFolder("Kies exportmap")
    If Len(folder) > 0 Then
        txtOutputFolder.Text = folder
    End If
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnRun_Click()
    If Len(Trim$(txtPrefix.Text)) = 0 Then
        MsgBox "Geef een prefix op voordat je start.", vbExclamation, "Pack & Rename"
        Exit Sub
    End If

    If Len(Trim$(txtOutputFolder.Text)) = 0 Then
        MsgBox "Selecteer een exportmap.", vbExclamation, "Pack & Rename"
        Exit Sub
    End If

    ExecutePackAndRename Application.SldWorks.ActiveDoc, _
                          Trim$(txtPrefix.Text), _
                          Trim$(txtOutputFolder.Text), _
                          chkIncludeDrawings.Value
    Unload Me
End Sub

Public Sub InitializeWithDefaults(ByVal swModel As SldWorks.ModelDoc2)
    Dim suggestedPrefix As String

    suggestedPrefix = GetSuggestedPrefix(swModel)

    txtPrefix.Text = suggestedPrefix
    chkIncludeDrawings.Value = True
    lblHelp.Caption = BuildHelpMessage(suggestedPrefix)
    txtOutputFolder.Text = GetDefaultFolder(swModel)
End Sub

Private Function GetDefaultFolder(ByVal swModel As SldWorks.ModelDoc2) As String
    Dim path As String
    path = swModel.GetPathName
    If InStrRev(path, "\") > 0 Then
        GetDefaultFolder = Left$(path, InStrRev(path, "\") - 1)
    Else
        GetDefaultFolder = CurDir$()
    End If
End Function

Private Function BrowseForFolder(prompt As String) As String
    Dim shellApp As Object
    Dim folder As Object

    Set shellApp = CreateObject("Shell.Application")
    Set folder = shellApp.BrowseForFolder(0, prompt, 0, 0)

    If Not folder Is Nothing Then
        BrowseForFolder = folder.Self.Path
    Else
        BrowseForFolder = ""
    End If
End Function

Private Function GetSuggestedPrefix(ByVal swModel As SldWorks.ModelDoc2) As String
    Dim baseName As String

    baseName = GetBaseName(swModel.GetPathName)

    If Len(baseName) = 0 Then
        baseName = "PRJ"
    End If

    GetSuggestedPrefix = baseName
End Function

Private Function BuildHelpMessage(ByVal prefix As String) As String
    BuildHelpMessage = "Exporteert alle parts (Pxx), subassemblies (Axx) en tekeningen naar één map." & vbCrLf & _
                      "Top-level assembly wordt " & prefix & "-A00; overige assemblies " & prefix & "-A01, A02, ..." & vbCrLf & _
                      "Parts volgen " & prefix & "-P01, P02, ...; tekeningen volgen het bijbehorende model en krijgen unieke namen."
End Function
