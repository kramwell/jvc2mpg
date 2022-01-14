VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "jvc2mpg 0.6.4"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3000
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Convert Folder"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Convert File"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "FileType"
      Height          =   855
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   975
      Begin VB.CheckBox Check1 
         Caption         =   "TOD"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Value           =   2  'Grayed
         Width           =   735
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   615
      Left            =   4080
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "GO"
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   2400
      Width           =   2775
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   6360
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4200
      TabIndex        =   0
      Top             =   3120
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Written by KramWell.com - 25/OCT/2014
'Program to convert TOD filetypes to MPG written for a friend that had a JVC camera, requires ffmpeg.exe
      
Option Explicit

Dim strFilename As String
Dim strFoldername As String
Dim GetFilenameFromPath As String

      Private Const BIF_RETURNONLYFSDIRS = 1
      Private Const BIF_DONTGOBELOWDOMAIN = 2
      Private Const MAX_PATH = 260

      Private Declare Function SHBrowseForFolder Lib "shell32" _
                                        (lpbi As BrowseInfo) As Long

      Private Declare Function SHGetPathFromIDList Lib "shell32" _
                                        (ByVal pidList As Long, _
                                        ByVal lpBuffer As String) As Long

      Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
                                        (ByVal lpString1 As String, ByVal _
                                        lpString2 As String) As Long

      Private Type BrowseInfo
         hWndOwner      As Long
         pIDLRoot       As Long
         pszDisplayName As Long
         lpszTitle      As Long
         ulFlags        As Long
         lpfnCallback   As Long
         lParam         As Long
         iImage         As Long
      End Type

      Private Sub Command2_Click()
      'Opens a Treeview control that displays the directories in a computer
         
         Dim lpIDList As Long
         Dim sBuffer As String
         Dim szTitle As String
         Dim tBrowseInfo As BrowseInfo

     '    szTitle = "Select a valid folder"
         With tBrowseInfo
            .hWndOwner = Me.hWnd
            .lpszTitle = lstrcat(szTitle, "")
            .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
         End With

         lpIDList = SHBrowseForFolder(tBrowseInfo)

         If (lpIDList) Then
            sBuffer = Space(MAX_PATH)
            SHGetPathFromIDList lpIDList, sBuffer
            sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
            strFoldername = sBuffer
         End If
         
If strFoldername <> "" Then
Call Command3_Click
End If
         
      End Sub





Private Sub Command1_Click()

CommonDialog.FileName = ""
CommonDialog.Filter = "Jvc (*.TOD)|*.TOD|All files (*.*)|*.*"
CommonDialog.DefaultExt = "TOD"
CommonDialog.DialogTitle = "Select the JVC file to convert"
CommonDialog.ShowOpen

strFilename = CommonDialog.FileName

GetFilenameFromPath = Right(strFilename, Len(strFilename) - InStrRev(strFilename, "\"))
GetFilenameFromPath = Replace(GetFilenameFromPath, ".TOD", "")

If strFilename <> "" Then
Call Command3_Click
End If

End Sub

Private Sub Command3_Click()

         Dim lpIDList As Long
         Dim sBuffer As String
         Dim szTitle As String
         Dim tBrowseInfo As BrowseInfo

      'Opens a Treeview control that displays the directories in a computer
         szTitle = "Where do you want it saved?"
         With tBrowseInfo
            .hWndOwner = Me.hWnd
            .lpszTitle = lstrcat(szTitle, "")
            .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
         End With

         lpIDList = SHBrowseForFolder(tBrowseInfo)

         If (lpIDList) Then
            sBuffer = Space(MAX_PATH)
            SHGetPathFromIDList lpIDList, sBuffer
            sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
         End If
      'Opens a Treeview control that displays the directories in a computer

If sBuffer <> "" Then

If strFilename = "" Then

Open App.Path & "\jvc.bat" For Output As #1
   Print #1, "@echo off"
   Print #1, "cls"
   Print #1, "ECHO."
   Print #1, "ECHO  -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-   "
   Print #1, "ECHO                         #####                            "
   Print #1, "ECHO        # #    #  ####  #     # #    # #####   ####       "
   Print #1, "ECHO        # #    # #    #       # ##  ## #    # #    #      "
   Print #1, "ECHO        # #    # #       #####  # ## # #    # #           "
   Print #1, "ECHO        # #    # #      #       #    # #####  #  ###      "
   Print #1, "ECHO   #    #  #  #  #    # #       #    # #      #    #      "
   Print #1, "ECHO    ####    ##    ####  ####### #    # #       ####       "
   Print #1, "ECHO                                        v0.6.4            "
   Print #1, "ECHO  -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-   "
   Print #1, "ECHO."
   Print #1, "ECHO Saving from: " & strFoldername
   Print #1, "ECHO Saving to: " & sBuffer
   'strFoldername strFilename
   
   Print #1, "pause"
   
  ' Print #1, "CD" & """strFoldername"""
   Print #1, "FOR %%a in (" & Chr(34) & strFoldername & "\*.TOD" & Chr(34) & ") do ffmpeg.exe -i " & Chr(34) & "%%a" & Chr(34) & " -vcodec copy -acodec copy """ & sBuffer & "\%%~na.MPG"""

Close #1

Shell App.Path & "\jvc.bat", vbNormalFocus


Else

'if file is selected only
Shell App.Path & "\ffmpeg.exe -i " & strFilename & " -vcodec copy -acodec copy " & sBuffer & "\" & GetFilenameFromPath & ".MPG", vbNormalFocus
'MsgBox "hello"

End If 'strFilename

'Kill App.Path & "\jvc.bat"
End If 'sBuffer is empty

End Sub

Private Sub Command4_Click()
Open App.Path & "\Test.bat" For Output As #1
   Print #1, "@echo off"
   Print #1, "cls"
   Print #1, "@echo Test file"
Close #1
End Sub
