Attribute VB_Name = "modIcons"
Option Explicit
Option Base 1

' Urbano DaGama (udgama@rocketmail.com)

' Drop me a line in case you need any help on this program or
' if you liked the code. That will encourage me to create more such
' programs.

Public Declare Function OSfCreateShellLink Lib "vb6stkit.dll" Alias "fCreateShellLink" _
         (ByVal lpstrFolderName As String, _
         ByVal lpstrLinkName As String, _
         ByVal lpstrLinkPath As String, _
         ByVal lpstrLinkArguments As String, _
         ByVal fPrivate As Long, _
         ByVal sParent As String) As Long

Public Const gstrQUOTE$ = """"


Public Sub CreateLink(iPath As String, iGroup As String, iTitle As String, iArgs As String)
On Error GoTo EH
Dim strProgramPath   As String   ' The path of the executable file
Dim strGroup         As String
Dim strProgramIconTitle As String
Dim strProgramArgs   As String
Dim sParent          As String

   
   strProgramPath = iPath
   strGroup = iGroup
   strProgramIconTitle = iTitle
   strProgramArgs = iArgs
   
   sParent = "$(Programs)"
   
   CreateShellLink strProgramPath, strGroup, strProgramArgs, strProgramIconTitle, True, sParent
   
   Exit Sub
EH:
   MsgBox Err.Description
   Exit Sub
End Sub

'-----------------------------------------------------------
' SUB: CreateShellLink
'
' Creates (or replaces) a link in either Start>Programs or
' any of its immediate subfolders in the Windows 95 shell.
'
' IN: [strLinkPath] - full path to the target of the link
'                     Ex: 'c:\Program Files\My Application\MyApp.exe"
'     [strLinkArguments] - command-line arguments for the link
'                     Ex: '-f -c "c:\Program Files\My Application\MyApp.dat" -q'
'     [strLinkName] - text caption for the link
'     [fLog] - Whether or not to write to the logfile (default
'                is true if missing)
'
' OUT:
'   The link will be created in the folder strGroupName
'-----------------------------------------------------------
'
Public Sub CreateShellLink(ByVal strLinkPath As String, _
         ByVal strGroupName As String, _
         ByVal strLinkArguments As String, _
         ByVal strLinkName As String, _
         ByVal fPrivate As Boolean, _
         sParent As String, _
         Optional ByVal fLog As Boolean = True)
Dim fSuccess As Boolean
Dim intMsgRet As Integer
Dim lREt       As Boolean
   strLinkName = strUnQuoteString(strLinkName)
   strLinkPath = strUnQuoteString(strLinkPath)
   
   If StrPtr(strLinkArguments) = 0 Then strLinkArguments = ""
   
   lREt = OSfCreateShellLink(strGroupName, strLinkName, strLinkPath, strLinkArguments, _
         fPrivate, sParent)    'the path should never be enclosed in double quotes

End Sub


Public Function strUnQuoteString(ByVal strQuotedString As String)
'
' This routine tests to see if strQuotedString is wrapped in quotation
' marks, and, if so, remove them.
'
    strQuotedString = Trim$(strQuotedString)

    If Mid$(strQuotedString, 1, 1) = gstrQUOTE Then
        If Right$(strQuotedString, 1) = gstrQUOTE Then
            '
            ' It's quoted.  Get rid of the quotes.
            '
            strQuotedString = Mid$(strQuotedString, 2, Len(strQuotedString) - 2)
        End If
    End If
    strUnQuoteString = strQuotedString
End Function

