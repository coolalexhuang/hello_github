Attribute VB_Name = "Modue_MDC"
'Created by huangjia(jianbin,huang@moodys.com) on 16/Nov/2012

Option Explicit
Option Compare Text

Public logFiles(30) As String
Public envList(30) As String

Public Enum DownloadFileDisposition
    OverwriteKill = 0
    OverwriteRecycle = 1
    DoNotOverwrite = 2
    PromptUser = 3
End Enum

Private Declare Function SHFileOperation Lib "shell32.dll" Alias _
    "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Private Declare Function PathIsNetworkPath Lib "shlwapi.dll" _
    Alias "PathIsNetworkPathA" ( _
    ByVal pszPath As String) As Long

Private Declare Function GetSystemDirectory Lib "kernel32" _
    Alias "GetSystemDirectoryA" ( _
    ByVal lpBuffer As String, _
    ByVal nSize As Long) As Long

Private Declare Function SHEmptyRecycleBin _
    Lib "shell32" Alias "SHEmptyRecycleBinA" _
    (ByVal hwnd As Long, _
     ByVal pszRootPath As String, _
     ByVal dwFlags As Long) As Long

Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_NOCONFIRMATION = &H10
Private Const MAX_PATH As Long = 260

Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type

Private Declare Function URLDownloadToFile Lib "urlmon" Alias _
  "URLDownloadToFileA" ( _
  ByVal pCaller As Long, _
  ByVal szURL As String, _
  ByVal szFileName As String, _
  ByVal dwReserved As Long, _
  ByVal lpfnCB As Long) As Long

Sub CreateNewMail(mailBody As String, environment As String)
    Dim myItem As Outlook.MailItem
    Dim myRecipient As Outlook.Recipient
    Dim content As String
    Set myItem = Application.CreateItem(olMailItem)
    'Set myRecipient = myItem.Recipients.Add("Dan Wilson")
    myItem.To = "MA Shenzhen MDC SE"
    myItem.CC = "Lu, Feng; Liu, Eric; Jiang, Lingyan Alinta; MA Shenzhen MDC QA"
    myItem.Subject = "Got server internal error on " & environment
    content = vbCrLf & vbCrLf & vbCrLf & vbCrLf & "-----------------------------------------------" & vbCrLf & mailBody
    myItem.BodyFormat = olFormatRichText
    myItem.Body = content
    myItem.Display
End Sub

Sub GetLog()
    UserForm_MDC.Show
    UserForm_MDC.TextBox1.SetFocus
End Sub

Public Function DownloadFile( _
  UrlFileName As String, _
  DestinationFileName As String, _
  Overwrite As DownloadFileDisposition, _
  ErrorText As String) As Boolean
  
Dim Disp As DownloadFileDisposition
Dim Res As VbMsgBoxResult
Dim B As Boolean
Dim S As String
Dim L As Long

ErrorText = vbNullString

If Dir(DestinationFileName, vbNormal) <> vbNullString Then
    Select Case Overwrite
        Case OverwriteKill
            On Error Resume Next
            Err.Clear
            Kill DestinationFileName
            If Err.Number <> 0 Then
                ErrorText = "Error Kill'ing file '" & DestinationFileName & "'." & vbCrLf & Err.Description
                DownloadFile = False
                Exit Function
            End If

        Case OverwriteRecycle
            On Error Resume Next
            Err.Clear
            B = RecycleFileOrFolder(DestinationFileName)
            If B = False Then
                ErrorText = "Error Recycle'ing file '" & DestinationFileName & "." & vbCrLf & Err.Description
                DownloadFile = False
                Exit Function
            End If

        Case DoNotOverwrite
            DownloadFile = False
            ErrorText = "File '" & DestinationFileName & "' exists and disposition is set to DoNotOverwrite."
            Exit Function

        'Case PromptUser
        Case Else
            S = "The destination file '" & DestinationFileName & "' already exists." & vbCrLf & _
                "Do you want to overwrite the existing file?"
            Res = MsgBox(S, vbYesNo, "Download File")
            If Res = vbNo Then
                ErrorText = "User selected not to overwrite existing file."
                DownloadFile = False
                Exit Function
            End If
            B = RecycleFileOrFolder(DestinationFileName)
            If B = False Then
                ErrorText = "Error Recycle'ing file '" & DestinationFileName & "." & vbCrLf & Err.Description
                DownloadFile = False
                Exit Function
            End If
    End Select
End If

L = URLDownloadToFile(0&, UrlFileName, DestinationFileName, 0&, 0&)
If L = 0 Then
    DownloadFile = True
Else
    ErrorText = "Buffer length invalid or not enough memory."
    DownloadFile = False
End If

End Function

Private Function RecycleFileOrFolder(FileSpec As String) As Boolean

    Dim FileOperation As SHFILEOPSTRUCT
    Dim lReturn As Long

    If (Dir(FileSpec, vbNormal) = vbNullString) And _
        (Dir(FileSpec, vbDirectory) = vbNullString) Then
        RecycleFileOrFolder = True
        Exit Function
    End If

    With FileOperation
        .wFunc = FO_DELETE
        .pFrom = FileSpec
        .fFlags = FOF_ALLOWUNDO
        ' Or
        .fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION
    End With

    lReturn = SHFileOperation(FileOperation)
    If lReturn = 0 Then
        RecycleFileOrFolder = True
    Else
        RecycleFileOrFolder = False
    End If
End Function


Function DownloadLog(Url As String) As Boolean
    Dim LocalFileName As String
    Dim ErrorText As String
    LocalFileName = "C:\temp\error.log"
    DownloadLog = DownloadFile(UrlFileName:=Url, _
                     DestinationFileName:=LocalFileName, _
                     Overwrite:=OverwriteKill, _
                     ErrorText:=ErrorText)
'    If DownloadLog = True Then
'        Debug.Print "下载成功"
'    Else
'        Debug.Print "下载失败: " & ErrorText
'    End If
End Function
