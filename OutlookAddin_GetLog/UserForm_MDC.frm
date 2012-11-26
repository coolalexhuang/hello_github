VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_MDC 
   Caption         =   "Error ID Search..."
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6975
   OleObjectBlob   =   "UserForm_MDC.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_MDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by huangjia(jianbin,huang@moodys.com) on 16/Nov/2012

Private Sub ComboBox1_Change()
    TextBox2.Text = logFiles(ComboBox1.ListIndex)
End Sub

Private Sub CommandButton1_Click()
    Dim env As String
    Dim guid As String
    Dim fileName As String
    
    Dim fso 'file system object
    Dim fo  'file object
    Dim output As String
    Dim line As String
    Dim preLine As String
    Dim index As Integer
    
    env = ComboBox1.SelText
    guid = Trim(TextBox1.Text)
    fileName = logFiles(ComboBox1.ListIndex)
    
    If guid = "" Then
        MsgBox ("Please input error id.")
        Exit Sub
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(fileName) Then
        If DownloadLog(fileName) = True Then
            Set fo = fso.OpenTextFile("c:\temp\error.log")
        Else
            MsgBox ("Log file is not found, please check the file path.")
            fso = Null
            Exit Sub
        End If
    Else
        Call UpdateSetting
        Set fo = fso.OpenTextFile(fileName)
    End If
    
    On Error GoTo NoFound
    preLine = Trim(fo.readline) 'Read the first line
    Do While True
        line = Trim(fo.readline)
        index = InStr(1, line, guid, 1)
        If (index <> 0) Then
            output = preLine + VBA.vbCrLf
            Do
                output = output & line & VBA.vbCrLf
                line = Trim(fo.readline)
            Loop While InStr(line, "Extended Properties") = 0
            output = output & line & VBA.vbCrLf & "-----------------------------------------------" 'Append the last line
            Exit Do
        End If
        preLine = line
    Loop
    fo.Close
    fo = Null
    Call CreateNewMail(output, env)
    'UserForm_MDC.Hide
    Exit Sub
    
NoFound:
    MsgBox ("Error id is not found in error log file.")
    fo.Close
    fo = Null
    'UserForm_MDC.Hide
End Sub

Private Sub TextBox2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim index As Integer
    logFiles(ComboBox1.ListIndex) = TextBox2.Text
End Sub

Private Sub UserForm_Initialize()
    Call InitSetting
    ComboBox1.Value = envList(0)
End Sub

Private Sub InitSetting()
    Dim fileName As String
    Dim fso As Object
    Dim fo As Object
    Dim rowIndex As Integer
    Set fso = CreateObject("Scripting.FileSystemObject")
    fileName = "c:\temp\olAddin.ini"
    'Write init setting if file is not found
    If Not fso.FileExists(fileName) Then
        Set fo = fso.CreateTextFile(fileName, True)
        fo.WriteLine ("CI_MDC,\\sz1-dev-mdc-w03\mdc_log\error.log")
        fo.WriteLine ("CI_MAC,\\sz1-dev-mdc-w03\mac_log\error.log")
        fo.WriteLine ("E2E_MDC,\\10.6.200.91\mdclogs\error.log")
        fo.WriteLine ("E2E_MAC,\\10.6.200.91\maclogs\error.log")
        fo.WriteLine ("Staging_MDC,http://10.6.129.170:8077/mdclogs/error.log")
        fo.WriteLine ("Staging_MAC,http://10.6.129.170:8077/maclogs/error.log")
        fo.WriteLine ("Staging_MDC1,http://10.6.129.171:8077/mdclogs/error.log")
        fo.WriteLine ("Staging_MAC1,http://10.6.129.171:8077/maclogs/error.log")
        fo.WriteLine ("More...,Please config in C:\temp\olAddin.ini (max to 30 lines)")
    End If
    
    'Read settings from file
    Set fo = fso.OpenTextFile(fileName)
    
    'logFiles = Array(Trim(fo.readline), Trim(fo.readline), Trim(fo.readline), Trim(fo.readline))
    rowIndex = 0
    Do While (fo.AtEndofStream = False)
        line = Trim(fo.readline)
        If line = "" Then
            Exit Do
        End If
        
        If rowIndex = 30 Then
            MsgBox ("More than 30 lines configed in c:\temp\olAddin.ini, will only load the first 30 lines.")
            Exit Do
        End If
        ComboBox1.AddItem (Trim(VBA.Split(line, ",")(0)))
        envList(rowIndex) = Trim(VBA.Split(line, ",")(0))
        logFiles(rowIndex) = Trim(VBA.Split(line, ",")(1))
        rowIndex = rowIndex + 1
    Loop
    fo.Close
    
End Sub

Private Sub UpdateSetting()
    Dim fileName As String
    Dim fso As Object
    Dim fo As Object
    Dim rowIndex As Integer
    fileName = "c:\temp\olAddin.ini"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fo = fso.CreateTextFile(fileName, True)
    rowIndex = 0
    Do While (rowIndex < ComboBox1.ListCount)
        fo.WriteLine (envList(rowIndex) & "," & logFiles(rowIndex))
        rowIndex = rowIndex + 1
    Loop
    fo.Close
End Sub

