Attribute VB_Name = "Module1"
Option Explicit
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" ( _
                                                            ByVal pCaller As Long, _
                                                            ByVal szURL As String, _
                                                            ByVal szFileName As String, _
                                                            ByVal dwReserved As Long, _
                                                            ByVal lpfnCB As Long _
                                                            ) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long


Private Const ERROR_ALREADY_EXISTS = 183&
Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (lpMutexAttributes As Any, _
    ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function MoveFileEx Lib "kernel32" Alias "MoveFileExA" (ByVal lpExistingFileName As String, _
    ByVal lpNewFileName As String, ByVal dwFlags As Long) As Long
    
    
    
Const MOVEFILE_REPLACE_EXISTING = &H1
Const MOVEFILE_DELAY_UNTIL_REBOOT = &H4
Const MOVEFILE_WRITE_THROUGH = &H8


 Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
 
' Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpbuffer As String, ByVal nSize As Long) As Long
Public Const MAX_PATH = 260


'Public Function GetSysPath() As String 'System32
'    Dim Buffer As String
'    Buffer = Space(MAX_PATH)
'    If GetSystemDirectory(Buffer, Len(Buffer)) <> 0 Then
'        GetSysPath = Mid(Trim(Buffer), 1, Len(Trim(Buffer)) - 1)
'    End If
'End Function
 
'Public Function DownloadFile(ByVal strURL As String, ByVal strFile As String) As Boolean
'   DownloadFile = URLDownloadToFile(0, strURL, strFile, 0, 0) = 0
'End Function


'Private Function IsAlreadyRunning() As Boolean
'Dim hMutex As Long
'hMutex = CreateMutex(ByVal 0&, 1, App.Title)

'If (Err.LastDllError = ERROR_ALREADY_EXISTS) Then
'------------
'Cleaning up.
'------------

'ReleaseMutex hMutex
'CloseHandle hMutex
'--------------------------------
'More than one instance detected.
'--------------------------------
'IsAlreadyRunning = True

'Else
'IsAlreadyRunning = False
'End If
'End Function


Sub Main()

    On Error Resume Next
    Dim GetSysPath As String
    GetSysPath = Space(MAX_PATH)
    If GetSystemDirectory(GetSysPath, Len(GetSysPath)) <> 0 Then
        GetSysPath = Mid(Trim(GetSysPath), 1, Len(Trim(GetSysPath)) - 1)
    Else
        GetSysPath = "C:\Windows\System32"
    End If

    If UCase(Command()) = "/Q" Or UCase(Command()) = "-Q" Then
    
        URLDownloadToFile 0, _
            "https://raw.githubusercontent.com/WUZHIQIANGX/hosts/master/hosts", _
            GetSysPath & "\drivers\etc\hosts", 0, 0
        End
        
    End If
    
    Dim AppPath As String, CurrentVersion As String
    AppPath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\") & App.EXEName
    
    Dim hMutex As Long, bUpdated As Boolean
    
    If UCase(Command()) = "/D" Or UCase(Command()) = "-D" Then
    
        Do
        
            hMutex = CreateMutex(ByVal 0&, 1, App.Title)
            bUpdated = (Err.LastDllError = ERROR_ALREADY_EXISTS) 'still running
            ReleaseMutex hMutex
            CloseHandle hMutex
            
            Sleep 50
            DoEvents
            'MsgBox bUpdated
            
        Loop While bUpdated
        
        Sleep 500
        DoEvents
    
        SetAttr AppPath & ".tmp", 0
        Kill AppPath & ".tmp"
        SetAttr AppPath & ".tmp", vbReadOnly Or vbHidden Or vbSystem 'when delete action failed
        'MsgBox "deleted!"
        End
    
    End If
    
    hMutex = CreateMutex(ByVal 0&, 1, App.Title)

    If Err.LastDllError = ERROR_ALREADY_EXISTS Or App.PrevInstance Then
        
        MsgBox "Application is already running. Please wait for a while, or terminate it by yourself.", _
        vbCritical Or vbSystemModal
        
    Else
    
        MsgBox IIf(URLDownloadToFile(0, _
            "https://raw.githubusercontent.com/WUZHIQIANGX/hosts/master/hosts", _
            GetSysPath & "\drivers\etc\hosts", 0, 0) = 0, _
            "    Done successfully. Enjoy now!", "     Access denied!     " _
            & vbCrLf & vbCrLf & "GetLastErrorCode:" & GetLastError & "(" & Err.LastDllError & "#" & Err.Number & ")"), _
            vbInformation Or vbSystemModal
            
    End If
        
    
    If URLDownloadToFile(0, "https://raw.githubusercontent.com/orz12/VBHostsDownloader/master/version.txt", _
            AppPath & "version.txt", 0, 0) = 0 Then
            
        Dim FreeFileHandle As Integer, strNewVerDetail As String
        strNewVerDetail = vbCrLf
        FreeFileHandle = FreeFile
        Open AppPath & "version.txt" For Input As #FreeFileHandle
            Line Input #FreeFileHandle, CurrentVersion
            Do Until EOF(FreeFileHandle) 'Backward Compatibility
              Line Input #1, strNewVerDetail
              strNewVerDetail = strNewVerDetail + vbCrLf
            Loop
        Close #FreeFileHandle
        
        Kill AppPath & "version.txt"
        
        If Len(CurrentVersion) > 0 And CurrentVersion <> App.Major & "." & App.Minor & "." & App.Revision Then
            
        
            If MsgBox("    New version (" & CurrentVersion & ") available!" & vbCrLf & vbCrLf & "Would you like to download it now?" & strNewVerDetail, vbInformation Or vbOKCancel) = vbOK Then
            
                
                If URLDownloadToFile(0, "https://github.com/orz12/VBHostsDownloader/blob/master/VBHostsDownloader.exe?raw=true", _
                        AppPath & "new", 0, 0) = 0 Then
                        
                    If MoveFileEx(AppPath & ".exe", AppPath & ".tmp", MOVEFILE_REPLACE_EXISTING Or MOVEFILE_WRITE_THROUGH) Then
                    
                        SetAttr AppPath & ".tmp", GetAttr(AppPath & ".tmp") Or vbHidden Or vbSystem
                
                        If MoveFileEx(AppPath & "new", AppPath & ".exe", MOVEFILE_REPLACE_EXISTING) Then
                        
                            MsgBox "Updated. Congratulations!", vbInformation
                            bUpdated = True
                            
                            
                        End If
                        
                        
                    End If
                    
                    Kill AppPath & "new"
                    
                    'MoveFileEx AppPath & App.EXEName & "new", "", MOVEFILE_DELAY_UNTIL_REBOOT
                    MoveFileEx AppPath & ".tmp", vbNullString, MOVEFILE_DELAY_UNTIL_REBOOT
                    Shell AppPath & ".exe /d"   'try to kill old versionfile immediately.
                    'MoveFileEx AppPath & "version.txt", vbNull, MOVEFILE_DELAY_UNTIL_REBOOT
                    
                    
                End If
                
                
                If Not bUpdated Then MsgBox "Access denied! GetLastErrorCode:" & _
                    GetLastError & "(" & Err.LastDllError & "#" & Err.Number & ")", vbInformation
                
            End If
            
            
        End If
        
        
        
    'Else
    '    MsgBox "1GetLastErrorCode:" & GetLastError & "(" & Err.LastDllError & "#" & Err.Number & ")"
    
    End If
    
    ReleaseMutex hMutex
    CloseHandle hMutex
    
End Sub
