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

    If UCase(Command()) = "/Q" Or UCase(Command()) = "-Q" Then URLDownloadToFile 0, _
            "https://raw.githubusercontent.com/WUZHIQIANGX/hosts/master/hosts", _
            "C:\Windows\System32\drivers\etc\hosts", 0, 0: End
            
    Dim hMutex As Long
    hMutex = CreateMutex(ByVal 0&, 1, App.Title)
    
    If Err.LastDllError = ERROR_ALREADY_EXISTS Or App.PrevInstance = True Then
    
        MsgBox "Application is already running. Please wait for a while, or terminate it by yourself.", _
        vbCritical Or vbSystemModal, "Hosts Downloader by LouizQ"
        
    Else
    
        MsgBox IIf(URLDownloadToFile(0, _
            "https://raw.githubusercontent.com/WUZHIQIANGX/hosts/master/hosts", _
            "C:\Windows\System32\drivers\etc\hosts", 0, 0) = 0, _
            "    Done successfully. Enjoy now!", "     Access denied!     " _
            & vbCrLf & vbCrLf & "GetLastErrorCode:" & GetLastError & "(" & Err.LastDllError & "#" & Err.Number & ")"), _
            vbInformation Or vbSystemModal, "Hosts Downloader by LouizQ"
            
    End If
    
    ReleaseMutex hMutex
    CloseHandle hMutex
    
    Dim AppPath As String, CurrentVersion As String
    
    AppPath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
    
    If URLDownloadToFile(0, "https://raw.githubusercontent.com/orz12/VBHostsDownloader/master/version.txt", _
            AppPath & "version.txt", 0, 0) = 0 Then
            
        On Error Resume Next
        
        Open AppPath & "version.txt" For Input As #1
        Line Input #1, CurrentVersion
        Close #1
        
        If Len(CurrentVersion) > 0 And CurrentVersion <> App.Major & "." & App.Minor & "." & App.Revision Then
        
            If MsgBox("    New version available!" & vbCrLf & vbCrLf & "Would you like to download it now?", vbInformation Or vbOKCancel) = vbOK Then
            
                Dim bUpdated As Boolean
                
                If URLDownloadToFile(0, "https://github.com/orz12/VBHostsDownloader/blob/master/VBHostsDownloader.exe?raw=true", _
                        AppPath & App.EXEName & "new", 0, 0) = 0 Then
                        
                    If MoveFileEx(AppPath & App.EXEName & ".exe", AppPath & App.EXEName & "backup", MOVEFILE_REPLACE_EXISTING Or MOVEFILE_WRITE_THROUGH) Then
                
                        If MoveFileEx(AppPath & App.EXEName & "new", AppPath & App.EXEName & ".exe", MOVEFILE_REPLACE_EXISTING) Then
                        
                            MsgBox "Updated. Congratulations!", vbInformation, "Hosts Downloader by LouizQ"
                            bUpdated = True
                            
                            
                        End If
                        
                    End If
                    
                    MoveFileEx AppPath & App.EXEName & "new", "", MOVEFILE_DELAY_UNTIL_REBOOT
                    MoveFileEx AppPath & App.EXEName & "backup", "", MOVEFILE_DELAY_UNTIL_REBOOT
                    'MoveFileEx AppPath & "version.txt", "", MOVEFILE_DELAY_UNTIL_REBOOT
                    
                    
                End If
                
            Else
                
                If Not bUpdated Then MsgBox "Access denied! GetLastErrorCode:" & GetLastError & "(" & Err.LastDllError & "#" & Err.Number & ")", vbInformation, "Hosts Downloader by LouizQ"
                
            End If
            
            
        End If
        
        
    'Else
    '    MsgBox "1GetLastErrorCode:" & GetLastError & "(" & Err.LastDllError & "#" & Err.Number & ")"
    
    End If
    
End Sub
