Attribute VB_Name = "Module1"
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
            "    Done successfully. Click OK to exit and enjoy now!", "     Access denied!     " _
            & vbCrLf & vbCrLf & "GetLastErrorCode:" & GetLastError & "(" & Err.LastDllError & "#" & Err.Number & ")"), _
            vbInformation Or vbSystemModal, "Hosts Downloader by LouizQ"
            
    End If
    
    ReleaseMutex hMutex
    CloseHandle hMutex
    
End Sub
