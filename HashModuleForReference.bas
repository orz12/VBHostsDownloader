Attribute VB_Name = "Module1"
Option Explicit
Public Sub Main()
'run this to test the file hasher. Select or comment lines as necessary
    'enter your own paths for the files to test
    'Set a reference to mscorlib 4.0 64-bit
    
    Dim sFPath As String, b64 As Boolean, sH As String
    
    'set output type and path to target file
    'b64 = False   'output hex
    b64 = True     'output base-64
    sFPath = "C:\Users\ms_ly\OneDrive\ÎÄµµ\GitHub\VBHostsDownloader\VBHostsDownloader.exe"
    
    'enable any one line to test hash
    'sh=FileToMD5(sFPath, b64)
    'sh=FileToSHA1(sFPath, b64)
    sH = FileToSHA256(sFPath, b64)
    'sH = FileToSHA384(sFPath, b64)
    'sH = FileToSHA512(sFPath, b64)
    
    Debug.Print sH & vbNewLine & Len(sH) & " characters in length"
    MsgBox sH & vbNewLine & Len(sH) & " characters in length"
End Sub

Private Sub TestFileHashes()
    

End Sub

Public Function FileToMD5(sFullPath As String, Optional bB64 As Boolean = False) As String
    'parameter full path with name of file returned in the function as an MD5 hash
    'Set a reference to mscorlib 4.0 64-bit
    
    Dim enc, bytes, outstr As String, pos As Integer
    
    Set enc = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    'Convert the string to a byte array and hash it
    bytes = GetFileBytes(sFullPath)
    bytes = enc.ComputeHash_2((bytes))
    
    If bB64 = True Then
       FileToMD5 = ConvToBase64String(bytes)
    Else
       FileToMD5 = ConvToHexString(bytes)
    End If
        
    Set enc = Nothing

End Function

Public Function FileToSHA1(sFullPath As String, Optional bB64 As Boolean = False) As String
    'parameter full path with name of file returned in the function as an SHA1 hash
    'Set a reference to mscorlib 4.0 64-bit
    
    Dim enc, bytes, outstr As String, pos As Integer
    
    Set enc = CreateObject("System.Security.Cryptography.SHA1CryptoServiceProvider")
    'Convert the string to a byte array and hash it
    bytes = GetFileBytes(sFullPath) 'returned as a byte array
    bytes = enc.ComputeHash_2((bytes))
    
    If bB64 = True Then
       FileToSHA1 = ConvToBase64String(bytes)
    Else
       FileToSHA1 = ConvToHexString(bytes)
    End If
        
    Set enc = Nothing

End Function

Public Function FileToSHA256(sFullPath As String, Optional bB64 As Boolean = False) As String
    'parameter full path with name of file returned in the function as an SHA2-256 hash
    'Set a reference to mscorlib 4.0 64-bit
    
    Dim enc, bytes, outstr As String, pos As Integer
    
    Set enc = CreateObject("System.Security.Cryptography.SHA256Managed")
    'Convert the string to a byte array and hash it
    bytes = GetFileBytes(sFullPath) 'returned as a byte array
    bytes = enc.ComputeHash_2((bytes))
    
    If bB64 = True Then
       FileToSHA256 = ConvToBase64String(bytes)
    Else
       FileToSHA256 = ConvToHexString(bytes)
    End If
        
    Set enc = Nothing

End Function

Public Function FileToSHA384(sFullPath As String, Optional bB64 As Boolean = False) As String
    'parameter full path with name of file returned in the function as an SHA2-384 hash
    'Set a reference to mscorlib 4.0 64-bit
    
    Dim enc, bytes, outstr As String, pos As Integer
    
    Set enc = CreateObject("System.Security.Cryptography.SHA384Managed")
    'Convert the string to a byte array and hash it
    bytes = GetFileBytes(sFullPath) 'returned as a byte array
    bytes = enc.ComputeHash_2((bytes))
    
    If bB64 = True Then
       FileToSHA384 = ConvToBase64String(bytes)
    Else
       FileToSHA384 = ConvToHexString(bytes)
    End If
    
    Set enc = Nothing

End Function

Public Function FileToSHA512(sFullPath As String, Optional bB64 As Boolean = False) As String
    'parameter full path with name of file returned in the function as an SHA2-512 hash
    'Set a reference to mscorlib 4.0 64-bit
    
    Dim enc, bytes, outstr As String, pos As Integer
    
    Set enc = CreateObject("System.Security.Cryptography.SHA512Managed")
    'Convert the string to a byte array and hash it
    bytes = GetFileBytes(sFullPath) 'returned as a byte array
    bytes = enc.ComputeHash_2((bytes))
    
    If bB64 = True Then
       FileToSHA512 = ConvToBase64String(bytes)
    Else
       FileToSHA512 = ConvToHexString(bytes)
    End If
    
    Set enc = Nothing

End Function

Private Function GetFileBytes(ByVal sPath As String) As Byte()
    'makes byte array from file
    'Set a reference to mscorlib 4.0 64-bit
    
    Dim lngFileNum As Long, bytRtnVal() As Byte, bTest
    
    lngFileNum = FreeFile
    
    If LenB(Dir(sPath)) Then ''// Does file exist?
        
        Open sPath For Binary Access Read As lngFileNum
        
        'a zero length file content will give error 9 here
        
        ReDim bytRtnVal(0 To LOF(lngFileNum) - 1&) As Byte
        Get lngFileNum, , bytRtnVal
        Close lngFileNum
    Else
        Err.Raise 53 'File not found
    End If
    
    GetFileBytes = bytRtnVal
    
    Erase bytRtnVal

End Function

Function ConvToBase64String(vIn As Variant) As Variant
    'used to produce a base-64 output
    'Set a reference to mscorlib 4.0 64-bit
    
    Dim oD As Object
      
    Set oD = CreateObject("MSXML2.DOMDocument")
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.base64"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToBase64String = Replace(oD.DocumentElement.Text, vbLf, "")
    
    Set oD = Nothing

End Function

Function ConvToHexString(vIn As Variant) As Variant
    'used to produce a hex output
    'Set a reference to mscorlib 4.0 64-bit
    
    Dim oD As Object
      
    Set oD = CreateObject("MSXML2.DOMDocument")
      
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.Hex"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToHexString = Replace(oD.DocumentElement.Text, vbLf, "")
    
    Set oD = Nothing

End Function
