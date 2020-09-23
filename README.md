<div align="center">

## Enumertate Windows NT Services


</div>

### Description

Populates a collection of installed Windows NT services, based upon type of service requested. Can be used to enumerate the services on a workstation or server. Requires Windows NT and administrator rights.
 
### More Info
 
SVC [output collection] = the collection to polulate, DisplayName [boolean] = return display names or service names.

Requires Windows NT and administrator rights.

Function returns the number of service found, and populates the collection.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matthew Ruffell](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matthew-ruffell.md)
**Level**          |Unknown
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matthew-ruffell-enumertate-windows-nt-services__1-3789/archive/master.zip)

### API Declarations

```
Private Type SERVICE_STATUS
 dwServiceType As Long
 dwCurrentState As Long
 dwControlsAccepted As Long
 dwWin32ExitCode As Long
 dwServiceSpecificExitCode As Long
 dwCheckPoint As Long
 dwWaitHint As Long
End Type
Private Type ENUM_SERVICE_STATUS
 lpServiceName As Long
 lpDisplayName As Long
 ServiceStatus As SERVICE_STATUS
End Type
Private Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function EnumServicesStatus Lib "advapi32.dll" Alias "EnumServicesStatusA" (ByVal hSCManager As Long, ByVal dwServiceType As Long, ByVal dwServiceState As Long, lpServices As Any, ByVal cbBufSize As Long, pcbBytesNeeded As Long, lpServicesReturned As Long, lpResumeHandle As Long) As Long
Private Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hSCObject As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
```


### Source Code

```
Public Function EnumerateServices(colSVC As Collection, bolDisplayName As Boolean, Optional lngServiceType As Variant, Optional lngServiceState As Variant, Optional strMachineName As Variant) As Long
 '// lngServiceType = 0 (win32 services)
 '// lngServiceType = 1 (driver services)
 '// lngServiceState = 0 (active and inactive services)
 '// lngServiceState = 1 (active services)
 '// lngServiceState = 2 (inactive services)
 Dim hSCM As Long
 Dim lngBytesNeeded As Long
 Dim lngResumeHandle As Long
 Dim lngServicesReturned As Long
 Dim lngStructsNeeded As Long
 Dim lngServiceStatusInfoBuffer As Long
 Dim lngSVCReturnCode As Long
 Dim lngI As Long
 Dim strSVCName As String * 250
 Dim lpEnumServiceStatus() As ENUM_SERVICE_STATUS
 On Error Resume Next
 If IsMissing(lngServiceType) = True Then lngServiceType = 0 Else lngServiceType = CLng(lngServiceType)
 If IsMissing(lngServiceState) = True Then lngServiceState = 0 Else lngServiceState = CLng(lngServiceState)
 If IsMissing(strMachineName) = True Then strMachineName = vbNullString Else strMachineName = CStr(strMachineName)
 If lngServiceType = 0 Then lngServiceType = 30
 If lngServiceType = 1 Then lngServiceType = 11
 If lngServiceState = 0 Then lngServiceState = 3
 If lngServiceState = 1 Then lngServiceState = &H1
 If lngServiceState = 2 Then lngServiceState = &H2
 '// Open the service manager
 hSCM = OpenSCManager(strMachineName, vbNullString, &H4)
 If hSCM = 0 Then Exit Function '// error opening
 '// Get buffer size (bytes) without passing a buffer
 Call EnumServicesStatus(hSCM, lngServiceType, lngServiceState, ByVal &H0, &H0, lngBytesNeeded, lngServicesReturned, lngResumeHandle)
 '// We should receive MORE_DATA error
 If Not Err.LastDllError = 234 Then
 Call CloseServiceHandle(hSCM)
 Exit Function
 End If
 '// Calculate the number of structures needed and redimention array
 lngStructsNeeded = lngBytesNeeded / Len(lpEnumServiceStatus(0)) + 1
 ReDim lpEnumServiceStatus(lngStructsNeeded - 1)
 '// Get buffer size in bytes
 lngServiceStatusInfoBuffer = lngStructsNeeded * Len(lpEnumServiceStatus(0))
 '// Get services information starting entry 0
 lngResumeHandle = 0
 lngSVCReturnCode = EnumServicesStatus(hSCM, lngServiceType, lngServiceState, lpEnumServiceStatus(0), lngServiceStatusInfoBuffer, lngBytesNeeded, lngServicesReturned, lngResumeHandle)
 If lngSVCReturnCode <> 0 Then
 For lngI = 0 To lngServicesReturned - 1
  If bolDisplayName = True Then
  Call lstrcpy(ByVal strSVCName, ByVal lpEnumServiceStatus(lngI).lpDisplayName)
  Else
  Call lstrcpy(ByVal strSVCName, ByVal lpEnumServiceStatus(lngI).lpServiceName)
  End If
  colSVC.Add StripTerminator(strSVCName)
 Next
 End If
 Call CloseServiceHandle(hSCM)
 EnumerateServices = colSVC.Count
End Function
Private Function StripTerminator(ByVal strString As String) As String
 If InStr(strString, Chr(0)) > 0 Then StripTerminator = Left(strString, InStr(strString, Chr(0)) - 1) Else StripTerminator = strString
End Function
```

