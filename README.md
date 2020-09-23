<div align="center">

## KillApp


</div>

### Description

Kill any application or process running if you know the .exe name. (Only Windows 95/98)
 
### More Info
 
myName: is the name of the app that wou want to kill (ex. "app.exe")

True if the application was killed correctly


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Fernando Vicente](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/fernando-vicente.md)
**Level**          |Unknown
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/fernando-vicente-killapp__1-2160/archive/master.zip)

### API Declarations

```
Const MAX_PATH& = 260
Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Type PROCESSENTRY32
 dwSize As Long
 cntUsage As Long
 th32ProcessID As Long
 th32DefaultHeapID As Long
 th32ModuleID As Long
 cntThreads As Long
 th32ParentProcessID As Long
 pcPriClassBase As Long
 dwFlags As Long
 szexeFile As String * MAX_PATH
End Type
```


### Source Code

```
Public Function KillApp(myName As String) As Boolean
 Const PROCESS_ALL_ACCESS = 0
 Dim uProcess As PROCESSENTRY32
 Dim rProcessFound As Long
 Dim hSnapshot As Long
 Dim szExename As String
 Dim exitCode As Long
 Dim myProcess As Long
 Dim AppKill As Boolean
 Dim appCount As Integer
 Dim i As Integer
 On Local Error GoTo Finish
 appCount = 0
 Const TH32CS_SNAPPROCESS As Long = 2&
 uProcess.dwSize = Len(uProcess)
 hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
 rProcessFound = ProcessFirst(hSnapshot, uProcess)
 Do While rProcessFound
 i = InStr(1, uProcess.szexeFile, Chr(0))
 szExename = LCase$(Left$(uProcess.szexeFile, i - 1))
 If Right$(szExename, Len(myName)) = LCase$(myName) Then
  KillApp = True
  appCount = appCount + 1
  myProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
  AppKill = TerminateProcess(myProcess, exitCode)
  Call CloseHandle(myProcess)
 End If
 rProcessFound = ProcessNext(hSnapshot, uProcess)
 Loop
 Call CloseHandle(hSnapshot)
Finish:
End Function
```

