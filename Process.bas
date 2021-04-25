Attribute VB_Name = "Process"
' Gestion des processus
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Const STILL_ACTIVE = &H103
Public Const PROCESS_QUERY_INFORMATION = &H400

' Lancement d'un Programme et attente de fin
Public Sub Shell32Bit(ByVal JobToDo As String)
Dim hProcess As Long
Dim RetVal As Long
  
'The next line launches JobToDo as icon,
'captures process ID
hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(JobToDo, vbMinimizedNoFocus))
Do
 'Get the status of the process
 GetExitCodeProcess hProcess, RetVal
 'Sleep command recommended as well
 'as DoEvents
 DoEvents
 Sleep 100
 'Loop while the process is active
Loop While RetVal = STILL_ACTIVE

End Sub
