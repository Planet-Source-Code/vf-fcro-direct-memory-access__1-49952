Attribute VB_Name = "Module1"
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function GlobalFree Lib "kernel32" (ByVal HMEM As Long) As Long

Public Enum Read_Write
[ReadMem]
[WriteMem]
End Enum


Public Type MEMORY
Data(2) As Long 'Access 12 byte as DWORD / LONG alignment
End Type




Public Sub AccessMemory(ByVal MemHandle As Long, ByVal RW As Read_Write)
CallWindowProc AddressOf DirectAccessMemory, MemHandle, RW, 0, 0
End Sub

Public Function DirectAccessMemory(MEMPOINT As MEMORY, ByVal RW As Read_Write, ByVal notused2 As Long, ByVal notused3 As Long) As Long

Select Case RW
Case [ReadMem]
'Directly read from memory pointer
MsgBox "Values:" & MEMPOINT.Data(0) & "," & MEMPOINT.Data(1) & "," & MEMPOINT.Data(2)


Case [WriteMem]
'Direct write to memory pointer
MEMPOINT.Data(0) = 989
MEMPOINT.Data(1) = 122
MEMPOINT.Data(2) = 525

End Select

End Function

