Attribute VB_Name = "DebugViewModule"
'// Project:        Planet Source Code Debug Viewer and Debug.Print in Compiled Mode
'// Name:           Debug.Print in Compiled mode and catching it.
'// Written by:     Alex Ionescu
'// Description:    Shows how to make your compiled message display a message in the Windows Debug Output Buffer
'// Remarks:        Copyright © 2003 Alex Ionescu. All Rights Reserved.

'//                 If you want a GUI, this *NEEDS* to be multi-threaded. Unfortunately, VB makes this hard.
'//                 This code is great for showing in a Message Box or saving to a log file however.
'//                 If you really need a GUI, I suggest using the award-winning MultiThreadVB code on PSC.

Option Explicit

' // ********************
' //   API DECLARATIONS
' // ********************

' // API to show a string from your program, equivalent to Debug.Print
Public Declare Sub OutputDebugString Lib "kernel32.dll" Alias "OutputDebugStringA" (ByVal lpOutputString As String)

' // APIs to catch the Debug Event
Public Declare Function CreateEvent Lib "kernel32.dll" Alias "CreateEventA" (ByVal lpEventAttributes As Long, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
Public Declare Function SetEvent Lib "kernel32.dll" (ByVal hEvent As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

' // API to receive/store the Debug Output Buffer
Public Declare Function CreateFileMapping Lib "kernel32.dll" Alias "CreateFileMappingA" (ByVal hFile As Long, ByVal lpFileMappigAttributes As Long, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Public Declare Function MapViewOfFile Lib "kernel32.dll" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long

' // General Memory Manipulation API
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' // ********************
' //      CONSTANTS
' // ********************

' // Constants for the File Mapping
Public Const PAGE_READWRITE As Long = &H4
Public Const FILE_MAP_WRITE As Long = &H2
Public Const INFINITE As Long = &HFFFFFFFF

' // ********************
' //    MAIN FUNCTION
' // ********************

Sub Main()
' // Catch OutputDebugString and show a Message Box

' // Declare all the variables
Dim hAckEvent As Long, hReadyEvent As Long, hSharedFile As Long
Dim pSharedMem As Long
Dim Buffer As String * 508
Dim PID As Long

' // Create the Events
hAckEvent = CreateEvent(0, 0, 0, "DBWIN_BUFFER_READY")
hReadyEvent = CreateEvent(0, 0, 0, "DBWIN_DATA_READY")

' // Create the shared buffer
hSharedFile = CreateFileMapping(-1, 0, PAGE_READWRITE, 0, 4096, "DBWIN_BUFFER")
pSharedMem = MapViewOfFile(hSharedFile, FILE_MAP_WRITE, 0, 0, 512)

' // Start the loop
Do
    SetEvent hAckEvent                                                          ' // We are ready
    WaitForSingleObject hReadyEvent, INFINITE                                   ' // Wait for a message
    CopyMemory ByVal VarPtr(PID), ByVal pSharedMem, 4                           ' // Copy the PID
    CopyMemory ByVal Buffer, ByVal pSharedMem + 4, 508                          ' // Copy the Message
    MsgBox PID & ": " & Left(Buffer, InStr(1, Buffer, vbNullChar) - 1)          ' // Display PID: Message
Loop
    
End Sub

