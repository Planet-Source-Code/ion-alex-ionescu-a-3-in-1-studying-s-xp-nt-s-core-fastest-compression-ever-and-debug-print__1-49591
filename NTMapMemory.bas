Attribute VB_Name = "MapMemoryModule"
'// Project:        Planet Source Code Undocumented NT Functions
'// Name:           NT Native Compression
'// Written by:     Alex Ionescu
'// Description:    Quick and easy compression (very fast too)
'// Remarks:        Copyright © 2003 Alex Ionescu. All Rights Reserved.

Option Explicit
' // ********************
' //      CONSTANTS
' // ********************

Public Const GENERIC_WRITE As Long = &H40000000
Public Const GENERIC_READ As Long = &H80000000
Public Const FILE_SHARE_WRITE As Long = &H2
Public Const FILE_SHARE_READ As Long = &H1
Public Const FILE_MAP_WRITE As Long = &H2
Public Const OPEN_ALWAYS As Long = 4
Public Const PAGE_READWRITE As Long = &H4
Public Const SEC_COMMIT As Long = &H8000000
Public Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Public Const SECTION_EXTEND_SIZE As Long = &H10
Public Const SECTION_MAP_EXECUTE As Long = &H8
Public Const SECTION_MAP_READ As Long = &H4
Public Const SECTION_MAP_WRITE As Long = &H2
Public Const SECTION_QUERY As Long = &H1
Public Const SECTION_ALL_ACCESS As Long = STANDARD_RIGHTS_REQUIRED Or SECTION_QUERY Or SECTION_MAP_WRITE Or SECTION_MAP_READ Or SECTION_MAP_EXECUTE Or SECTION_EXTEND_SIZE

' // ********************
' //      API CALLS
' // ********************

' // File Mapping APIs
Public Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function NtCreateSection Lib "ntdll.dll" (Handle As Long, ByVal DesiredAcess As Long, ObjectAttributes As Any, SectionSize As Any, ByVal Protect As Long, ByVal Attributes As Long, ByVal FileHandle As Long) As Long
Public Declare Function NtMapViewOfSection Lib "ntdll.dll" (ByVal Handle As Long, ByVal ProcessHandle As Long, BaseAddress As Long, ByVal ZeroBits As Long, ByVal CommitSize As Long, SectionOffset As Any, ViewSize As Long, ByVal InheritDisposition As Long, ByVal AllocaitonType As Long, ByVal Protect As Long) As Long
Public Declare Function NtUnmapViewOfSection Lib "ntdll.dll" (ByVal ProcessHandle As Long, ByVal Handle As Long) As Long
Public Declare Function NtClose Lib "ntdll.dll" (ByVal hObject As Long) As Long

' // File Size APIs
Public Declare Function GetFileSize Lib "kernel32.dll" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Public Declare Function SetEndOfFile Lib "kernel32.dll" (ByVal hFile As Long) As Long
Public Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, ByVal liDistanceToMove As Long, ByVal lpNewFilePointer As Long, ByVal dwMoveMethod As Long) As Long

' // Pointer de-referencing and general memory copy API
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' // ********************
' //     MAP FUNCTION
' // ********************

Public Function OpenFile(File As String, Size As Long, FileHandle As Long, MemoryHandle As Long) As Long
' // Opens a File and maps in into memory, returning the pointer to the starting address
' // This is the fastest file I/O for anything that uses Memory Pointers
' // IN: File to load
' // OUT: Pointer to file loaded in memory, size of file, and handle

    Dim BaseAddress As Long
    
    FileHandle = CreateFile(File, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_ALWAYS, 0&, 0&)  ' // Open the file from disk
    If Size = 0 Then Size = GetFileSize(FileHandle, 0)                                                                          ' // Get size of the file if not there already
    NtCreateSection MemoryHandle, SECTION_ALL_ACCESS, ByVal 0&, Size, PAGE_READWRITE, SEC_COMMIT, FileHandle                    ' // Load file to memory
    NtMapViewOfSection MemoryHandle, -1, BaseAddress, 0&, Size, 0&, Size, 1, 0, PAGE_READWRITE                                  ' // Map it into memory
    OpenFile = BaseAddress                                                                                                      ' // Return the memory pointer
End Function
