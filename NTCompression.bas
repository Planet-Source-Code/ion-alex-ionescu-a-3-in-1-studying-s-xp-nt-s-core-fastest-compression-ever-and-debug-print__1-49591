Attribute VB_Name = "NTCompression"
'// Project:        Planet Source Code Undocumented NT Functions
'// Name:           NT Native Compression
'// Written by:     Alex Ionescu
'// Description:    Quick and easy compression (very fast too)
'// Remarks:        Copyright © 2003 Alex Ionescu. All Rights Reserved.
'//                 234MB ISO File (Heavily compressed): 20 seconds, 2% ratio with this application.
'//                 234MB ISO File (Heavily compressed): 38 seconds, 2% ratio with WinZip.
'//                 Compression Speed: 94MBit/s

'//                 5MB BMP file (non-compressed): 90% ratio with this application (under 100ms).
'//                 5MB BMP file (non-compressed): 98% ratio with this Winzip (200ms-300ms).

'//                 Code needed to use WinZip Libraries (Infozip) in VB: 300kb worth of DLLs, plus
'//                 big class modules or wrapper bas modules filled with Callbacks and complex code.
'//                 Code needed to use NT Compression in VB: This small module. PERIOD.

'//                 Compressing a memory pointer (anything in memory) with zip/rar/any library: IMPOSSIBLE
'//                 Compressing a memory pointer with VB compression written: SLOW
'//                 Compressing a memory pointer with this module: FAST

Option Explicit

' // ********************
' //      CONSTANTS
' // ********************

' // Compression Format Constants
Public Enum CompressionFormats
    COMPRESSION_FORMAT_NONE = &H0
    COMPRESSION_FORMAT_DEFAULT = &H1
    COMPRESSION_FORMAT_LZNT1 = &H2
    COMPRESSION_FORMAT_NS3 = &H3
    COMPRESSION_FORMAT_NS15 = &HF
End Enum

' // Compression Engine Constants
Public Enum CompressionEngines
    COMPRESSION_ENGINE_STANDARD = &H0
    COMPRESSION_ENGINE_MAXIMUM = &H100
    COMPRESSION_ENGINE_HIBER = &H200
End Enum

' // Buffer Manipulation Constants
Public Const MEM_COMMIT As Long = &H1000
Public Const MEM_DECOMMIT As Long = &H4000
Public Const PAGE_EXECUTE_READWRITE As Long = &H40

' // ********************
' //      API CALLS
' // ********************

' // Compression API
Declare Function RtlCompressBuffer Lib "NTDLL" ( _
    ByVal CompressionFormatAndEngine As Integer, _
    ByVal UnCompressedBuffer As Long, _
    ByVal UnCompressedBufferSize As Long, _
    ByVal CompressedBuffer As Long, _
    ByVal CompressedBufferSize As Long, _
    ByVal UncompressedChunkSize As Long, _
    FinalCompressedSize As Long, _
    ByVal Workspace As Long) As Long

' // Decompression API
Declare Function RtlDecompressBuffer Lib "NTDLL" ( _
    ByVal CompressionFormat As Integer, _
    ByVal UnCompressedBufferPtr As Long, _
    ByVal UnCompressedBufferSize As Long, _
    ByVal CompressedBuffer As Long, _
    ByVal CompressedBufferSize As Long, _
    FinalCompressedSize As Long) As Long
    
' // Initialize Compression API
Declare Function RtlGetCompressionWorkSpaceSize Lib "NTDLL" ( _
    ByVal CompressionFormatAndEngine As Integer, _
    CompressBufferWorkSpaceSize As Long, _
    CompressFragmentWorkSpaceSize As Long) As Long
    
' // Buffer Allocation and Deallocation APIs
Declare Function NtAllocateVirtualMemory Lib "ntdll.dll" ( _
    ByVal ProcessHandle As Long, _
    BaseAddress As Long, _
    ByVal ZeroBits As Long, _
    regionsize As Long, _
    ByVal AllocationType As Long, _
    ByVal Protect As Long) As Long

Declare Function NtFreeVirtualMemory Lib "ntdll.dll" ( _
    ByVal ProcessHandle As Long, _
    BaseAddress As Long, _
    regionsize As Long, _
    ByVal FreeType As Long) As Long
      
' // ********************
' //  WRAPPER FUNCTIONS
' // ********************

Public Function CreateWorkSpace(Format As CompressionFormats, Engine As CompressionEngines) As Long
' // IN: Format and Engine
' // OUT: Pointer to Buffer
' // About CompressionFormatAnd Engine: This is an integer, which means 4 bytes, or in hex: 0xYYYY where Y can be 0 to F. The first two bytes are the format
' // and the last two bytes are the engine. Therefore, LZNT1 which is 0x0002 and high compression which is 0x0100 would become 0x0102, basically we OR the two
    
    ' // Variable Declarations
    Dim CompressionType         As Integer                                                                      ' // Holds our two ORed values
    Dim WorkSpaceBuffer         As String                                                                       ' // Holds the WorkSpace Buffer
    Dim WorkSpaceSize           As Long                                                                         ' // Return value from API call
    Dim FragmentWorkSpaceSize   As Long                                                                         ' // We don't care about this one
    
    ' // Create the Workspace
    CompressionType = Format Or Engine                                                                          ' // Calculate the Format+Engine Value
    Call RtlGetCompressionWorkSpaceSize(CompressionType, WorkSpaceSize, FragmentWorkSpaceSize)                  ' // Call the API to get our Workspace Size
    NtAllocateVirtualMemory -1, CreateWorkSpace, 0, WorkSpaceSize, MEM_COMMIT, PAGE_EXECUTE_READWRITE           ' // Return a pointer to the WorkSpace Buffer
    
End Function
Public Function Compress(Format As CompressionFormats, Engine As CompressionEngines, UnCompressedBuffer As Long, ByVal UnCompressedBufferSize As Long, FinalSize As Long, NewFile As String, ByVal Workspace As Long) As Long
' // IN: Compression Format and Engine, Data to compress as a Byte Array, Pointer to Workspace Buffer,
' // and ChunkSize (&H1000 is recommended, but can be 0)
' // OUT: Compressed Data, Final Size of the Compressed Data

    ' // Variable Declarations
    Dim CompressionType         As Integer                                                                      ' // Holds our two ORed values
    Dim CompressedBuffer        As Long                                                                         ' // Holds the buffer to receive
    Dim CompressedBufferSize    As Long                                                                         ' // Holds the size of the buffer to receive
    Dim CompressedBufferHandle  As Long
    Dim CompressedBufferHandle2 As Long
    
    ' // Open Destination File
    CompressedBufferSize = UnCompressedBufferSize * 1.13 + 4                                                    ' // Size of compressed buffer can never be bigger
    CompressedBuffer = OpenFile(NewFile, CompressedBufferSize, CompressedBufferHandle, CompressedBufferHandle2) ' // Allocate it
    CompressionType = Format Or Engine                                                                          ' // Calculate the Format+Engine Value
      
    ' // Do the Call
    Compress = RtlCompressBuffer(CompressionType, UnCompressedBuffer, UnCompressedBufferSize, CompressedBuffer, CompressedBufferSize, 0&, FinalSize, Workspace)
    
    ' // Write the new file
    Call NtUnmapViewOfSection(-1, CompressedBuffer)
    Call NtClose(CompressedBufferHandle2)
    Call SetFilePointer(CompressedBufferHandle, FinalSize, 0, 0)
    Call SetEndOfFile(CompressedBufferHandle)
    Call NtClose(CompressedBufferHandle)
    
    ' // Empty the Workspace Buffer
    Debug.Print NtFreeVirtualMemory(-1, Workspace, 0, MEM_DECOMMIT)
End Function
Public Function DeCompress(Format As CompressionFormats, CompressedBuffer As Long, CompressedBufferSize As Long, FinalSize As Long, NewFile As String) As Long
' // IN: Compression Format, Data to decompress as a Byte Array
' // OUT: Decompressed Data, Final Size of the Compressed Buffer

    ' // Variable Declarations
    Dim UnCompressedBuffer      As Long                                                                         ' // Holds the buffer to send
    Dim UnCompressedBufferSize  As Long                                                                         ' // Holds the size of the buffer to send
    Dim OriginalBufferHandle    As Long
    Dim OriginalBufferHandle2   As Long
       
    '// Calculations needed for the API Call
    UnCompressedBufferSize = CompressedBufferSize * 12.5                                                        ' // Max compression possible (92%)
    UnCompressedBuffer = OpenFile(NewFile, UnCompressedBufferSize, OriginalBufferHandle, OriginalBufferHandle2) ' // Pointer to the compressed buffer
      
    ' // Do the Call
    DeCompress = RtlDecompressBuffer(Format, UnCompressedBuffer, UnCompressedBufferSize, CompressedBuffer, CompressedBufferSize, FinalSize)
    
    ' // Write the new file
    Call NtUnmapViewOfSection(-1, UnCompressedBuffer)
    Call NtClose(OriginalBufferHandle2)
    Call SetFilePointer(OriginalBufferHandle, FinalSize, 0, 0)
    Call SetEndOfFile(OriginalBufferHandle)
    Call NtClose(OriginalBufferHandle)

End Function



