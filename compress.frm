VERSION 5.00
Begin VB.Form frmCompress 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "NT Native Compression [Beta Version] - Alex Ionescu"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5415
   Icon            =   "compress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "NT Native Encryption"
      Height          =   1695
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   5415
      Begin VB.TextBox txtOriginalFile 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   16
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Frame Frame2 
         Caption         =   "Level"
         Height          =   450
         Left            =   3600
         TabIndex        =   15
         Top             =   115
         Width           =   1695
         Begin VB.OptionButton optHigh 
            Caption         =   "&High"
            Height          =   255
            Left            =   960
            TabIndex        =   8
            Top             =   160
            Width           =   680
         End
         Begin VB.OptionButton optNormal 
            Caption         =   "&Normal"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   160
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.OptionButton optNS15 
         Caption         =   "NS1&5"
         Height          =   255
         Left            =   2280
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optNS3 
         Caption         =   "NS&3"
         Height          =   255
         Left            =   1560
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optLZNT1 
         Caption         =   "&LZNT1"
         Height          =   255
         Left            =   720
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtInputFile 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   600
         Width           =   2655
      End
      Begin VB.CommandButton cmdDecompress 
         Caption         =   "&Decompress!"
         Height          =   255
         Left            =   4200
         TabIndex        =   3
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdCompress 
         Caption         =   "&Compress!"
         Height          =   255
         Left            =   4200
         TabIndex        =   1
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtOutputFile 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label lblStatus 
         Height          =   255
         Left            =   4200
         TabIndex        =   14
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblOriginal 
         Caption         =   "File to Open:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblCompressed 
         Caption         =   "Compessed File:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblDecompressed 
         Caption         =   "Original File:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblFormat 
         Caption         =   "Format:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmCompress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Workspace As Long

Private Sub cmdCompress_Click()
    Dim CompressionFormat As CompressionFormats
    Dim CompressionLevel As CompressionEngines
    Dim pFile As Long, FileSize As Long
    Dim FileHandle As Long, MemoryFileHandle As Long
    Dim FinalSize As Long

    If optLZNT1.Value = True Then
        CompressionFormat = COMPRESSION_FORMAT_LZNT1
    ElseIf optNS3.Value = True Then
        CompressionFormat = COMPRESSION_FORMAT_NS3
    Else
        CompressionFormat = COMPRESSION_FORMAT_NS15
    End If
    
    If optHigh.Value = False Then CompressionLevel = COMPRESSION_ENGINE_STANDARD Else CompressionLevel = COMPRESSION_ENGINE_MAXIMUM
           
    If Workspace = 0 Then Workspace = CreateWorkSpace(CompressionFormat, CompressionLevel)
    pFile = OpenFile(txtInputFile, FileSize, FileHandle, MemoryFileHandle)
    
    Call Compress(CompressionFormat, CompressionLevel, pFile, FileSize, FinalSize, txtOutputFile, Workspace)
    lblStatus = 100 - Int((FinalSize / FileSize) * 100) & "%"
    Call NtUnmapViewOfSection(-1, pFile)
    Call NtClose(FileHandle)
    Call NtClose(MemoryFileHandle)
    
End Sub

Private Sub cmdDecompress_Click()
    Dim CompressionFormat As CompressionFormats
    Dim pFile As Long, FileSize As Long
    Dim FileHandle As Long, MemoryFileHandle As Long
    Dim FinalSize As Long
    
    If optLZNT1.Value = True Then
        CompressionFormat = COMPRESSION_FORMAT_LZNT1
    ElseIf optNS3.Value = True Then
        CompressionFormat = COMPRESSION_FORMAT_NS3
    Else
        CompressionFormat = COMPRESSION_FORMAT_NS15
    End If
    
    pFile = OpenFile(txtOutputFile, FileSize, FileHandle, MemoryFileHandle)
    
    Call DeCompress(CompressionFormat, pFile, FileSize, FinalSize, txtOriginalFile)
    lblStatus = 100 - Int((FinalSize / FileSize) * 100) & "%"
    Call NtUnmapViewOfSection(-1, pFile)
    Call NtClose(FileHandle)
    Call NtClose(MemoryFileHandle)
End Sub

