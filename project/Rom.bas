Attribute VB_Name = "Rom"

Option Explicit

' File formats
Public Const BIN_FILE As Boolean = False
Public Const SMD_FILE As Boolean = True

' Format info
Public Const HEADER_SIZE As Integer = &H200
Public Const BLOCK_SIZE As Integer = &H4000

' Section offsets
Private Enum Offsets
    CONSOLE_OFFSET = &H100
    COMPANY_OFFSET = &H110
    COPYRIGHT_OFFSET = &H118
    LOCAL_NAME_OFFSET = &H120
    INTL_NAME_OFFSET = &H150
    PRODUCT_TYPE_OFFSET = &H180
    PRODUCT_CODE_OFFSET = &H183
    CHECKSUM_OFFSET = &H18E
    IO_DEVICES_OFFSET = &H190
    ROM_START_OFFSET = &H1A0
    ROM_STOP_OFFSET = &H1A4
    RAM_START_OFFSET = &H1A8
    RAM_STOP_OFFSET = &H1AC
    SRAM_DATA_OFFSET = &H1B0
    SRAM_START_OFFSET = &H1B4
    SRAM_STOP_OFFSET = &H1B8
    MODEM_DATA_OFFSET = &H1BC
    COMMENT_OFFSET = &H1C8
    REGIONS_OFFSET = &H1F0
End Enum

' Section lengths
Public Enum Lengths
    CONSOLE_LENGTH = 16
    COMPANY_LENGTH = 8
    COPYRIGHT_LENGTH = 8
    NAME_LENGTH = 48
    PRODUCT_TYPE_LENGTH = 2
    PRODUCT_CODE_LENGTH = 11
    CHECKSUM_LENGTH = 2
    IO_DEVICES_LENGTH = 16
    ADDRESS_LENGTH = 4
    MODEM_DATA_LENGTH = 12
    COMMENT_LENGTH = 40
    REGIONS_LENGTH = 3
End Enum

Public Type ROMFile
    Contents() As Byte
    Path As String
    FileSize As Long
    CRC32 As Long
    Format As Boolean
    Valid As Boolean
    
    Console As String
    Company As String
    Copyright As String
    LocalName As String
    IntlName As String
    ProductType As String
    ProductCode As String
    ModemData As String
    Comment As String
    
    IODevices As String
    Regions As String
    
    ROMStart As Long
    ROMStop As Long
    RAMStart As Long
    RAMStop As Long
    SRAMData As String
    SRAMStart As Long
    SRAMStop As Long
    
    Checksum As Long
    CalculatedChecksum As Long
End Type

Public Function OpenROM(Path As String) As ROMFile
    Dim File As ROMFile
    
    Dim FileNum As Integer: FileNum = FreeFile()
    
    With File
        Open Path For Binary As #FileNum
            ReDim .Contents(LOF(FileNum) - 1)
            Get #FileNum, , .Contents
        Close #FileNum
        
        .Path = Path
    End With
    
    OpenROM = ReadROM(File)
End Function

Private Function ReadROM(File As ROMFile) As ROMFile
    With File
        ' Calculates CRC32 before doing any conversions
        .CRC32 = CalculateCRC32(.Contents)
        
        ' Determines binary format
        If .Contents(&H1) = &H3 _
        And .Contents(&H8) = &HAA _
        And .Contents(&H9) = &HBB _
        And .Contents(&HA) = &H6 Then
            .Format = SMD_FILE
            
            ' Converts to bin format if interleaved
            .Contents = DeInterleave(.Contents)
        Else
            .Format = BIN_FILE
        End If
        
        .FileSize = UBound(.Contents) + 1
        
        ' Subtracts smd header
        If .Format = SMD_FILE Then
            .FileSize = .FileSize - HEADER_SIZE
        End If
        
        ' Breaks file into sections
        .Console = ReadString(.Contents, CONSOLE_OFFSET, CONSOLE_LENGTH)
        .Company = ReadString(.Contents, COMPANY_OFFSET, COMPANY_LENGTH)
        .Copyright = ReadString(.Contents, COPYRIGHT_OFFSET, COPYRIGHT_LENGTH)
        .LocalName = ReadString(.Contents, LOCAL_NAME_OFFSET, NAME_LENGTH)
        .IntlName = ReadString(.Contents, INTL_NAME_OFFSET, NAME_LENGTH)
        .ProductType = ReadString(.Contents, PRODUCT_TYPE_OFFSET, PRODUCT_TYPE_LENGTH)
        .ProductCode = ReadString(.Contents, PRODUCT_CODE_OFFSET, PRODUCT_CODE_LENGTH)
        .ModemData = ReadString(.Contents, MODEM_DATA_OFFSET, MODEM_DATA_LENGTH)
        .Comment = ReadString(.Contents, COMMENT_OFFSET, COMMENT_LENGTH)
        
        .IODevices = ReadString(.Contents, IO_DEVICES_OFFSET, IO_DEVICES_LENGTH)
        .Regions = ReadString(.Contents, REGIONS_OFFSET, REGIONS_LENGTH)
        
        .ROMStart = ReadNumber(.Contents, ROM_START_OFFSET, ADDRESS_LENGTH)
        .ROMStop = ReadNumber(.Contents, ROM_STOP_OFFSET, ADDRESS_LENGTH)
        .RAMStart = ReadNumber(.Contents, RAM_START_OFFSET, ADDRESS_LENGTH)
        .RAMStop = ReadNumber(.Contents, RAM_STOP_OFFSET, ADDRESS_LENGTH)
        .SRAMData = ReadString(.Contents, SRAM_DATA_OFFSET, ADDRESS_LENGTH)
        .SRAMStart = ReadNumber(.Contents, SRAM_START_OFFSET, ADDRESS_LENGTH)
        .SRAMStop = ReadNumber(.Contents, SRAM_STOP_OFFSET, ADDRESS_LENGTH)
        
        .Checksum = ReadNumber(.Contents, CHECKSUM_OFFSET, CHECKSUM_LENGTH)
        .CalculatedChecksum = CalculateChecksum(.Contents)
        
        ' Checks for string at $100 (same check done by TMSS)
        .Valid = InStr(.Console, "SEGA") = 1 Or InStr(.Console, " SEGA") = 1
    End With
    
    ReadROM = File
End Function

Public Sub SaveROM(File As ROMFile)
    With File
        ' Validates fields and writes to byte array
        .Contents = WriteString(.Contents, .Console, CONSOLE_OFFSET, CONSOLE_LENGTH)
        .Contents = WriteString(.Contents, .Company, COMPANY_OFFSET, COMPANY_LENGTH)
        .Contents = WriteString(.Contents, .Copyright, COPYRIGHT_OFFSET, COPYRIGHT_LENGTH)
        .Contents = WriteString(.Contents, .LocalName, LOCAL_NAME_OFFSET, NAME_LENGTH)
        .Contents = WriteString(.Contents, .IntlName, INTL_NAME_OFFSET, NAME_LENGTH)
        .Contents = WriteString(.Contents, .ProductType, PRODUCT_TYPE_OFFSET, PRODUCT_TYPE_LENGTH)
        .Contents = WriteString(.Contents, .ProductCode, PRODUCT_CODE_OFFSET, PRODUCT_CODE_LENGTH)
        .Contents = WriteString(.Contents, .ModemData, MODEM_DATA_OFFSET, MODEM_DATA_LENGTH)
        .Contents = WriteString(.Contents, .Comment, COMMENT_OFFSET, COMMENT_LENGTH)
        
        .Contents = WriteString(.Contents, .IODevices, IO_DEVICES_OFFSET, IO_DEVICES_LENGTH)
        .Contents = WriteString(.Contents, .Regions, REGIONS_OFFSET, REGIONS_LENGTH)
        
        .Contents = WriteNumber(.Contents, .Checksum, CHECKSUM_OFFSET, CHECKSUM_LENGTH)
        
        ' Converts to smd format if flagged for conversion
        If .Format = SMD_FILE Then
            .Contents = Interleave(.Contents)
        End If
        
        ' Recalculates CRC32
        .CRC32 = CalculateCRC32(.Contents)
    End With
    
    ' Deletes file if it already exists
    On Error Resume Next
    Kill File.Path
    
    Dim FileNum As Integer: FileNum = FreeFile()
    
    Open File.Path For Binary As #FileNum
        Put #FileNum, , File.Contents
    Close #FileNum
End Sub

' Converts from bin to smd
Private Function Interleave(Contents As Variant) As Variant
    Dim Length As Long: Length = UBound(Contents) + 1
    Dim NumBlocks As Integer: NumBlocks = Length \ BLOCK_SIZE
    
    Dim Converted() As Byte
    ReDim Converted(Length + HEADER_SIZE - 1)
    
    ' Creates header, all other values are 0
    ' First byte is number of blocks in file
    Converted(0) = NumBlocks
    ' Fixed values
    Converted(1) = &H3
    Converted(8) = &HAA
    Converted(9) = &HBB
    Converted(10) = &H6
    
    Dim i As Integer
    
    For i = 0 To NumBlocks - 1
        Dim Start As Long: Start = CLng(i) * BLOCK_SIZE
        Dim Offset As Integer: Offset = 0
        
        Dim j As Integer
        
        For j = 0 To BLOCK_SIZE \ 2 - 1
            ' Even bytes
            Converted(HEADER_SIZE + Start + j + BLOCK_SIZE \ 2) = Contents(Start + Offset)
            ' Odd bytes
            Converted(HEADER_SIZE + Start + j) = Contents(Start + Offset + 1)
            
            Offset = Offset + 2
        Next
    Next
    
    Interleave = Converted
End Function

' Converts from smd to bin
Private Function DeInterleave(Contents As Variant) As Variant
    Dim Length As Long: Length = UBound(Contents) - HEADER_SIZE + 1
    Dim NumBlocks As Integer: NumBlocks = Length \ BLOCK_SIZE
    
    Dim Converted() As Byte
    ReDim Converted(Length - 1)
    
    Dim i As Integer
    
    For i = 0 To NumBlocks - 1
        Dim Start As Long: Start = CLng(i) * BLOCK_SIZE
        Dim Offset As Integer: Offset = 0
        
        Dim j As Integer
        
        For j = 0 To BLOCK_SIZE \ 2 - 1
            ' Even bytes
            Converted(Start + Offset) = Contents(HEADER_SIZE + Start + j + BLOCK_SIZE \ 2)
            ' Odd bytes
            Converted(Start + Offset + 1) = Contents(HEADER_SIZE + Start + j)
            
            Offset = Offset + 2
        Next
    Next
    
    DeInterleave = Converted
End Function

Private Function CalculateChecksum(Contents As Variant) As Long
    Dim Checksum As Long: Checksum = 0
    Dim i As Long
    
    ' Adds up entire ROM after header one word at a time
    For i = HEADER_SIZE To UBound(Contents) Step 2
        ' Shifts first byte left
        Checksum = Checksum + Contents(i) * &H100 + Contents(i + 1)
        
        ' Prevents overflow
        Checksum = Checksum Mod &H10000
    Next
    
    CalculateChecksum = Checksum
End Function

Private Function ReadNumber(Contents As Variant, Start As Integer, Length As Integer) As Long
    Dim Number As Long: Number = 0
    Dim Offset As Integer: Offset = Length - 1
    Dim i As Long
    
    For i = Start To Start + Length - 1
        Number = Number + Contents(i) * 2 ^ (8 * Offset)
        Offset = Offset - 1
    Next
    
    ReadNumber = Number
End Function

Private Function ReadString(Contents As Variant, Start As Integer, Length As Integer) As String
    Dim Slice As String: Slice = ""
    Dim i As Integer
    
    For i = Start To Start + Length - 1
        Slice = Slice & Chr(Contents(i))
    Next
    
    ReadString = RTrim(Slice)
End Function

Private Function ReadArray(Contents As Variant, Start As Integer, Length As Integer) As Variant
    Dim Buffer() As Byte
    ReDim Buffer(Length - 1)
    
    Dim Offset As Integer: Offset = 0
    Dim i As Integer
    
    For i = Start To Start + Length - 1
        Buffer(Offset) = Contents(i)
        Offset = Offset + 1
    Next
    
    ReadArray = Buffer
End Function

Private Function WriteNumber(Contents As Variant, Number As Long, Start As Integer, Length As Integer)
    Dim Offset As Integer: Offset = Length - 1
    Dim i As Long
    
    For i = Start To Start + Length - 1
        Dim Shift As Long: Shift = 2 ^ (8 * Offset)
        Contents(i) = (Number And &HFF * Shift) \ Shift
        
        Offset = Offset - 1
    Next
    
    WriteNumber = Contents
End Function

Private Function WriteString(Contents As Variant, Text As String, Start As Integer, Length As Integer)
    ' Pads string to length with zeroes
    Text = Text & String(Length - Len(Text), " ")
    ' Truncates string if longer than length
    Text = Left(Text, Length)
    ' Forces string to uppercase
    Text = UCase(Text)
    
    Dim Offset As Integer: Offset = 1
    Dim i As Integer
    
    For i = Start To Start + Length - 1
        Contents(i) = Asc(Mid(Text, Offset, 1))
        Offset = Offset + 1
    Next
    
    WriteString = Contents
End Function

Private Function WriteArray(Contents As Variant, Bytes As Variant, Start As Integer, Length As Integer)
    Dim Offset As Integer: Offset = 0
    Dim i As Integer
    
    For i = Start To Start + Length - 1
        Contents(i) = Bytes(Offset)
        Offset = Offset + 1
    Next
    
    WriteArray = Contents
End Function

Public Function PatchROM(File As ROMFile, Patch As IPSPatch) As ROMFile
    File.Contents = ApplyPatch(File.Contents, Patch)
    PatchROM = ReadROM(File)
End Function
