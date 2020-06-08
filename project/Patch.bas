Attribute VB_Name = "Patch"

Option Explicit

Private Const IPS_HEADER As String = "PATCH"
Private Const IPS_FOOTER As String = "EOF"

Private Const POS_LENGTH As Integer = 3
Private Const SIZE_LENGTH As Integer = 2
Private Const RLE_LENGTH As Integer = 2

Public Type IPSPatch
    Contents() As Byte
    Path As String
    Valid As Boolean
    
    Records As Integer
    HighestPos As Long
End Type

Public Function OpenPatch(Path As String) As IPSPatch
    Dim Patch As IPSPatch
    Dim FileNum As Integer: FileNum = FreeFile()
    
    With Patch
        Open Path For Binary As #FileNum
            ReDim .Contents(LOF(FileNum) - 1)
            Get #FileNum, , .Contents
        Close #FileNum
        
        Dim Length As Long: Length = UBound(.Contents) + 1
        
        If Length >= Len(IPS_HEADER) + Len(IPS_FOOTER) Then
            .Valid = True
            
            Dim Header As String: Header = Chr(.Contents(0)) _
                & Chr(.Contents(1)) _
                & Chr(.Contents(2)) _
                & Chr(.Contents(3)) _
                & Chr(.Contents(4))
            Dim Footer As String: Footer = Chr(.Contents(Length - 3)) _
                & Chr(.Contents(Length - 2)) _
                & Chr(.Contents(Length - 1))
            
            .Valid = Header = IPS_HEADER And Footer = IPS_FOOTER
        End If
        
        .Path = Path
    End With
    
    OpenPatch = ReadPatch(Patch)
End Function

Private Function ReadPatch(Patch As IPSPatch) As IPSPatch
    With Patch
        Dim Offset As Long: Offset = Len(IPS_HEADER)
        Dim Length As Long: Length = UBound(.Contents) - Len(IPS_FOOTER) + 1
        
        Dim Count As Integer
        Dim HighestPos As Long
        Dim LocalMax As Long
        
        Dim Pos As Long
        Dim Size As Long
        Dim RLESize As Long
        Dim RLEPos As Long
        
        While Offset < Length
            Count = Count + 1
            Pos = CLng(.Contents(Offset)) * &H10000 + _
                CLng(.Contents(Offset + 1)) * &H100 + _
                .Contents(Offset + 2)
            Size = CLng(.Contents(Offset + POS_LENGTH)) * &H100 + _
                .Contents(Offset + POS_LENGTH + 1)
            Offset = Offset + POS_LENGTH + SIZE_LENGTH
            
            If Size = 0 Then
                ' Run-length encoding
                RLESize = CLng(.Contents(Offset)) * &H100 + _
                    .Contents(Offset + 1)
                RLEPos = Pos + RLESize
                LocalMax = RLEPos
                Offset = Offset + RLE_LENGTH + 1
            Else
                LocalMax = Pos
                Offset = Offset + Size
            End If
            
            HighestPos = IIf(HighestPos > LocalMax, HighestPos, LocalMax)
        Wend
        
        .Records = Count
        .HighestPos = HighestPos
    End With
    
    ReadPatch = Patch
End Function

Public Function ApplyPatch(Contents As Variant, Patch As IPSPatch) As Variant
    With Patch
        Dim ContentsSize As Long: ContentsSize = UBound(Contents) + 1
        ReDim Preserve Contents(IIf(.HighestPos > ContentsSize, .HighestPos - 1, ContentsSize))
        
        Dim Offset As Long: Offset = Len(IPS_HEADER)
        Dim Length As Long: Length = UBound(.Contents) - Len(IPS_FOOTER) + 1
        
        Dim Pos As Long
        Dim Size As Long
        Dim RLESize As Long
        Dim RLEByte As Byte
        
        While Offset < Length
            Pos = CLng(.Contents(Offset)) * &H10000 + _
                CLng(.Contents(Offset + 1)) * &H100 + _
                .Contents(Offset + 2)
            Size = CLng(.Contents(Offset + POS_LENGTH)) * &H100 + _
                .Contents(Offset + POS_LENGTH + 1)
            Offset = Offset + POS_LENGTH + SIZE_LENGTH
            
            Dim i As Long
            
            If Size = 0 Then
                ' Run-length encoding
                RLESize = CLng(.Contents(Offset)) * &H100 + _
                    .Contents(Offset + 1)
                RLEByte = .Contents(Offset + RLE_LENGTH)
                
                For i = 0 To RLESize - 1
                    Contents(Pos + i) = RLEByte
                Next
                
                Offset = Offset + RLE_LENGTH + 1
            Else
                For i = 0 To Size - 1
                    Contents(Pos + i) = .Contents(Offset + i)
                Next
                
                Offset = Offset + Size
            End If
        Wend
    End With
    
    ApplyPatch = Contents
End Function
