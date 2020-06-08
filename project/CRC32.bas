Attribute VB_Name = "CRC32"

Option Explicit

Private CRC32Table() As Long
Private TableGenerated As Boolean

Public Function CalculateCRC32(Contents As Variant) As Long
    If Not TableGenerated Then
        CRC32Table = GenerateCRC32Table
        TableGenerated = True
    End If
    
    Dim Checksum As Long: Checksum = &HFFFFFFFF
    Dim Index As Integer: Index = 0
    Dim i As Long
    
    For i = 0 To UBound(Contents)
          Index = (Checksum And &HFF) Xor Contents(i)
          Checksum = ((Checksum And &HFFFFFF00) \ &H100) And &HFFFFFF
          Checksum = Checksum Xor CRC32Table(Index)
    Next
    
    CalculateCRC32 = Not Checksum
End Function

Private Function GenerateCRC32Table() As Variant
    Dim Polynomial As Long: Polynomial = &HEDB88320
    Dim CRC As Long: CRC = 0
    
    Dim CRC32Table() As Long
    ReDim CRC32Table(&HFF)
    
    Dim i As Integer
    
    For i = 0 To UBound(CRC32Table)
        CRC = i
        
        Dim j As Integer
        
        For j = 8 To 1 Step -1
            If (CRC And 1) Then
                CRC = ((CRC And &HFFFFFFFE) \ 2) And &H7FFFFFFF
                CRC = CRC Xor Polynomial
            Else
                CRC = ((CRC And &HFFFFFFFE) \ 2) And &H7FFFFFFF
            End If
        Next
        
        CRC32Table(i) = CRC
    Next
    
    GenerateCRC32Table = CRC32Table
End Function
