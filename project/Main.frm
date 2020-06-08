VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7740
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9915
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Main"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6179.64
   ScaleMode       =   0  'User
   ScaleWidth      =   7107.526
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCRC32 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Left            =   720
      TabIndex        =   64
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Frame fraAddresses 
      Caption         =   "Addresses"
      Height          =   1335
      Left            =   120
      TabIndex        =   49
      Top             =   5760
      Width           =   9615
      Begin VB.TextBox txtAddresses 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   55
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtAddresses 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   54
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtAddresses 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   4560
         TabIndex        =   53
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtAddresses 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   4560
         TabIndex        =   52
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtAddresses 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   375
         Index           =   4
         Left            =   8040
         TabIndex        =   51
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtAddresses 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   375
         Index           =   5
         Left            =   8040
         TabIndex        =   50
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblAddresses 
         Caption         =   "ROM Start:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   61
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblAddresses 
         Caption         =   "ROM Stop:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   60
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblAddresses 
         Caption         =   "RAM Start:"
         Height          =   255
         Index           =   2
         Left            =   3480
         TabIndex        =   59
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblAddresses 
         Caption         =   "RAM Stop:"
         Height          =   255
         Index           =   3
         Left            =   3480
         TabIndex        =   58
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblAddresses 
         Caption         =   "Backup RAM Start:"
         Height          =   255
         Index           =   4
         Left            =   6480
         TabIndex        =   57
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblAddresses 
         Caption         =   "Backup RAM Stop:"
         Height          =   255
         Index           =   5
         Left            =   6480
         TabIndex        =   56
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.Frame fraGeneral 
      Caption         =   "General"
      Height          =   2295
      Left            =   120
      TabIndex        =   41
      Top             =   120
      Width           =   9615
      Begin VB.TextBox txtGeneral 
         Height          =   375
         Index           =   5
         Left            =   8040
         MaxLength       =   8
         TabIndex        =   5
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtGeneral 
         Height          =   375
         Index           =   0
         Left            =   1680
         MaxLength       =   16
         TabIndex        =   0
         Top             =   240
         Width           =   3495
      End
      Begin VB.TextBox txtGeneral 
         Height          =   375
         Index           =   1
         Left            =   1680
         MaxLength       =   48
         TabIndex        =   1
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox txtGeneral 
         Height          =   375
         Index           =   2
         Left            =   1680
         MaxLength       =   48
         TabIndex        =   2
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox txtGeneral 
         Height          =   375
         Index           =   4
         Left            =   6480
         MaxLength       =   8
         TabIndex        =   4
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtGeneral 
         Height          =   375
         Index           =   6
         Left            =   7920
         MaxLength       =   2
         TabIndex        =   7
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtGeneral 
         Height          =   375
         Index           =   3
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   3
         Top             =   1680
         Width           =   3495
      End
      Begin VB.TextBox txtGeneral 
         Height          =   375
         Index           =   8
         Left            =   6480
         MaxLength       =   12
         TabIndex        =   9
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox txtGeneral 
         Height          =   375
         Index           =   7
         Left            =   8400
         MaxLength       =   11
         TabIndex        =   8
         Top             =   1200
         Width           =   1095
      End
      Begin VB.ComboBox cboProductType 
         Height          =   315
         ItemData        =   "Main.frx":1FF84
         Left            =   6480
         List            =   "Main.frx":1FF86
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblCompany 
         Height          =   255
         Left            =   6480
         TabIndex        =   63
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label lblGeneral 
         Caption         =   "Copyright:"
         Height          =   255
         Index           =   5
         Left            =   5280
         TabIndex        =   62
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblGeneral 
         Caption         =   "Console name:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   48
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblGeneral 
         Caption         =   "Domestic name:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   47
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblGeneral 
         Caption         =   "International name:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   46
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblGeneral 
         Caption         =   "Company:"
         Height          =   255
         Index           =   4
         Left            =   5280
         TabIndex        =   45
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblGeneral 
         Caption         =   "Product code:"
         Height          =   255
         Index           =   6
         Left            =   5280
         TabIndex        =   44
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblGeneral 
         Caption         =   "Comment:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   43
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label lblGeneral 
         Caption         =   "Modem data:"
         Height          =   255
         Index           =   7
         Left            =   5280
         TabIndex        =   42
         Top             =   1800
         Width           =   975
      End
   End
   Begin VB.TextBox txtSize 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   37
      Top             =   7200
      Width           =   855
   End
   Begin VB.TextBox txtChecksum 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   35
      Top             =   7200
      Width           =   615
   End
   Begin VB.CommandButton cmdFixChecksum 
      Caption         =   "&Correct"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   31
      Top             =   7200
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   9240
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraIODevices 
      Caption         =   "I/O Devices"
      Height          =   1815
      Left            =   120
      TabIndex        =   39
      Tag             =   "6"
      Top             =   2520
      Width           =   9615
      Begin VB.CheckBox chkIODevices 
         Caption         =   "Serial (RS-232)"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   22
         Tag             =   "R"
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CheckBox chkIODevices 
         Caption         =   "Analogue Joystick"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Tag             =   "A"
         Top             =   720
         Width           =   2175
      End
      Begin VB.CheckBox chkIODevices 
         Caption         =   "Activator"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   18
         Tag             =   "L"
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CheckBox chkIODevices 
         Caption         =   "Mega Mouse"
         Height          =   255
         Index           =   10
         Left            =   4920
         TabIndex        =   20
         Tag             =   "M"
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CheckBox chkIODevices 
         Caption         =   "Floppy Disk Drive"
         Height          =   255
         Index           =   13
         Left            =   2520
         TabIndex        =   23
         Tag             =   "F"
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CheckBox chkIODevices 
         Caption         =   "CD-ROM Drive"
         Height          =   255
         Index           =   14
         Left            =   4920
         TabIndex        =   24
         Tag             =   "C"
         Top             =   1440
         Width           =   2295
      End
      Begin VB.CheckBox chkIODevices 
         Caption         =   "Paddle Controller"
         Height          =   255
         Index           =   3
         Left            =   7440
         TabIndex        =   13
         Tag             =   "V"
         Top             =   360
         Width           =   2055
      End
      Begin VB.CheckBox chkIODevices 
         Caption         =   "Drawing Tablet"
         Height          =   255
         Index           =   6
         Left            =   4920
         TabIndex        =   16
         Tag             =   "T"
         Top             =   720
         Width           =   2295
      End
      Begin VB.CheckBox chkIODevices 
         Caption         =   "Light Gun"
         Height          =   255
         Index           =   7
         Left            =   7440
         TabIndex        =   17
         Tag             =   "G"
         Top             =   720
         Width           =   2055
      End
      Begin VB.CheckBox chkIODevices 
         Caption         =   "Master System Control Pad"
         Height          =   255
         Index           =   2
         Left            =   4920
         TabIndex        =   12
         Tag             =   "0"
         Top             =   360
         Width           =   2295
      End
      Begin VB.CheckBox chkIODevices 
         Caption         =   "Team Player"
         Height          =   255
         Index           =   5
         Left            =   2520
         TabIndex        =   15
         Tag             =   "4"
         Top             =   720
         Width           =   2175
      End
      Begin VB.CheckBox chkIODevices 
         Caption         =   "Trackball"
         Height          =   255
         Index           =   11
         Left            =   7440
         TabIndex        =   21
         Tag             =   "B"
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CheckBox chkIODevices 
         Caption         =   "Printer"
         Height          =   255
         Index           =   15
         Left            =   7440
         TabIndex        =   25
         Tag             =   "P"
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CheckBox chkIODevices 
         Caption         =   "Keyboard"
         Height          =   255
         Index           =   9
         Left            =   2520
         TabIndex        =   19
         Tag             =   "K"
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CheckBox chkIODevices 
         Caption         =   "6-button Control Pad"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   11
         Tag             =   "6"
         Top             =   360
         Width           =   2175
      End
      Begin VB.CheckBox chkIODevices 
         Caption         =   "3-Button Control Pad"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Tag             =   "J"
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame fraRegions 
      Caption         =   "Regions"
      Height          =   1215
      Left            =   120
      TabIndex        =   40
      Top             =   4440
      Width           =   9615
      Begin VB.OptionButton optRegionFormat 
         Caption         =   "New format"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Tag             =   "1"
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton optRegionFormat 
         Caption         =   "Old format"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Tag             =   "0"
         Top             =   360
         Width           =   1815
      End
      Begin VB.CheckBox chkRegions 
         Caption         =   "United States (U)"
         Height          =   255
         Index           =   1
         Left            =   4920
         TabIndex        =   29
         Tag             =   "U"
         Top             =   480
         Width           =   1815
      End
      Begin VB.CheckBox chkRegions 
         Caption         =   "Japan (J)"
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   28
         Tag             =   "J"
         Top             =   480
         Width           =   1815
      End
      Begin VB.CheckBox chkRegions 
         Caption         =   "Europe (E)"
         Height          =   255
         Index           =   2
         Left            =   7440
         TabIndex        =   30
         Tag             =   "E"
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Label lblCRC32 
      Caption         =   "CRC32:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   65
      Top             =   7320
      Width           =   615
   End
   Begin VB.Label lblSizeDescription 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   8280
      TabIndex        =   38
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Label lblSize 
      Caption         =   "Size:"
      Height          =   255
      Index           =   1
      Left            =   6840
      TabIndex        =   36
      Top             =   7320
      Width           =   375
   End
   Begin VB.Label lblChecksumStatus 
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   4800
      TabIndex        =   34
      Top             =   7200
      Width           =   375
   End
   Begin VB.Label lblChecksumStatus 
      Height          =   255
      Index           =   1
      Left            =   5280
      TabIndex        =   33
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Label lblChecksum 
      Caption         =   "Checksum:"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   32
      Top             =   7320
      Width           =   855
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
         Enabled         =   0   'False
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuFileSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuShowAsEnglish 
         Caption         =   "Display Field as &English"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuShowAsJapanese 
         Caption         =   "Display Field as &Japanese"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuConvertToBin 
         Caption         =   "Convert to &Bin"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuConvertToSMD 
         Caption         =   "Convert to &SMD"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuToolsSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFixChecksum 
         Caption         =   "Fix &Checksum"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuToolsSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuApplyPatch 
         Caption         =   "Apply IPS &Patch..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuWebSite 
         Caption         =   "Visit &Web Site"
      End
      Begin VB.Menu mnuHelpSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Path As String
Private File As ROMFile

Private FileOpened As Boolean
Private FileModified As Boolean
Private FormatChanged As Boolean
Private SelectedField As Integer

' Text fields
Private Enum Fields
    CONSOLE_INDEX
    LOCAL_NAME_INDEX
    INTL_NAME_INDEX
    COMMENT_INDEX
    COMPANY_INDEX
    COPYRIGHT_INDEX
    PRODUCT_TYPE_INDEX
    PRODUCT_CODE_INDEX
    MODEM_DATA_INDEX
End Enum

' Fonts
Private Const DEFAULT_ENGLISH_FONT As String = "Tahoma"
Private Const DEFAULT_JAPANESE_FONT As String = "MS Gothic"

' Common dialogue filters
Private Const ALL_SUPPORTED_FILE_FILTER As String = "All supported files (*.bin;*.gen;*.32x;*.smd)|*.bin;*.gen;*.32x;*.smd"
Private Const BIN_FILE_FILTER As String = "Sega Genesis ROM files (*.bin;*.gen;*.32x)|*.bin;*.gen;*.32x"
Private Const SMD_FILE_FILTER As String = "Super Magic Drive files (*.smd)|*.smd"
Private Const IPS_FILE_FILTER As String = "IPS patch files (*.ips)|*.ips"
Private Const ALL_FILE_FILTER As String = "All files (*.*)|*.*"
Private Const BIN_DEFAULT_EXTENSION As String = "bin"
Private Const SMD_DEFAULT_EXTENSION As String = "smd"
Private Const IPS_DEFAULT_EXTENSION As String = "ips"

Private Function OpenFile(Path As String) As Boolean
    On Error GoTo READ_ERROR
    
    ' Checks if file exists
    If Dir(Path) = "" Then
        GoTo READ_ERROR
    End If
    
    Dim FileSize As Long: FileSize = FileLen(Path)
    
    ' Checks if file is large enough
    If FileSize <= HEADER_SIZE Then
        GoTo READ_ERROR
    End If
    
    File = OpenROM(Path)
    
    ' Checks if file passes TMSS check
    If Not File.Valid Then
        GoTo READ_ERROR
    End If
    
    LoadFile File
    
    FileOpened = True
    FileModified = False
    FormatChanged = False
    
    SetTitle Path
    
    mnuClose.Enabled = True
    mnuSave.Enabled = True
    mnuSaveAs.Enabled = True
    mnuApplyPatch.Enabled = True
    
    If File.Format = SMD_FILE Then
        mnuConvertToBin.Enabled = True
        mnuConvertToSMD.Enabled = False
    Else
        mnuConvertToBin.Enabled = False
        mnuConvertToSMD.Enabled = True
    End If
    
    OpenFile = True
    Exit Function
    
READ_ERROR:
    MsgBox "The selected file is not a valid Sega Genesis ROM file.", _
        vbInformation, _
        "Error"
    OpenFile = False
    Exit Function
End Function

Private Sub LoadFile(File As ROMFile)
    With File
        txtGeneral(CONSOLE_INDEX).Text = .Console
        txtGeneral(PRODUCT_CODE_INDEX).Text = .ProductCode
        txtGeneral(MODEM_DATA_INDEX).Text = .ModemData
        txtGeneral(COMMENT_INDEX).Text = .Comment
        
        SetName LOCAL_NAME_INDEX, .LocalName
        SetName INTL_NAME_INDEX, .IntlName
        
        SetCopyright .Company, .Copyright
        SetProductType .ProductType
        SetIODevices .IODevices
        SetRegions .Regions
        
        SetAddresses 0, .ROMStart, .ROMStop
        SetAddresses 2, .RAMStart, .RAMStop
        SetAddresses 4, .SRAMStart, .SRAMStop
        SetChecksum .Checksum, .CalculatedChecksum
        SetCRC32 .CRC32
        SetSize .FileSize
    End With
End Sub

Private Function SaveFile() As Boolean
    On Error GoTo SAVE_ERROR
    
    With File
        .Console = txtGeneral(CONSOLE_INDEX).Text
        .Company = txtGeneral(COMPANY_INDEX).Text
        .Copyright = txtGeneral(COPYRIGHT_INDEX).Text
        .LocalName = txtGeneral(LOCAL_NAME_INDEX).Text
        .IntlName = txtGeneral(INTL_NAME_INDEX).Text
        .ProductType = txtGeneral(PRODUCT_TYPE_INDEX).Text
        .ProductCode = txtGeneral(PRODUCT_CODE_INDEX).Text
        .ModemData = txtGeneral(MODEM_DATA_INDEX).Text
        .Comment = txtGeneral(COMMENT_INDEX).Text
        
        .IODevices = ReadCheckboxes(chkIODevices)
        .Regions = ReadRegions
    End With
    
    SaveROM File
    
    FileModified = False
    FormatChanged = False
    
    SetCRC32 File.CRC32
    SetTitle File.Path
    
    SaveFile = True
    Exit Function
    
SAVE_ERROR:
    MsgBox "There was an error writing the file.", _
        vbInformation, _
        "Error"
    SaveFile = False
    Exit Function
End Function

Private Function Save() As Boolean
    Dim Status As Boolean
    
    If FormatChanged Then
        ' Forces Save As if format changed
        Status = SaveAs
    Else
        Status = SaveFile
    End If
    
    Save = Status
End Function

Private Function SaveAs() As Boolean
    On Error GoTo CANCEL_ERROR
    
    Dim Filter As String
    Dim DefaultExt As String
    
    If File.Format = SMD_FILE Then
        Filter = SMD_FILE_FILTER
        DefaultExt = SMD_DEFAULT_EXTENSION
    Else
        Filter = BIN_FILE_FILTER
        DefaultExt = BIN_DEFAULT_EXTENSION
    End If
    
    Dim FileName As String
    
    If FormatChanged Then
        FileName = GetFileNameWithoutExtension(File.Path)
    Else
        FileName = GetFileName(File.Path)
    End If
    
    With cdlg
        .Filter = Filter
        .DefaultExt = DefaultExt
        .FileName = FileName
        .CancelError = True
        .Flags = cdlOFNOverwritePrompt
        .ShowSave
    End With
    
    File.Path = cdlg.FileName
    SaveFile
    
    FormatChanged = False
    
    SetTitle File.Path
    
    SaveAs = True
    Exit Function
    
CANCEL_ERROR:
    SaveAs = False
    Exit Function
End Function

Private Sub CloseFile()
    Dim Element As Control
    
    For Each Element In txtGeneral
        Element.Text = ""
    Next
    
    For Each Element In txtAddresses
        Element.Text = ""
    Next
    
    For Each Element In chkIODevices
        Element.Value = 0
    Next
    
    For Each Element In chkRegions
        Element.Value = 0
    Next
    
    For Each Element In optRegionFormat
        Element.Value = IIf(Element.Tag > 0, False, True)
    Next
    
    txtGeneral(LOCAL_NAME_INDEX).Font = DEFAULT_ENGLISH_FONT
    txtGeneral(INTL_NAME_INDEX).Font = DEFAULT_ENGLISH_FONT
    mnuShowAsEnglish.Enabled = False
    mnuShowAsJapanese.Enabled = False
    
    lblCompany.Caption = ""
    cboProductType.ListIndex = 0
    
    txtChecksum.Text = ""
    cmdFixChecksum.Enabled = False
    mnuFixChecksum.Enabled = False
    mnuApplyPatch.Enabled = False
    lblChecksumStatus(0).Caption = ""
    lblChecksumStatus(1).Caption = ""
    
    txtSize.Text = ""
    lblSizeDescription.Caption = ""
    
    Me.Caption = App.Title
    
    mnuClose.Enabled = False
    mnuSave.Enabled = False
    mnuSaveAs.Enabled = False
    
    FileOpened = False
End Sub

Private Function PromptForSave() As Boolean
    Dim Cancel As Boolean
    
    If FileOpened And FileModified Then
        Dim Answer As Integer
        Answer = MsgBox("The ROM file has been modified. Do you want to save the changes?", _
            vbYesNoCancel Or vbExclamation, _
            "Save Changes")
        
        If Answer = vbYes Then
            ' Only closes if save is successful
            Cancel = Not Save
        ElseIf Answer = vbCancel Then
            Cancel = True
        End If
    End If
    
    PromptForSave = Cancel
End Function

Private Function GetFileName(Path) As String
    GetFileName = Right(Path, Len(Path) - InStrRev(Path, "\"))
End Function

Private Function GetFileNameWithoutExtension(Path) As String
    Dim FileName As String: FileName = GetFileName(Path)
    GetFileNameWithoutExtension = Left(FileName, InStrRev(FileName, ".") - 1)
End Function

Private Sub SetModified()
    If FileOpened And Not FileModified Then
        FileModified = True
        Me.Caption = "*" & Me.Caption
    End If
End Sub

Private Sub SetTitle(Path)
    Me.Caption = GetFileName(Path) & " - " & App.Title
End Sub

Private Function FormatBytes(Number As Long, Length As Integer) As String
    Dim HexString As String: HexString = Hex$(Number)
    Dim FormattedString As String: FormattedString = ""
    
    Dim i As Integer
    
    ' Length is in bytes
    For i = 1 To Length * 2 - Len(HexString)
        HexString = "0" & HexString
    Next
    
    For i = 1 To Len(HexString)
        If i Mod 2 > 0 Then
            FormattedString = FormattedString & " "
        End If
        
        FormattedString = FormattedString & Mid(HexString, i, 1)
    Next
    
    FormatBytes = FormattedString
End Function

Private Function ReadCheckboxes(Elements As Variant)
    Dim Value As String: Value = ""
    Dim Element As Control
    
    For Each Element In Elements
        If Element.Value > 0 Then
            Value = Value & Element.Tag
        End If
    Next
    
    ReadCheckboxes = Value
End Function

Private Function ReadRegions()
    Dim Value As String: Value = ""
    
    ' Old format
    If optRegionFormat(0).Value Then
        Value = ReadCheckboxes(chkRegions)
    ' New format
    Else
        Dim HexRegion As Integer: HexRegion = 0
        Dim Element As Control
        
        For Each Element In chkRegions
            If Element.Value > 0 Then
                Select Case Element.Tag
                    Case "J"
                        HexRegion = HexRegion + 1
                    Case "U"
                        HexRegion = HexRegion + 4
                    Case "E"
                        HexRegion = HexRegion + 8
                End Select
            End If
        Next
                
        ' Converts number to ASCII character code
        ' 0-9
        If HexRegion < 10 Then
            Value = Chr(HexRegion + 48)
        ' A-F
        Else
            Value = Chr(HexRegion + 55)
        End If
    End If
    
    ReadRegions = Value
End Function

Private Sub MaxChecked(Elements As Variant, Index As Integer, Max As Integer)
    Dim Count As Integer: Count = 0
    Dim Element As Control
    
    For Each Element In Elements
        If Element.Value > 0 Then
            Count = Count + 1
        End If
    Next
    
    ' Unchecks if exceeding max selected values
    If Count > Max Then
        Elements(Index) = 0
    End If
End Sub

Private Sub FixChecksum()
    With File
        .Checksum = .CalculatedChecksum
        SetChecksum .Checksum, .CalculatedChecksum
    End With
    
    SetModified
End Sub

Private Sub Form_Load()
    Me.Caption = App.Title
    mnuAbout.Caption = "&About " & App.Title
    
    SelectedField = -1
    
    cboProductType.AddItem "(Other)"
    cboProductType.AddItem "Game"
    cboProductType.AddItem "Education"
    cboProductType.AddItem "BIOS (TMSS)"
    cboProductType.AddItem "BIOS (Sega CD)"
    cboProductType.ListIndex = 0
    
    optRegionFormat(0).Value = True
    
    If Command <> "" Then
        ' Splits arguments string into files
        Dim Arguments As String: Arguments = Mid(Command, 2, Len(Command) - 2)
        Dim Files() As String: Files = Split(Arguments, """ """)
        
        ' Opens first argument
        OpenFile Files(0)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = PromptForSave
End Sub

Private Sub cmdFixChecksum_Click()
    FixChecksum
End Sub

Private Sub mnuOpen_Click()
    If PromptForSave Then
        GoTo CANCEL_ERROR
    End If
    
    On Error GoTo CANCEL_ERROR
    
    With cdlg
        .Filter = ALL_SUPPORTED_FILE_FILTER & "|" _
                  & BIN_FILE_FILTER & "|" _
                  & SMD_FILE_FILTER & "|" _
                  & ALL_FILE_FILTER
        .DefaultExt = BIN_DEFAULT_EXTENSION
        .CancelError = True
        .Flags = cdlOFNHideReadOnly
        .ShowOpen
    End With
    
    OpenFile cdlg.FileName
    
    Exit Sub
    
CANCEL_ERROR:
    Exit Sub
End Sub

Private Sub mnuSave_Click()
    Save
End Sub

Private Sub mnuSaveAs_Click()
    SaveAs
End Sub

Private Sub mnuClose_Click()
    If Not PromptForSave Then
        CloseFile
    End If
End Sub

Private Sub mnuExit_Click()
    If Not PromptForSave Then
        End
    End If
End Sub

Private Sub mnuConvertToBin_Click()
    FormatChanged = Not FormatChanged
    File.Format = BIN_FILE
    
    mnuConvertToBin.Enabled = False
    mnuConvertToSMD.Enabled = True
    
    SetModified
End Sub

Private Sub mnuConvertToSMD_Click()
    FormatChanged = Not FormatChanged
    File.Format = SMD_FILE
    
    mnuConvertToBin.Enabled = True
    mnuConvertToSMD.Enabled = False
    
    SetModified
End Sub

Private Sub mnuFixChecksum_Click()
    FixChecksum
End Sub

Private Sub mnuApplyPatch_Click()
    On Error GoTo CANCEL_ERROR
    
    With cdlg
        .Filter = IPS_FILE_FILTER & "|" & ALL_FILE_FILTER
        .DefaultExt = IPS_DEFAULT_EXTENSION
        .CancelError = True
        .Flags = cdlOFNHideReadOnly
        .ShowOpen
    End With
    
    Dim Patch As IPSPatch: Patch = OpenPatch(cdlg.FileName)
    
    If Not Patch.Valid Then
        GoTo READ_ERROR
    End If
    
    File = PatchROM(File, Patch)
    
    LoadFile File
    SetModified
    
    Exit Sub
    
CANCEL_ERROR:
    Exit Sub
    
READ_ERROR:
    MsgBox "The selected file is not a valid IPS patch file.", _
        vbInformation, _
        "Error"
    Exit Sub
End Sub

Private Sub mnuShowAsEnglish_Click()
    If SelectedField = LOCAL_NAME_INDEX _
    Or SelectedField = INTL_NAME_INDEX Then
        txtGeneral(SelectedField).Font = DEFAULT_ENGLISH_FONT
        
        mnuShowAsEnglish.Enabled = False
        mnuShowAsJapanese.Enabled = True
    End If
End Sub

Private Sub mnuShowAsJapanese_Click()
    If SelectedField = LOCAL_NAME_INDEX _
    Or SelectedField = INTL_NAME_INDEX Then
        txtGeneral(SelectedField).Font = DEFAULT_JAPANESE_FONT
        
        mnuShowAsEnglish.Enabled = True
        mnuShowAsJapanese.Enabled = False
    End If
End Sub

Private Sub mnuWebSite_Click()
    frmAbout.OpenWebSite
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub cboProductType_Click()
    Dim Code As String
    
    Select Case cboProductType.ListIndex
        Case 1
            Code = "GM"
        Case 2
            Code = "AL"
        Case 3
            Code = "OS"
        Case 4
            Code = "BR"
        Case Else
            Code = ""
    End Select
    
    txtGeneral(PRODUCT_TYPE_INDEX).Text = Code
End Sub

Private Sub txtGeneral_Change(Index As Integer)
    Dim i As Integer: i = txtGeneral(Index).SelStart
    
    txtGeneral(Index).Text = UCase(txtGeneral(Index).Text)
    txtGeneral(Index).SelStart = i
    
    SetModified
End Sub

Private Sub txtGeneral_Changed()
    SetModified
End Sub

Private Sub txtGeneral_GotFocus(Index As Integer)
    SelectedField = Index
    
    If Index = LOCAL_NAME_INDEX _
    Or Index = INTL_NAME_INDEX Then
        ToggleLanguage Index
    End If
End Sub

Private Sub txtGeneral_LostFocus(Index As Integer)
    SelectedField = -1
    
    If Index = LOCAL_NAME_INDEX _
    Or Index = INTL_NAME_INDEX Then
        mnuShowAsEnglish.Enabled = False
        mnuShowAsJapanese.Enabled = False
    ElseIf Index = PRODUCT_TYPE_INDEX Then
        SetProductType txtGeneral(Index).Text
    End If
End Sub

Private Sub chkIODevices_Click(Index As Integer)
    MaxChecked chkIODevices, Index, IO_DEVICES_LENGTH
    SetModified
End Sub

Private Sub chkRegions_Click(Index As Integer)
    MaxChecked chkRegions, Index, REGIONS_LENGTH
    SetModified
End Sub

Private Sub ToggleLanguage(Index As Integer)
    Dim Lang As Boolean: Lang = txtGeneral(Index).Font = DEFAULT_ENGLISH_FONT
    
    mnuShowAsEnglish.Enabled = Not Lang
    mnuShowAsJapanese.Enabled = Lang
End Sub

Private Sub SetName(Index As Integer, Name As String)
    Dim Found As Boolean
    Dim i As Integer
    
    For i = 1 To Len(Name):
        ' Looks for non-ASCII characters
        If Asc(Mid(Name, i, 1)) >= &H7F Then
            Found = True
            Exit For
        End If
    Next
    
    txtGeneral(Index).Text = Name
    txtGeneral(Index).Font = IIf(Found, DEFAULT_JAPANESE_FONT, DEFAULT_ENGLISH_FONT)
    
    If SelectedField = Index Then
        ToggleLanguage SelectedField
    End If
End Sub

Private Sub SetCopyright(Company As String, Copyright As String)
    txtGeneral(COMPANY_INDEX).Text = Company
    txtGeneral(COPYRIGHT_INDEX).Text = Copyright
    
    Dim Slice As String: Slice = LCase(Company)
    Dim CompanyName As String
    
    If InStr(Slice, "sega") > 0 Then
        CompanyName = "Sega"
    ElseIf InStr(Slice, "acld") > 0 Then
        CompanyName = "Ballistic"
    ElseIf InStr(Slice, "asci") > 0 Then
        CompanyName = "Asciiware"
    ElseIf InStr(Slice, "inf") > 0 Then
        CompanyName = "Infogrames"
    ElseIf InStr(Slice, "rsi") > 0 Then
        CompanyName = "Razorsoft"
    ElseIf InStr(Slice, "trec") > 0 Then
        CompanyName = "Treco"
    ElseIf InStr(Slice, "vrgn") > 0 Then
        CompanyName = "Virgin Games"
    ElseIf InStr(Slice, "wstn") > 0 Then
        CompanyName = "Westone"
    ElseIf InStr(Copyright, "t-snk 95-feb") > 0 Then
        CompanyName = "Hi-Tech Entertainment"
    ElseIf InStr(Slice, "100") > 0 Then
        CompanyName = "THQ Software"
    ElseIf InStr(Slice, "101") > 0 Then
        CompanyName = "TecMagik"
    ElseIf InStr(Slice, "112") > 0 Then
        CompanyName = "Designer Software"
    ElseIf InStr(Slice, "113") > 0 Then
        CompanyName = "Psygnosis"
    ElseIf InStr(Slice, "119") > 0 Then
        CompanyName = "Accolade"
    ElseIf InStr(Slice, "120") > 0 Then
        CompanyName = "Codemasters"
    ElseIf InStr(Slice, "125") > 0 Then
        CompanyName = "Interplay"
    ElseIf InStr(Slice, "130") > 0 Then
        CompanyName = "Activision"
    ElseIf InStr(Slice, "132") > 0 Then
        CompanyName = "Shiny or Playmates"
    ElseIf InStr(Slice, "144") > 0 Then
        CompanyName = "Atlus"
    ElseIf InStr(Slice, "151") > 0 Then
        CompanyName = "Infogrames"
    ElseIf InStr(Slice, "161") > 0 Then
        CompanyName = "Fox Interactive"
    ElseIf InStr(Slice, "239") > 0 Then
        CompanyName = "Disney Interactive"
    ElseIf InStr(Slice, "10") > 0 Then
        CompanyName = "Takara"
    ElseIf InStr(Slice, "11") > 0 Then
        CompanyName = "Taito or Accolade"
    ElseIf InStr(Slice, "12") > 0 Then
        CompanyName = "Capcom"
    ElseIf InStr(Slice, "13") > 0 Then
        CompanyName = "Data East"
    ElseIf InStr(Slice, "14") > 0 Then
        CompanyName = "Namco or Tengen"
    ElseIf InStr(Slice, "15") > 0 Then
        CompanyName = "Sunsoft"
    ElseIf InStr(Slice, "16") > 0 Then
        CompanyName = "Bandai"
    ElseIf InStr(Slice, "17") > 0 Then
        CompanyName = "Dempa"
    ElseIf InStr(Slice, "18") > 0 Or InStr(Slice, "19") > 0 Then
        CompanyName = "Technosoft"
    ElseIf InStr(Slice, "20") > 0 Then
        CompanyName = "Asmik"
    ElseIf InStr(Slice, "22") > 0 Then
        CompanyName = "Micronet"
    ElseIf InStr(Slice, "23") > 0 Then
        CompanyName = "Vic Tokai"
    ElseIf InStr(Slice, "24") > 0 Then
        CompanyName = "American Sammy"
    ElseIf InStr(Slice, "29") > 0 Then
        CompanyName = "Kyugo"
    ElseIf InStr(Slice, "32") > 0 Then
        CompanyName = "Wolf Team"
    ElseIf InStr(Slice, "33") > 0 Then
        CompanyName = "Kaneko"
    ElseIf InStr(Slice, "35") > 0 Then
        CompanyName = "Toaplan"
    ElseIf InStr(Slice, "36") > 0 Then
        CompanyName = "Tecmo"
    ElseIf InStr(Slice, "40") > 0 Then
        CompanyName = "Toaplan"
    ElseIf InStr(Slice, "42") > 0 Then
        CompanyName = "UFL Company Limited"
    ElseIf InStr(Slice, "43") > 0 Then
        CompanyName = "Human"
    ElseIf InStr(Slice, "45") > 0 Then
        CompanyName = "Game Arts"
    ElseIf InStr(Slice, "47") > 0 Then
        CompanyName = "Sage's Creation"
    ElseIf InStr(Slice, "48") > 0 Then
        CompanyName = "Tengen"
    ElseIf InStr(Slice, "49") > 0 Then
        CompanyName = "Renovation or Telenet"
    ElseIf InStr(Slice, "50") > 0 Then
        CompanyName = "Electronic Arts"
    ElseIf InStr(Slice, "56") > 0 Then
        CompanyName = "Razorsoft"
    ElseIf InStr(Slice, "58") > 0 Then
        CompanyName = "Mentrix"
    ElseIf InStr(Slice, "60") > 0 Then
        CompanyName = "Victor Musical Industries"
    ElseIf InStr(Slice, "69") > 0 Then
        CompanyName = "Arena"
    ElseIf InStr(Slice, "70") > 0 Then
        CompanyName = "Virgin Games"
    ElseIf InStr(Slice, "73") > 0 Then
        CompanyName = "Soft Vision"
    ElseIf InStr(Slice, "74") > 0 Then
        CompanyName = "Palsoft"
    ElseIf InStr(Slice, "76") > 0 Then
        CompanyName = "Koei"
    ElseIf InStr(Slice, "79") > 0 Then
        CompanyName = "U.S. Gold"
    ElseIf InStr(Slice, "81") > 0 Then
        CompanyName = "Acclaim or Flying Edge"
    ElseIf InStr(Slice, "83") > 0 Then
        CompanyName = "Gametek"
    ElseIf InStr(Slice, "86") > 0 Then
        CompanyName = "Absolute"
    ElseIf InStr(Slice, "93") > 0 Then
        CompanyName = "Sony"
    ElseIf InStr(Slice, "95") > 0 Then
        CompanyName = "Konami"
    ElseIf InStr(Slice, "97") > 0 Then
        CompanyName = "Tradewest"
    Else
        CompanyName = "Unknown"
    End If
    
    lblCompany.Caption = CompanyName
End Sub

Private Sub SetProductType(ProductType As String)
    Dim Index As Integer
    
    Select Case ProductType
        Case "GM"
            Index = 1
        Case "AL"
            Index = 2
        Case "OS"
            Index = 3
        Case "BR"
            Index = 4
        Case Else
            Index = 0
    End Select
    
    cboProductType.ListIndex = Index
    txtGeneral(PRODUCT_TYPE_INDEX).Text = ProductType
End Sub

Private Sub SetCheckboxes(Elements As Variant, Text As String)
    Dim Element As Control
    
    For Each Element In Elements
        Dim Found As Boolean: Found = InStr(Text, Element.Tag) > 0
        Element.Value = IIf(Found, 1, 0)
    Next
End Sub

Private Sub SetIODevices(IODevices As String)
    SetCheckboxes chkIODevices, IODevices
End Sub

Private Sub SetRegions(Regions As String)
    Dim RegionFormat As Boolean
    
    If Len(Regions) = 1 Then
        Dim AsciiRegion As String: AsciiRegion = Asc(Regions)
        
        ' Region is considered to be in the new format if
        ' only one character and between 0-9 or A-F except E
        RegionFormat = (AsciiRegion >= 48 And AsciiRegion <= 57) _
            Or (AsciiRegion >= 65 And AsciiRegion <= 68) _
            Or AsciiRegion = 70
    End If
    
    ' Old format
    If Not RegionFormat Then
        SetCheckboxes chkRegions, Regions
    ' New format
    Else
        Dim HexRegion As Integer
        
        ' Converts ASCII character code to actual number
        If AsciiRegion >= 48 And AsciiRegion <= 57 Then
            HexRegion = AsciiRegion - 48
        Else
            HexRegion = AsciiRegion - 55
        End If
        
        Dim ConvertedRegions As String: ConvertedRegions = ""
        
        ' Japan (NTSC)
        If HexRegion And 1 Then
            ConvertedRegions = ConvertedRegions & "J"
        End If
        
        ' Overseas (NTSC)
        If HexRegion And 4 Then
            ConvertedRegions = ConvertedRegions & "U"
        End If
        
        ' Overseas (PAL)
        If HexRegion And 8 Then
            ConvertedRegions = ConvertedRegions & "E"
        End If
        
        SetCheckboxes chkRegions, ConvertedRegions
    End If
    
    Dim Element As Control
    
    For Each Element In optRegionFormat
        Element.Value = IIf(Element.Tag > 0, RegionFormat, Not RegionFormat)
    Next
End Sub

Private Sub SetAddresses(Index As Integer, StartAddr As Long, StopAddr As Long)
    txtAddresses(Index).Text = FormatBytes(StartAddr, 4)
    txtAddresses(Index + 1).Text = FormatBytes(StopAddr, 4)
End Sub

Private Sub SetChecksum(Checksum As Long, CalculatedChecksum As Long)
    txtChecksum.Text = FormatBytes(Checksum, 2)
    
    If Checksum = CalculatedChecksum Then
        lblChecksumStatus(0).Caption = Chr(&HFC)
        lblChecksumStatus(0).ForeColor = RGB(0, 192, 0)
        lblChecksumStatus(1).Caption = "Checksum passed."
        
        cmdFixChecksum.Enabled = False
        mnuFixChecksum.Enabled = False
    Else
        lblChecksumStatus(0).Caption = Chr(&HFB)
        lblChecksumStatus(0).ForeColor = RGB(192, 0, 0)
        lblChecksumStatus(1).Caption = "Checksum failed."
        
        cmdFixChecksum.Enabled = True
        mnuFixChecksum.Enabled = True
    End If
End Sub

Private Sub SetCRC32(CRC32 As Long)
    txtCRC32.Text = FormatBytes(CRC32, 4)
End Sub

Private Sub SetSize(FileSize As Long)
    Dim DescriptionInBytes As String: DescriptionInBytes = ""
    Dim DescriptionInBits As String: DescriptionInBits = ""
    
    If FileSize >= &H100000 Then
        DescriptionInBytes = FileSize / &H100000 & " MB"
    Else
        DescriptionInBytes = FileSize \ &H400 & " KB"
    End If
    
    DescriptionInBits = (FileSize / &H100000) * 8 & " Mbit"
    
    txtSize.Text = FormatBytes(FileSize, 3)
    lblSizeDescription.Caption = DescriptionInBytes & " (" & DescriptionInBits & ")"
End Sub
