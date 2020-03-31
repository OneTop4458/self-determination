VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDiskIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// 이 모듈은 수학쟁이님의 모듈입니다.
Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" ( _
    ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByRef lpSecurityAttributes As Any, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long _
) As Long
Private Declare Function SetFilePointer Lib "kernel32.dll" ( _
    ByVal hFile As Long, _
    ByVal lDistanceToMove As Long, _
    ByRef lpDistanceToMoveHigh As Long, _
    ByVal dwMoveMethod As Long _
) As Long
Private Declare Function ReadFile Lib "kernel32.dll" ( _
    ByVal hFile As Long, _
    ByRef lpBuffer As Any, _
    ByVal nNumberOfBytesToRead As Long, _
    ByRef lpNumberOfBytesRead As Long, _
    ByRef lpOverlapped As Any _
) As Long
Private Declare Function WriteFile Lib "kernel32.dll" ( _
    ByVal hFile As Long, _
    ByRef lpBuffer As Any, _
    ByVal nNumberOfBytesToWrite As Long, _
    ByRef lpNumberOfBytesWritten As Long, _
    ByRef lpOverlapped As Any _
) As Long
Private Declare Function DeviceIoControl Lib "kernel32.dll" ( _
    ByVal hDevice As Long, _
    ByVal dwIoControlCode As Long, _
    ByRef lpInBuffer As Any, _
    ByVal nInBufferSize As Long, _
    ByRef lpOutBuffer As Any, _
    ByVal nOutBufferSize As Long, _
    ByRef lpBytesReturned As Long, _
    ByRef lpOverlapped As Any _
) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" ( _
    ByVal hObject As Long _
) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" ( _
    ByRef Destination As Any, _
    ByRef Source As Any, _
    ByVal Length As Long _
)
Private Const GENERIC_READ& = &H80000000
Private Const GENERIC_WRITE& = &H40000000
Private Const INVALID_HANDLE_VALUE& = &HFFFFFFFF
Private Const MasterBootRecord_Signature = &HAA55
Private Const FILE_SHARE_READ& = 1&
Private Const FILE_SHARE_WRITE& = 2&
Private Const OPEN_EXISTING& = 3&
Private Const FILE_BEGIN& = 0&
Private Const FILE_CURRENT& = 1&
Private Const FILE_END& = 2&
Private Const IOCTL_DISK_UPDATE_PROPERTIES = &H70140
Private ptrLow As Long, ptrHigh As Long, lLen As Long, hHardDrive As Long

Public Property Get Handle() As Long
    Handle = hHardDrive
End Property

Public Property Get LowPtr() As Long
    LowPtr = ptrLow
End Property

Public Property Let LowPtr(ByVal NewPtr As Long)
    ptrLow = NewPtr
    SetFilePointer hHardDrive, (ptrLow \ 512) * 512, ptrHigh, FILE_BEGIN
End Property

Public Property Get HighPtr() As Long
    HighPtr = ptrHigh
End Property

Public Property Let HighPtr(ByVal NewPtr As Long)
    ptrHigh = NewPtr
    SetFilePointer hHardDrive, (ptrLow \ 512) * 512, ptrHigh, FILE_BEGIN
End Property

Public Property Get Length() As Long
    Length = lLen
End Property

Public Property Let Length(ByVal NewLen As Long)
    lLen = NewLen
End Property

Public Sub MoveRelatively(ByVal LowPtr As Long, Optional ByVal HighPtr As Long = 0&)
    Dim Temp As Currency
    SetFilePointer hHardDrive, (LowPtr \ 512) * 512, HighPtr, FILE_CURRENT
    
    ptrHigh = ptrHigh + HighPtr
    Temp = CCur(ptrLow) + CCur(LowPtr)
    ptrHigh = ptrHigh + CLng(Temp / 4294987296@)
    ptrLow = CLng(Temp - CCur(CLng(Temp / 4294987296@)) * 4294987296@)
End Sub

Public Function ReadBytes2(Optional ByVal OffsetLow As Long = -1&, _
                           Optional ByVal OffsetHigh As Long = 0&, _
                           Optional ByVal cbLength As Long = 512&, _
                           Optional ByVal TemporaryOffset As Boolean = False) As Byte()
    Dim LowStore As Long, HighStore As Long, LengthStore As Long
    If TemporaryOffset Then LowStore = LowPtr: HighStore = HighPtr
    LengthStore = Length
    If OffsetLow > -1& Then
        LowPtr = OffsetLow: HighPtr = OffsetHigh
    End If
    Length = cbLength
    
    ReadBytes2 = ReadBytes(False)
    If TemporaryOffset Then
        LowPtr = LowStore: HighPtr = HighStore
    Else
        MoveRelatively cbLength
    End If
    Length = LengthStore
End Function

Public Function WriteBytes2(ByRef ByteArray() As Byte, _
                            Optional ByVal OffsetLow As Long = -1&, _
                            Optional ByVal OffsetHigh As Long = 0&, _
                            Optional ByVal cbLength As Long = -1&, _
                            Optional ByVal TemporaryOffset As Boolean = False) As Boolean
    Dim LowStore As Long, HighStore As Long, LengthStore As Long
    If TemporaryOffset Then LowStore = LowPtr: HighStore = HighPtr
    LengthStore = Length
    If OffsetLow > -1& Then
        LowPtr = OffsetLow: HighPtr = OffsetHigh
    End If
    If cbLength > -1& Then
        Length = cbLength
    End If
    
    WriteBytes2 = WriteBytes(ByteArray(), False)
    If TemporaryOffset Then
        LowPtr = LowStore: HighPtr = HighStore
    Else
        MoveRelatively cbLength
    End If
    If cbLength > -1& Then
        Length = LengthStore
    End If
End Function

Public Function ReadBytes(Optional ByVal MoveAfterRead As Boolean = True) As Byte()
    On Error GoTo RuntimeError
    Dim Buffer(0 To 511) As Byte, BufferResult() As Byte, Total As Long, CurrentPtr As Long, DataChunkSize As Long
    
    ReDim BufferResult(lLen - 1)
    CurrentPtr = 0
    Total = 0
    Do While Total < lLen
        ReadFile hHardDrive, Buffer(0), 512, 0&, ByVal 0&
        If Total = 0 Then
            ' 처음이므로
            If (ptrLow Mod 512) + lLen > 512 Then
                RtlMoveMemory BufferResult(CurrentPtr), _
                              Buffer(ptrLow Mod 512), _
                              512 - (ptrLow Mod 512)
                              
                DataChunkSize = 512 - (ptrLow Mod 512)
            Else
                RtlMoveMemory BufferResult(CurrentPtr), _
                              Buffer(ptrLow Mod 512), _
                              lLen
                DataChunkSize = lLen
            End If
        Else
            If lLen - Total < 512 Then
                ' 적당히 잘라야한다.
                RtlMoveMemory BufferResult(CurrentPtr), _
                              Buffer(0), _
                              lLen - Total
                DataChunkSize = lLen - Total
            Else
                ' 512 바이트 그대로 옮겨넣는다.
                RtlMoveMemory BufferResult(CurrentPtr), _
                              Buffer(0), _
                              512
                DataChunkSize = 512
            End If
        End If
        CurrentPtr = CurrentPtr + DataChunkSize
        Total = Total + DataChunkSize
    Loop

    SetFilePointer hHardDrive, ptrLow, ptrHigh, FILE_BEGIN
    If MoveAfterRead Then MoveRelatively Length
    ReadBytes = BufferResult()
RuntimeError:
End Function

Public Function WriteBytes(ByRef ByteArray() As Byte, Optional ByVal MoveAfterWrite As Boolean = True) As Boolean
    On Error GoTo RuntimeError
    Dim Buffer(0 To 511) As Byte, Total As Long, CurrentPtr As Long, DataChunkSize As Long
    
    CurrentPtr = 0
    Total = 0
    Do While Total < lLen
        ReadFile hHardDrive, Buffer(0), 512, 0&, ByVal 0&
        If Total = 0 Then
            ' 처음이므로
            If (ptrLow Mod 512) + lLen > 512 Then
                RtlMoveMemory Buffer(ptrLow Mod 512), _
                              ByteArray(CurrentPtr), _
                              512 - (ptrLow Mod 512)
                              
                DataChunkSize = 512 - (ptrLow Mod 512)
            Else
                RtlMoveMemory Buffer(ptrLow Mod 512), _
                              ByteArray(CurrentPtr), _
                              lLen
                DataChunkSize = lLen
            End If
        Else
            If lLen - Total < 512 Then
                ' 적당히 잘라야한다.
                RtlMoveMemory Buffer(0), _
                              ByteArray(CurrentPtr), _
                              lLen - Total
                DataChunkSize = lLen - Total
            Else
                ' 512 바이트 그대로 옮겨넣는다.
                RtlMoveMemory Buffer(0), _
                              ByteArray(CurrentPtr), _
                              512
                DataChunkSize = 512
            End If
        End If
        SetFilePointer hHardDrive, -512&, ByVal 0&, FILE_CURRENT
        WriteFile hHardDrive, Buffer(0), 512, 0&, ByVal 0&
        CurrentPtr = CurrentPtr + DataChunkSize
        Total = Total + DataChunkSize
    Loop

    SetFilePointer hHardDrive, ptrLow, ptrHigh, FILE_BEGIN
    If MoveAfterWrite Then MoveRelatively Length
    WriteBytes = True
RuntimeError:
End Function

Public Function OpenDrive(Optional ByVal Number As Long = 0) As Boolean
    CloseDrive
    hHardDrive = CreateFile("\\.\PhysicalDrive" & Number, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    If hHardDrive <> INVALID_HANDLE_VALUE Then OpenDrive = True
End Function

Public Sub CloseDrive()
    If hHardDrive = INVALID_HANDLE_VALUE Then Exit Sub
    CloseHandle hHardDrive
    hHardDrive = INVALID_HANDLE_VALUE
End Sub

Public Sub UpdateDisk()
    DeviceIoControl hHardDrive, IOCTL_DISK_UPDATE_PROPERTIES, ByVal 0&, 0, ByVal 0&, 0, 0&, ByVal 0&
End Sub

Private Sub Class_Initialize()
    hHardDrive = INVALID_HANDLE_VALUE
    lLen = 512&
End Sub

Private Sub Class_Terminate()
    CloseDrive
End Sub
