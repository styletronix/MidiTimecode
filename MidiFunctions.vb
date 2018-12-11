'SysEx Send and Midi provided by
'https://github.com/microDRUM/md-config-tool/blob/master/microDRUM_Utility/NAudio/MidiOutEx.cs
'Source code by Styletronix.net 2013 / 2014

Imports System.Runtime.InteropServices
Imports System.Collections

Public Class MidiFunctions
#Region "Declarations"
    ' MIDI input device capabilities structure
    <StructLayout(LayoutKind.Sequential)> _
    Public Structure MIDIINCAPS
        Dim wMid As Short ' Manufacturer ID
        Dim wPid As Short ' Product ID
        Dim vDriverVersion As Integer ' Driver version
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)> Dim szPname As String ' Product Name
        Dim dwSupport As Integer ' Supported extras
    End Structure

    ' MIDI output device capabilities structure
    <StructLayout(LayoutKind.Sequential)> _
    Public Structure MIDIOUTCAPS
        Dim wMid As Short      ' Manufacturer ID
        Dim wPid As Short      ' Product ID
        Dim vDriverVersion As Integer      ' Driver version
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)> Dim szPname As String      ' Product Name
        Dim wTechnology As Short      ' Device type
        Dim wVoices As Short      ' n. of voices (internal synth only)
        Dim wNotes As Short      ' max n. of notes (internal synth only)
        Dim wChannelMask As Short      ' n. of Midi channels (internal synth only)
        Dim dwSupport As Integer      ' Supported extra controllers (volume, etc)
    End Structure

    'MIDI data block header
    'Structure MIDIHDR provided by
    'https://github.com/microDRUM/md-config-tool/blob/master/microDRUM_Utility/NAudio/MidiOutEx.cs
    <StructLayout(LayoutKind.Sequential)> _
    Public Structure MIDIHDR
        Public data As IntPtr
        Public bufferLength As Integer
        Public bytesRecorded As Integer
        Public user As Integer
        Public flags As Integer
        Public lpNext As IntPtr
        Public reserved As Integer
        Public offset As Integer
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=4)> _
        Public dwReserved As Integer()
    End Structure


    'Input functions
    Declare Function midiInGetNumDevs Lib "winmm.dll" () As Short
    Declare Function midiInGetDevCaps Lib "winmm.dll" Alias "midiInGetDevCapsA" (ByVal uDeviceID As Integer, ByRef lpCaps As MIDIINCAPS, ByVal uSize As Integer) As Integer
    Declare Function midiInGetErrorText Lib "winmm.dll" Alias "midiInGetErrorTextA" (ByVal err_Renamed As Integer, ByVal lpText As String, ByVal uSize As Integer) As Integer
    Declare Function midiInOpen Lib "winmm.dll" (ByRef lphMidiIn As IntPtr, ByVal uDeviceID As Integer, ByVal dwCallback As MidiDelegate, ByVal dwInstance As Integer, ByVal dwFlags As Integer) As Integer
    Declare Function midiInClose Lib "winmm.dll" (ByVal hMidiIN As IntPtr) As Integer
    'Declare Function midiInPrepareHeader Lib "winmm.dll" (ByVal hMidiIN As IntPtr, ByRef lpMidiInHdr As MIDIHDR, ByVal uSize As Integer) As Integer
    'Declare Function midiInUnprepareHeader Lib "winmm.dll" (ByVal hMidiIN As IntPtr, ByRef lpMidiInHdr As MIDIHDR, ByVal uSize As Integer) As Integer
    'Declare Function midiInAddBuffer Lib "winmm.dll" (ByVal hMidiIN As IntPtr, ByRef lpMidiInHdr As MIDIHDR, ByVal uSize As Integer) As Integer
    Declare Function midiInReset Lib "winmm.dll" (ByVal hMidiIN As IntPtr) As Integer
    Declare Function midiInStart Lib "winmm.dll" (ByVal hMidiIN As IntPtr) As Integer
    Declare Function midiInStop Lib "winmm.dll" (ByVal hMidiIN As IntPtr) As Integer
    Public Delegate Function MidiDelegate(ByVal MidiHandle As IntPtr, ByVal wMsg As UInteger, ByVal Instance As IntPtr, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Boolean

    'Declare Function midiInPrepareHeader Lib "winmm.dll" (ByVal hMidiIN As IntPtr, ByRef lpMidiInHdr As IntPtr, ByVal uSize As Integer) As Integer
    <DllImport("winmm.dll", EntryPoint:="midiInPrepareHeader")> _
    Public Shared Function midiInPrepareHeader(handle As IntPtr, headerPtr As IntPtr, sizeOfMidiHeader As Integer) As Integer
    End Function


    Declare Function midiInUnprepareHeader Lib "winmm.dll" (ByVal hMidiIN As IntPtr, ByRef lpMidiInHdr As IntPtr, ByVal uSize As Integer) As Integer
    Declare Function midiInAddBuffer Lib "winmm.dll" (ByVal hMidiIN As IntPtr, ByRef lpMidiInHdr As IntPtr, ByVal uSize As Integer) As Integer


    'Output functions
    Declare Function midiOutGetNumDevs Lib "winmm.dll" () As Short
    Declare Function midiOutGetDevCaps Lib "winmm.dll" Alias "midiOutGetDevCapsA" (ByVal uDeviceID As Integer, ByRef lpCaps As MIDIOUTCAPS, ByVal uSize As Integer) As Integer
    Declare Function midiOutGetErrorText Lib "winmm.dll" Alias "midiOutGetErrorTextA" (ByVal errcode As Integer, ByVal lpText As String, ByVal uSize As Integer) As Integer
    Declare Function midiOutOpen Lib "winmm.dll" (ByRef lphMidiOut As IntPtr, ByVal uDeviceID As Integer, ByVal dwCallback As Integer, ByVal dwInstance As Integer, ByVal dwFlags As Integer) As Integer
    Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As IntPtr) As Integer
    Declare Function midiOutReset Lib "winmm.dll" (ByVal hMidiOut As IntPtr) As Long
    Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As IntPtr, ByVal dwMsg As Integer) As Integer


    'Following declarationsprovided by
    'https://github.com/microDRUM/md-config-tool/blob/master/microDRUM_Utility/NAudio/MidiOutEx.cs
    <DllImport("winmm.dll", EntryPoint:="midiOutLongMsg")> _
    Public Shared Function midiOutLongMsg(handle As IntPtr, headerPtr As IntPtr, sizeOfMidiHeader As Integer) As Integer
    End Function
    <DllImport("winmm.dll", EntryPoint:="midiOutPrepareHeader")> _
    Public Shared Function MidiOutPrepareHeader(handle As IntPtr, headerPtr As IntPtr, sizeOfMidiHeader As Integer) As Integer
    End Function
    <DllImport("winmm.dll", EntryPoint:="midiOutUnprepareHeader")> _
    Public Shared Function MidiOutUnprepareHeader(handle As IntPtr, headerPtr As IntPtr, sizeOfMidiHeader As Integer) As Integer
    End Function


    ' Callback Function constants
    Public Const CALLBACK_FUNCTION As Integer = &H30000      ' dwCallback is a FARPROC
    Private Const MIDI_IO_STATUS = &H20&   ' dwCallback is a FARPROC

    Public Const MIM_OPEN As Short = &H3C1S      ' MIDI In Port Opened
    Public Const MIM_CLOSE As Short = &H3C2S      ' MIDI In Port Closed
    Public Const MIM_DATA As Short = &H3C3S      ' MIDI In Short Data (e.g. Notes & CC)
    Public Const MIM_LONGDATA As Short = &H3C4S      ' MIDI In Long Data (i.e. SYSEX)
    Public Const MIM_ERROR As Short = &H3C5S      ' MIDI In Error
    Public Const MIM_LONGERROR As Short = &H3C6S      ' MIDI In Long Error
    Public Const MIM_MOREDATA As Short = &H3CCS      ' MIDI Header Buffer is Full
    Public Const MOM_OPEN As Short = &H3C7S      ' MIDI Out Port Opened
    Public Const MOM_CLOSE As Short = &H3C8S      ' MIDI Out Port Closed
    Public Const MOM_DONE As Short = &H3C9S      ' MIDI Out Data sending completed
    Public Const MOM_POSITIONCB As Short = &HCAS      ' MIDI Out Position requested

    ' Midi Error Constants
    Public Const MMSYSERR_NOERROR As Short = 0
    Public Const MMSYSERR_ERROR As Short = 1
    Public Const MMSYSERR_BADDEVICEID As Short = 2
    Public Const MMSYSERR_NOTENABLED As Short = 3
    Public Const MMSYSERR_ALLOCATED As Short = 4
    Public Const MMSYSERR_INVALHANDLE As Short = 5
    Public Const MMSYSERR_NODRIVER As Short = 6
    Public Const MMSYSERR_NOMEM As Short = 7
    Public Const MMSYSERR_NOTSUPPORTED As Short = 8
    Public Const MMSYSERR_BADERRNUM As Short = 9
    Public Const MMSYSERR_INVALFLAG As Short = 10
    Public Const MMSYSERR_INVALPARAM As Short = 11
    Public Const MMSYSERR_HANDLEBUSY As Short = 12
    Public Const MMSYSERR_INVALIDALIAS As Short = 13
    Public Const MMSYSERR_BADDB As Short = 14
    Public Const MMSYSERR_KEYNOTFOUND As Short = 15
    Public Const MMSYSERR_READERROR As Short = 16
    Public Const MMSYSERR_WRITEERROR As Short = 17
    Public Const MMSYSERR_DELETEERROR As Short = 18
    Public Const MMSYSERR_VALNOTFOUND As Short = 19
    Public Const MMSYSERR_NODRIVERCB As Short = 20
    Public Const MMSYSERR_LASTERROR As Short = 20
    Public Const MIDIERR_UNPREPARED As Short = 64 ' header not prepared
    Public Const MIDIERR_STILLPLAYING As Short = 65 ' still something playing
    Public Const MIDIERR_NOMAP As Short = 66 ' no current map
    Public Const MIDIERR_NOTREADY As Short = 67 ' hardware is still busy
    Public Const MIDIERR_NODEVICE As Short = 68 ' port no longer connected
    Public Const MIDIERR_INVALIDSETUP As Short = 69 ' invalid setup
    Public Const MIDIERR_LASTERROR As Short = 69 ' last error in range

    ' Midi Header flags
    Public Const MHDR_DONE As Short = 1 ' Set by the device driver to indicate that it is finished with the buffer and is returning it to the application.
    Public Const MHDR_PREPARED As Short = 2 ' Set by Windows to indicate that the buffer has been prepared
    Public Const MHDR_INQUEUE As Short = 4 ' Set by Windows to indicate that the buffer is queued for playback
    Public Const MHDR_ISSTRM As Short = 8 ' Set to indicate that the buffer is a stream buffer

    Const MIDIMAPPER = (-1)

    Private MidiOut As IntPtr
    Private MidiIn As IntPtr
    Public DeviceID As Integer
    Public DeviceIDIn As Integer
#End Region

#Region "Settings"
    Public Class Settings
        Public Shared Property DisableMTCOutFilter As Boolean
        Public Shared Property DropSomeMTCQuarterFrames As Boolean
        Public Shared Property FollowMidiSMPTESpecifications As Boolean = True
    End Class
#End Region

#Region "MTC Common"
    Public Class MTCFrame
        Public Sub New()
            Me.FrameRate = MTCFrameRate.fps30
        End Sub
        Public Sub New(Hour As Integer, Minutes As Integer, Seconds As Integer, Frame As Integer, SMPTEQuarterFrame As Integer, FrameRate As MTCFrameRate)
            Me.Hour = Hour
            Me.Minutes = Minutes
            Me.Seconds = Seconds
            Me.Frame = Frame
            Me.FrameRate = FrameRate
            Me.SMPTEQuarterFrame = SMPTEQuarterFrame
        End Sub
        Public Sub New(Time As TimeSpan, FrameRate As MTCFrameRate)
            Me.FrameRate = FrameRate
            Me.SetTime(Time)
        End Sub

        Public Property Frame As Int16
        Public Property Seconds As Int16
        Public Property Minutes As Int16
        Public Property Hour As Int16
        Public Property FrameRate As MTCFrameRate
        Public Property SMPTEQuarterFrame As Int16

        Public Function GetTime() As TimeSpan
            Return New TimeSpan(0, Hour, Minutes, Seconds, (1000 / GetFPS() * Frame) + (1000 / GetFPS() / 4 * SMPTEQuarterFrame))
        End Function
        Public Function GetTimeWithoutSMPTEpart() As TimeSpan
            Return New TimeSpan(0, Hour, Minutes, Seconds, (1000 / GetFPS() * Frame))
        End Function
        Public Function GetFPS() As Byte
            Select Case FrameRate
                Case MTCFrameRate.fps24
                    Return 24
                Case MTCFrameRate.fps25
                    Return 25
                Case MTCFrameRate.fps29Drop
                    Return 29.97
                Case MTCFrameRate.fps30
                    Return 30
            End Select

            Return 24
        End Function
        Public Shadows Function ToString() As String
            Return Hour.ToString("D2") & ":" & _
                Minutes.ToString("D2") & ":" & _
                Seconds.ToString("D2") & ":" & _
                Frame.ToString("D2") & " @ " & _
                FrameRate.ToString
        End Function
        Public Shadows Function ToStringWithSMPTE() As String
            Return Hour.ToString("D2") & ":" & _
                Minutes.ToString("D2") & ":" & _
                Seconds.ToString("D2") & ":" & _
                Frame.ToString("D2") & ":" & _
                SMPTEQuarterFrame.ToString & " @ " & _
                FrameRate.ToString
        End Function
        Public Function Clone() As MTCFrame
            Return New MTCFrame(Me.Hour, Me.Minutes, Me.Seconds, Me.Frame, Me.SMPTEQuarterFrame, Me.FrameRate)
        End Function

        Public Sub SetTime(Time As TimeSpan)
            Hour = Time.Hours
            Minutes = Time.Minutes
            Seconds = Time.Seconds
            Frame = (Math.Truncate(Time.Milliseconds / 1000 * GetFPS()))
            SMPTEQuarterFrame = Math.Truncate(((Time.Milliseconds - (1000 / GetFPS() * Frame)) / (1000 / GetFPS())) * 4)
        End Sub
        Public Function AddFrame() As MTCFrame
            SMPTEQuarterFrame = 0
            Frame += 1
            If Frame >= GetFPS() Then
                Frame = 0
                Seconds += 1
            End If
            If Seconds >= 60 Then
                Seconds = 0
                Minutes += 1
            End If
            If Minutes >= 60 Then
                Minutes = 0
                Hour += 1
            End If
            If Hour >= 24 Then
                Hour = 0
            End If

            Return Me
        End Function
        Public Function SubtractFrame() As MTCFrame
            SMPTEQuarterFrame = 0

            Frame -= 1
            If Frame < 0 Then
                Frame = GetFPS() - 1
                Seconds -= 1
            End If
            If Seconds < 0 Then
                Seconds = 59
                Minutes -= 1
            End If
            If Minutes < 0 Then
                Minutes = 59
                Hour -= 1
            End If
            If Hour < 0 Then
                Hour = 23
            End If

            Return Me
        End Function
    End Class
    Public Enum MTCFrameRate
        <System.ComponentModel.Description("24 FPS")> _
        fps24 = 0
        <System.ComponentModel.Description("25 FPS")> _
        fps25 = 1
        <System.ComponentModel.Description("29.97 FPS (drop)")> _
        fps29Drop = 2
        <System.ComponentModel.Description("30 FPS")> _
        fps30 = 3
    End Enum
#End Region

#Region "MTC Receiver"
    Public Event MTCReceived(sender As Object, e As MTCReceivedEventArgs) 'Raised after new Frame has received
    Public Event MTCTick(sender As Object, e As MTCReceivedEventArgs) 'Raised after each Quarterframe received (4 times per 1 FPS)
    Public Event MTCStatusChanged(sender As Object, e As MTCStatusChangedEventArgs)
    Public Event BeatTickReceived(sender As Object, e As EventArgs)

    Public Class MTCReceivedEventArgs
        Inherits EventArgs
        Public Property MTC As MTCFrame
    End Class
    Public Class MTCStatusChangedEventArgs
        Inherits EventArgs

        Public Property Status As MTCStatus
    End Class
    Public Enum MTCStatus
        Running
        Stopped
        WaitingForFirstFrame
    End Enum

    Public ReadOnly Property ReceivedMTC As MTCFrame
        Get
            SyncLock MTCFrameData_Sync
                Return MTCFrameData.Clone
            End SyncLock
        End Get
    End Property
    Public ReadOnly Property SendingMTC As MTCFrame
        Get
            SyncLock LastMTCFrameSend_SyncLock
                Return LastMTCFrameSend.Clone
            End SyncLock
        End Get
    End Property

    Private BeatTickCounter As Int16 = 0

    Private TemporaryMTCFrameData As New MTCFrame
    Private MTCFrameData_Sync As New Object
    Private MTCFrameData As New MTCFrame

    Private FrameCounter As Byte
    Public MTCCounter As Int64
    Public MTCIncompleteCounter As Int64
    Private _MTCDirectionBackward As Boolean
    Private Property MTCDirectionBackward As Boolean
        Get
            Return _MTCDirectionBackward
        End Get
        Set(value As Boolean)
            If Not _MTCDirectionBackward = value Then
                MTCFrameIncompleted = True
                _MTCDirectionBackward = value
            End If
        End Set
    End Property
    Private MTCFrameIncompleted As Boolean = True

    Private Function MTCFrameCompleteDetection(FrameNo As Byte) As Boolean
        'Detect Time direction
        If FrameNo = 0 And FrameCounter = 7 Then
            MTCDirectionBackward = False
        ElseIf FrameNo = 7 And FrameCounter = 0 Then
            MTCDirectionBackward = True
        ElseIf FrameNo - 1 = FrameCounter Then
            MTCDirectionBackward = False
        ElseIf FrameNo + 1 = FrameCounter Then
            MTCDirectionBackward = True
        Else
            MTCFrameIncompleted = True
        End If

        FrameCounter = FrameNo

        If FrameCounter = 0 Then
            If MTCFrameIncompleted Then
                Debug.WriteLine("Incomplete MTC Frame detected!")
                MTCFrameIncompleted = False
                Return False
            Else
                Return True
            End If

        Else
            Return False

        End If
    End Function
    Private MTCFrameCompleted As Boolean
    Public MTCSenderIsNotMIDIconform As Boolean

    Public DebugList As New Dictionary(Of Int64, String)
    Private _MTCReceiveDebugging As Boolean
    Private DebugID As Int64
    Public Property MTCReceiveDebugging As Boolean
        Get
            Return _MTCReceiveDebugging
        End Get
        Set(value As Boolean)
            Me._MTCReceiveDebugging = value
            If value Then
                DebugList.Clear()
                MTCReceiveStopwatch = New Stopwatch
            Else
                MTCReceiveStopwatch.Stop()
                MTCReceiveStopwatch = Nothing
            End If
        End Set
    End Property
    Private MTCReceiveStopwatch As Stopwatch
    Private _CurrentMTCStatus As MTCStatus
    Public ReadOnly Property MTCReceivingStatus As MTCStatus
        Get
            Return _CurrentMTCStatus
        End Get
    End Property
    Private Property CurrentMTCStatus As MTCStatus
        Get
            Return Me._CurrentMTCStatus
        End Get
        Set(value As MTCStatus)
            If Not Me._CurrentMTCStatus = value Then
                Me._CurrentMTCStatus = value

                'TODO: Make RaiseEvent Asynchronous
                If value = MTCStatus.Running Then
                    'Updating StopDetectionFrame to avoid unexptected STOP detections right after start
                    If Me.MTCDirectionBackward Then
                        Me.MTCStopDetectionLastFrame = Me.ReceivedMTC.Clone.AddFrame
                    Else
                        Me.MTCStopDetectionLastFrame = Me.ReceivedMTC.Clone.SubtractFrame
                    End If
                End If

                RaiseEvent MTCStatusChanged(Me, New MTCStatusChangedEventArgs With {.Status = value})

                Debug.Print(value.ToString)
            End If
        End Set
    End Property
    Private MTCStopDetectionTimer As System.Timers.Timer

    Private Sub ParceMTCData(Data As Byte)
        'Get MIDI Data
        Dim dat = GetHighLow(Data)
        Dim MTCFrameNo = dat.High
        Dim MTCFrameDataByte = dat.Low

        If Not MTCSenderIsNotMIDIconform Then
            If Me.CurrentMTCStatus = MTCStatus.Stopped Then
                Me.CurrentMTCStatus = MTCStatus.WaitingForFirstFrame
            End If
        End If

        'Check previous frame for complete message
        If MTCFrameCompleteDetection(MTCFrameNo) Then
            If Me.MTCDirectionBackward Then
                SyncLock MTCFrameData_Sync
                    Me.MTCFrameData = Me.TemporaryMTCFrameData.Clone
                End SyncLock
            Else
                'Add two Frames based on MIDI definition. On Forward direction, 0 is the base reference and after message 4 on frame is elapsed, after message 7, 2 frames have elapsed , so at 0, the MTC Timecode is two frames behind.
                SyncLock MTCFrameData_Sync
                    If MTCSenderIsNotMIDIconform Then
                        Me.MTCFrameData = Me.TemporaryMTCFrameData.Clone
                    Else
                        Me.MTCFrameData = Me.TemporaryMTCFrameData.Clone.AddFrame.AddFrame
                    End If
                End SyncLock
            End If

            'TODO: Make the RaiseEvent Asynchronous to avoid delays in processing of new incomming midi messages.
            RaiseEvent MTCReceived(Me, New MTCReceivedEventArgs With {.MTC = Me.MTCFrameData.Clone})
            If Not MTCSenderIsNotMIDIconform Then Me.CurrentMTCStatus = MTCStatus.Running
            MTCCounter += 1
        End If

        'At 30FPS one Quarterframe should be received every 8.4ms
        Select Case MTCFrameNo
            Case 0 ' after 0 ms on forward running timecode
                Me.TemporaryMTCFrameData.Frame = MTCFrameDataByte

            Case 1 ' after 8.33 ms on forward running timecode
                Me.TemporaryMTCFrameData.Frame += (MTCFrameDataByte << 4)

            Case 2 ' after 16.66 ms on forward running timecode
                Me.TemporaryMTCFrameData.Seconds = MTCFrameDataByte

            Case 3 ' after 25 ms on forward running timecode
                Me.TemporaryMTCFrameData.Seconds += (MTCFrameDataByte << 4)

            Case 4 ' after 33.33 ms on forward running timecode
                Me.TemporaryMTCFrameData.Minutes = MTCFrameDataByte

            Case 5 ' after 41.66 ms on forward running timecode
                Me.TemporaryMTCFrameData.Minutes += (MTCFrameDataByte << 4)

            Case 6 ' after 50 ms on forward running timecode
                Me.TemporaryMTCFrameData.Hour = MTCFrameDataByte

            Case 7 ' after 58.33 ms on forward running timecode
                Me.TemporaryMTCFrameData.Hour += ((MTCFrameDataByte And 1) << 4)
                Me.TemporaryMTCFrameData.FrameRate = (MTCFrameDataByte >> 1) And 3

        End Select


        'TODO: Detect if MTC Sender sends Quarterframe every 8.4ms thus skipping 1 Frame in every 8 Quartermessages as by MIDI definition. 
        'If not, it is not MIDI conform and MTC Frame should not be incremented at this point.
        If Not MTCSenderIsNotMIDIconform Then
            If MTCFrameNo = 4 Then
                'After 4 QuarterFrames, One SMPTEFrame has passed (every 33.3ms) So SMPTEQuarterFrame = 4 equals to the next SMTPE Frame received
                SyncLock MTCFrameData_Sync
                    If Me.MTCDirectionBackward Then
                        Me.MTCFrameData.SubtractFrame()
                    Else
                        Me.MTCFrameData.AddFrame()
                    End If

                    'TODO: Make the RaiseEvent Asynchronous to avoid delays in processing of new incomming midi messages.
                    'TODO: Implement a integer representation of the MTC Frame to avoid the need for synclocks while accessing MTC Objects
                    SyncLock MTCFrameData_Sync
                        RaiseEvent MTCReceived(Me, New MTCReceivedEventArgs With {.MTC = Me.MTCFrameData.Clone})
                    End SyncLock
                End SyncLock


            ElseIf MTCFrameNo > 4 Then
                SyncLock MTCFrameData_Sync
                    Me.MTCFrameData.SMPTEQuarterFrame = MTCFrameNo - 4
                End SyncLock
            Else
                SyncLock MTCFrameData_Sync
                    Me.MTCFrameData.SMPTEQuarterFrame = MTCFrameNo
                End SyncLock
            End If
        End If


        'Event which is raised for each Beat-Tick = 4 times per FPS
        'TODO: Make the RaiseEvent Asynchronous to avoid delays in processing of new incomming midi messages.
        'TODO: Implement a integer representation of the MTC Frame to avoid the need for synclocks while accessing MTC Objects
        SyncLock MTCFrameData_Sync
            RaiseEvent MTCTick(Me, New MTCReceivedEventArgs With {.MTC = Me.MTCFrameData.Clone})
            If Me.TimeNotifySource = TimeNotifySourceEnum.MTC_IN Then
                VerifyTimelineTick(Me.MTCFrameData.GetTime.TotalMilliseconds)
            End If
        End SyncLock

        'For debugging only
        If MTCReceiveDebugging Then
            If Not MTCReceiveStopwatch.IsRunning Then MTCReceiveStopwatch.Start()
            Try
                DebugID += 1
                DebugList.Add(DebugID, MTCReceiveStopwatch.Elapsed.TotalMilliseconds & " - " & Me.MTCFrameData.ToStringWithSMPTE & " - Byte " & MTCFrameNo)
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private MTCStopDetectionLastFrame As New MTCFrame
    Private MTCStopDetectionCounter As Int64
    Private Sub MTCStopDetectionElapsed(sender As Object, e As System.Timers.ElapsedEventArgs)
        Dim NewMTC = Me.ReceivedMTC.Clone

        If MTCDirectionBackward Then
            If Not NewMTC.GetTime < Me.MTCStopDetectionLastFrame.GetTime Then
                If MTCSenderIsNotMIDIconform OrElse CurrentMTCStatus = MTCStatus.WaitingForFirstFrame Then
                    'Stop of non MIDI conform senders takes more time. MIDI conform stops are detectet within 30 ms while non conform takes about 90 ms
                    If MTCStopDetectionCounter > 10 Then
                        Me.CurrentMTCStatus = MTCStatus.Stopped
                    Else
                        MTCStopDetectionCounter += 1
                    End If
                Else
                    Me.CurrentMTCStatus = MTCStatus.Stopped
                End If
            Else
                'Start of Non MIDI conform Senders are detected here, while MTC conform senders are detected during Parsing which results in no delay during start if sender is midi conform.
                If MTCSenderIsNotMIDIconform Then
                    Me.CurrentMTCStatus = MTCStatus.Running
                End If
                Me.MTCStopDetectionCounter = 0
                Me.MTCStopDetectionLastFrame = NewMTC
            End If
        Else
            If Not NewMTC.GetTime > Me.MTCStopDetectionLastFrame.GetTime Then
                If MTCSenderIsNotMIDIconform OrElse CurrentMTCStatus = MTCStatus.WaitingForFirstFrame Then
                    'Stop of non MIDI conform senders takes more time. MIDI conform stops are detectet within 30 ms while non conform takes about 90 ms
                    If MTCStopDetectionCounter > 10 Then
                        Me.CurrentMTCStatus = MTCStatus.Stopped
                    Else
                        MTCStopDetectionCounter += 1
                    End If
                Else
                    Me.CurrentMTCStatus = MTCStatus.Stopped
                End If
            Else
                'Start of Non MIDI conform Senders are detected here, while MTC conform senders are detected during Parsing which results in no delay during start if sender is midi conform.
                If MTCSenderIsNotMIDIconform Then
                    Me.CurrentMTCStatus = MTCStatus.Running
                End If
                Me.MTCStopDetectionCounter = 0
            End If
        End If

        If Me.CurrentMTCStatus = MTCStatus.Stopped Then Me.MTCStopDetectionCounter = 0

        If MTCSenderIsNotMIDIconform Then
            If Me.MTCStopDetectionCounter = 0 Then
                Me.MTCStopDetectionLastFrame = NewMTC
            End If
        Else
            Me.MTCStopDetectionLastFrame = NewMTC
        End If

    End Sub
#End Region

#Region "MTC Sender"
    Private LastMTCFrameSend As New MTCFrame
    Private LastMTCFrameSend_SyncLock As New Object
    Private MTCFullFrameTimer As System.Timers.Timer
    Public Property SendMTCFullFrameDuringStopAtRegularInterval
    Public Property SysExChannel As Int16 = 127

    Private Sub SendMTC(time As MTCFrame, QuadFrameNo As Int16)
        Select Case QuadFrameNo
            Case 0
                midi_outshort(&HF1, &H0 + GetHighLow(time.Frame).Low, 0)
            Case 1
                midi_outshort(&HF1, &H10 + GetHighLow(time.Frame).High, 0)
            Case 2
                midi_outshort(&HF1, &H20 + GetHighLow(time.Seconds).Low, 0)
            Case 3
                midi_outshort(&HF1, &H30 + GetHighLow(time.Seconds).High, 0)
            Case 4
                midi_outshort(&HF1, &H40 + GetHighLow(time.Minutes).Low, 0)
            Case 5
                midi_outshort(&HF1, &H50 + GetHighLow(time.Minutes).High, 0)
            Case 6
                midi_outshort(&HF1, &H60 + GetHighLow(time.Hour).Low, 0)
            Case 7
                midi_outshort(&HF1, &H70 + (time.FrameRate << 1) + GetHighLow(time.Hour).High, 0)
        End Select
    End Sub
    Private Sub SendMTCFullFrame(MTC As MTCFrame)
        Dim dat As New List(Of Byte)
        dat.Add(&HF0) 'Start
        dat.Add(&H7F) 'Realtime
        dat.Add(Me.SysExChannel) 'SysExChannel 127 -> default
        dat.Add(&H1) 'Midi Time Code
        dat.Add(&H1) 'Full Message
        dat.Add((MTC.FrameRate << 5) + MTC.Hour)
        dat.Add(MTC.Minutes)
        dat.Add(MTC.Seconds)
        dat.Add(MTC.Frame)
        dat.Add(&HF7) 'End

        SendSysEx(dat.ToArray)
    End Sub

    Public Structure GetHighLowResult
        Dim Low As Int16
        Dim High As Int16
    End Structure
    Public Shared Function GetHighLow(value As Integer) As GetHighLowResult
        Dim ret As New GetHighLowResult

        Dim b As New BitArray(New Integer() {value})

        Dim bLow As New BitArray(4)
        Dim bHigh As New BitArray(4)


        bHigh.Item(0) = b.Item(7)
        bHigh.Item(1) = b.Item(6)
        bHigh.Item(2) = b.Item(5)
        bHigh.Item(3) = b.Item(4)

        bLow.Item(0) = b.Item(3)
        bLow.Item(1) = b.Item(2)
        bLow.Item(2) = b.Item(1)
        bLow.Item(3) = b.Item(0)

        ret.Low = ToNumeral(bLow, 4)
        ret.High = ToNumeral(bHigh, 4)

        Return ret
    End Function
    Public Shared Function ToNumeral(binary As BitArray, length As Integer) As Integer
        Dim numeral As Integer = 0
        For i As Integer = 0 To length - 1
            If binary(i) Then
                numeral = numeral Or (CInt(1) << (length - 1 - i))
            End If
        Next
        Return numeral
    End Function
    Public Property MTCsenderHighPrecisionMode As Boolean

    Private MTCRunning As Boolean
    Public ReadOnly Property isMTCSenderRunning As Boolean
        Get
            Return Me.MTCRunning
        End Get
    End Property

    Private NewMTCPosition As MTCFrame
    Public Sub MTCSenderStart()
        If Me.MTCRunning = True Then Return

        Me.MTCRunning = True

        If Not Me.MTCSendingThread Is Nothing AndAlso Me.MTCSendingThread.IsAlive Then Me.MTCSendingThread.Abort()
        Me.MTCSendingThread = New System.Threading.Thread(AddressOf Me.MTCWorker)
        Me.MTCSendingThread.Start()
    End Sub
    Public Sub MTCSenderStart(Time As MTCFrame)
        Me.NewMTCPosition = Time.Clone
        MTCSenderStart()
    End Sub
    Public Sub MTCSenderStop()
        Me.MTCRunning = False
    End Sub
    Public Sub MTCSenderGoto(time As MTCFrame)
        Me.NewMTCPosition = time.Clone
    End Sub
    Public Property BPMtoSend As Double
    Public Property SendStartAndStopWithMTC As Boolean
    Private MTCSendingThread As System.Threading.Thread
    Public Event BeatTickSender(sender As Object, e As EventArgs)
    Private SendContinueDelayCounter As Int16

    'Benachrichtigung für Timeline
    Private _NextTimeNotify As Double = -1
    Private SuspendNotify As Boolean
    Public Property NextTimeNotify As Double
        Get
            Return _NextTimeNotify
        End Get
        Set(value As Double)
            _NextTimeNotify = value
        End Set
    End Property
    Public Event TimeNotify(sender As Object, e As TimeNotifyEventArgs)
    Public Class TimeNotifyEventArgs
        Inherits EventArgs

        Public Property NotifyTime As Double
        Public Property CurrentTime As Double
        Public Property NextTime As Double
    End Class
    Public Property TimeNotifySource As TimeNotifySourceEnum = TimeNotifySourceEnum.MTC_OUT
    Public Enum TimeNotifySourceEnum
        MTC_IN
        MTC_OUT
    End Enum
    Private LastVerifyTime As Double
    Private Sub VerifyTimelineTick(CurrentTimeInMS As Double)
        'Detect rewind
        If LastVerifyTime > CurrentTimeInMS Then
            If _NextTimeNotify = -1 Or _NextTimeNotify > CurrentTimeInMS Then
                _NextTimeNotify = CurrentTimeInMS
            End If
        End If
        LastVerifyTime = CurrentTimeInMS

        If Not SuspendNotify AndAlso _NextTimeNotify >= 0 AndAlso CurrentTimeInMS >= _NextTimeNotify Then
            SuspendNotify = True
            Dim t_notify As New System.Threading.Thread(Sub()
                                                            Dim dat = New TimeNotifyEventArgs With {.CurrentTime = CurrentTimeInMS,
                                                                                                    .NotifyTime = _NextTimeNotify,
                                                                                                    .NextTime = -1}
                                                            RaiseEvent TimeNotify(Me, dat)
                                                            _NextTimeNotify = dat.NextTime
                                                            SuspendNotify = False
                                                        End Sub)
            t_notify.Start()
        End If
    End Sub

    Private Sub MTCWorker()
        Dim NextBeatClock As Double
        Dim BeatTick As Int16

        Dim NextQuadFrame As Double
        Dim CurrentTime As Double
        Dim CurrentTimeNotify As Double
        Dim ResetCounter As Double

        Dim CurrentQuadFrameNo = 0
        Dim TotalQuadFrameNo = 1
        Dim incrementingQuadFrameNo = 0

        Dim CurrentMTCFrame As New MTCFrame
        CurrentMTCFrame.FrameRate = MTCFrameRate.fps30

        If SendStartAndStopWithMTC Then
            SendStart()
        End If

        Dim t As New Stopwatch
        t.Start()

        Do While Me.MTCRunning
            CurrentTime = t.Elapsed.TotalMilliseconds

            'Benachrichtigung für Timeline
            If Me.TimeNotifySource = TimeNotifySourceEnum.MTC_OUT Then
                CurrentTimeNotify = CurrentTime + ResetCounter
                VerifyTimelineTick(CurrentTimeNotify)

                'If Not SuspendNotify AndAlso _NextTimeNotify >= 0 AndAlso CurrentTimeNotify >= _NextTimeNotify Then
                '    SuspendNotify = True
                '    Dim t_notify As New System.Threading.Thread(Sub()
                '                                                    Dim dat = New TimeNotifyEventArgs With {.CurrentTime = CurrentTimeNotify,
                '                                                                                            .NotifyTime = _NextTimeNotify,
                '                                                                                            .NextTime = -1}
                '                                                    RaiseEvent TimeNotify(Me, dat)
                '                                                    _NextTimeNotify = dat.NextTime
                '                                                    SuspendNotify = False
                '                                                End Sub)
                '    t_notify.Start()
                'End If
            End If


            If CurrentTime >= NextQuadFrame Then
                If CurrentTime >= 10000000 Then
                    ResetCounter += CurrentTime
                    'Restart Stopwatch after 2 Hour to avoid buffer overflow after abrox 4,7 hours caused by the following calculations.
                    'Currently, there may be a jitter of aprox 1/24 second in bpm send during the Stopwatch restart because "NexBeatClock" is set to 0.
                    t.Restart()
                    TotalQuadFrameNo = 1
                    NextBeatClock = 0
                End If

                If CurrentQuadFrameNo = 0 Then
                    If Not NewMTCPosition Is Nothing Then
                        CurrentMTCFrame = NewMTCPosition.Clone
                        NewMTCPosition = Nothing
                        SendMTCFullFrame(CurrentMTCFrame)
                    End If
                End If

                'Sending Part of MTCFrame
                SendMTC(CurrentMTCFrame, CurrentQuadFrameNo)

                SyncLock Me.LastMTCFrameSend_SyncLock
                    If CurrentQuadFrameNo = 0 Then Me.LastMTCFrameSend = CurrentMTCFrame.Clone

                    'Adding SMPTEQuarterFrame to LastMTCFrameSend which results in a 4 times higher resolution than FPS.
                    If CurrentQuadFrameNo = 3 Then
                        Me.LastMTCFrameSend.AddFrame()
                    ElseIf CurrentQuadFrameNo > 3 Then
                        Me.LastMTCFrameSend.SMPTEQuarterFrame = CurrentQuadFrameNo - 4
                    Else
                        Me.LastMTCFrameSend.SMPTEQuarterFrame = CurrentQuadFrameNo
                    End If
                End SyncLock

                'Prepare next step
                TotalQuadFrameNo += 1
                CurrentQuadFrameNo += 1
                incrementingQuadFrameNo += 1

                'Set Time for next Frame to send
                NextQuadFrame = (1000 * TotalQuadFrameNo) / CurrentMTCFrame.GetFPS / 4

                'Add two FPS to currentMTC when 8 Quarterframes have been send. (4 Quarterframes per FPS)
                'Normally this would happen once at frame 0 and once at frame 4, but this is the temporary MTC Object which has just submitted.
                If CurrentQuadFrameNo > 7 Then
                    CurrentQuadFrameNo = 0
                    CurrentMTCFrame.AddFrame.AddFrame()
                End If
            End If


            'Send BPM (MIDI Clock)
            If BPMtoSend > 0 Then
                If CurrentTime >= NextBeatClock Then
                    SendBeatClock()
                    NextBeatClock += 60000 / (BPMtoSend * 24)

                    BeatTick += 1

                    If BeatTick >= 24 Then
                        If SendContinueDelayCounter >= 10 Then
                            SendContinue()
                            SendContinueDelayCounter = 0
                        Else
                            SendContinueDelayCounter += 1
                        End If

                        RaiseEvent BeatTickSender(Me, New EventArgs)
                        BeatTick = 0
                    End If
                End If
            Else
                NextBeatClock = CurrentTime
            End If

            If MTCsenderHighPrecisionMode Then
                System.Threading.Thread.Sleep(0) 'Sleep(0) simply gives up your thread's current timeslice to another waiting thread.
                'This causes some CPU load on one processor bit gives accuracy of up to < 0.05 ms
            Else
                System.Threading.Thread.Sleep(1) 'Thread sleeps for aprox. 1 to 5 ms depending on cpu load.
            End If
        Loop


        If SendStartAndStopWithMTC Then
            SendStop()
        End If
    End Sub

    Private Sub MTCFullFrameTimerElapsed(sender As Object, e As System.Timers.ElapsedEventArgs)
        If Me.CurrentMTCStatus = MTCStatus.Stopped Then
            If Not Me.NewMTCPosition Is Nothing Then
                Me.LastMTCFrameSend = Me.NewMTCPosition.Clone
                SendMTCFullFrame(Me.NewMTCPosition.Clone)
                If Not SendMTCFullFrameDuringStopAtRegularInterval Then
                    Me.NewMTCPosition = Nothing
                End If
            End If
        End If
    End Sub
#End Region

#Region "MIDI IN"
    Private MidiInHandler = New MidiDelegate(AddressOf MidiInProc)
    Private HDR As MIDIHDR
    Private HDRPtr As IntPtr

    Public Sub OpenMIDI_In()
        Dim midiOpenError As Long = midiInOpen(Me.MidiIn, Me.DeviceIDIn, MidiInHandler, 0, CALLBACK_FUNCTION + MIDI_IO_STATUS)
        If midiOpenError Then
            MsgBox("Der MIDI Mapper kann nicht geöffnet werden. ", 48, "Fehler bei OpenMIDI")
            CloseMIDI_In()
        Else
            'Me.HDR = New MIDIHDR
            'Dim size As Integer = Marshal.SizeOf(GetType(MIDIHDR))
            'HDRPtr = Marshal.AllocHGlobal(size)

            'midiInPrepareHeader(Me.MidiIn, HDRPtr, size)
            'midiInAddBuffer(Me.MidiIn, HDRPtr, size)
            midiInStart(Me.MidiIn)
        End If
    End Sub
    Public Enum MidiTypeEnum
        NoteOn
        NoteOff
        ControlChange
    End Enum
    Public Class MidiReceivedEventArgs
        Inherits EventArgs

        Public Property MidiType As MidiTypeEnum
        Public Property Channel As Byte
        Public Property Command As Byte
        Public Property Value As Byte
    End Class
    Public Event MidiReceived(sender As Object, e As MidiReceivedEventArgs)

    Structure MIDIMSG
        Dim MIDIStatus As Byte
        Dim MIDIByte1 As Byte
        Dim MIDIByte2 As Byte
        Dim Garbage As Byte
    End Structure
    Private SysXFlag As Int16 = 0

    Public Function MidiInProc(ByVal hmidiIn As IntPtr, ByVal wMsg As Integer, ByVal dwInstance As IntPtr, ByVal dwParam1 As IntPtr, ByVal dwParam2 As IntPtr) As Boolean
        If wMsg = MIM_DATA Then
            Dim Bytes = BitConverter.GetBytes(dwParam1.ToInt32)
            Dim MidiMessage As New MIDIMSG
            MidiMessage.MIDIStatus = Bytes(0)
            MidiMessage.MIDIByte1 = Bytes(1)
            MidiMessage.MIDIByte2 = Bytes(2)

            If MidiMessage.MIDIStatus = &HF1 Then
                'MTC Quarterframe
                ParceMTCData(MidiMessage.MIDIByte1)

            ElseIf MidiMessage.MIDIStatus >= &HB0 And MidiMessage.MIDIStatus <= &HBF Then
                'Control Changed
                RaiseEvent MidiReceived(Me, New MidiReceivedEventArgs With {.MidiType = MidiTypeEnum.NoteOn,
                                                                           .Channel = (MidiMessage.MIDIStatus - &HB0) + 1,
                                                                           .Command = MidiMessage.MIDIByte1,
                                                                           .Value = MidiMessage.MIDIByte2}
                                                                           )

            ElseIf MidiMessage.MIDIStatus >= &H90 And MidiMessage.MIDIStatus <= &H9F Then
                'Note ON
                RaiseEvent MidiReceived(Me, New MidiReceivedEventArgs With {.MidiType = MidiTypeEnum.NoteOn,
                                                                            .Channel = (MidiMessage.MIDIStatus - &H90) + 1,
                                                                            .Command = MidiMessage.MIDIByte1,
                                                                            .Value = MidiMessage.MIDIByte2}
                                                                            )
                'Debug.WriteLine("Note ON")

            ElseIf MidiMessage.MIDIStatus >= &H80 And MidiMessage.MIDIStatus <= &H8F Then
                'Note OFF
                RaiseEvent MidiReceived(Me, New MidiReceivedEventArgs With {.MidiType = MidiTypeEnum.NoteOff,
                                                                            .Channel = (MidiMessage.MIDIStatus - &H80) + 1,
                                                                            .Command = MidiMessage.MIDIByte1,
                                                                            .Value = MidiMessage.MIDIByte2}
                                                                            )
                'Debug.WriteLine("Note OFF")

            ElseIf MidiMessage.MIDIStatus = &HF8 Then
                BeatTickCounter += 1
                If BeatTickCounter >= 24 Then
                    RaiseEvent BeatTickReceived(Me, New EventArgs)
                    BeatTickCounter = 0
                End If

            ElseIf MidiMessage.MIDIStatus = &HFA Then
                'Start
                BeatTickCounter = 24

            ElseIf MidiMessage.MIDIStatus = &HFB Then
                'Continue

            ElseIf MidiMessage.MIDIStatus = &HFC Then
                'Stop

            Else
                Debug.WriteLine("Unknown Message: " & MidiMessage.MIDIStatus.ToString())

            End If


        ElseIf wMsg = MIM_LONGDATA Then
            Debug.WriteLine("MIM_LONGDATA")

            Dim mydata As New MIDIHDR
            Marshal.PtrToStructure(dwParam1, mydata)

            If mydata.bufferLength > 0 Then
                Dim Data(mydata.bufferLength - 1) As Byte
                Marshal.PtrToStructure(mydata.data, Data)

                Dim x = Data
            End If

        ElseIf wMsg = MIM_OPEN Then
            Debug.WriteLine("MIM_OPEN")

        ElseIf wMsg = MIM_CLOSE Then
            Debug.WriteLine("MIM_CLOSE")

        ElseIf wMsg = MIM_ERROR Then
            Debug.WriteLine("MIM_ERROR")

        ElseIf wMsg = MIM_LONGERROR Then
            Debug.WriteLine("MIM_LONGERROR")

        ElseIf wMsg = MIM_MOREDATA Then
            Debug.WriteLine("MIM_MOREDATA")


        End If


        Return True
    End Function
    Public Sub CloseMIDI_In()
        midiInStop(Me.MidiIn)
        midiInReset(Me.MidiIn)
        ' midiInUnprepareHeader(Me.MidiIn, Me.HDR, Len(Me.HDR))
        midiInClose(Me.MidiIn)

        Me.MidiIn = 0
    End Sub
    Public Sub Reset_In()
        midiInReset(MidiIn)
    End Sub
    Public Sub SetDevice_In(Device As String)
        Dim midicaps As MIDIINCAPS = Nothing

        ' Test for MIDI mapper
        If midiInGetDevCaps(MIDIMAPPER, midicaps, Len(midicaps)) = 0 Then ' OK
            If midicaps.szPname = Device Then
                Me.DeviceIDIn = MIDIMAPPER
                Return
            End If
        End If

        ' Add other devs
        For i As Integer = 0 To midiInGetNumDevs() - 1
            If midiInGetDevCaps(i, midicaps, Len(midicaps)) = 0 Then ' OK
                If midicaps.szPname = Device Then
                    Me.DeviceIDIn = i
                    Return
                End If
            End If
        Next
    End Sub
    Public Function getDeviceList_In() As List(Of String)
        Dim ret As New List(Of String)
        Dim midicaps As MIDIINCAPS = Nothing

        ' Test for MIDI mapper
        If midiInGetDevCaps(MIDIMAPPER, midicaps, Len(midicaps)) = 0 Then ' OK
            ret.Add(midicaps.szPname)
        End If

        ' Add other devs
        For i As Integer = 0 To midiInGetNumDevs() - 1
            If midiInGetDevCaps(i, midicaps, Len(midicaps)) = 0 Then ' OK
                ret.Add(midicaps.szPname)
            End If
        Next
        Return ret
    End Function
#End Region

#Region "MIDI OUT"
    Public Sub OpenMIDI()
        Dim midiOpenError As Long = midiOutOpen(Me.MidiOut, Me.DeviceID, 0, 0, &H0)
        If midiOpenError Then
            MsgBox("Der MIDI Mapper kann nicht geöffnet werden. ", 48, "Fehler bei OpenMIDI")
            CloseMIDI()
        End If
    End Sub
    Public Sub CloseMIDI()
        midiOutClose(Me.MidiOut)
        Me.MidiOut = 0
    End Sub
    Public Sub Reset()
        midiOutReset(MidiOut)
    End Sub
    Public Sub SetDevice(Device As String)
        Dim midicaps As MIDIOUTCAPS = Nothing

        ' Test for MIDI mapper
        If midiOutGetDevCaps(MIDIMAPPER, midicaps, Len(midicaps)) = 0 Then ' OK
            If midicaps.szPname = Device Then
                Me.DeviceID = MIDIMAPPER
                Return
            End If
        End If

        ' Add other devs
        For i As Integer = 0 To midiOutGetNumDevs() - 1
            If midiOutGetDevCaps(i, midicaps, Len(midicaps)) = 0 Then ' OK
                If midicaps.szPname = Device Then
                    Me.DeviceID = i
                    Return
                End If
            End If
        Next
    End Sub
    Public Function getDeviceList() As List(Of String)
        Dim ret As New List(Of String)
        Dim midicaps As MIDIOUTCAPS = Nothing

        ' Test for MIDI mapper
        If midiOutGetDevCaps(MIDIMAPPER, midicaps, Len(midicaps)) = 0 Then ' OK
            ret.Add(midicaps.szPname)
        End If

        ' Add other devs
        For i As Integer = 0 To midiOutGetNumDevs() - 1
            If midiOutGetDevCaps(i, midicaps, Len(midicaps)) = 0 Then ' OK
                ret.Add(midicaps.szPname)
            End If
        Next
        Return ret
    End Function

    Public Sub note_off(ch As Integer, ByVal kk As Integer)
        midi_outshort(&H80 + ch, kk, 0)
    End Sub
    Public Sub note_on(ch As Integer, ByVal kk As Integer, v As Integer)
        ch = ch + 0
        midi_outshort(&H90 + ch, kk, v)
    End Sub
    Public Sub control_change(ch As Integer, ccnr As Integer, ByVal v As Integer)
        midi_outshort(&HB0 + ch, ccnr, v)
    End Sub
    Public Sub SendMidiOut(ByVal midiData1 As Long, ByVal midiData2 As Long, ByVal midiMessageOut As Long, ByVal Kanal As Integer)
        If Me.MidiOut = 0 Then
            OpenMIDI()
        End If
        Dim midiMessage = Kanal + midiMessageOut + midiData1 * &H100 + midiData2 * &H10000
        Dim Res = midiOutShortMsg(Me.MidiOut, midiMessage)
    End Sub

    Public Sub SendSysEx(data As Byte())
        Dim result As Integer
        Dim ptr As IntPtr
        Dim size As Integer = Marshal.SizeOf(GetType(MIDIHDR))
        Dim header As New MIDIHDR
        header.data = Marshal.AllocHGlobal(data.Length)
        For i As Integer = 0 To data.Length - 1
            Marshal.WriteByte(header.data, i, data(i))
        Next
        header.bufferLength = data.Length
        header.bytesRecorded = data.Length
        header.flags = 0

        Try
            ptr = Marshal.AllocHGlobal(Marshal.SizeOf(GetType(MIDIHDR)))
        Catch generatedExceptionName As Exception
            Marshal.FreeHGlobal(header.data)
            Throw
        End Try

        Try
            Marshal.StructureToPtr(header, ptr, False)
        Catch generatedExceptionName As Exception
            Marshal.FreeHGlobal(header.data)
            Marshal.FreeHGlobal(ptr)
            Throw
        End Try

        result = MidiOutPrepareHeader(MidiOut, ptr, size)
        If result = 0 Then
            result = midiOutLongMsg(MidiOut, ptr, size)
        End If
        If result = 0 Then
            result = MidiOutUnprepareHeader(MidiOut, ptr, size)
        End If

        Marshal.FreeHGlobal(header.data)
        Marshal.FreeHGlobal(ptr)

    End Sub
    Public Sub SendBeatClock()
        midi_outshort(&HF8)
    End Sub
    Public Sub SendStart()
        midi_outshort(&HFA)
    End Sub
    Public Sub SendContinue()
        midi_outshort(&HFB)
    End Sub
    Public Sub SendStop()
        midi_outshort(&HFC)
    End Sub

    Private Sub midi_outshort(b1 As Integer, b2 As Integer, b3 As Integer)
        Dim midi_error As Integer
        If MidiOut = 0 Then Return

        midi_error = midiOutShortMsg(Me.MidiOut, packdword(0, b3, b2, b1))
        If Not midi_error = 0 Then
            MsgBox(midi_error)
        End If
    End Sub
    Private Sub midi_outshort(b1 As Integer)
        Dim midi_error As Integer
        If MidiOut = 0 Then Return

        midi_error = midiOutShortMsg(Me.MidiOut, b1)
        If Not midi_error = 0 Then
            MsgBox(midi_error)
        End If
    End Sub
    Private Function packdword(i1 As Integer, i2 As Integer, i3 As Integer, i4 As Integer) As Long
        Return i2 * &H10000 + i3 * &H100 + i4
    End Function
#End Region

    Public Sub New()
        Me._CurrentMTCStatus = MTCStatus.Stopped

        Me.MTCStopDetectionTimer = New System.Timers.Timer(30)
        Me.MTCStopDetectionTimer.AutoReset = True
        AddHandler Me.MTCStopDetectionTimer.Elapsed, AddressOf MTCStopDetectionElapsed
        Me.MTCStopDetectionTimer.Start()

        Me.MTCFullFrameTimer = New System.Timers.Timer(200)
        Me.MTCFullFrameTimer.AutoReset = True
        AddHandler Me.MTCFullFrameTimer.Elapsed, AddressOf MTCFullFrameTimerElapsed
        Me.MTCFullFrameTimer.Start()
    End Sub
End Class