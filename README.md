<div align="center">

## Access CDDB with VB5


</div>

### Description

This code reads a CD's Identification Number and then access the CDDB for a list of Tracks and information about the CD.
 
### More Info
 
CDDB Information

Works best with VB6 and could cause VB5 to crash (But it very rarely does that)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Brian](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brian.md)
**Level**          |Unknown
**User Rating**    |3.7 (67 globes from 18 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/brian-access-cddb-with-vb5__1-1979/archive/master.zip)

### API Declarations

```
'Put the following code in a CLASS MODULE named CCD
Option Explicit
' *********************************************************
' Various types needed for use by mciSendCommand() function
' *********************************************************
' Structure needed for opening the CDROM device
Private Type MCI_OPEN_PARMS
  dwCallback As Long
  wDeviceID As Long
  lpstrDeviceType As String
  lpstrElementName As String
  lpstrAlias As String
End Type
' This structure is used when setting time format to be returned
Private Type MCI_SET_PARMS
  dwCallback As Long
  dwTimeFormat As Long
  dwAudio As Long
End Type
' This structure is used when accessing various status information
Private Type MCI_STATUS_PARMS
  dwCallback As Long
  dwReturn As Long
  dwItem As Long
  dwTrack As Integer
End Type
' The actual API function used when accessing the CD drive
Private Declare Function mciSendCommand Lib "winmm.dll" Alias "mciSendCommandA" _
  (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As Long, ByRef dwParam2 As Any) As Long
' Error codes
Private Const MMSYSERR_NOERROR = 0
' Constants used together with mciSendCommand
Private Const MCI_CLOSE = &H804
Private Const MCI_FORMAT_MSF = 2
Private Const MCI_OPEN = &H803
Private Const MCI_OPEN_ELEMENT = &H200&
Private Const MCI_OPEN_TYPE = &H2000&
Private Const MCI_SET = &H80D
Private Const MCI_SET_TIME_FORMAT = &H400&
Private Const MCI_STATUS_ITEM = &H100&
Private Const MCI_STATUS_LENGTH = &H1&
Private Const MCI_STATUS_NUMBER_OF_TRACKS = &H3&
Private Const MCI_STATUS_POSITION = &H2&
Private Const MCI_TRACK = &H10&
Private Const MCI_STATUS = &H814
' Some instances of the structures declared above
Private mciOpenParms As MCI_OPEN_PARMS
Private mciSetParms As MCI_SET_PARMS
Private mciStatusParms As MCI_STATUS_PARMS
' Some own types needed
Private Type TTrackInfo
  Minutes As Long
  Seconds As Long
  Frames As Long
End Type
' Private storage
Private m_Error As Long     ' Error code from API call
Private m_CID As String     ' Computed disc id
Private m_Drive As String    ' Drive letter
Private m_DeviceID As Long    ' Device Id
Private m_NTracks As Integer   ' Number of tracks in CD
Private m_Length As Long     ' Length of CD in seconds
Private m_Tracks() As TTrackInfo ' Track info for each and every track on the CD
                 ' Zero based. Last index used for storing lead-out
                 ' position information.
' ******************************************************************
' Initialize the class
' ******************************************************************
Private Sub Class_Initialize()
  m_CID = "(unavailable)"
  m_Drive = ""
  m_Error = 0
  m_DeviceID = -1
  m_NTracks = 0
End Sub
' ******************************************************************
' DiscID
' ******************************************************************
Public Property Get DiscID() As String
  DiscID = m_CID
End Property
' ******************************************************************
' ErrorCode
' ******************************************************************
Public Property Get ErrorCode() As Long
  Error = m_Error
End Property
' ******************************************************************
' Init - Initialize the new object. This will open the device
'    and retrieve the information we want
' ******************************************************************
Public Sub Init(sDrive As String)
  Dim p1 As Integer
  m_Error = MMSYSERR_NOERROR
  m_Drive = sDrive
  ' Open the CD
  If OpenCD Then
   Call LoadCDInfo
   CloseCD
  End If
End Sub
' ******************************************************************
' Class_Terminate
' ******************************************************************
Private Sub Class_Terminate()
  If m_DeviceID <> -1 Then
   CloseCD
  End If
End Sub
' *************************************************************************
' OpenCD - Open the CD Driver for use, we also set the time format
'      Returns device id for the opened CD
' *************************************************************************
Private Function OpenCD() As Bool
```


### Source Code

```
'Add 2 command buttons to your form (Call them btnCalc and btnExit
'Add a Combobox called cboDrives and a Textbox called txtID
Option Explicit
Private Sub btnCalc_Click()
  Dim MyCD As New CCD
  MyCD.Init cboDrives.Text
  txtID.Text = MyCD.DiscID
End Sub
Private Sub btnExit_Click()
  Unload Me
End Sub
Private Sub Form_Load()
  cboDrives.AddItem "D:"
  cboDrives.AddItem "E:"
  cboDrives.AddItem "F:"
  cboDrives.AddItem "G:"
  cboDrives.AddItem "H:"
  cboDrives.AddItem "I:"
  cboDrives.ListIndex = 0
End Sub
```

