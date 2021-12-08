Attribute VB_Name = "Module1"
Option Explicit

Declare Function auxGetNumDevs Lib "winmm.dll" () As Long
Type AUXCAPS
  wMid As Integer               'Manufacturer ID
  wPid As Integer               'Product ID
  vDriverVersion As Integer     'Version of the Driver
  szPname As String * 80        'Product Name (NULL-Terminated String)
  wTechnology As Integer        'Type of Device
  dwSupport As Long             'Functionality Supported by Driver
End Type
  'Flags for the wTechnology Field in the AUXCAPS Structure.
  Global Const AUXCAPS_CDAUDIO = 1        'Audio output from an internal CD-ROM drive.
  Global Const AUXCAPS_AUXIN = 2          'Audio output from auxiliary input jacks.
  'Flags for the dwSupport Field in the AUXCAPS Structure.
  Global Const AUXCAPS_VOLUME = &H1       'Supports volume control.
  Global Const AUXCAPS_LRVOLUME = &H2     'Supports separate left and right volume control.

Global Const MMSYSERR_BASE = 0
Global Const MMSYSERR_BADDEVICEID = (MMSYSERR_BASE + 2)     'Specified device ID is out of range.
Global Const MMSYSERR_NODRIVER = (MMSYSERR_BASE + 6)        'The driver failed to install.

Declare Function auxGetDevCaps Lib "winmm.dll" Alias "auxGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As AUXCAPS, ByVal uSize As Long) As Long

Declare Function auxGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long

Declare Function auxSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long

Type CnvrtIVolumesType
  LeftVolume As Integer
  RightVolume As Integer
End Type

Type CnvrtLVolumeType
  Volume As Long
End Type

