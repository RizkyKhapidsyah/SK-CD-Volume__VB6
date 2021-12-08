VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmVolume 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CD Audio"
   ClientHeight    =   3315
   ClientLeft      =   5820
   ClientTop       =   2595
   ClientWidth     =   4935
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3315
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar chkLock 
      Height          =   135
      Left            =   360
      TabIndex        =   6
      Top             =   2280
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   2760
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "&OK"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   2520
      Width           =   855
   End
   Begin VB.VScrollBar RScroll 
      Height          =   1815
      LargeChange     =   10
      Left            =   1200
      Max             =   0
      Min             =   100
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.VScrollBar LScroll 
      Height          =   1815
      LargeChange     =   10
      Left            =   120
      Max             =   0
      Min             =   100
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblRight 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "R"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1245
      TabIndex        =   4
      Top             =   2040
      Width           =   150
   End
   Begin VB.Label lblLeft 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "L"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   195
      TabIndex        =   3
      Top             =   2040
      Width           =   120
   End
   Begin VB.Line lnRight 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   1020
      X2              =   1020
      Y1              =   1800
      Y2              =   240
   End
   Begin VB.Line lnTop 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   555
      X2              =   1025
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line lnLeft 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   1025
      X2              =   555
      Y1              =   1800
      Y2              =   240
   End
End
Attribute VB_Name = "frmVolume"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

    'General Declarations:
    Option Explicit

    Dim FirstFlag As Integer

    Dim DevID As Integer
    Dim RVol As Long, LVol As Long

    Dim LScrollPrevValue As Integer, RScrollPrevValue As Integer

    Dim lk As Integer

Private Sub cmdOK_Click()

    Unload frmVolume

End Sub

Private Function Cnvrt2Volume(LeftVol As Long, RightVol As Long) As Long
  Dim IVolumes As CnvrtIVolumesType
  Dim LVolume As CnvrtLVolumeType

  IVolumes.LeftVolume = CInt("&H" & Hex$(LeftVol))
  IVolumes.RightVolume = CInt("&H" & Hex$(RightVol))
  LSet LVolume = IVolumes
  Cnvrt2Volume = LVolume.Volume
End Function

Private Sub DisplayErrorAndExit(ErrNum As Integer)
  Select Case ErrNum
    Case MMSYSERR_BADDEVICEID
      MsgBox "Specified device ID is out of range."
    Case MMSYSERR_NODRIVER
      MsgBox "The driver failed to install."
  End Select
  End
End Sub

Private Sub Form_Load()
    'Center Volume Form
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    
    'Declare variables
    Dim IniVol As Long
    Dim RetVal As Integer

  FirstFlag = True

  DevID = GetSoundBoardCDAudioID()

  MMControl1.DeviceType = "CDAudio"
  MMControl1.Command = "Open"

  RetVal = auxGetVolume(DevID, IniVol)
  If RetVal <> 0 Then Call DisplayErrorAndExit(RetVal)
  LVol = IniVol And &HFFFF&
  RVol = (IniVol And &HFFFF0000) / &H10000

  LScroll.Value = Abs(CInt((LVol * 100&) / &HFFFF&))
  LScrollPrevValue = LScroll.Value
  RScroll.Value = Abs(CInt((RVol * 100&) / &HFFFF&))
  RScrollPrevValue = RScroll.Value

  FirstFlag = False

   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MMControl1.Command = "Close"
End Sub

Private Function GetSoundBoardCDAudioID()
  Dim NumAuxDevs As Integer
  Dim DevNum As Integer
  Dim DevCaps As AUXCAPS
  Dim RetVal As Integer, Pos As Integer

  NumAuxDevs = auxGetNumDevs()
  GetSoundBoardCDAudioID = 0
  For DevNum = 0 To (NumAuxDevs - 1)
    RetVal = auxGetDevCaps(DevNum, DevCaps, 92)
    If RetVal <> 0 Then Call DisplayErrorAndExit(RetVal)
    Pos = InStr(DevCaps.szPname, "CD")
    If Pos <> 0 Then
      GetSoundBoardCDAudioID = DevNum
      Exit Function
    End If
  Next DevNum
End Function

Private Sub LScroll_Change()
    If chkLock.Value = True Then
        RScroll.Value = LScroll.Value
    End If
    
    Dim DeltaLScroll As Integer, DeltaVolume As Long
    Dim RetVal As Integer

    If Not FirstFlag Then
        DeltaLScroll = LScroll.Value - LScrollPrevValue
        LScrollPrevValue = LScroll.Value
        DeltaVolume = ((&HFFFF& * CLng(Abs(DeltaLScroll))) / 100&)
    If DeltaLScroll < 0 Then
        LVol = LVol - DeltaVolume
        If LVol < &H0& Then LVol = &H0&
        Else
        LVol = LVol + DeltaVolume
        If LVol > &HFFFF& Then LVol = &HFFFF&
        End If
        RetVal = auxSetVolume(DevID, Cnvrt2Volume(LVol, RVol))
        If RetVal <> 0 Then Call DisplayErrorAndExit(RetVal)
    End If

End Sub

Private Sub RScroll_Change()
    If chkLock.Value = True Then
        LScroll.Value = RScroll.Value
    End If

    
    Dim DeltaRScroll As Integer, DeltaVolume As Long
    Dim RetVal As Integer

  If Not FirstFlag Then
    DeltaRScroll = RScroll.Value - RScrollPrevValue
    RScrollPrevValue = RScroll.Value
    DeltaVolume = ((&HFFFF& * CLng(Abs(DeltaRScroll))) / 100&)
    If DeltaRScroll < 0 Then
      RVol = RVol - DeltaVolume
      If RVol < &H0& Then RVol = &H0&
    Else
      RVol = RVol + DeltaVolume
      If RVol > &HFFFF& Then RVol = &HFFFF&
    End If
    RetVal = auxSetVolume(DevID, Cnvrt2Volume(LVol, RVol))
    If RetVal <> 0 Then Call DisplayErrorAndExit(RetVal)
  End If


End Sub

