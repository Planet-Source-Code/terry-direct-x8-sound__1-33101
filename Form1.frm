VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Direct X 8 Sound Program"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar hsbVolume 
      Enabled         =   0   'False
      Height          =   255
      Left            =   240
      Max             =   0
      Min             =   -2500
      TabIndex        =   8
      Top             =   2520
      Width           =   5175
   End
   Begin VB.HScrollBar hsbPan 
      Enabled         =   0   'False
      Height          =   255
      Left            =   240
      Max             =   1000
      TabIndex        =   6
      Top             =   1800
      Value           =   1000
      Width           =   5175
   End
   Begin VB.HScrollBar hsbFrequency 
      Enabled         =   0   'False
      Height          =   255
      Left            =   240
      Max             =   1000
      Min             =   1
      TabIndex        =   4
      Top             =   1080
      Value           =   1
      Width           =   5175
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdPlayloop 
      Caption         =   "Play Loop"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2520
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblVolume 
      Caption         =   "Volume"
      Height          =   285
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lblPan 
      Caption         =   "Pan"
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblFrequency 
      Caption         =   "Frequency"
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_dx As New DirectX8
Private m_ds As DirectSound8
Private m_dsb As DirectSoundSecondaryBuffer8

Private Sub Form_Load()

On Error GoTo Label_Error

' Initialize DirectSound.
Set m_ds = m_dx.DirectSoundCreate(vbNullString)
m_ds.SetCooperativeLevel Me.hWnd, DSSCL_PRIORITY
Exit Sub

Label_Error:
    MsgBox "Error. Cannot initialize DirectSound."
    Unload Me

End Sub

Private Sub cmdOpen_Click()

' Display an Open File dialog box.
CommonDialog1.FileName = ""
CommonDialog1.Filter = "WAV Files(*.WAV)|*.WAV|All files (*.*)|*.*"
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Then Exit Sub

' Stop playback and reset the DirectSound buffer.
If Not (m_dsb Is Nothing) Then m_dsb.Stop
Set m_dsb = Nothing

' Create a DirectSound buffer from the WAV file the user selected.
Dim dsBufDesc As DSBUFFERDESC
dsBufDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME
On Error Resume Next
Set m_dsb = m_ds.CreateSoundBufferFromFile(CommonDialog1.FileName, dsBufDesc)
If Err Then
   MsgBox "Could not create sound buffer from the file: " + CommonDialog1.FileName
   Exit Sub
End If
On Error GoTo 0
        
' Initialize the three scroll bars (hsbFrequency, hsbPan, and hsbVolume).
Me.hsbFrequency.Value = dsBufDesc.fxFormat.lSamplesPerSec / 100
Me.hsbPan.Value = 0
Me.hsbVolume.Value = 0

' Set the Caption of Form1 to the WAV filename selected by the user.
Me.Caption = CommonDialog1.FileTitle + " - The My DirectSound Program."

' Enable the Play, Play Loop, and Stop buttons and enable the three scroll bars.
Me.cmdPlay.Enabled = True
Me.cmdPlayloop.Enabled = True
Me.cmdStop.Enabled = True
Me.hsbFrequency.Enabled = True
Me.hsbPan.Enabled = True
Me.hsbVolume.Enabled = True

End Sub
Private Sub cmdPlay_Click()
    
' Start the playback.
m_dsb.Play 0
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

' Stop playback, and reset the m_dsb, m_ds, and m_dx objects.
If Not (m_dsb Is Nothing) Then m_dsb.Stop
Set m_dsb = Nothing
Set m_ds = Nothing
Set m_dx = Nothing

End Sub
Private Sub cmdPlayLoop_Click()
   
   ' Start the playback (in loop).
   m_dsb.Play DSBPLAY_LOOPING

End Sub
Private Sub cmdStop_Click()

' Stop the playback and reset the playback position to the start of the WAV file.
If Not (m_dsb Is Nothing) Then
   m_dsb.Stop
   m_dsb.SetCurrentPosition 0
End If

End Sub
Private Sub hsbFrequency_Change()

' Set the playback frequency according to the current value of the hsbFrequency scroll bar.
m_dsb.SetFrequency 100 * CLng(Me.hsbFrequency.Value)

End Sub
Private Sub hsbPan_Change()

' Set the pan according to the current value of the hsbPan scroll bar.
m_dsb.SetPan hsbPan.Value

End Sub
Private Sub hsbVolume_Change()

' Set the playback volume according to the current value of the hsbVolume scrollbar.
m_dsb.SetVolume hsbVolume.Value

End Sub

