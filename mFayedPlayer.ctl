VERSION 5.00
Begin VB.UserControl FayedVedioPlayer 
   BackColor       =   &H00000000&
   ClientHeight    =   2265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2865
   ScaleHeight     =   151
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   191
End
Attribute VB_Name = "FayedVedioPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'**********************************************************
'                     <<   Video Palyer   >>
'
'                    By : Mohammed Samir Fayed
'
'**********************************************************

Option Explicit

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function mciSendString Lib "winmm" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Const WS_CHILD = &H40000000

Public Enum AudioChannel
    LeftOnly = 1
    RightOnly = 2
    LeftRight = 3
End Enum

Public AB_IsOpen As Boolean

Private RetVal As Long
Private mWidth As Single, mHeight As Single


Public Sub AA_Open(mFile As String)
    ' ÝÊÍ ÇáãáÝ
    Dim CommandString As String, mShortName As String * 255
    Dim Ext As String, MedTp As String
    Ext = Right$(Trim$(mFile), 3)

    Select Case Ext
        Case "cda"
            MedTp = "videodisc"
        'Case "avi"
        '    MedTp = "avivideo"
        Case Else
            MedTp = "MPEGVideo"
    End Select

    RetVal = GetShortPathName(mFile, mShortName, 254) 'COnvert to short path ( c:\Progra~1\... )
    mShortName = Left$(mShortName, RetVal) 'Reassign the path

    CommandString = "Open " & Trim(mShortName) & " type " & MedTp & " alias AVIFile parent " & CStr(UserControl.hWnd) & " style " & CStr(WS_CHILD)
    RetVal = mciSendString(CommandString, vbNullString, 0, 0)
    If RetVal = 0 Then AB_IsOpen = True
    
    ' áãÚÑÝÉ ÇáÚÑÖ æÇáÅÑÊÝÇÚ ááãáÝ

    Dim RetString As String
    Dim res() As String
    
    RetString = Space$(256)
    CommandString = "where AVIFile destination"
    RetVal = mciSendString(CommandString, RetString, Len(RetString), 0)

    RetString = Left$(RetString, InStr(RetString & vbNullChar, vbNullChar) - 1)
    res() = Split(RetString)
    mWidth = CSng(res(2))
    mHeight = CSng(res(3))

End Sub

Public Sub AA_Close()
    ' áÛáÞ ÇáãáÝ
    RetVal = mciSendString("close AVIFile", 0, 0, 0)
    If RetVal = 0 Then AB_IsOpen = False
    mWidth = 0
    mHeight = 0
End Sub

Public Sub AA_Play()
    ' ÊÔÛíá ÇáãáÝ
    RetVal = mciSendString("play AVIFile", 0, 0, 0)
End Sub

Public Sub AA_Pause()
    ' ááÅäÊÙÇÑ ÇáãÄÞÊ
    RetVal = mciSendString("pause AVIFile", 0, 0, 0)
End Sub

Public Sub AA_Stop()
    ' ááÊæÞÝ æÇáÚæÏÉ áÃæá ÇáãáÝ
    RetVal = mciSendString("stop AVIFile", 0, 0, 0)
    RetVal = mciSendString("seek AVIFile to start", 0, 0, 0)
End Sub

Public Property Get AB_TotalTime() As Long
    ' áãÚÑÝÉ Øæá ÇáÝÊÑÉ ÇáÒãäíÉ ááãáÝ
    Dim strReturn As String * 255
    RetVal = mciSendString("set AVIFile time format milliseconds", 0, 0, 0)
    RetVal = mciSendString("status AVIFile length", strReturn, 255, 0)
    AB_TotalTime = Val(strReturn)
End Property

Public Property Get AB_CurrentPosition() As Long
    ' áãÚÑÝÉ ãßÇä ÇáÊÔÛíá ÇáÍÇáí ÏÇÎá ÇáãáÝ
    Dim strReturn As String * 255
    RetVal = mciSendString("set AVIFile time format milliseconds", 0, 0, 0)
    RetVal = mciSendString("status AVIFile position", strReturn, 255, 0)
    AB_CurrentPosition = Val(strReturn)
End Property

Public Property Let AB_CurrentPosition(mPosInMiliseconds As Long)
    ' ááÅäÊÞÇá Åáì äÞØÉ ÏÇÎá ÇáãáÝ
    Dim RetVal As Long
    RetVal = mciSendString("set AVIFile time format milliseconds", 0, 0, 0)
    RetVal = mciSendString("seek AVIFile to " & mPosInMiliseconds, 0, 0, 0)
End Property


Public Sub AB_SetAudioChannel(mChannel As AudioChannel)
    ' áÊÍÏíÏ ãÎÑÌ ÇáÕæÊ
    ' ÇáÓãÇÚÉ ÇáíÓÑì ¡ Çáíãäì Ãã ÇáÅËäíä
    Select Case mChannel
        Case AudioChannel.LeftOnly    ' LEFT_ONLY
            RetVal = mciSendString("set AVIFile audio right off", 0, 0, 0)
            RetVal = mciSendString("set AVIFile audio left on", 0, 0, 0)
        Case AudioChannel.RightOnly  ' RIGHT_ONLY
            RetVal = mciSendString("set AVIFile audio right on", 0, 0, 0)
            RetVal = mciSendString("set AVIFile audio left off", 0, 0, 0)
        Case AudioChannel.LeftRight  ' RIGHT_LEFT
            RetVal = mciSendString("set AVIFile audio right on", 0, 0, 0)
            RetVal = mciSendString("set AVIFile audio left on", 0, 0, 0)
        End Select
End Sub

Public Sub AB_SetAudioVolume(m_Volume As Integer)
    ' áÖÈØ ÇáÕæÊ
    Dim Ret As Long
    Ret = mciSendString("setaudio AVIFile volume to " & Str(m_Volume), 0, 0, 0)

End Sub


Public Sub AB_Fill(mFill As Boolean)
    ' áÅÚÇÏÉ ÊÍÌíã ÇáãáÝ ÏÇÎá ÇáäÇÝÐÉ
    If mFill = False Then Exit Sub
    RetVal = mciSendString("Put AVIFile window at 0 0 " & CStr(UserControl.ScaleWidth) & " " & CStr(UserControl.ScaleHeight), vbNullString, 0, 0)
End Sub

Public Function AB_GetPlayHieght() As Single
    ' áãÚÑÝÉ ÇáÅÑÊÝÇÚ ÇáÐí íÚãá Úáíå ÇáãáÝ ÇáÂä
    AB_GetPlayHieght = mHeight
End Function

Public Function AB_GetPlayWidth() As Single
    ' áãÚÑÝÉ ÇáÚÑÖ ÇáÐí íÚãá Úáíå ÇáãáÝ ÇáÂä
     AB_GetPlayWidth = mWidth
End Function

Private Sub UserControl_Terminate()
    AA_Close
End Sub
