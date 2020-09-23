VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMainClient 
   BackColor       =   &H00800080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10125
   Icon            =   "frmMainClient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   492
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   675
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock3 
      Left            =   120
      Top             =   990
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   525
      Left            =   9300
      RightToLeft     =   -1  'True
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   12
      Top             =   4290
      Visible         =   0   'False
      Width           =   1245
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   120
      Top             =   570
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5640
      Top             =   5940
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox txtSend 
      Height          =   855
      Left            =   60
      TabIndex        =   7
      Top             =   6000
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   1508
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmMainClient.frx":030A
   End
   Begin ProChatting.lvButtons_H cmdConnect 
      Height          =   435
      Left            =   8730
      TabIndex        =   4
      Top             =   60
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   767
      Caption         =   "Connect"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16711680
      cFHover         =   16711680
      cBhover         =   16777088
      cGradient       =   16777088
      Gradient        =   1
      Mode            =   0
      Value           =   0   'False
      cBack           =   12640511
   End
   Begin VB.TextBox txtRemotePort 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3240
      TabIndex        =   2
      Text            =   "2501"
      Top             =   -390
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox txtRemoteIP 
      Height          =   315
      Left            =   7290
      TabIndex        =   0
      Text            =   "127.0.0.1"
      Top             =   120
      Width           =   1245
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   150
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin ProChatting.lvButtons_H Command1 
      Height          =   435
      Left            =   8670
      TabIndex        =   5
      Top             =   6810
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   767
      Caption         =   "Close"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16711680
      cFHover         =   16711680
      cBhover         =   16777088
      cGradient       =   16777088
      Gradient        =   1
      Mode            =   0
      Value           =   0   'False
      cBack           =   12640511
   End
   Begin ProChatting.lvButtons_H cmdSend 
      Height          =   435
      Left            =   180
      TabIndex        =   6
      Top             =   6870
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   767
      Caption         =   "Send"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16711680
      cFHover         =   16711680
      cBhover         =   16776960
      cGradient       =   16776960
      Gradient        =   1
      Mode            =   0
      Value           =   0   'False
      cBack           =   16777215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4980
      Top             =   5340
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainClient.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainClient.frx":0734
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainClient.frx":0ACE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainClient.frx":0E68
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainClient.frx":1202
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainClient.frx":159C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainClient.frx":1936
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainClient.frx":1CD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainClient.frx":206A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainClient.frx":2404
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainClient.frx":279E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainClient.frx":2AB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainClient.frx":82AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainClient.frx":8404
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar 
      Height          =   600
      Left            =   60
      TabIndex        =   8
      Top             =   5310
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   1058
      ButtonWidth     =   847
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Font"
            Key             =   "Font"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Color"
            Key             =   "Color"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Size"
            Key             =   "Size"
            ImageIndex      =   6
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   7
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "s8"
                  Text            =   "8"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "s9"
                  Text            =   "9"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "s10"
                  Text            =   "10"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "s11"
                  Text            =   "11"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "s12"
                  Text            =   "12"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "s13"
                  Text            =   "13"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "s14"
                  Text            =   "14"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox DisTxt 
      Height          =   5205
      Left            =   60
      TabIndex        =   9
      Top             =   90
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   9181
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmMainClient.frx":879E
   End
   Begin ProChatting.lvButtons_H cmdStartVideo 
      Height          =   435
      Left            =   8520
      TabIndex        =   10
      Top             =   3510
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   767
      Caption         =   "Start Video"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16711680
      cFHover         =   16711680
      cBhover         =   16777088
      cGradient       =   16777088
      Gradient        =   1
      Mode            =   0
      Value           =   0   'False
      cBack           =   12640511
   End
   Begin ProChatting.lvButtons_H cmdStop 
      Height          =   435
      Left            =   8520
      TabIndex        =   11
      Top             =   4020
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   767
      Caption         =   "Stop Video"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16711680
      cFHover         =   16711680
      cBhover         =   16777088
      cGradient       =   16777088
      Gradient        =   1
      Mode            =   0
      Value           =   0   'False
      Enabled         =   0   'False
      cBack           =   12640511
   End
   Begin ProChatting.lvButtons_H cmdTalk 
      Height          =   435
      Left            =   3270
      TabIndex        =   13
      Top             =   6870
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   767
      Caption         =   "&Talk"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16711680
      cFHover         =   16711680
      cBhover         =   16776960
      cGradient       =   16776960
      Gradient        =   1
      Mode            =   0
      Value           =   0   'False
      Enabled         =   0   'False
      cBack           =   16777215
   End
   Begin VB.Image Image2 
      Height          =   2775
      Left            =   6180
      Stretch         =   -1  'True
      Top             =   600
      Width           =   3405
   End
   Begin VB.Image Image1 
      Height          =   1845
      Left            =   6180
      Stretch         =   -1  'True
      Top             =   3450
      Width           =   2265
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remote Port"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2310
      TabIndex        =   3
      Top             =   180
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IP"
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Left            =   7020
      TabIndex        =   1
      Top             =   180
      Width           =   150
   End
End
Attribute VB_Name = "frmMainClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mCapHwnd As Long ' Capture Video
Private Rec As String ' Reciving Video
Private wStream As WaveStream ' Recording Wave
Private mReciveVideo As Boolean

Private Sub cmdConnect_Click()
    On Error GoTo myErr
    
    If cmdConnect.Caption = "Connect" Then
        If IsNumeric(Trim(txtRemotePort.Text)) = False Then MsgBox "ÃÏÎá ÑÞã ááãäÝÐ ...", vbExclamation, Me.Caption: Exit Sub
        
        If Winsock1.State <> 7 Then
            Winsock1.Close
            Winsock1.RemotePort = CInt(Trim(txtRemotePort.Text))
            Winsock1.RemoteHost = txtRemoteIP.Text
            Winsock1.Connect
        End If
        
        If Winsock2.State <> 7 Then
            Winsock2.Close
            Winsock2.RemotePort = CInt(Trim(txtRemotePort.Text)) + 1
            Winsock2.RemoteHost = txtRemoteIP.Text
            Winsock2.Connect
        End If

        If Winsock3.State <> 7 Then
            Winsock3.RemotePort = CInt(Trim(txtRemotePort.Text)) + 2
            Winsock3.RemoteHost = txtRemoteIP.Text
            Winsock3.Connect
        End If
    
        cmdConnect.Caption = "Disconnect"
        cmdSend.Enabled = True
    Else
        Winsock1.Close
        Winsock2.Close
        Winsock3.Close
        cmdConnect.Caption = "Connect"
        cmdSend.Enabled = False
    End If
    
Exit Sub
myErr:
    MsgBox Err.Number & "   " & Err.Description & "   " & Err.Source

End Sub

Private Sub cmdSend_Click()
    If Winsock1.State <> sckConnected Then Exit Sub
    m_SendText
End Sub

Private Sub cmdStartVideo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    cmdStop.Enabled = True
    cmdStartVideo.Enabled = False
    
    ' setup a capture window
    mCapHwnd = capCreateCaptureWindowA("WebCap", 0, 0, 0, 352, 288, Me.hWND, 0)
    DoEvents
    
    If mCapHwnd = 0 Then
        ' Camera Not Found
        MsgBox "Error in Camera .", vbExclamation, Me.Caption
    End If
    
    ' connect to the capture device
    Call SendMessage(mCapHwnd, WM_CAP_CONNECT, 0, 0)
    DoEvents
    
    Call SendMessage(mCapHwnd, WM_CAP_SET_PREVIEW, 0, 0)
    
    Call SendMessage(mCapHwnd, WM_CAP_GET_FRAME, 0, 0)

    ' copy the frame to the clipboard
    Call SendMessage(mCapHwnd, WM_CAP_COPY, 0, 0)

    Image1.Picture = Clipboard.GetData
    
If Winsock1.State = 7 Then Winsock1.SendData "msf-Vid-Req"
    
End Sub

Private Sub cmdStop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' disconnect from the video source
    On Error Resume Next
    DoEvents
    cmdStop.Enabled = False
    cmdStartVideo.Enabled = True
    
    If Winsock1.State = 7 Then Winsock1.SendData "msf-Vid-StopFrame"
    Call SendMessage(mCapHwnd, WM_CAP_DISCONNECT, 0, 0)
    mCapHwnd = 0
End Sub



Private Sub Command1_Click()
    Unload Me
    End
End Sub


Private Sub Form_Load()
On Error Resume Next
    Set wStream = New WaveStream
    Call wStream.InitACMCodec(WAVE_FORMAT_GSM610, TIMESLICE)
    
'   Call wStream.InitACMCodec(WAVE_FORMAT_ADPCM, TIMESLICE)
'   Call wStream.InitACMCodec(WAVE_FORMAT_MSN_AUDIO, TIMESLICE)
'   Call wStream.InitACMCodec(WAVE_FORMAT_PCM, TIMESLICE)

End Sub

Private Sub cmdTalk_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Activates Audio Recording...
On Error Resume Next
    Dim rc As Long                         ' Return Code Variable
    
    If (Not wStream.Playing And _
        Not wStream.Recording And _
            wStream.RecDeviceFree And _
            wStream.PlayDeviceFree) Then   ' Check Audio Device Status
        
        cmdTalk.Caption = "&Talking"                    ' Update Button Status To "Talking"
        Screen.MousePointer = vbHourglass               ' Set Hourglass
        
        wStream.Recording = True           ' Set Recording Flag
        rc = wStream.RecordWave(Me.hWND, Winsock3)     ' Record voice and send to all connected sockets
        
        Screen.MousePointer = vbDefault                 ' Reset Mouse Pointer
        cmdTalk.Caption = "&Talk"                       ' Reset Button Status
        
    End If

End Sub

Private Sub cmdTalk_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    wStream.Recording = False
End Sub


Private Sub TBar_ButtonClick(ByVal Button As MSComctlLib.Button)
   On Error Resume Next
   Select Case LCase(Button.Key)
       Case "font"
          CD1.Flags = &H1
          CD1.ShowFont
          If CD1.FontName <> "" Then
             txtSend.SelFontName = CD1.FontName
          End If
          txtSend.SelBold = CD1.FontBold
          txtSend.SelItalic = CD1.FontItalic
          txtSend.SelFontSize = CD1.FontSize
       
       Case "color"
          CD1.ShowColor
          If CD1.Color <> "" Then
             txtSend.SelColor = CD1.Color
          End If
   End Select
End Sub

Private Sub TBar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
       Case "s8"
         txtSend.SelFontSize = 8
       Case "s9"
         txtSend.SelFontSize = 9
       Case "s10"
         txtSend.SelFontSize = 10
       Case "s11"
         txtSend.SelFontSize = 11
       Case "s12"
         txtSend.SelFontSize = 12
       Case "s13"
         txtSend.SelFontSize = 13
       Case "s14"
         txtSend.SelFontSize = 14
   End Select
End Sub

Private Sub Winsock1_Close()
    cmdConnect.Caption = "Connect"
    cmdSend.Enabled = False
End Sub

Private Sub Winsock1_Connect()
    cmdSend.Enabled = True
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
    
    Dim StrData As String
    Winsock1.GetData StrData, vbString, bytesTotal
    
    StrData = Trim(StrData)
    
    If Left(StrData, 8) = "msf-Vid-" Then
    
        If Trim(StrData) = "msf-Vid-Req" Then
            If MsgBox("Do you want to recive Video ? ", vbInformation + vbYesNo, Me.Caption) = vbYes Then
                mReciveVideo = True
                Winsock1.SendData "msf-Vid-GetFrame"
            End If
                
        ElseIf Trim(StrData) = "msf-Vid-StopFrame" Then
            mReciveVideo = False

        ElseIf Trim(StrData) = "msf-Vid-GetFrame" Then
            ' Send Picture Frame
            Call SendMessage(mCapHwnd, WM_CAP_GET_FRAME, 0, 0)
            Call SendMessage(mCapHwnd, WM_CAP_COPY, 0, 0)
            Image1.Picture = Clipboard.GetData
    
            Dim mPic As String
            Picture1.Picture = Image1.Picture
            SAVEJPEG App.Path & "\sTemp.jpg", 50, Me.Picture1
            Open App.Path & "\sTemp.jpg" For Binary As #1
                mPic = Space(LOF(1))
                Get #1, , mPic
            Close #1
            Winsock2.SendData mPic & "@CAM@"

        End If
    Exit Sub
    End If
        
    DisTxt.SelStart = Len(DisTxt.Text)
    DisTxt.SelText = "Server says : "
    DisTxt.SelRTF = StrData & vbNewLine
End Sub


Private Sub m_SendText()
On Error Resume Next
    Winsock1.SendData txtSend.TextRTF
    DisTxt.SelStart = Len(DisTxt.Text)
    DisTxt.SelText = Winsock1.LocalIP & "  says : "
    DisTxt.SelRTF = txtSend.TextRTF & vbNewLine
    txtSend.Text = ""
End Sub


'Private Sub Winsock2_ConnectionRequest(ByVal requestID As Long)
'    Winsock2.Close
'    Winsock2.Accept requestID
'    Winsock1.SendData "msf-Vid-GetFrame"
'    MsgBox "Pic Come"
'End Sub


Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next

    Dim mPic As String
    Dim file As String
    
    Winsock2.GetData mPic
    
    If InStr(Right(mPic, 20), "@CAM@") Then
        Dim S1 As Variant
        S1 = Split(mPic, "@CAM@") '
        file = S1(0)
        Rec = Rec & file
            
        Open App.Path & "\rPic2.jpg" For Binary As #2
        Put #2, , Rec
        Rec = ""
        Close #2
        
        Image2.Picture = LoadPicture(App.Path & "\rPic2.jpg")
        If mReciveVideo = True Then Winsock1.SendData "msf-Vid-GetFrame"
    
    Else
        Rec = Rec & mPic
    End If
    
    
    
End Sub

Private Sub Winsock3_Close()
    cmdTalk.Enabled = False
End Sub

Private Sub Winsock3_Connect()
    cmdTalk.Enabled = True
End Sub


Private Sub Winsock3_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
' Incomming Buffer On...
'--------------------------------------------------------------
    Dim rc As Long                                      ' Return Code Variable
    Dim WaveData() As Byte                              ' Byte array of wave data
    Static ExBytes As Long                      ' Extra bytes in frame buffer
    Static ExData() As Byte              ' Extra bytes from frame buffer
'--------------------------------------------------------------
With wStream
    If (Winsock3.BytesReceived > 0) Then        ' Validate that bytes where actually received
        Do While (Winsock3.BytesReceived > 0)   ' While data available...
            If (ExBytes = 0) Then               ' Was there leftover data from last time
                If (.waveChunkSize <= Winsock3.BytesReceived) Then ' Can we get and entire wave buffer of data
                    Call Winsock3.GetData(WaveData, vbByte + vbArray, .waveChunkSize) ' Get 1 wave buffer of data
                    Call .SaveStreamBuffer(1, WaveData) ' Save wave data to buffer
                    Call .AddStreamToQueue(1)       ' Queue current stream for playback
                Else
                    ExBytes = Winsock3.BytesReceived ' Save Extra bytes
                    Call Winsock3.GetData(ExData, vbByte + vbArray, ExBytes) ' Get Extra data
                End If
            Else
                Call Winsock3.GetData(WaveData, vbByte + vbArray, .waveChunkSize - ExBytes) ' Get leftover bits
                ExData = MidB(ExData, 1) & MidB(WaveData, 1) ' Sync wave bits...
                Call .SaveStreamBuffer(1, ExData) ' Save the current wave data to the wave buffer
                Call .AddStreamToQueue(1)           ' Queue the current wave stream
                ExBytes = 0                      ' Clear Extra byte count
                ExData(1) = ""                      ' Clear Extra data buffer
            End If
        Loop                                            ' Look for next Data Chunk
        
        If (Not .Playing And .PlayDeviceFree And _
            Not .Recording And .RecDeviceFree) Then     ' Check Audio Device Status
            Call m_PlaySound                          ' Start PlayBack..
        End If
    End If
End With
End Sub

Private Sub m_PlaySound()
On Error Resume Next
'--------------------------------------------------------------
    Dim rc As Long                                      ' Return Code Variable
    Dim iPort As Integer                                ' Local Port
    Dim itm As Integer                                  ' Current listitem
'--------------------------------------------------------------
    If (Not wStream.Playing And wStream.PlayDeviceFree And _
        Not wStream.Recording And wStream.RecDeviceFree) Then ' Validate Audio Device Status
        wStream.Playing = True                          ' Turn Playing Status On
        cmdTalk.Caption = "&Playing"                    ' Modify Button Status Caption
        Screen.MousePointer = vbHourglass               ' Set Pointer To HourGlass
        
        iPort = wStream.StreamInQueue
        Do While (iPort <> NULLPORTID)                  ' While socket ports have data to playback
            rc = wStream.PlayWave(Me.hWND, iPort)       ' Play wave data in iPort...
            Call wStream.RemoveStreamFromQueue(iPort)   ' Remove PortID From PlayWave Queue
            iPort = wStream.StreamInQueue
        Loop                                            ' Search for next socket in playback queue
        
        Screen.MousePointer = vbDefault                 ' Set Pointer To Normal
        cmdTalk.Caption = "&Talk"                       ' Modify Button Status Caption
        wStream.Playing = False                         ' Turn Playing Status Off
    End If
'--------------------------------------------------------------
End Sub

