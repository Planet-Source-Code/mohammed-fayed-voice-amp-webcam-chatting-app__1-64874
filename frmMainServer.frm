VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMainServer 
   BackColor       =   &H00800080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10230
   Icon            =   "frmMainServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   492
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   682
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock3 
      Left            =   30
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtlocalIP 
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   150
      Width           =   1245
   End
   Begin VB.PictureBox Picture1 
      Height          =   525
      Left            =   8955
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   5
      Top             =   3600
      Visible         =   0   'False
      Width           =   1245
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   0
      Top             =   810
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtPort 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2850
      TabIndex        =   0
      Text            =   "2501"
      Top             =   113
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   -30
      Top             =   330
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5670
      Top             =   5970
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox txtSend 
      Height          =   855
      Left            =   90
      TabIndex        =   2
      Top             =   5880
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   1508
      _Version        =   393217
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmMainServer.frx":030A
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5010
      Top             =   5310
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
            Picture         =   "frmMainServer.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainServer.frx":0734
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainServer.frx":0ACE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainServer.frx":0E68
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainServer.frx":1202
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainServer.frx":159C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainServer.frx":1936
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainServer.frx":1CD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainServer.frx":206A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainServer.frx":2404
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainServer.frx":279E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainServer.frx":2AB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainServer.frx":82AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainServer.frx":8404
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar 
      Height          =   540
      Left            =   90
      TabIndex        =   3
      Top             =   5325
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   953
      ButtonWidth     =   847
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
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
      Height          =   5235
      Left            =   90
      TabIndex        =   4
      Top             =   60
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   9234
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmMainServer.frx":879E
   End
   Begin ProChatting.lvButtons_H cmdListen 
      Height          =   435
      Left            =   8250
      TabIndex        =   8
      Top             =   90
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   767
      Caption         =   "Listen"
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
      Gradient        =   2
      Mode            =   0
      Value           =   0   'False
      cBack           =   12640511
   End
   Begin ProChatting.lvButtons_H Command1 
      Height          =   435
      Left            =   8280
      TabIndex        =   9
      Top             =   6900
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
      Gradient        =   2
      Mode            =   0
      Value           =   0   'False
      cBack           =   12640511
   End
   Begin ProChatting.lvButtons_H cmdStartVideo 
      Height          =   435
      Left            =   8610
      TabIndex        =   10
      Top             =   3630
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
      Gradient        =   2
      Mode            =   0
      Value           =   0   'False
      cBack           =   12640511
   End
   Begin ProChatting.lvButtons_H cmdStop 
      Height          =   435
      Left            =   8640
      TabIndex        =   11
      Top             =   4140
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
      Gradient        =   2
      Mode            =   0
      Value           =   0   'False
      cBack           =   12640511
   End
   Begin ProChatting.lvButtons_H cmdTalk 
      Height          =   435
      Left            =   120
      TabIndex        =   12
      Top             =   6810
      Width           =   3075
      _ExtentX        =   5424
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
      cBhover         =   16777088
      cGradient       =   16777088
      Gradient        =   1
      Mode            =   0
      Value           =   0   'False
      cBack           =   12640511
   End
   Begin ProChatting.lvButtons_H cmdSend 
      Height          =   435
      Left            =   3210
      TabIndex        =   13
      Top             =   6810
      Width           =   2955
      _ExtentX        =   5212
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
      cBhover         =   16777088
      cGradient       =   16777088
      Gradient        =   1
      Mode            =   0
      Value           =   0   'False
      cBack           =   12640511
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IP"
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Left            =   6630
      TabIndex        =   7
      Top             =   210
      Width           =   150
   End
   Begin VB.Image Image2 
      Height          =   1845
      Left            =   6210
      Stretch         =   -1  'True
      Top             =   3540
      Width           =   2265
   End
   Begin VB.Image Image1 
      Height          =   2775
      Left            =   6210
      Stretch         =   -1  'True
      Top             =   660
      Width           =   3405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Port"
      ForeColor       =   &H000040C0&
      Height          =   195
      Left            =   2460
      TabIndex        =   1
      Top             =   180
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "frmMainServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Rec As String
Private mCapHwnd As Long
Private wStream As WaveStream
Private mReciveVideo As Boolean

Private Sub cmdListen_Click()
On Error GoTo myErr:

    If cmdListen.Caption = "Listen" Then
        If IsNumeric(Trim(txtPort.Text)) = False Then MsgBox "ÃÏÎá ÑÞã ..", vbInformation, Me.Caption: Exit Sub
        If Winsock1.State <> 7 Then
            Winsock1.Close
            Winsock1.LocalPort = CInt(Trim(txtPort.Text))
            Winsock1.Listen
        End If
        
        If Winsock2.State <> 7 Then
            Winsock2.Close
            Winsock2.LocalPort = CInt(Trim(txtPort.Text)) + 1
            Winsock2.Listen
        End If
        
        If Winsock3.State <> 7 Then
            Winsock3.Close
            Winsock3.LocalPort = CInt(Trim(txtPort.Text)) + 2
            Winsock3.Listen
        End If
        cmdListen.Caption = "Stop"
        cmdSend.Enabled = True
    Else
        Winsock1.Close
        Winsock2.Close
        Winsock3.Close
        cmdListen.Caption = "Listen"
        cmdSend.Enabled = False
    End If

Exit Sub
myErr:
    MsgBox Err.Number & "  " & Err.Description & "  " & Err.Source, vbCritical, Me.Caption

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
    Image2.Picture = Clipboard.GetData
    
    Winsock1.SendData "msf-Vid-Req"
    
End Sub

Private Sub cmdStop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' disconnect from the video source
    On Error Resume Next
    DoEvents
    cmdStop.Enabled = False
    cmdStartVideo.Enabled = True

    Winsock1.SendData "msf-Vid-StopFrame"
    Call SendMessage(mCapHwnd, WM_CAP_DISCONNECT, 0, 0)
    mCapHwnd = 0
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


Private Sub Command1_Click()
    Unload Me
    End
End Sub

Private Sub Form_Load()
On Error Resume Next
    txtlocalIP.Text = Winsock1.LocalIP
    
    Set wStream = New WaveStream
    Call wStream.InitACMCodec(WAVE_FORMAT_GSM610, TIMESLICE)

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
    cmdListen.Caption = "Listen"
    cmdSend.Enabled = False
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    On Error Resume Next
    'If MsgBox("Do you want to Accept Connection From " & vbCrLf & " request ID : " & requestID, vbInformation + vbYesNo, Me.Caption) = vbNo Then Exit Sub
    Winsock1.Close
    Winsock1.Accept requestID
    Winsock1.SendData "server : Connection Complete."
    DisTxt.SelStart = Len(DisTxt.Text)
    DisTxt.Text = "server : Connection Complete." & vbCrLf
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
                'Winsock1.SendData "msf-Vid-Req-Yes"
                mReciveVideo = True
                Winsock1.SendData "msf-Vid-GetFrame"
                'If Winsock2.State <> 7 Then Winsock2.Close
                'Winsock2.LocalPort = 2502
                'Winsock2.Listen
            'Else
            '    Winsock1.SendData "msf-Vid-Req-NO"
            End If
        ElseIf Trim(StrData) = "msf-Vid-StopFrame" Then
            mReciveVideo = False
            'If Winsock2.State <> 7 Then
            '    Winsock2.Close
            '    Winsock2.RemotePort = 2502
            '    Winsock2.RemoteHost = Winsock1.RemoteHostIP
            '    Winsock2.Connect
            'End If
                
        ElseIf Trim(StrData) = "msf-Vid-GetFrame" Then
            ' Send Frame
            
            Call SendMessage(mCapHwnd, WM_CAP_GET_FRAME, 0, 0)
            Call SendMessage(mCapHwnd, WM_CAP_COPY, 0, 0)
            Image2.Picture = Clipboard.GetData
    
            Dim mPic As String
            'If Dir(App.Path & "\sTemp.jpg") <> "" Then Kill App.Path & "\sTemp.jpg"
            Picture1.Picture = Image2.Picture
            SAVEJPEG App.Path & "\sTemp2.jpg", 50, Me.Picture1
            Open App.Path & "\sTemp2.jpg" For Binary As #1
                mPic = Space(LOF(1))
                Get #1, , mPic
            Close #1
            Winsock2.SendData mPic & "@CAM@"
            
        End If
    
    Exit Sub
    End If
    
    DisTxt.SelStart = Len(DisTxt.Text)
    DisTxt.SelText = Winsock1.RemoteHostIP & "  Says : "
    DisTxt.SelRTF = StrData & vbNewLine

End Sub


Private Sub m_SendText()
On Error Resume Next
    Winsock1.SendData txtSend.TextRTF
    DisTxt.SelStart = Len(DisTxt.Text)
    DisTxt.SelText = "Server says : "
    DisTxt.SelRTF = txtSend.TextRTF & vbNewLine
    txtSend.Text = ""
End Sub

Private Sub Winsock2_ConnectionRequest(ByVal requestID As Long)
    ' For Video Send & Recive
    Winsock2.Close
    Winsock2.Accept requestID
End Sub


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
            
        Open App.Path & "\rPic.jpg" For Binary As #2
        Put #2, , Rec
        Rec = ""
        Close #2
        
        Image1.Picture = LoadPicture(App.Path & "\rPic.jpg")
        If mReciveVideo = True Then Winsock1.SendData "msf-Vid-GetFrame"
    
    Else
        Rec = Rec & mPic
    End If
    
    
    
End Sub

Private Sub Winsock3_Close()
    cmdTalk.Enabled = False
End Sub

Private Sub Winsock3_ConnectionRequest(ByVal requestID As Long)
    Winsock3.Close
    Winsock3.Accept requestID
    cmdTalk.Enabled = True
End Sub

Private Sub Winsock3_DataArrival(ByVal bytesTotal As Long)
' Incomming Buffer On...
On Error Resume Next
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
