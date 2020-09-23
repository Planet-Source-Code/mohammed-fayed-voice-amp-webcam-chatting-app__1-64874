Attribute VB_Name = "Module1"
Public Const TIMESLICE = 0.2            ' Time Slicing 1/5 Second

    
  Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWND As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
  Public Declare Function capCreateCaptureWindowA Lib "avicap32.dll" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal nID As Long) As Long
   
      
        
    Public Const WM_USER As Long = 1024
    
    Public Const WM_CAP_CONNECT As Long = 1034
    Public Const WM_CAP_DISCONNECT As Long = 1035
    Public Const WM_CAP_GET_FRAME As Long = 1084
    Public Const WM_CAP_COPY As Long = 1054
    
    Public Const WM_CAP_START As Long = WM_USER
    
    Public Const WM_CAP_DLG_VIDEOFORMAT As Long = WM_CAP_START + 41
    Public Const WM_CAP_DLG_VIDEOSOURCE As Long = WM_CAP_START + 42
    Public Const WM_CAP_DLG_VIDEODISPLAY As Long = WM_CAP_START + 43
    Public Const WM_CAP_GET_VIDEOFORMAT As Long = WM_CAP_START + 44
    Public Const WM_CAP_SET_VIDEOFORMAT As Long = WM_CAP_START + 45
    Public Const WM_CAP_DLG_VIDEOCOMPRESSION As Long = WM_CAP_START + 46
    Public Const WM_CAP_SET_PREVIEW As Long = WM_CAP_START + 50
    

Public Sub main()
Dim str As String
str = Command$
If Trim(str) = "" Then
    frmMainServer.Show
Else
    
    frmMainClient.Show
    frmMainClient.txtRemoteIP = str
End If
End Sub
