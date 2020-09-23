Attribute VB_Name = "modJPG"

Private Declare Function BMPToJPG Lib "converter.dll" (ByVal InputFilename As String, ByVal OutputFilename As String, ByVal Quality As Long) As Integer

Public Sub SAVEJPEG(FSpec As String, ByVal TheQuality As Long, APIC As PictureBox)
   
   On Error Resume Next
   If Dir(App.Path & "\proT_876.bmp") <> "" Then Kill App.Path & "\proT_876.bmp"
   SavePicture APIC.Picture, App.Path & "\proT_876.bmp"
   DoEvents
   
   Call BMPToJPG(App.Path & "\proT_876.bmp", FSpec, TheQuality)
   DoEvents

End Sub
