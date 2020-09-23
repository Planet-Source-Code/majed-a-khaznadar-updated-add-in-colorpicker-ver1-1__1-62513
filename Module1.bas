Attribute VB_Name = "Module1"
' All Rights Reserved For Majed A.Khaznadar
'==============================================
Option Explicit

Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Type POINTAPI
        X As Long
        Y As Long
End Type

Public hColor As String
Public RVal As String
Public GVal As String
Public BVal As String


Public Function GetDcColor() As Double
Dim DeskHdc&, ret&
Dim Pxy As POINTAPI
    DeskHdc = GetDC(0)
    GetCursorPos Pxy
    GetDcColor = GetPixel(DeskHdc, Pxy.X, Pxy.Y)
    ret& = ReleaseDC(0&, DeskHdc)
End Function

Private Sub RGBValue(Color As Double)
    
    RVal = Color And &HFF
    GVal = (Color \ &H100) And &HFF
    BVal = (Color \ &H10000) And &HFF
    
End Sub

Public Function RGBColor(Color As Double) As String

    RGBValue (Color)
    RGBColor = RVal & " " & GVal & " " & BVal

End Function

Public Sub Gradient(TheObject As Object, ByVal RedVal As Long, ByVal GreenVal As Long, ByVal BlueVal As Long, ByVal Direction As Integer)
    Dim Step As Integer, Reps As Integer, FillTop As Integer
    Dim FillLeft As Integer, FillRight As Integer, FillBottom As Integer
    If Direction < 1 Or Direction > 4 Then Direction = 1 ' Bit of error checking
    FillTop = 0
    FillLeft = 0
    If Direction < 3 Then
       Step = (TheObject.Height / 60)
       If Direction = 2 Then FillTop = TheObject.Height - Step
       FillBottom = FillTop + Step
       FillRight = TheObject.Width
    Else
       Step = (TheObject.Width / 60)
       If Direction = 4 Then FillLeft = TheObject.Width - Step
       FillRight = FillLeft + Step
       FillBottom = TheObject.Height
    End If
    For Reps = 1 To 100
       If Direction = 2 And Reps = 100 Then FillTop = 0  ' Need this to get rid of
       If Direction = 4 And Reps = 100 Then FillLeft = 0 ' drift in Step division
       RedVal = RedVal - 3
       GreenVal = GreenVal - 3
       BlueVal = BlueVal - 3
       If RedVal <= 0 Then RedVal = 0
       If GreenVal <= 0 Then GreenVal = 0
       If BlueVal <= 0 Then BlueVal = 0
       TheObject.Line (FillLeft, FillTop)-(FillRight, FillBottom), RGB(RedVal, GreenVal, BlueVal), BF
       If Direction < 3 Then
          If Direction = 1 Then
             FillTop = FillBottom
          Else
             FillTop = FillTop - Step
          End If
          FillBottom = FillTop + Step
       Else
          If Direction = 3 Then
             FillLeft = FillRight
          Else
             FillLeft = FillLeft - Step
          End If
          FillRight = FillLeft + Step
      End If
    Next Reps
End Sub


Public Function SetTopMostWindow(hWnd As Long, Topmost As Boolean) As Long
If Topmost = True Then
    SetTopMostWindow = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
Else
    SetTopMostWindow = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
    SetTopMostWindow = False
End If
End Function


