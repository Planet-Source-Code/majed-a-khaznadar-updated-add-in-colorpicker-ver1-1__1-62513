VERSION 5.00
Begin VB.Form frmAddIn 
   BackColor       =   &H00E0F0F0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Khaznadar Color Picker Ver1.1"
   ClientHeight    =   3960
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   6030
   Icon            =   "frmAddIn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicRGB 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   2
      Left            =   5460
      ScaleHeight     =   255
      ScaleWidth      =   360
      TabIndex        =   30
      Top             =   2835
      Width           =   390
   End
   Begin VB.PictureBox PicRGB 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   1
      Left            =   5460
      ScaleHeight     =   255
      ScaleWidth      =   360
      TabIndex        =   29
      Top             =   2520
      Width           =   390
   End
   Begin VB.PictureBox PicRGB 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   5460
      ScaleHeight     =   255
      ScaleWidth      =   360
      TabIndex        =   28
      Top             =   2205
      Width           =   390
   End
   Begin VB.TextBox txtRBG 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   4785
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   2205
      Width           =   645
   End
   Begin VB.TextBox txtRBG 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   4785
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   2520
      Width           =   645
   End
   Begin VB.TextBox txtRBG 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   4785
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   2835
      Width           =   645
   End
   Begin VB.PictureBox picColorpicker 
      Height          =   330
      Left            =   4680
      Picture         =   "frmAddIn.frx":08CA
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   24
      Top             =   1050
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   4770
      MousePointer    =   99  'Custom
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   70
      TabIndex        =   23
      ToolTipText     =   "Click and Hold the Click, then move the pointer wherever you want to capture the color"
      Top             =   1245
      Width           =   1080
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0FFC0&
      Caption         =   "S&ave"
      Height          =   690
      Left            =   4785
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3150
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0F0F0&
      Caption         =   "Favorits Color Memory Panel"
      Height          =   2775
      Left            =   60
      TabIndex        =   6
      Top             =   1080
      Width           =   4500
      Begin VB.CommandButton Command6 
         BackColor       =   &H000080FF&
         Caption         =   "Current-->Favorite"
         Height          =   330
         Left            =   915
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2220
         Width           =   1980
      End
      Begin VB.PictureBox Pic5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   3000
         ScaleHeight     =   330
         ScaleWidth      =   1335
         TabIndex        =   19
         Top             =   2190
         Width           =   1365
         Begin VB.Shape Shape5 
            BorderColor     =   &H00C0C0C0&
            Height          =   330
            Left            =   0
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H000080FF&
         Caption         =   "Current-->Favorite"
         Height          =   330
         Left            =   915
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1725
         Width           =   1980
      End
      Begin VB.PictureBox Pic4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   3000
         ScaleHeight     =   315
         ScaleWidth      =   1335
         TabIndex        =   16
         Top             =   1725
         Width           =   1365
         Begin VB.Shape Shape4 
            BorderColor     =   &H00C0C0C0&
            Height          =   315
            Left            =   0
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H000080FF&
         Caption         =   "Current-->Favorite"
         Height          =   330
         Left            =   915
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1260
         Width           =   1980
      End
      Begin VB.PictureBox Pic3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3000
         ScaleHeight     =   345
         ScaleWidth      =   1335
         TabIndex        =   13
         Top             =   1260
         Width           =   1365
         Begin VB.Shape Shape3 
            BorderColor     =   &H00C0C0C0&
            Height          =   345
            Left            =   0
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H000080FF&
         Caption         =   "Current-->Favorite"
         Height          =   330
         Left            =   915
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   795
         Width           =   1980
      End
      Begin VB.PictureBox Pic2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   3000
         ScaleHeight     =   315
         ScaleWidth      =   1335
         TabIndex        =   10
         Top             =   810
         Width           =   1365
         Begin VB.Shape Shape2 
            BorderColor     =   &H00C0C0C0&
            Height          =   315
            Left            =   0
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H000080FF&
         Caption         =   "Current-->Favorite"
         Height          =   330
         Left            =   915
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   345
         Width           =   1980
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   3000
         ScaleHeight     =   315
         ScaleWidth      =   1335
         TabIndex        =   7
         Top             =   375
         Width           =   1365
      End
      Begin VB.Shape Shape12 
         BorderColor     =   &H00C0C0C0&
         Height          =   360
         Left            =   900
         Top             =   2205
         Width           =   2010
      End
      Begin VB.Shape Shape11 
         BorderColor     =   &H00C0C0C0&
         Height          =   360
         Left            =   885
         Top             =   1710
         Width           =   2010
      End
      Begin VB.Shape Shape10 
         BorderColor     =   &H00C0C0C0&
         Height          =   360
         Left            =   900
         Top             =   1245
         Width           =   2010
      End
      Begin VB.Shape Shape9 
         BorderColor     =   &H00C0C0C0&
         Height          =   360
         Left            =   900
         Top             =   780
         Width           =   2010
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H00C0C0C0&
         Height          =   360
         Left            =   900
         Top             =   330
         Width           =   2010
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C0C0&
         Height          =   375
         Left            =   2985
         Top             =   360
         Width           =   1395
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color 5"
         Height          =   195
         Left            =   225
         TabIndex        =   20
         Top             =   2295
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color 4"
         Height          =   195
         Left            =   210
         TabIndex        =   17
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color 3"
         Height          =   195
         Left            =   210
         TabIndex        =   14
         Top             =   1335
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color 2"
         Height          =   195
         Left            =   210
         TabIndex        =   11
         Top             =   870
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color 1"
         Height          =   195
         Left            =   225
         TabIndex        =   8
         Top             =   420
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H0040BDA3&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   6000
      TabIndex        =   3
      Top             =   0
      Width           =   6030
      Begin VB.CommandButton Command1 
         Caption         =   "About"
         Height          =   240
         Left            =   4260
         TabIndex        =   5
         Top             =   30
         Width           =   1680
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "© All Rights Reserved For Majed A.Khaznadar 2005"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   90
         TabIndex        =   4
         Top             =   60
         Width           =   3735
      End
   End
   Begin VB.TextBox txtColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   75
      TabIndex        =   2
      Top             =   360
      Width           =   4500
   End
   Begin VB.CommandButton CancelButton 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4725
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   810
      Width           =   1200
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H00CA7736&
      Caption         =   "&Custom Color"
      Height          =   375
      Left            =   4710
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   390
      Width           =   1215
   End
   Begin VB.Shape Shape7 
      Height          =   885
      Left            =   4800
      Top             =   1290
      Width           =   1065
   End
   Begin VB.Shape Shape6 
      Height          =   660
      Left            =   60
      Top             =   375
      Width           =   4515
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VBInstance As VBIDE.VBE
Public Connect As Connect

Option Explicit
Private Type ChooseColor
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    FLAGS As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long


Private Sub CancelButton_Click()
    Connect.Hide
End Sub

Private Sub Command1_Click()
Form1.Show
End Sub

Private Sub Command2_Click()
On Error Resume Next
Pic1.BackColor = Val(txtColor)
End Sub

Private Sub Command3_Click()
On Error Resume Next
Pic2.BackColor = Val(txtColor)
End Sub

Private Sub Command4_Click()
On Error Resume Next
Pic3.BackColor = Val(txtColor)
End Sub

Private Sub Command5_Click()
On Error Resume Next
Pic4.BackColor = Val(txtColor)

End Sub

Private Sub Command6_Click()
On Error Resume Next
Pic5.BackColor = Val(txtColor)

End Sub

Private Sub Command7_Click()
SaveSetting "K-ColorPicker", "Favorites", "1", Pic1.BackColor
SaveSetting "K-ColorPicker", "Favorites", "2", Pic2.BackColor
SaveSetting "K-ColorPicker", "Favorites", "3", Pic3.BackColor
SaveSetting "K-ColorPicker", "Favorites", "4", Pic4.BackColor
SaveSetting "K-ColorPicker", "Favorites", "5", Pic5.BackColor
Unload Me
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_Load()
On Error GoTo zobb
Dim a, b, c, d, e
a = GetSetting("K-ColorPicker", "Favorites", "1", Pic1.BackColor)
b = GetSetting("K-ColorPicker", "Favorites", "2", Pic2.BackColor)
c = GetSetting("K-ColorPicker", "Favorites", "3", Pic3.BackColor)
d = GetSetting("K-ColorPicker", "Favorites", "4", Pic4.BackColor)
e = GetSetting("K-ColorPicker", "Favorites", "5", Pic5.BackColor)
Pic1.BackColor = a
Pic2.BackColor = b
Pic3.BackColor = c
Pic4.BackColor = d
Pic5.BackColor = e
picColor.MouseIcon = picColorpicker.Picture
Exit Sub
zobb:
MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub OKButton_Click()
    Dim cc As ChooseColor
        Dim CustColor(16) As Long
        Dim hColor As String
        cc.lStructSize = Len(cc)
        cc.hwndOwner = frmAddIn.hWnd
        cc.hInstance = App.hInstance
        cc.FLAGS = 0
        cc.lpCustColors = String$(16 * 4, 0)
        Dim a
        Dim X
        Dim c1
        Dim c2
        Dim c3
        Dim c4
        a = ChooseColor(cc)
        Cls
        If (a) Then
        hColor = "&H" & Hex(Str$(cc.rgbResult)) & "&"
        txtColor.Text = hColor
        Clipboard.SetText txtColor
        End If
End Sub

Private Sub Pic1_Click()
txtColor = "&H" & Hex(Pic1.BackColor) & "&"
Mem
End Sub

Private Sub Pic2_Click()
txtColor = "&H" & Hex(Pic2.BackColor) & "&"
Mem
End Sub

Private Sub Pic3_Click()
txtColor = "&H" & Hex(Pic3.BackColor) & "&"
Mem
End Sub

Private Sub Pic4_Click()
txtColor = "&H" & Hex(Pic4.BackColor) & "&"
Mem
End Sub

Private Sub Pic5_Click()
txtColor = "&H" & Hex(Pic5.BackColor) & "&"
Mem
End Sub

Sub Mem()
Clipboard.SetText txtColor
End Sub

Private Sub picColor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
picColor.BackColor = GetDcColor
hColor = "&H" & Hex(picColor.BackColor) & "&"
txtColor.Text = hColor
RGBColor (GetDcColor)
txtRBG(0).Text = Val(RVal)
txtRBG(1).Text = Val(GVal)
txtRBG(2).Text = Val(BVal)
PicRGB(0).BackColor = RGB(RVal, 0, 0)
PicRGB(1).BackColor = RGB(0, GVal, 0)
PicRGB(2).BackColor = RGB(0, 0, BVal)
'Gradient picGradient, Val(RVal), Val(GVal), Val(BVal), 4
End If
End Sub

Private Sub picColorpicker_Click()

End Sub
