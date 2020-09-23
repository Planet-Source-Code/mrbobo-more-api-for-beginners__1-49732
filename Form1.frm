VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Using the API to handle File Paths"
   ClientHeight    =   9495
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9495
   ScaleWidth      =   14340
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "The Code"
      Height          =   3375
      Left            =   300
      TabIndex        =   32
      Top             =   5820
      Width           =   13695
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         Height          =   435
         Left            =   11820
         TabIndex        =   34
         Top             =   2760
         Width           =   1575
      End
      Begin RichTextLib.RichTextBox RTF 
         Height          =   2235
         Left            =   240
         TabIndex        =   33
         Top             =   420
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   3942
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         RightMargin     =   2.00000e5
         TextRTF         =   $"Form1.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "The Funtions"
      Height          =   4095
      Left            =   300
      TabIndex        =   6
      Top             =   1500
      Width           =   8955
      Begin VB.OptionButton OptFunction 
         Caption         =   "AddBackslash"
         Height          =   315
         Index           =   0
         Left            =   300
         TabIndex        =   30
         Top             =   420
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.OptionButton OptFunction 
         Caption         =   "AddExtension"
         Height          =   315
         Index           =   1
         Left            =   300
         TabIndex        =   29
         Top             =   900
         Width           =   1395
      End
      Begin VB.OptionButton OptFunction 
         Caption         =   "DriveLetterFromPath"
         Height          =   315
         Index           =   2
         Left            =   300
         TabIndex        =   28
         Top             =   1395
         Width           =   2415
      End
      Begin VB.OptionButton OptFunction 
         Caption         =   "FitPath"
         Height          =   315
         Index           =   3
         Left            =   300
         TabIndex        =   27
         Top             =   1875
         Width           =   915
      End
      Begin VB.OptionButton OptFunction 
         Caption         =   "FileExists"
         Height          =   315
         Index           =   4
         Left            =   300
         TabIndex        =   26
         Top             =   2370
         Width           =   2415
      End
      Begin VB.OptionButton OptFunction 
         Caption         =   "IsNetworkPath"
         Height          =   315
         Index           =   5
         Left            =   300
         TabIndex        =   25
         Top             =   2850
         Width           =   2415
      End
      Begin VB.OptionButton OptFunction 
         Caption         =   "IsURL"
         Height          =   315
         Index           =   6
         Left            =   300
         TabIndex        =   24
         Top             =   3330
         Width           =   2415
      End
      Begin VB.OptionButton OptFunction 
         Caption         =   "GetDriveNumber"
         Height          =   315
         Index           =   7
         Left            =   3060
         TabIndex        =   23
         Top             =   405
         Width           =   2415
      End
      Begin VB.OptionButton OptFunction 
         Caption         =   "IsFolder"
         Height          =   315
         Index           =   8
         Left            =   3060
         TabIndex        =   22
         Top             =   885
         Width           =   2415
      End
      Begin VB.OptionButton OptFunction 
         Caption         =   "IsFolderEmpty"
         Height          =   315
         Index           =   9
         Left            =   3060
         TabIndex        =   21
         Top             =   1380
         Width           =   2415
      End
      Begin VB.OptionButton OptFunction 
         Caption         =   "IsFileExtenstion"
         Height          =   315
         Index           =   10
         Left            =   3060
         TabIndex        =   20
         Top             =   1860
         Width           =   1455
      End
      Begin VB.OptionButton OptFunction 
         Caption         =   "IconNumberOnly"
         Height          =   315
         Index           =   11
         Left            =   3060
         TabIndex        =   19
         Top             =   2340
         Width           =   1575
      End
      Begin VB.OptionButton OptFunction 
         Caption         =   "QuotePath"
         Height          =   315
         Index           =   12
         Left            =   3060
         TabIndex        =   18
         Top             =   2835
         Width           =   1155
      End
      Begin VB.OptionButton OptFunction 
         Caption         =   "RemoveBackslash"
         Height          =   315
         Index           =   13
         Left            =   3060
         TabIndex        =   17
         Top             =   3315
         Width           =   1695
      End
      Begin VB.OptionButton OptFunction 
         Caption         =   "RemoveExtension"
         Height          =   315
         Index           =   14
         Left            =   6180
         TabIndex        =   16
         Top             =   390
         Width           =   1635
      End
      Begin VB.OptionButton OptFunction 
         Caption         =   "ChangeExtension"
         Height          =   315
         Index           =   15
         Left            =   6180
         TabIndex        =   15
         Top             =   870
         Width           =   1635
      End
      Begin VB.OptionButton OptFunction 
         Caption         =   "FileOnly"
         Height          =   315
         Index           =   16
         Left            =   6180
         TabIndex        =   14
         Top             =   1350
         Width           =   915
      End
      Begin VB.OptionButton OptFunction 
         Caption         =   "PathOnly"
         Height          =   315
         Index           =   17
         Left            =   6180
         TabIndex        =   13
         Top             =   1845
         Width           =   1035
      End
      Begin VB.OptionButton OptFunction 
         Caption         =   "DriveOnly"
         Height          =   315
         Index           =   18
         Left            =   6180
         TabIndex        =   12
         Top             =   2325
         Width           =   1035
      End
      Begin VB.OptionButton OptFunction 
         Caption         =   "UnQuotePath"
         Height          =   315
         Index           =   19
         Left            =   6180
         TabIndex        =   11
         Top             =   2820
         Width           =   1335
      End
      Begin VB.TextBox txtExtension 
         Appearance      =   0  'Flat
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   10
         Text            =   "txt"
         Top             =   900
         Width           =   435
      End
      Begin VB.TextBox txtExtension 
         Appearance      =   0  'Flat
         Height          =   255
         Index           =   1
         Left            =   4680
         TabIndex        =   9
         Text            =   "*.txt;*.bmp;*.f?m"
         Top             =   1860
         Width           =   1275
      End
      Begin VB.TextBox txtExtension 
         Appearance      =   0  'Flat
         Height          =   255
         Index           =   2
         Left            =   8040
         TabIndex        =   8
         Text            =   "bmp"
         Top             =   840
         Width           =   495
      End
      Begin VB.ComboBox cboChars 
         Height          =   315
         ItemData        =   "Form1.frx":0080
         Left            =   1860
         List            =   "Form1.frx":0093
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1860
         Width           =   915
      End
      Begin VB.Label Label3 
         Caption         =   "Chars."
         Height          =   255
         Left            =   1380
         TabIndex        =   31
         Top             =   1920
         Width           =   495
      End
   End
   Begin VB.ComboBox cboPath 
      Height          =   315
      ItemData        =   "Form1.frx":00AC
      Left            =   300
      List            =   "Form1.frx":00EC
      TabIndex        =   3
      Top             =   360
      Width           =   8955
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "Do it"
      Height          =   375
      Left            =   7380
      TabIndex        =   0
      Top             =   1020
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   $"Form1.frx":02CF
      Height          =   1215
      Left            =   9600
      TabIndex        =   35
      Top             =   300
      Width           =   4455
   End
   Begin VB.Image imgBobo 
      Height          =   2190
      Left            =   10440
      Picture         =   "Form1.frx":0426
      Top             =   1800
      Width           =   2760
   End
   Begin VB.Label Label1 
      Caption         =   "File path:"
      Height          =   315
      Left            =   300
      TabIndex        =   4
      Top             =   120
      Width           =   3555
   End
   Begin VB.Label lblDescript 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1395
      Left            =   9780
      TabIndex        =   2
      Top             =   4140
      Width           =   3975
   End
   Begin VB.Label lblReturn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   300
      TabIndex        =   1
      Top             =   1020
      Width           =   6855
   End
   Begin VB.Label Label2 
      Caption         =   "Returned:"
      Height          =   315
      Left            =   300
      TabIndex        =   5
      Top             =   780
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Action As Integer

Private Sub cmdCopy_Click()
    Clipboard.Clear
    Clipboard.SetText RTF.Text, vbCFText
End Sub

Private Sub cmdFunction_Click()
    Select Case Action
        Case 0 'AddBackslash
            lblReturn.Caption = AddBackSlash(cboPath.Text)
        Case 1 'AddExtension
            lblReturn.Caption = AddExtension(cboPath.Text, txtExtension(0).Text)
        Case 2 'DriveLetterFromPath
            lblReturn.Caption = DriveLetterFromPath(cboPath.Text)
        Case 3 'FitPath
            lblReturn.Caption = FitPath(cboPath.Text, CLng(cboChars.Text))
        Case 4 'FileExists
            lblReturn.Caption = FileExists(cboPath.Text)
        Case 5 'IsNetworkPath
            lblReturn.Caption = IsNetworkPath(cboPath.Text)
        Case 6 'IsURL
            lblReturn.Caption = IsURL(cboPath.Text)
        Case 7 'GetDriveNumber
            lblReturn.Caption = GetDriveNumber(cboPath.Text)
        Case 8 'IsFolder
            lblReturn.Caption = IsFolder(cboPath.Text)
        Case 9 'IsFolderEmpty
            lblReturn.Caption = IsFolderEmpty(cboPath.Text)
        Case 10 'IsFileExtenstion
            lblReturn.Caption = IsFileExtenstion(cboPath.Text, txtExtension(1).Text)
        Case 11 'IconNumberOnly
            lblReturn.Caption = IconNumberOnly(cboPath.Text)
        Case 12 'QuotePath
            lblReturn.Caption = QuotePath(cboPath.Text)
        Case 13 'RemoveBackslash
            lblReturn.Caption = RemoveBackSlash(cboPath.Text)
        Case 14 'RemoveExtension
            lblReturn.Caption = RemoveExtension(cboPath.Text)
        Case 15 'ChangeExtension
            lblReturn.Caption = ChangeExtension(cboPath.Text, txtExtension(2).Text)
        Case 16 'FileOnly
            lblReturn.Caption = FileOnly(cboPath.Text)
        Case 17 'PathOnly
            lblReturn.Caption = PathOnly(cboPath.Text)
        Case 18 'DriveOnly
            lblReturn.Caption = DriveOnly(cboPath.Text)
        Case 19 'UnQuotePath
            lblReturn.Caption = UnQuotePath(cboPath.Text)
    End Select
End Sub

Private Sub Form_Load()
    OptFunction_Click 0
    cboChars.ListIndex = 1
End Sub

Private Sub OptFunction_Click(Index As Integer)
    Action = Index
    lblDescript.Caption = LoadResString(Index + 100)
    ResourceToRTF RTF, Index + 100
    If Trim(cboPath.Text) = "" Then cboPath.ListIndex = Action
End Sub

Public Sub ResourceToRTF(mRTF As RichTextBox, CustRes As Integer)
    Dim bytResourceData() As Byte
    bytResourceData = LoadResData(CustRes, "Custom")
    mRTF.TextRTF = StrConv(bytResourceData, vbUnicode)
End Sub
