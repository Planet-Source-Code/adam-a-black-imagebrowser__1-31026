VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Image Browser"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   ScaleHeight     =   455
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   701
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picThumb 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      Height          =   1152
      Left            =   8880
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   98
      TabIndex        =   6
      Top             =   6240
      Visible         =   0   'False
      Width           =   1536
   End
   Begin VB.PictureBox PicSrc 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   855
      Left            =   7800
      ScaleHeight     =   795
      ScaleWidth      =   1035
      TabIndex        =   5
      Top             =   6240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   7200
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Frame fraSelection 
      Caption         =   "Image Browser"
      Height          =   6375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   10095
      Begin VB.DirListBox dlb 
         Height          =   5265
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   2895
      End
      Begin VB.DriveListBox drv 
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   2895
      End
      Begin MSComctlLib.ListView flb 
         Height          =   5650
         Left            =   3240
         TabIndex        =   1
         Top             =   480
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   9975
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "imgList"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Items in List"
         Height          =   195
         Left            =   3240
         TabIndex        =   7
         Top             =   240
         Width           =   825
      End
      Begin VB.Label lblPreview 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Double click to preview"
         Height          =   195
         Left            =   6960
         TabIndex        =   4
         Top             =   840
         Width           =   1665
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const MAX_PATH = 260

Private Type FILETIME
       dwLowDateTime As Long
       dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
       dwFileAttributes As Long
       ftCreationTime As FILETIME
       ftLastAccessTime As FILETIME
       ftLastWriteTime As FILETIME
       nFileSizeHigh As Long
       nFileSizeLow As Long
       dwReserved0 As Long
       dwReserved1 As Long
       cFileName As String * MAX_PATH
       cAlternate As String * 14
End Type

Dim EnablePreview As Boolean
Dim Filename As String
Dim INIPath As String
Dim lstFilesFocus As Boolean

Dim flbList As New Collection

Private Sub dlb_Change()
    Dim i As Long
    Dim FN As String
    Dim hHeight As Double, hWidth As Double
    
    For i = flbList.Count To 1 Step -1
        flbList.Remove (i)
    Next
    
    flb.Icons = Nothing
    ImgList.ListImages.Clear
    
    flb.ListItems.Clear
    flb.Refresh
    
    GetFiles dlb.Path
    
    For i = flbList.Count To 1 Step -1
        FN = LCase$(Right$(flbList.Item(i), 3))
        If FN <> "jpg" And FN <> "bmp" And FN <> "cur" And FN <> "ico" Then
            flbList.Remove (i)
        End If
    Next
    
    For i = 1 To flbList.Count
        PicSrc.Picture = LoadPicture(flbList(i))
        
        hWidth = PicSrc.Width
        hHeight = PicSrc.Height
        
        If hHeight > 76.8 Then
            hWidth = 76.8 * PicSrc.Width / PicSrc.Height
            hHeight = 76.8
        End If
        
        If hWidth > 102.4 Then
            hHeight = 102.4 * PicSrc.Height / PicSrc.Width
            hWidth = 102.4
        End If
        
        picThumb.PaintPicture PicSrc, (picThumb.Width - hWidth) / 2, (picThumb.Height - hHeight) / 2, hWidth, hHeight
        ImgList.ListImages.Add , , picThumb.Image
        If flb.Icons Is Nothing Then flb.Icons = ImgList
        flb.ListItems.Add , , GetFileName(flbList(i)), i
        
        picThumb.Cls
        
        Caption = "GENERATING PREVIEWS  " & Format(Round(i / flbList.Count * 100, 2), "###.00") & "%"
    Next
    
    flb.Arrange = lvwAutoTop
    lblInfo.Caption = flb.ListItems.Count & " items in list"
    Caption = "Image Browser"
End Sub

Private Sub drv_Change()
    On Error GoTo Err
    dlb.Path = drv.Drive

Exit Sub
Err:
If Err.Number = 68 Then
    drv.Drive = "C:"
End If
End Sub

Private Sub GetFiles(Path As String)
   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long, fPath As String, fName As String
   Dim colFiles As Collection
   Dim varFile As Variant
   
   fPath = AddBackslash(Path)
   fName = fPath & "*.*"
   Set colFiles = New Collection
   
   hFile = FindFirstFile(fName, WFD)
   If (hFile > 0) And ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
       colFiles.Add fPath & StripNulls(WFD.cFileName)
   End If
   
   While FindNextFile(hFile, WFD)
       If ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
           colFiles.Add fPath & StripNulls(WFD.cFileName)
       End If
   Wend
   
   FindClose hFile
   
   For Each varFile In colFiles
       flbList.Add varFile
   Next
   Set colFiles = Nothing
End Sub

Private Function StripNulls(f As String) As String
   StripNulls = Left$(f, InStr(1, f, Chr$(0)) - 1)
End Function

Private Function AddBackslash(S As String) As String
   If Len(S) Then
      If Right$(S, 1) <> "\" Then
         AddBackslash = S & "\"
      Else
         AddBackslash = S
      End If
   Else
      AddBackslash = "\"
   End If
End Function

Private Function GetFileName(File As String) As String
    Dim i As Integer
    For i = Len(File) To 1 Step -1
        If Mid$(File, i, 1) = "\" Then
            i = i + 1
            Exit For
        End If
    Next
    
    GetFileName = Mid$(File, i)
End Function

Private Sub flb_DblClick()
    Dim Filename As String
    
    Filename = AddBackslash(dlb.Path)
    Filename = Filename & flb.SelectedItem
    
    ShellExecute Form1.hwnd, "", Filename, "", dlb.Path, 0
End Sub
