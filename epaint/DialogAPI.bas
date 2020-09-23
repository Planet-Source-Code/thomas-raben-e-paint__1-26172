Attribute VB_Name = "DialogAPI"
Option Compare Text
Option Explicit

Public Const FILTERS = "All picture files" & vbNullChar & "*.bmp;*.gif;*.jpeg;*.jpg" & vbNullChar & _
                       "Bitmap files" & vbNullChar & "*.bmp" & vbNullChar & _
                       "Jpeg files" & vbNullChar & "*.jpg;*.jpeg" & vbNullChar & _
                       "Gif files" & vbNullChar & "*.gif" & vbNullChar & _
                       "All files" & vbNullChar & "*.*"

Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800

Private Const OFN_LONGNAMES = &H200000
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_NODEREFERENCELINKS = &H100000

Private Type OPENFILENAME
    lStructSize As Long
    hInstance As Long
    hwndOwner As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustomFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type


Private Declare Function GetSaveFileName Lib "comdlg32.dll" _
       Alias "GetSaveFileNameA" (lpofn As OPENFILENAME) As Long

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias _
 "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Function GetSaveName(Optional ByVal WindowTitle As String = "Save File", _
         Optional ByVal FilterStr As String = FILTERS) As String
    Dim DlgInfo As OPENFILENAME
    Dim RET As Long
    
    With DlgInfo
        .lStructSize = Len(DlgInfo)
        .hwndOwner = 0
        .lpstrFilter = FilterStr
        .nFilterIndex = 1
        .lpstrFile = Space(512) & vbNullChar
        .nMaxFile = Len(.lpstrFile)
        .lpstrFileTitle = Space(256) & vbNullChar
        .nMaxFileTitle = Len(.lpstrFileTitle)
        .lpstrInitialDir = CurDir & vbNullChar
        .lpstrTitle = WindowTitle & vbNullChar
        .Flags = OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT
        .nMaxCustomFilter = 0
        .nFileOffset = 0
        .nFileExtension = 0
        .lCustData = 0
        .lpfnHook = 0
        .hInstance = 0
    End With
    
    RET = GetSaveFileName(DlgInfo)
    If Not RET = 0 Then
        GetSaveName = left(DlgInfo.lpstrFile, InStr(DlgInfo.lpstrFile, vbNullChar) - 1)
    Else
        GetSaveName = ""
    End If
End Function

Function GetOpenName(Optional ByVal WindowTitle As String = "Load File", _
                     Optional ByVal FILTERS As String = FILTERS, _
                     Optional ByVal DefaultFileName As String = "")
 Dim RET As Long
 Dim DlgInfo As OPENFILENAME
 
 With DlgInfo
      .lStructSize = Len(DlgInfo)
      .hwndOwner = 0
      .lpstrFilter = FILTERS
      .nFilterIndex = 1
      .lpstrFile = DefaultFileName & Space$(1024) & vbNullChar & vbNullChar
      .nMaxFile = Len(.lpstrFile)
      .lpstrDefExt = vbNullChar & vbNullChar
      .lpstrFileTitle = vbNullChar & Space$(512) & vbNullChar & vbNullChar
      .nMaxFileTitle = Len(.lpstrFileTitle)
      .lpstrInitialDir = CurDir + vbNullChar
      .lpstrTitle = WindowTitle
      .Flags = OFN_LONGNAMES Or OFN_CREATEPROMPT Or OFN_NODEREFERENCELINKS
 End With
  GetOpenName = GetOpenFileName(DlgInfo)
  GetOpenName = left(DlgInfo.lpstrFile, InStr(DlgInfo.lpstrFile, vbNullChar) - 1)
End Function

