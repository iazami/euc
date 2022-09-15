Attribute VB_Name = "modShare"
Option Compare Database
Option Explicit

'*************************************************************************************************************
'-------------------------------------------------------------------------------------------------------------
    'date           contents
    '20220907       登録

'-------------------------------------------------------------------------------------------------------------
'*************************************************************************************************************

Dim DB900 As Database
Dim RS900 As Recordset
Dim i As Integer
Type BROWSEINFO
  hWndOwner As Long
  pidlRoot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfn As Long
  lParam As Long
  iImage As Long
End Type

Declare PtrSafe Function SHBrowseForFolder Lib "SHELL32" (lpbi As BROWSEINFO) As Long
Declare PtrSafe Function SHGetPathFromIDList Lib "SHELL32" (ByVal pIDL As Long, ByVal pszPath As String) As Long

'コンピュータ名を取得する
Private Declare PtrSafe Function GetComputerName Lib "kernel32" Alias _
"GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'メニューやコマンドボタンから閉じられないようにする。
Public Declare PtrSafe Function GetSystemMenu Lib "user32" _
(ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare PtrSafe Function DeleteMenu Lib "user32" _
(ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare PtrSafe Function DrawMenuBar Lib "user32" _
(ByVal hwnd As Long) As Long
Public Const MF_BYCOMMAND = &H0
Public Const SC_CLOSE = &HF060

Public Function GetMyPcName() As String
    Dim Buf As String
    Dim lngRet As Long
    Dim lngn As Long
    
    Buf = Space$(255)
    
    'API関数によってコンピュータ名を取得。
    lngRet = GetComputerName(Buf, 255)
    
    lngn = InStr(1, Buf, vbNullChar)
    If lngn <> 0 Then
        'NULL削除
        GetMyPcName = Left(Buf, lngn - 1)
    Else
        GetMyPcName = Buf
    End If

End Function

Public Function GetFileName(拡張子 As String, strDir As String)
    Const ENABLE_WIZHOOK = 51488399
    Const DISABLE_WIZHOOK = 0
    Dim strFile   As String
    Dim intResult As Integer
    WizHook.KEY = ENABLE_WIZHOOK    ' WizHook 有効化
    intResult = WizHook.GetFileName( _
                    0, "", "", "", strFile, strDir, _
                    拡張子 & "ファイル (*." & 拡張子 & ")|*." & 拡張子, _
                    0, 0, 0, True _
                    )
    WizHook.KEY = DISABLE_WIZHOOK    ' WizHook 無効化
    GetFileName = strFile
End Function

Public Function GetMultiFileName(拡張子 As String, strDir As String)
    Const ENABLE_WIZHOOK = 51488399
    Const DISABLE_WIZHOOK = 0
    Dim strFile   As String
    Dim intResult As Integer
    
    WizHook.KEY = ENABLE_WIZHOOK    ' WizHook 有効化
    intResult = WizHook.GetFileName( _
                    0, "", "", "", strFile, strDir, _
                    拡張子 & "ファイル (*." & 拡張子 & ")|*." & 拡張子, _
                    0, 0, 8, True _
                    )
    WizHook.KEY = DISABLE_WIZHOOK    ' WizHook 無効化
    GetMultiFileName = strFile
End Function

