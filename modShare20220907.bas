Attribute VB_Name = "modShare"
Option Compare Database
Option Explicit

'*************************************************************************************************************
'-------------------------------------------------------------------------------------------------------------
    'date           contents
    '20220907       �o�^

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

'�R���s���[�^�����擾����
Private Declare PtrSafe Function GetComputerName Lib "kernel32" Alias _
"GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'���j���[��R�}���h�{�^����������Ȃ��悤�ɂ���B
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
    
    'API�֐��ɂ���ăR���s���[�^�����擾�B
    lngRet = GetComputerName(Buf, 255)
    
    lngn = InStr(1, Buf, vbNullChar)
    If lngn <> 0 Then
        'NULL�폜
        GetMyPcName = Left(Buf, lngn - 1)
    Else
        GetMyPcName = Buf
    End If

End Function

Public Function GetFileName(�g���q As String, strDir As String)
    Const ENABLE_WIZHOOK = 51488399
    Const DISABLE_WIZHOOK = 0
    Dim strFile   As String
    Dim intResult As Integer
    WizHook.KEY = ENABLE_WIZHOOK    ' WizHook �L����
    intResult = WizHook.GetFileName( _
                    0, "", "", "", strFile, strDir, _
                    �g���q & "�t�@�C�� (*." & �g���q & ")|*." & �g���q, _
                    0, 0, 0, True _
                    )
    WizHook.KEY = DISABLE_WIZHOOK    ' WizHook ������
    GetFileName = strFile
End Function

Public Function GetMultiFileName(�g���q As String, strDir As String)
    Const ENABLE_WIZHOOK = 51488399
    Const DISABLE_WIZHOOK = 0
    Dim strFile   As String
    Dim intResult As Integer
    
    WizHook.KEY = ENABLE_WIZHOOK    ' WizHook �L����
    intResult = WizHook.GetFileName( _
                    0, "", "", "", strFile, strDir, _
                    �g���q & "�t�@�C�� (*." & �g���q & ")|*." & �g���q, _
                    0, 0, 8, True _
                    )
    WizHook.KEY = DISABLE_WIZHOOK    ' WizHook ������
    GetMultiFileName = strFile
End Function

