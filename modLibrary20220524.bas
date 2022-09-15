Attribute VB_Name = "modLibrary"
Option Compare Database
Option Explicit

Dim DB900 As Database
Dim RS900 As Recordset
Dim TD As TableDef
Dim QD As QueryDef
Dim Criteria As String
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
    Dim intFlags As Integer
    
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

Public Function DelTbl(src As String)
'Table delete
        
    Set DB900 = DBEngine.Workspaces(0).Databases(0)
    
    On Error GoTo errHandler
    
    DB900.TableDefs.Refresh
    For i = 0 To DB900.TableDefs.Count - 1
        If (DB900.TableDefs(i).Name = src) Then
            DB900.TableDefs.Delete src
            Exit For
        End If
    Next
    DB900.TableDefs.Refresh
    DelTbl = True
    GoTo prcEnd
    
errHandler:
    MsgBox Error(Err)
    DelTbl = False

prcEnd:

End Function

Public Function IsExistTbl(src As String)
'Check Is Exist Table
        
    Set DB900 = DBEngine.Workspaces(0).Databases(0)
    
    On Error GoTo errHandler
    
    DB900.TableDefs.Refresh
    
    IsExistTbl = False
    For i = 0 To DB900.TableDefs.Count - 1
        If (DB900.TableDefs(i).Name = src) Then
            IsExistTbl = True
            Exit For
        End If
    Next
    GoTo prcEnd
    
errHandler:
    MsgBox Error(Err)
    IsExistTbl = False

prcEnd:

End Function


Public Function CpyTbl(src As String, dest As String)

    CpyTbl = True
    On Error GoTo errHandler
    
    DoCmd.SetWarnings False
    DoCmd.CopyObject , dest, acTable, src
    DoCmd.SetWarnings True

    GoTo prcEnd
errHandler:
    MsgBox Error(Err)
    CpyTbl = False

prcEnd:
    
End Function

Public Function DelQry(src As String)
'Query delete
       
    Set DB900 = DBEngine.Workspaces(0).Databases(0)
    
    On Error GoTo errHandler
    
    DB900.QueryDefs.Refresh
    For i = 0 To DB900.QueryDefs.Count - 1
        If (DB900.QueryDefs(i).Name = src) Then
            DB900.QueryDefs.Delete src
            Exit For
        End If
    Next
    DelQry = True
    GoTo prcEnd
    
errHandler:
    MsgBox Error(Err)
    DelQry = False

prcEnd:

End Function

Public Function GetFieldData2(src As String, Criteria As String, fld As String)
    Dim QD9 As QueryDef
    Dim MySQL As String
    
    On Error GoTo errHandler
    
    DelQry ("Q_900")
    MySQL = "SELECT " & src & ".*" _
               & " FROM " & src _
                & " WHERE " & Criteria & ";"
    Set QD9 = CurrentDb.CreateQueryDef("Q_900", MySQL)

    Set DB900 = DBEngine.Workspaces(0).Databases(0)
    Set RS900 = DB900.OpenRecordset("Q_900", dbOpenSnapshot)
    
    If RS900.RecordCount = 0 Then
        GetFieldData2 = ""
    Else
        'Renew Record
        RS900.MoveFirst
        GetFieldData2 = RS900(fld)
    End If
    GoTo prcEnd
    
errHandler:
    MsgBox Error(Err)
        GetFieldData2 = ""

prcEnd:
    RS900.Close

End Function


Public Function TransferTable2Csv(src As String, dest As String, Optional withTitle As Boolean)
    
    Dim i As Integer
    Dim Filenumber As Double
    Dim rs_list As Recordset
    Dim fieldname As String
    
    TransferTable2Csv = False
    
    On Error GoTo errHandler
    
    Set rs_list = CurrentDb.OpenRecordset(src, dbOpenDynaset)
    
    If rs_list.RecordCount > 0 Then
        
        Filenumber = FreeFile
        Open dest For Output As Filenumber
        
        If withTitle Then
            '項目名出力（テーブルフィールド名とファイルフィールド名が異なる場合戻す）
            For i = 0 To rs_list.fields.Count - 2
                fieldname = adjTablefieldname((rs_list.fields(i).Name))
                Print #1, Nz(fieldname); ",";
            Next
            fieldname = adjTablefieldname((rs_list.fields(i).Name))
            Print #1, Nz(fieldname)
        End If
        
        'データ出力
        rs_list.MoveFirst
        Do Until rs_list.EOF
            For i = 0 To rs_list.fields.Count - 2
                Print #1, Nz(rs_list.fields(i)); ",";
            Next
            Print #1, Nz(rs_list.fields(i))
        
            rs_list.MoveNext
        Loop
        Close Filenumber
    
    Else
        MsgBox ("データは0件です")
    End If
    TransferTable2Csv = True
    GoTo prcEnd

errHandler:
    MsgBox Error(Err)
    
prcEnd:

End Function


Public Function TransferTable2DQCsvWith(src As String, dest As String, header As String, trailer As String)
    
    Dim i As Integer
    Dim Filenumber As Double
    Dim rs_list As Recordset
    
    TransferTable2DQCsvWith = False
    
    If header = "" Or trailer = "" Then
        MsgBox ("ヘッダ・トレーラーを設定してください")
        GoTo prcEnd
    End If
    
    On Error GoTo errHandler
    
    Set rs_list = CurrentDb.OpenRecordset(src, dbOpenDynaset)
    
    If rs_list.RecordCount > 0 Then
        
        Filenumber = FreeFile
        Open dest For Output As Filenumber
        
        Print #1, """" & header & """"
            
        'データ出力
        rs_list.MoveFirst
        Do Until rs_list.EOF
            For i = 0 To rs_list.fields.Count - 2
                Print #1, """" & Nz(rs_list.fields(i)) & """"; ",";
            Next
            Print #1, """" & Nz(rs_list.fields(i)) & """"
        
            rs_list.MoveNext
        Loop
        
        Print #1, """" & trailer & """"
                
        Close Filenumber
    
    Else
        MsgBox ("データは0件です")
    End If
    TransferTable2DQCsvWith = True
    GoTo prcEnd

errHandler:
    MsgBox Error(Err)
    
prcEnd:

End Function

Public Function IsRecExist2(tbl As String, Criteria As String)
'nomatch ---> return False
'match   ---> return True

    Set DB900 = DBEngine.Workspaces(0).Databases(0)
    Set RS900 = DB900.OpenRecordset(tbl, dbOpenSnapshot)
    
    On Error GoTo errHandler
    If RS900.RecordCount = 0 Then
        IsRecExist2 = False
    Else
        'Renew Record
        RS900.FindFirst Criteria
        If RS900.NoMatch Then
            IsRecExist2 = False
        Else
            IsRecExist2 = True
        End If
            
    End If
    GoTo prcEnd
    
errHandler:
    MsgBox Error(Err)
    IsRecExist2 = False

prcEnd:
    RS900.Close

End Function

Public Function SetFieldData(src As String, fld As String, dat As String)
'Set All Record Data in Field
    
    Set DB900 = DBEngine.Workspaces(0).Databases(0)
    Set RS900 = DB900.OpenRecordset(src, dbOpenDynaset)
    
    On Error GoTo errHandler
    
    If RS900.RecordCount > 0 Then
        RS900.MoveFirst
        Do Until RS900.EOF
            RS900.Edit
            RS900(fld).Value = dat
            RS900.Update
            RS900.MoveNext
        Loop
    End If
    SetFieldData = True
    GoTo prcEnd
    
errHandler:
    MsgBox Error(Err)
    SetFieldData = False

prcEnd:
    RS900.Close
    
End Function

Public Function MergeTbl(src As String, dest As String)
    
    'SrcをDestにあわせる
    On Error GoTo errHandler
    Dim QD As QueryDef
    Dim MySQL As String
    
    MergeTbl = False
    DoCmd.SetWarnings False
    
    DelQry ("Q_TMP")
    MySQL = "INSERT INTO " & dest & " SELECT " & src & ".* FROM " & src & ";"
    Set QD = CurrentDb.CreateQueryDef("Q_TMP", MySQL)
    DoCmd.OpenQuery "Q_TMP", acNormal, acEdit
    
    MergeTbl = True
    GoTo prcEnd
    
errHandler:
    MsgBox Error(Err)
    MergeTbl = False

prcEnd:
    DoCmd.SetWarnings True
    
End Function


Public Function adjTablefieldname(fieldname As Variant)
    Dim strField As Variant
    
    If fieldname = "預り番号直近" Then
        strField = Nz(GetFieldData2("tableschema", "tablefieldname='" & fieldname & "'", "filefieldname"))
        If strField = "" Then
            adjTablefieldname = fieldname
        Else
            adjTablefieldname = strField
        End If
    Else
        adjTablefieldname = fieldname
    End If
    
End Function

Public Function DelTblWithAster(src As String)
'Table delete with '$' or '?'  20220727
    Dim i As Integer
    Dim tbls As DAO.TableDefs
    Dim tbl  As DAO.TableDef
    Dim DelTables As Collection
 
    On Error GoTo errHandler
    
    'テーブル一覧取得
    Set tbls = CurrentDb.TableDefs
 
    Set DelTables = New Collection
    For Each tbl In tbls
        If tbl.Name Like src Then
            DelTables.Add tbl.Name
        End If
    Next
 
    For i = 1 To DelTables.Count
        'テーブル削除
        tbls.Delete (DelTables(i))
    Next
    
    DelTblWithAster = True
    GoTo prcEnd
    
errHandler:
    MsgBox Error(Err)
    DelTblWithAster = False

prcEnd:

End Function
