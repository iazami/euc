Attribute VB_Name = "modHoliday"
Option Compare Database
'*************************************************************************************************************
'-------------------------------------------------------------------------------------------------------------
    'date           contents
    '2011           make�j�����X�g()����
    '20190501       make�j�����X�g()�̓V�c�a�����ύX
    '20220826       make�j�����X�g()�ɎR�̓��i8/11�j��ǉ��A�̈�̓����X�|�[�c�̓��ɕύX
'-------------------------------------------------------------------------------------------------------------
'*************************************************************************************************************

Public Function make�j�����X�g(YYYY As Long)
    '�ύX����
    '2011       ����
    '2019.5.1   �V�c�a�����ύX
    '2022.8.26  �R�̓��i8/11�j��ǉ��A�̈�̓����X�|�[�c�̓��ɕύX
    
    Dim SQL As String
    
    Dim dt�O�N��A�� As Date
    Dim dt���� As Date
    Dim dt���l�̓� As Date
    Dim dt�����L�O�̓� As Date
    Dim dt�V�c�a���� As Date
    Dim dt�t���̓� As Date
    Dim dt���a�̓� As Date
    Dim dt���@�L�O�� As Date
    Dim dt�݂ǂ�̓� As Date
    Dim dt���ǂ��̓� As Date
    Dim dt�C�̓� As Date
    Dim dt�R�̓� As Date
    Dim dt�h�V�̓� As Date
    Dim dt�H���̓� As Date
    Dim dt�X�|�[�c�̓� As Date
    Dim dt�����̓� As Date
    Dim dt�ΘJ���ӂ̓� As Date
    
    Dim dt�U�֋x�� As Date
    
    Dim dt��A�� As Date
    Dim dt���N���� As Date
    
    On Error GoTo errHandler
    
    If YYYY > 1972 And YYYY < 2099 Then
        MsgBox ("����28�N���݁i�R�̓��V�݁j���߂��Ă��鍑���̏j�����瓖�N�̏j�����X�g���쐬���܂�")
        
        dt�O�N��A�� = DateSerial(YYYY - 1, 12, 31) '12/31
        dt���� = DateSerial(YYYY, 1, 1) '1/1
            '����1���̗j�������߂�
        dt���l�̓� = dtWeekday(YYYY, 1, 2, 2) '1����2���j��
        dt�����L�O�̓� = DateSerial(YYYY, 2, 11)    '2/11
        dt�t���̓� = get�t���̓�(YYYY) '���N�����V���䂪���\
        dt���a�̓� = DateSerial(YYYY, 4, 29)    '4/29
        dt���@�L�O�� = DateSerial(YYYY, 5, 3)   '5/3
        dt�݂ǂ�̓� = DateSerial(YYYY, 5, 4)   '5/4
        dt���ǂ��̓� = DateSerial(YYYY, 5, 5)   '5/5
        dt�C�̓� = dtWeekday(YYYY, 7, 3, 2)   '7����3���j��
        dt�R�̓� = DateSerial(YYYY, 8, 11)    '8/11
        dt�h�V�̓� = dtWeekday(YYYY, 9, 3, 2) '9����3���j��
        dt�H���̓� = get�H���̓�(YYYY) ''���N�����V���䂪���\
        dt�X�|�[�c�̓� = dtWeekday(YYYY, 10, 2, 2) '10����2���j��
        dt�����̓� = DateSerial(YYYY, 11, 3)    '11/3
        dt�ΘJ���ӂ̓� = DateSerial(YYYY, 11, 23)   '11/23
        
        If YYYY >= 2019 Then    '�ߘa
            dt�V�c�a���� = DateSerial(YYYY, 2, 23) '2/23
        Else                    '����
            dt�V�c�a���� = DateSerial(YYYY, 12, 23) '12/23
        End If
        
        dt��A�� = DateSerial(YYYY, 12, 31) '12/31
        dt���N���� = DateSerial(YYYY + 1, 1, 1) '��1/1
        
        i = CpyTbl("tbl�j�����^", "tbl�j��")
        SQL = "INSERT INTO tbl�j�� ( �j��, ���� ) SELECT #" & dt���� & "#, '����';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl�j�� ( �j��, ���� ) SELECT #" & dt���l�̓� & "#, '���l�̓�';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl�j�� ( �j��, ���� ) SELECT #" & dt�����L�O�̓� & "#, '�����L�O�̓�';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl�j�� ( �j��, ���� ) SELECT #" & dt�t���̓� & "#, '�t���̓�';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl�j�� ( �j��, ���� ) SELECT #" & dt���a�̓� & "#, '���a�̓�';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl�j�� ( �j��, ���� ) SELECT #" & dt���@�L�O�� & "#, '���@�L�O��';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl�j�� ( �j��, ���� ) SELECT #" & dt�݂ǂ�̓� & "#, '�݂ǂ�̓�';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl�j�� ( �j��, ���� ) SELECT #" & dt���ǂ��̓� & "#, '���ǂ��̓�';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl�j�� ( �j��, ���� ) SELECT #" & dt�C�̓� & "#, '�C�̓�';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl�j�� ( �j��, ���� ) SELECT #" & dt�R�̓� & "#, '�R�̓�';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl�j�� ( �j��, ���� ) SELECT #" & dt�h�V�̓� & "#, '�h�V�̓�';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl�j�� ( �j��, ���� ) SELECT #" & dt�H���̓� & "#, '�H���̓�';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl�j�� ( �j��, ���� ) SELECT #" & dt�X�|�[�c�̓� & "#, '�X�|�[�c�̓�';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl�j�� ( �j��, ���� ) SELECT #" & dt�����̓� & "#, '�����̓�';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl�j�� ( �j��, ���� ) SELECT #" & dt�ΘJ���ӂ̓� & "#, '�ΘJ���ӂ̓�';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl�j�� ( �j��, ���� ) SELECT #" & dt�V�c�a���� & "#, '�V�c�a����';"
        CurrentDb.Execute SQL, dbFailOnError
                
        
        '�U�֋x���̃`�F�b�N�ƃZ�b�g
        If Weekday(dt�����L�O�̓�) = 1 Then '�j�������j�Ȃ�U�֋x��
            dt�U�֋x�� = dt�����L�O�̓� + 1
            Do While IsRecExist2("tbl�j��", "�j��=#" & dt�U�֋x�� & "#")   '�U�ւ��j�������j�Ȃ玟�̓����`�F�b�N
                dt�U�֋x�� = dt�U�֋x�� + 1
            Loop
            SQL = "INSERT INTO tbl�j�� ( �j��, ���� ) SELECT #" & dt�U�֋x�� & "#, '�U�֋x��';"
            CurrentDb.Execute SQL, dbFailOnError
        End If
        
        If Weekday(dt�t���̓�) = 1 Then '�j�������j�Ȃ�U�֋x��
            dt�U�֋x�� = dt�t���̓� + 1
            Do While IsRecExist2("tbl�j��", "�j��=#" & dt�U�֋x�� & "#")   '�U�ւ��j�������j�Ȃ玟�̓����`�F�b�N
                dt�U�֋x�� = dt�U�֋x�� + 1
            Loop
            SQL = "INSERT INTO tbl�j�� ( �j��, ���� ) SELECT #" & dt�U�֋x�� & "#, '�U�֋x��';"
            CurrentDb.Execute SQL, dbFailOnError
        End If
        
        If Weekday(dt���a�̓�) = 1 Then '�j�������j�Ȃ�U�֋x��
            dt�U�֋x�� = dt���a�̓� + 1
            Do While IsRecExist2("tbl�j��", "�j��=#" & dt�U�֋x�� & "#")   '�U�ւ��j�������j�Ȃ玟�̓����`�F�b�N
                dt�U�֋x�� = dt�U�֋x�� + 1
            Loop
            SQL = "INSERT INTO tbl�j�� ( �j��, ���� ) SELECT #" & dt�U�֋x�� & "#, '�U�֋x��';"
            CurrentDb.Execute SQL, dbFailOnError
        End If
        
        If Weekday(dt���@�L�O��) = 1 Then '�j�������j�Ȃ�U�֋x��
            dt�U�֋x�� = dt���@�L�O�� + 1
            Do While IsRecExist2("tbl�j��", "�j��=#" & dt�U�֋x�� & "#")   '�U�ւ��j�������j�Ȃ玟�̓����`�F�b�N
                dt�U�֋x�� = dt�U�֋x�� + 1
            Loop
            SQL = "INSERT INTO tbl�j�� ( �j��, ���� ) SELECT #" & dt�U�֋x�� & "#, '�U�֋x��';"
            CurrentDb.Execute SQL, dbFailOnError
        End If
        
        If Weekday(dt�݂ǂ�̓�) = 1 Then '�j�������j�Ȃ�U�֋x��
            dt�U�֋x�� = dt�݂ǂ�̓� + 1
            Do While IsRecExist2("tbl�j��", "�j��=#" & dt�U�֋x�� & "#")   '�U�ւ��j�������j�Ȃ玟�̓����`�F�b�N
                dt�U�֋x�� = dt�U�֋x�� + 1
            Loop
            SQL = "INSERT INTO tbl�j�� ( �j��, ���� ) SELECT #" & dt�U�֋x�� & "#, '�U�֋x��';"
            CurrentDb.Execute SQL, dbFailOnError
        End If
       
        If Weekday(dt���ǂ��̓�) = 1 Then '�j�������j�Ȃ�U�֋x��
            dt�U�֋x�� = dt���ǂ��̓� + 1
            Do While IsRecExist2("tbl�j��", "�j��=#" & dt�U�֋x�� & "#")   '�U�ւ��j�������j�Ȃ玟�̓����`�F�b�N
                dt�U�֋x�� = dt�U�֋x�� + 1
            Loop
            SQL = "INSERT INTO tbl�j�� ( �j��, ���� ) SELECT #" & dt�U�֋x�� & "#, '�U�֋x��';"
            CurrentDb.Execute SQL, dbFailOnError
        End If

        If Weekday(dt�R�̓�) = 1 Then '�j�������j�Ȃ�U�֋x��
            dt�U�֋x�� = dt�R�̓� + 1
            Do While IsRecExist2("tbl�j��", "�j��=#" & dt�U�֋x�� & "#")   '�U�ւ��j�������j�Ȃ玟�̓����`�F�b�N
                dt�U�֋x�� = dt�U�֋x�� + 1
            Loop
            SQL = "INSERT INTO tbl�j�� ( �j��, ���� ) SELECT #" & dt�U�֋x�� & "#, '�U�֋x��';"
            CurrentDb.Execute SQL, dbFailOnError
        End If
       
        If Weekday(dt�H���̓�) = 1 Then '�j�������j�Ȃ�U�֋x��
            dt�U�֋x�� = dt�H���̓� + 1
            Do While IsRecExist2("tbl�j��", "�j��=#" & dt�U�֋x�� & "#")   '�U�ւ��j�������j�Ȃ玟�̓����`�F�b�N
                dt�U�֋x�� = dt�U�֋x�� + 1
            Loop
            SQL = "INSERT INTO tbl�j�� ( �j��, ���� ) SELECT #" & dt�U�֋x�� & "#, '�U�֋x��';"
            CurrentDb.Execute SQL, dbFailOnError
        End If
        
        If Weekday(dt�����̓�) = 1 Then '�j�������j�Ȃ�U�֋x��
            dt�U�֋x�� = dt�����̓� + 1
            Do While IsRecExist2("tbl�j��", "�j��=#" & dt�U�֋x�� & "#")   '�U�ւ��j�������j�Ȃ玟�̓����`�F�b�N
                dt�U�֋x�� = dt�U�֋x�� + 1
            Loop
            SQL = "INSERT INTO tbl�j�� ( �j��, ���� ) SELECT #" & dt�U�֋x�� & "#, '�U�֋x��';"
            CurrentDb.Execute SQL, dbFailOnError
        End If
        
        If Weekday(dt�ΘJ���ӂ̓�) = 1 Then '�j�������j�Ȃ�U�֋x��
            dt�U�֋x�� = dt�ΘJ���ӂ̓� + 1
            Do While IsRecExist2("tbl�j��", "�j��=#" & dt�U�֋x�� & "#")   '�U�ւ��j�������j�Ȃ玟�̓����`�F�b�N
                dt�U�֋x�� = dt�U�֋x�� + 1
            Loop
            SQL = "INSERT INTO tbl�j�� ( �j��, ���� ) SELECT #" & dt�U�֋x�� & "#, '�U�֋x��';"
            CurrentDb.Execute SQL, dbFailOnError
        End If
        
        If Weekday(dt�V�c�a����) = 1 Then '�j�������j�Ȃ�U�֋x��
            dt�U�֋x�� = dt�V�c�a���� + 1
            Do While IsRecExist2("tbl�j��", "�j��=#" & dt�U�֋x�� & "#")   '�U�ւ��j�������j�Ȃ玟�̓����`�F�b�N
                dt�U�֋x�� = dt�U�֋x�� + 1
            Loop
            SQL = "INSERT INTO tbl�j�� ( �j��, ���� ) SELECT #" & dt�U�֋x�� & "#, '�U�֋x��';"
            CurrentDb.Execute SQL, dbFailOnError
        End If
        
        SQL = "INSERT INTO tbl�j�� ( �j��, ���� ) SELECT #" & dt�O�N��A�� & "#, '��A��';"
        CurrentDb.Execute SQL, dbFailOnError
        
        SQL = "INSERT INTO tbl�j�� ( �j��, ���� ) SELECT #" & dt���� + 1 & "#, '����';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl�j�� ( �j��, ���� ) SELECT #" & dt���� + 2 & "#, '����';"
        CurrentDb.Execute SQL, dbFailOnError
        
        SQL = "INSERT INTO tbl�j�� ( �j��, ���� ) SELECT #" & dt��A�� & "#, '��A��';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl�j�� ( �j��, ���� ) SELECT #" & dt���N���� & "#, '���N����';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl�j�� ( �j��, ���� ) SELECT #" & dt���N���� + 1 & "#, '���N����';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl�j�� ( �j��, ���� ) SELECT #" & dt���N���� + 2 & "#, '���N����';"
        CurrentDb.Execute SQL, dbFailOnError
                
    Else
        MsgBox ("�j�����X�g�͍쐬����܂���")
    End If
    make�j�����X�g = True
    GoTo prcEnd

errHandler:
    MsgBox Error(Err)
    
prcEnd:

End Function

Public Function dtWeekday(YYYY As Long, MM As Integer, weekno As Integer, weekcode As Integer) As Date
    
    'YYYY�NMM���̑�no�T�ڂ�weekcode�j��(���j��:1�A���j��:2�A�E�E�E�y�j��:7�j�̓��t���v�Z����
    
    Dim DD As Integer           '���߂���t�̓���
    Dim dt���� As Date          '�����̓��t
    Dim weekday���� As Integer  '�����̗j���ԍ�
    
    On Error GoTo errHandler
    
    dt���� = DateSerial(YYYY, MM, 1)
    weekday���� = Weekday(dt����)   '
    If weekday���� <= weekcode Then
        DD = (weekcode - weekday����) + (weekno - 1) * 7 + 1
        dtWeekday = DateSerial(YYYY, MM, DD)
    Else
        DD = (weekcode - weekday���� + 7) + (weekno - 1) * 7 + 1
        dtWeekday = DateSerial(YYYY, MM, DD)
    End If
    
    GoTo prcEnd

errHandler:
    MsgBox Error(Err)
    
prcEnd:

End Function

Public Function get�t���̓�(YYYY As Long) As Date
    
    '����N���w�肵�āA���̔N�̏t���̓����v�Z���Ԃ��i1900-2099�j
    
    Dim intWrk As Integer
    Dim DD As Integer           '���߂���t�̓���
    Dim MM As Integer           '����
    
    On Error GoTo errHandler
    
    If YYYY < 1900 Or YYYY > 2099 Then
        MsgBox ("�t���̓��쐬�ΏۊO�̔N�x�ł�")
        GoTo prcEnd
    End If
    
    MM = 3
    
    intWrk = YYYY Mod 4
    Select Case intWrk
        Case 0
            If YYYY <= 1956 Then
                DD = 21
            ElseIf YYYY <= 2088 Then
                DD = 20
            ElseIf YYYY <= 2096 Then
                DD = 19
            End If
            
        Case 1
            If YYYY <= 1989 Then
                DD = 21
            ElseIf YYYY <= 2097 Then
                DD = 20
            End If
            
        Case 2
            If YYYY <= 2022 Then
                DD = 21
            ElseIf YYYY <= 2098 Then
                DD = 20
            End If
            
        Case 3
            If YYYY <= 1923 Then
                DD = 22
            ElseIf YYYY <= 2055 Then
                DD = 21
            ElseIf YYYY <= 2099 Then
                DD = 20
            End If
            
    End Select
    
    get�t���̓� = DateSerial(YYYY, MM, DD)
    
    GoTo prcEnd

errHandler:
    MsgBox Error(Err)
    
prcEnd:

End Function

Public Function get�H���̓�(YYYY As Long) As Date
    
    '����N���w�肵�āA���̔N�̏H���̓����v�Z���Ԃ��i1900-2099�j
    
    Dim intWrk As Integer
    Dim DD As Integer           '���߂���t�̓���
    Dim MM As Integer           '����
    
    On Error GoTo errHandler
    
    If YYYY < 1900 Or YYYY > 2099 Then
        MsgBox ("�H���̓��쐬�ΏۊO�̔N�x�ł�")
        GoTo prcEnd
    End If
    
    MM = 9
    
    intWrk = YYYY Mod 4
    Select Case intWrk
        Case 0
            If YYYY <= 2008 Then
                DD = 23
            ElseIf YYYY <= 2096 Then
                DD = 22
            End If
            
        Case 1
            If YYYY <= 1917 Then
                DD = 24
            ElseIf YYYY <= 2041 Then
                DD = 23
            ElseIf YYYY <= 2097 Then
                DD = 22
            End If
            
        Case 2
            If YYYY <= 1949 Then
                DD = 24
            ElseIf YYYY <= 2074 Then
                DD = 23
            ElseIf YYYY <= 2098 Then
                DD = 22
            End If
            
        Case 3
            If YYYY <= 1979 Then
                DD = 24
            ElseIf YYYY <= 2099 Then
                DD = 23
            End If
            
    End Select
    
    get�H���̓� = DateSerial(YYYY, MM, DD)
    
    GoTo prcEnd

errHandler:
    MsgBox Error(Err)
    
prcEnd:

End Function

Public Function �j��fromXls(xlHoliday As String)
    'xlHoliday����j���e�[�u���itbl�j��)���쐬����
    
    Dim rs As Recordset
    Dim i As Integer
    Dim SQL As String
    
    DoCmd.Hourglass True
    
    �j��fromXls = False
    If IsRecExist2(xlHoliday, True) Then
    
        Set rs = CurrentDb.OpenRecordset(xlHoliday, dbOpenDynaset)
        If Not (rs.fields(0).Name = "�j��" And rs.fields(1).Name = "����") Then
            rs.Close
'            MsgBox ("�G�N�Z���̍��ږ����قȂ�܂��i�j���A���́j�B")
            GoTo prcEnd
        End If
        
        i = CpyTbl("tbl�j�����^", "tbl�j��")
          
        On Error GoTo errHandler
        
        If rs.RecordCount() > 0 Then
            rs.MoveFirst
            Do Until rs.EOF
                SQL = "INSERT INTO tbl�j�� ( �j��, ���� ) SELECT #" & rs!�j�� & "#,'" & rs!���� & "';"
                CurrentDb.Execute SQL, dbFailOnError
                rs.MoveNext
            Loop
            �j��fromXls = True
        End If
        rs.Close
      
    End If
    
    GoTo prcEnd
    
errHandler:
    MsgBox Error(Err)
    �j��fromXls = False

prcEnd:
    DoCmd.Hourglass False
    
End Function

'-----------------------------------------------------------------------------------------------------------

Public Function get���c�Ɠ�(�W�v�� As Date)
    Dim ��� As Date
    ��� = �W�v��
    
    ��� = ��� + 1
    If Weekday(���) = 1 Then '1:���j��
        ��� = ��� + 1
    ElseIf Weekday(���) = 7 Then '7:�y�j��
        ��� = ��� + 2
    End If
    
    Do While IsRecExist2("tbl�j��", "Format(�j��, 'YYYYMMDD') =" & Format(���, "YYYYMMDD"))
        ��� = ��� + 1
        If Weekday(���) = 1 Then
            ��� = ��� + 1
        ElseIf Weekday(���) = 7 Then
            ��� = ��� + 2
        End If
    Loop
    get���c�Ɠ� = Format(���, "YYYYMMDD")

End Function

Public Function get�O�c�Ɠ�(�W�v�� As Date)
    'usage:get�O�c�Ɠ�(#2012/10/09#)
    
    Dim ��� As Date
    ��� = �W�v��
    
    ��� = ��� - 1
    If Weekday(���) = 1 Then '1:���j��
        ��� = ��� - 2
    ElseIf Weekday(���) = 7 Then '7:�y�j��
        ��� = ��� - 1
    End If
    
    Do While IsRecExist2("tbl�j��", "Format(�j��, 'YYYYMMDD') =" & Format(���, "YYYYMMDD"))
        ��� = ��� - 1
        If Weekday(���) = 1 Then
            ��� = ��� - 2
        ElseIf Weekday(���) = 7 Then
            ��� = ��� - 1
        End If
    Loop
    get�O�c�Ɠ� = Format(���, "YYYYMMDD")

End Function

