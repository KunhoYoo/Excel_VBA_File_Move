Option Explicit

Sub move_Files_In_Folder()

    Dim strPath As String          '파일 가져올 폴더 경로를 넣을 변수
    Dim strTarget As String       '파일이 이동될 폴더 경로를 넣을 변수
    Dim i As Integer                   '반복구문에 사용될 변수
    Dim fileName As String       '각 파일 이름을 넣을 변수
    Dim strExt As String            '파일 확장자를 넣을 변수
    Dim rngC As Range            'A열 각 셀을 넣을 변수
    Dim msgYN As String          '메시지 박스 선택결과 넣을 변수
    
    Application.ScreenUpdating = False      '화면 업데이트 (일시) 정지
    
    '-----------------------------------------------
    ' 파일 가져올 폴더 선택하는 코드
    '-----------------------------------------------
    With Application.FileDialog(msoFileDialogFolderPicker)  '폴더선택 창에서
        .Title = "파일을 이동해올 폴더 선택"    '폴더창 타이틀
        .Show                                                    '폴더 선택창 띄우기
 
        If .SelectedItems.Count = 0 Then          '취소 선택 시
            Exit Sub                                             '매크로 종료
        Else                                                       '폴더 선택시
            strPath = .SelectedItems(1) & "\"      '폴더 경로를 변수에 넣음
        End If
    End With
    
    '-----------------------------------------------
    ' 파일 이동될 폴더 선택하는 코드
    '-----------------------------------------------
    With Application.FileDialog(msoFileDialogFolderPicker)  '폴더선택 창에서
        .Title = "파일이 이동될 폴더 선택"       '폴더창 타이틀
        .Show                                                    '폴더 선택창 띄우기
 
        If .SelectedItems.Count = 0 Then          '취소 선택 시
            Exit Sub                                             '매크로 종료
        Else                                                       '폴더 선택시
            strTarget = .SelectedItems(1) & "\"    '폴더 경로를 변수에 넣음
        End If
    End With
    
    For i = 1 To 9                                             '7회 반복
        strExt = Choose(i, "*.PCM", "*.wav", "*.jpg", "*.png", "*.xls*", "*.txt*", "*mp3", "*JPG", "*PNG")       '파일 확장자를 선택
        
        fileName = Dir(strPath & strExt)           '(폴더내)파일을 변수에 넣음
    
        Do While fileName <> ""                        '이름이 없지 않다면, 즉, 엑셀파일이 존재하면
            
            For Each rngC In Range("B1", Cells(Rows.Count, "B").End(3)) 'B열 각 셀을 순환
            
                If fileName = rngC Then                 '파일이름과 셀의 이름이 일치하면
                        
                    On Error Resume Next              '에러 발생해도 다음코드 진행
                    If FileLen(strTarget & fileName) > 0 Then  '이동될 폴더에 동일파일 존재여부 확인
                            
                        If Err <> 0 Then                       '동일한 파일이 없을 경우(에러가 난 경우)
                            Name strPath & fileName As strTarget & fileName '파일을 이동
                        Else                                         '동일한 파일이 존재할 경우
                            Err.Clear                             '발생한 에러 초기화
                            
                            '--------------------------------------------------------------------
                            ' 이동될 폴더에 동일 파일 존재할 경우의 처리코드
                            '--------------------------------------------------------------------
                            msgYN = MsgBox(fileName & " 파일이 존재합니다. 이동할까요?", 3, "파일존재")
                            
                            Select Case msgYN            '메시지박스 결과를 선택
                                Case vbYes:                    'Yes 선택 시
                                    Kill strTarget & fileName  '기존파일 삭제
                                    Name strPath & fileName As strTarget & fileName '파일을 이동
                                Case vbNo:                      'No 선택 시 덮어쓰지 않음
                                Case vbCancel: Exit Sub '취소시 메크로 종료
                            End Select
                        End If
                        On Error GoTo 0                      '에러 검출기능 복원
                        
                    End If
                End If
            Next rngC
            
            fileName = Dir                                      '다음 파일을 파일이름에 넣음
        Loop                                                          '무한 반복
    Next i
End Sub

