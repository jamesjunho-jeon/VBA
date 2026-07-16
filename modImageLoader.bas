'===============================================================
' modImageLoader
'
' 셀에 입력된 "폴더경로 + 파일명"을 기준으로 이미지를 찾아
' 워크시트에 삽입/교체하는 재사용 가능한 라이브러리 모듈.
'
' [사용법]
' 1) 이 파일을 VBA 편집기(Alt+F11)에서
'    파일 > 파일 내보내기/가져오기(Import File)로 각 통합문서에 추가하거나,
'    개인용 매크로 통합 문서(Personal.xlsb) / 추가 기능(.xlam)에 넣어두면
'    여러 엑셀 파일에서 공용으로 사용할 수 있습니다.
' 2) 아무 시트에서 매크로 실행(F5) > InstallImageLoader 를 실행하면
'    "ImageConfig" 시트와 업데이트 버튼이 자동으로 생성됩니다.
' 3) ImageConfig 시트에 폴더경로 / 파일명 / 표시위치(셀주소) / 표시시트명을
'    행 단위로 입력한 뒤 "이미지 업데이트" 버튼을 누르면 이미지가 삽입됩니다.
' 4) 폴더경로나 파일명을 바꾼 뒤 다시 버튼을 누르면 해당 위치의 이미지만
'    새 파일로 교체됩니다.
' 5) 이미지를 더 추가하려면 ImageConfig 시트에 행을 추가하면 됩니다
'    (이름이 다른 이미지를 원하는 개수만큼 추가 가능).
'===============================================================
Option Explicit

Public Const IMG_CONFIG_SHEET As String = "ImageConfig"

' 지원 확장자 (필요 시 추가)
Private Function SupportedExtensions() As Variant
    SupportedExtensions = Array("jpg", "jpeg", "png", "gif", "bmp", "tif", "tiff")
End Function

'---------------------------------------------------------------
' 설치: ImageConfig 시트와 업데이트 버튼을 자동으로 만들어 준다.
' 새 통합문서에서 한 번만 실행하면 된다.
'---------------------------------------------------------------
Sub InstallImageLoader()
    Dim wb As Workbook
    Dim ws As Worksheet

    Set wb = ActiveWorkbook

    On Error Resume Next
    Set ws = wb.Worksheets(IMG_CONFIG_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = IMG_CONFIG_SHEET
    End If

    With ws
        .Range("A1").Value = "No"
        .Range("B1").Value = "폴더경로"
        .Range("C1").Value = "파일명(확장자 생략가능)"
        .Range("D1").Value = "표시위치(셀주소)"
        .Range("E1").Value = "표시시트명(비우면 이 시트)"
        .Range("F1").Value = "너비(pt, 선택)"
        .Range("G1").Value = "높이(pt, 선택)"
        .Range("H1").Value = "상태"
        .Range("A1:H1").Font.Bold = True
        .Columns("A:H").AutoFit
    End With

    If ws.Cells(2, "B").Value = "" Then
        ws.Cells(2, "A").Value = 1
        ws.Cells(2, "B").Value = "C:\Images"
        ws.Cells(2, "C").Value = "sample1"
        ws.Cells(2, "D").Value = "E2"
        ws.Cells(2, "E").Value = ws.Name
    End If

    On Error Resume Next
    ws.Buttons("btnUpdateImages").Delete
    On Error GoTo 0

    Dim btn As Button
    Set btn = ws.Buttons.Add(ws.Range("J1").Left, ws.Range("J1").Top, 130, 32)
    btn.Name = "btnUpdateImages"
    btn.OnAction = "UpdateAllImages"
    btn.Characters.Text = "이미지 업데이트"

    MsgBox "설치가 완료되었습니다." & vbNewLine & _
           "'" & IMG_CONFIG_SHEET & "' 시트에 폴더경로/파일명/표시위치를 입력한 뒤" & vbNewLine & _
           "'이미지 업데이트' 버튼을 눌러주세요.", vbInformation
End Sub

'---------------------------------------------------------------
' 메인 업데이트: ImageConfig 시트의 모든 행을 읽어 이미지를 삽입/교체한다.
' 버튼(OnAction)에 연결해서 사용.
'---------------------------------------------------------------
Sub UpdateAllImages()
    Dim wb As Workbook
    Dim wsConfig As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long, r As Long
    Dim folderPath As String, fileNameRaw As String
    Dim anchorAddr As String, targetSheetName As String
    Dim widthPt As Variant, heightPt As Variant
    Dim filePath As String
    Dim successCount As Long, failCount As Long
    Dim failList As String

    Set wb = ActiveWorkbook

    On Error Resume Next
    Set wsConfig = wb.Worksheets(IMG_CONFIG_SHEET)
    On Error GoTo 0

    If wsConfig Is Nothing Then
        MsgBox "'" & IMG_CONFIG_SHEET & "' 시트를 찾을 수 없습니다." & vbNewLine & _
               "InstallImageLoader 매크로를 먼저 실행해주세요.", vbExclamation
        Exit Sub
    End If

    lastRow = wsConfig.Cells(wsConfig.Rows.Count, "B").End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Application.ScreenUpdating = False

    For r = 2 To lastRow
        folderPath = Trim(wsConfig.Cells(r, "B").Value)
        fileNameRaw = Trim(wsConfig.Cells(r, "C").Value)
        anchorAddr = Trim(wsConfig.Cells(r, "D").Value)
        targetSheetName = Trim(wsConfig.Cells(r, "E").Value)
        widthPt = wsConfig.Cells(r, "F").Value
        heightPt = wsConfig.Cells(r, "G").Value

        If folderPath = "" Or fileNameRaw = "" Or anchorAddr = "" Then GoTo ContinueLoop

        Set wsTarget = Nothing
        If targetSheetName = "" Then
            Set wsTarget = wsConfig
        Else
            On Error Resume Next
            Set wsTarget = wb.Worksheets(targetSheetName)
            On Error GoTo 0
        End If

        If wsTarget Is Nothing Then
            failCount = failCount + 1
            failList = failList & vbNewLine & r & "행: 시트 '" & targetSheetName & "' 없음"
            wsConfig.Cells(r, "H").Value = "실패: 시트없음"
            GoTo ContinueLoop
        End If

        filePath = FindImageFile(folderPath, fileNameRaw)

        If filePath = "" Then
            failCount = failCount + 1
            failList = failList & vbNewLine & r & "행: 파일 없음 (" & fileNameRaw & ")"
            wsConfig.Cells(r, "H").Value = "실패: 파일없음"
        Else
            If PlaceImage(wsTarget, anchorAddr, filePath, widthPt, heightPt) Then
                successCount = successCount + 1
                wsConfig.Cells(r, "H").Value = "성공 " & Format(Now, "hh:nn:ss")
            Else
                failCount = failCount + 1
                failList = failList & vbNewLine & r & "행: 삽입 실패"
                wsConfig.Cells(r, "H").Value = "실패: 삽입오류"
            End If
        End If

ContinueLoop:
    Next r

    Application.ScreenUpdating = True

    Dim msg As String
    msg = "업데이트 완료" & vbNewLine & "성공: " & successCount & " / 실패: " & failCount
    If failCount > 0 Then msg = msg & vbNewLine & failList
    MsgBox msg, vbInformation
End Sub

'---------------------------------------------------------------
' 폴더 안에서 확장자를 모르는 파일명에 대해 지원 확장자를 순서대로 탐색.
' fileNameRaw 에 이미 확장자가 포함되어 있으면 그것을 우선 확인한다.
'---------------------------------------------------------------
Function FindImageFile(ByVal folderPath As String, ByVal fileNameRaw As String) As String
    Dim exts As Variant
    exts = SupportedExtensions()

    Dim i As Long
    Dim testPath As String
    Dim dotPos As Long
    Dim possibleExt As String

    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    dotPos = InStrRev(fileNameRaw, ".")
    If dotPos > 0 Then
        possibleExt = LCase(Mid(fileNameRaw, dotPos + 1))
        If IsInArray(possibleExt, exts) Then
            testPath = folderPath & fileNameRaw
            If Dir(testPath) <> "" Then
                FindImageFile = testPath
                Exit Function
            End If
        End If
    End If

    For i = LBound(exts) To UBound(exts)
        testPath = folderPath & fileNameRaw & "." & exts(i)
        If Dir(testPath) <> "" Then
            FindImageFile = testPath
            Exit Function
        End If
        testPath = folderPath & fileNameRaw & "." & UCase(exts(i))
        If Dir(testPath) <> "" Then
            FindImageFile = testPath
            Exit Function
        End If
    Next i

    FindImageFile = ""
End Function

Private Function IsInArray(ByVal val As String, ByRef arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If LCase(arr(i)) = LCase(val) Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function

'---------------------------------------------------------------
' 지정 시트의 anchorAddr(셀주소) 위치에 이미지를 삽입한다.
' 같은 위치에 이미 삽입된 이미지가 있으면 이름으로 찾아 삭제 후 재삽입(교체).
'---------------------------------------------------------------
Private Function PlaceImage(ByVal ws As Worksheet, ByVal anchorAddr As String, _
                             ByVal filePath As String, ByVal widthPt As Variant, _
                             ByVal heightPt As Variant) As Boolean
    Dim anchorCell As Range
    Dim shapeName As String
    Dim shp As Shape
    Dim pic As Shape

    On Error GoTo Fail

    Set anchorCell = ws.Range(anchorAddr)
    shapeName = "IMG_" & Replace(anchorAddr, "$", "")

    For Each shp In ws.Shapes
        If shp.Name = shapeName Then
            shp.Delete
            Exit For
        End If
    Next shp

    Set pic = ws.Shapes.AddPicture(filePath, msoFalse, msoTrue, _
        anchorCell.Left, anchorCell.Top, -1, -1)

    pic.Name = shapeName
    pic.Placement = xlMoveAndSize

    If IsNumeric(widthPt) And widthPt > 0 Then
        pic.LockAspectRatio = msoFalse
        pic.Width = CDbl(widthPt)
        If IsNumeric(heightPt) And heightPt > 0 Then
            pic.Height = CDbl(heightPt)
        Else
            pic.LockAspectRatio = msoTrue
        End If
    ElseIf IsNumeric(heightPt) And heightPt > 0 Then
        pic.LockAspectRatio = msoTrue
        pic.Height = CDbl(heightPt)
    Else
        pic.LockAspectRatio = msoFalse
        pic.Width = anchorCell.Width
        pic.Height = anchorCell.Height
    End If

    pic.Top = anchorCell.Top
    pic.Left = anchorCell.Left

    PlaceImage = True
    Exit Function

Fail:
    PlaceImage = False
End Function
