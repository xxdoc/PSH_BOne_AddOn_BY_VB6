VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP982"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File               : PS_PP982.cls
'//  Module             : 생산관리>풍산홀딩스애드온>생산리포트
'//  Desc               : 반별가동률 조회(개별, 집계)
'//  FormType           : PS_PP982
'//  Create Date(Start) :
'//  Create Date(End)   :
'//  Creator            : Song Myoung gyu
'//  Modified Date      :
'//  Modifier           :
'//  Company            : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01  As String
Public oForm01          As SAPbouiCOM.Form

Public oGrid01          As SAPbouiCOM.Grid
Public oGrid02          As SAPbouiCOM.Grid
Public oGrid03          As SAPbouiCOM.Grid
Public oMat01           As SAPbouiCOM.Matrix

Public oDS_PS_PP982L    As SAPbouiCOM.DataTable
Public oDS_PS_PP982M    As SAPbouiCOM.DataTable
Public oDS_PS_PP982N    As SAPbouiCOM.DataTable
Public oDS_PS_PP982O    As SAPbouiCOM.DBDataSource

Private oLastItemUID01  As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01   As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01   As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
Private CheckBoxCount   As Integer '반 체크박스 최대수량

Private oLast_Mode&

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error

    Dim i           As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01   As New MSXML2.DOMDocument
    
'    Dim ProgBar01 As SAPbouiCOM.ProgressBar
'    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("로딩 중...", 100, False)
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP982.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    
    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_PP982_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
     
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때

    oForm01.Freeze True
    
    Call PS_PP982_CreateItems
    Call PS_PP982_ComboBox_Setting
    Call PS_PP982_CF_ChooseFromList
    Call PS_PP982_FormItemEnabled
    Call PS_PP982_EnableMenus
    Call PS_PP982_SetDocument(oFromDocEntry01)
    Call PS_PP982_FormResize
    Call PS_PP982_FormReset
    
'    Call PS_PP982_Add_MatrixRow01(0, True)
    
    oForm01.EnableMenu ("1283"), False '// 삭제
    oForm01.EnableMenu ("1286"), False '// 닫기
    oForm01.EnableMenu ("1287"), False '// 복제
    oForm01.EnableMenu ("1285"), False '// 복원
    oForm01.EnableMenu ("1284"), True  '// 취소
    oForm01.EnableMenu ("1293"), False '// 행삭제
    oForm01.EnableMenu ("1281"), False
    oForm01.EnableMenu ("1282"), True

    CheckBoxCount = 12 '반 체크박스 최대 수량 초기화

    oForm01.Update
    oForm01.Freeze False
    
'    '기준일자
'    oForm01.Items("StdDt01").Specific.VALUE = Format(Now, "YYYYMMDD")
'    oForm01.Items("StdDt02").Specific.VALUE = Format(Now, "YYYYMMDD")
'    oForm01.Items("StdDt03").Specific.VALUE = Format(Now, "YYYYMMDD")
    
    oForm01.Items("Folder01").Specific.Select '폼이 로드 될 때 Folder01이 선택됨
    
    oForm01.Visible = True
    Set oXmlDoc01 = Nothing
    
'    ProgBar01.VALUE = 100
'    ProgBar01.Stop
'    Set ProgBar01 = Nothing
    
    Exit Sub
LoadForm_Error:
    oForm01.Update
    oForm01.Freeze False
    Set oXmlDoc01 = Nothing
    Set oForm01 = Nothing
    
'    ProgBar01.VALUE = 100
'    ProgBar01.Stop
'    Set ProgBar01 = Nothing
    
    Sbo_Application.SetStatusBarMessage "Form_Load Error:" & Err.Description, bmt_Short, True
End Sub

Sub PS_PP982_Add_MatrixRow01(ByVal oRow As Long, Optional RowIserted As Boolean)
'******************************************************************************
'Function ID : PS_PP982_Add_MatrixRow01()
'해당모듈    : PS_PP982
'기능        : 메트릭스 Row추가
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP982_Add_MatrixRow01_Error

    If RowIserted = False Then '//행추가여부
        oDS_PS_PP982O.InsertRecord (oRow)
    End If

    oMat01.AddRow
    oDS_PS_PP982O.offset = oRow
    oDS_PS_PP982O.setValue "U_LineNum", oRow, oRow + 1

    oMat01.LoadFromDataSource
    Exit Sub

PS_PP982_Add_MatrixRow01_Error:
    MDC_Com.MDC_GF_Message "PS_PP982_Add_MatrixRow01_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub PS_PP982_SelectGrid01()
'******************************************************************************
'Function ID : PS_PP982_SelectGrid01()
'해당모듈    : PS_PP982
'기능        : 반별자료 조회
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP982_SelectGrid01_Error

    Call oForm01.Freeze(True)
    
    Dim loopCount As Integer
    Dim Query01   As String
    Dim ErrNum    As String
    
'    Dim RecordSet01 As SAPbobsCOM.Recordset
'    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim BPLID      As String
    Dim FrDt       As String
    Dim ToDt       As String
    Dim CntcCode   As String
    Dim ClsCodeChk As String
    Dim CheckBoxID As String
    
    BPLID = Trim(oForm01.Items("BPLID").Specific.VALUE)
    FrDt = Trim(oForm01.Items("FrDt").Specific.VALUE)
    ToDt = Trim(oForm01.Items("ToDt").Specific.VALUE)
    CntcCode = Trim(oForm01.Items("CntcCode").Specific.VALUE)
    ClsCodeChk = ""
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    For loopCount = 1 To CheckBoxCount
    
        If loopCount >= 1 And loopCount < 10 Then
            CheckBoxID = "Check0" & loopCount
        Else
            CheckBoxID = "Check" & loopCount
        End If
    
        If oForm01.DataSources.UserDataSources.Item(CheckBoxID).VALUE = "Y" Then
    
            ClsCodeChk = ClsCodeChk & Left(Right(MDC_PS_Common.Split_String(oForm01.Items(CheckBoxID).Specific.Caption, "-", 1), 5), 4) & ","
        
        End If
    
    Next
        
    Query01 = "         EXEC PS_PP982_01 '"
    Query01 = Query01 & BPLID & "','"
    Query01 = Query01 & FrDt & "','"
    Query01 = Query01 & ToDt & "','"
    Query01 = Query01 & CntcCode & "','"
    Query01 = Query01 & ClsCodeChk & "'"
    
    Call oGrid01.DataTable.Clear
    Call oDS_PS_PP982L.ExecuteQuery(Query01)
'    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("DataTable")

    oGrid01.Columns(7).RightJustified = True
    oGrid01.Columns(8).RightJustified = True
    oGrid01.Columns(9).RightJustified = True
    oGrid01.Columns(10).RightJustified = True
    oGrid01.Columns(11).RightJustified = True
    oGrid01.Columns(12).RightJustified = True
    
'    oGrid01.Columns(12).BackColor = RGB(255, 255, 125) '[결산]계, 노랑
'    oGrid01.Columns(19).BackColor = RGB(255, 255, 125) '[계산]계, 노랑
'    oGrid01.Columns(26).BackColor = RGB(255, 255, 125) '[완료]계, 노랑
    
'    oGrid01.Columns(9).BackColor = RGB(255, 255, 125) '품의일, 노랑
'    oGrid01.Columns(10).BackColor = RGB(255, 255, 125) '가입고일, 노랑
'    oGrid01.Columns(11).BackColor = RGB(0, 210, 255) '차이(품의-가입고), 하늘
'    oGrid01.Columns(12).BackColor = RGB(255, 255, 125) '검수입고일, 노랑
'    oGrid01.Columns(13).BackColor = RGB(0, 210, 255) '차이(가입고-품의), 하늘
'    oGrid01.Columns(14).BackColor = RGB(255, 167, 167) '총소요일, 빨강
        
    If oGrid01.Rows.Count = 1 Then
        ErrNum = 1
        GoTo PS_PP982_SelectGrid01_Error:
    End If
        
    oGrid01.AutoResizeColumns
    oForm01.Update
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    'Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP982_SelectGrid01_Exit:
    'Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP982_SelectGrid01_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
    Else
        Sbo_Application.SetStatusBarMessage "PS_PP982_SelectGrid01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
    
    'Set RecordSet01 = Nothing
    
    Call oForm01.Freeze(False)
    
End Sub

Public Sub PS_PP982_SelectGrid02()
'******************************************************************************
'Function ID : PS_PP982_SelectGrid02()
'해당모듈    : PS_PP982
'기능        : 집계자료(주별)
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP982_SelectGrid02_Error

    '//메트릭스에 데이터 로드
    Call oForm01.Freeze(True)
    
    Dim loopCount As Integer
    Dim Query01   As String
    Dim ErrNum    As String
    
'    Dim RecordSet01 As SAPbobsCOM.Recordset
'    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLID      As String
    Dim FrDt       As String
    Dim ToDt       As String
    Dim CntcCode   As String
    Dim ClsCodeChk As String
    Dim CheckBoxID As String
    
    BPLID = Trim(oForm01.Items("BPLID").Specific.VALUE)
    FrDt = Trim(oForm01.Items("FrDt").Specific.VALUE)
    ToDt = Trim(oForm01.Items("ToDt").Specific.VALUE)
    CntcCode = Trim(oForm01.Items("CntcCode").Specific.VALUE)
    ClsCodeChk = ""
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    For loopCount = 1 To CheckBoxCount
    
        If loopCount >= 1 And loopCount < 10 Then
            CheckBoxID = "Check0" & loopCount
        Else
            CheckBoxID = "Check" & loopCount
        End If
    
        If oForm01.DataSources.UserDataSources.Item(CheckBoxID).VALUE = "Y" Then
    
            ClsCodeChk = ClsCodeChk & Left(Right(MDC_PS_Common.Split_String(oForm01.Items(CheckBoxID).Specific.Caption, "-", 1), 5), 4) & ","
        
        End If
    
    Next
        
    Query01 = "         EXEC PS_PP982_02 '"
    Query01 = Query01 & BPLID & "','"
    Query01 = Query01 & FrDt & "','"
    Query01 = Query01 & ToDt & "','"
    Query01 = Query01 & CntcCode & "','"
    Query01 = Query01 & ClsCodeChk & "'"
    
    Call oGrid02.DataTable.Clear
    Call oDS_PS_PP982M.ExecuteQuery(Query01)
'    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("DataTable")

    oGrid02.Columns(2).RightJustified = True
    oGrid02.Columns(3).RightJustified = True
    oGrid02.Columns(4).RightJustified = True
    oGrid02.Columns(5).RightJustified = True
    oGrid02.Columns(6).RightJustified = True
    oGrid02.Columns(7).RightJustified = True
    oGrid02.Columns(8).RightJustified = True
'    oGrid02.Columns(13).RightJustified = True
'    oGrid02.Columns(14).RightJustified = True
'    oGrid02.Columns(15).RightJustified = True
'    oGrid02.Columns(16).RightJustified = True
'    oGrid02.Columns(17).RightJustified = True

'    oGrid01.Columns(12).BackColor = RGB(255, 255, 125) '[결산]계, 노랑
'    oGrid01.Columns(19).BackColor = RGB(255, 255, 125) '[계산]계, 노랑
'    oGrid01.Columns(26).BackColor = RGB(255, 255, 125) '[완료]계, 노랑
    
'    oGrid01.Columns(9).BackColor = RGB(255, 255, 125) '품의일, 노랑
'    oGrid01.Columns(10).BackColor = RGB(255, 255, 125) '가입고일, 노랑
'    oGrid01.Columns(11).BackColor = RGB(0, 210, 255) '차이(품의-가입고), 하늘
'    oGrid01.Columns(12).BackColor = RGB(255, 255, 125) '검수입고일, 노랑
'    oGrid01.Columns(13).BackColor = RGB(0, 210, 255) '차이(가입고-품의), 하늘
'    oGrid01.Columns(14).BackColor = RGB(255, 167, 167) '총소요일, 빨강
        
    If oGrid02.Rows.Count = 1 Then
        ErrNum = 1
        GoTo PS_PP982_SelectGrid02_Error:
    End If
        
    oGrid02.AutoResizeColumns
    oForm01.Update
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
'    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP982_SelectGrid02_Exit:
'    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP982_SelectGrid02_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    If ErrNum = 1 Then
    
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
    
    Else

        Sbo_Application.SetStatusBarMessage "PS_PP982_SelectGrid02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    
    End If
    
'    Set RecordSet01 = Nothing
    
    Call oForm01.Freeze(False)
    
End Sub

Public Sub PS_PP982_SelectMatrix01()
'******************************************************************************
'Function ID : PS_PP982_SelectMatrix01
'해당모듈    : PS_PP982
'기능        : 예외작번 관리 조회
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP982_SelectMatrix01_Error

    Dim i       As Integer
    Dim sQry    As String
    
    Dim ErrNum  As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLID     As String '사업장
    
    BPLID = Trim(oForm01.Items("BPLID").Specific.VALUE) '사업장
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    sQry = "EXEC [PS_PP982_04] '"
    sQry = sQry & BPLID & "'"
    
    Call oRecordSet01.DoQuery(sQry)
    
    oMat01.Clear
    oDS_PS_PP982O.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    If (oRecordSet01.RecordCount = 0) Then
    
        ErrNum = 1
        
        oForm01.Mode = fm_ADD_MODE
        
        Call PS_PP982_Add_MatrixRow01(0, True)
        GoTo PS_PP982_SelectMatrix01_Error
        
        Exit Sub
    End If
    
    For i = 0 To oRecordSet01.RecordCount - 1
    
        If i + 1 > oDS_PS_PP982O.Size Then
            oDS_PS_PP982O.InsertRecord (i)
        End If
    
        oMat01.AddRow
        oDS_PS_PP982O.offset = i
    
        Call oDS_PS_PP982O.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_PP982O.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("Select").VALUE)) '선택
        Call oDS_PS_PP982O.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("CntcCode").VALUE)) '사번
        Call oDS_PS_PP982O.setValue("U_ColReg03", i, Trim(oRecordSet01.Fields("CntcName").VALUE)) '성명
        Call oDS_PS_PP982O.setValue("U_ColReg04", i, Trim(oRecordSet01.Fields("TeamName").VALUE)) '팀
        Call oDS_PS_PP982O.setValue("U_ColReg05", i, Trim(oRecordSet01.Fields("RspName").VALUE)) '담당
        Call oDS_PS_PP982O.setValue("U_ColReg06", i, Trim(oRecordSet01.Fields("ClsName").VALUE)) '반
        Call oDS_PS_PP982O.setValue("U_ColReg07", i, Trim(oRecordSet01.Fields("Comments").VALUE)) '비고
        
        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
        
    Next
    
    Call PS_PP982_Add_MatrixRow01(oMat01.VisualRowCount)
    
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Call oForm01.Freeze(False)
  
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub

PS_PP982_SelectMatrix01_Error:

    Call oForm01.Freeze(False)
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_PP982_SelectMatrix01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Function PS_PP982_InsertMatrix01() As Boolean
'******************************************************************************
'Function ID : PS_PP982_InsertMatrix01()
'해당모듈    : PS_PP982
'기능        : 저장
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP982_InsertMatrix01_Error

    Dim loopCount As Integer
    Dim sQry      As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLID       As String '사업장
    Dim CntcCode    As String '사번
    Dim CntcName    As String '성명
    Dim Comments    As String '비고
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("저장 중...", oMat01.VisualRowCount - 1, False)
    
    BPLID = Trim(oForm01.Items("BPLID").Specific.VALUE)
    
    Call oMat01.FlushToDataSource
    For loopCount = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PS_PP982O.GetValue("U_ColReg01", loopCount)) = "Y" Then
            
            CntcCode = Trim(oDS_PS_PP982O.GetValue("U_ColReg02", loopCount)) '사번
            CntcName = Trim(oDS_PS_PP982O.GetValue("U_ColReg03", loopCount)) '성명
            Comments = Trim(oDS_PS_PP982O.GetValue("U_ColReg07", loopCount)) '비고
            
            sQry = "      EXEC [PS_PP982_05] '"
            sQry = sQry & BPLID & "','" '사업장
            sQry = sQry & CntcCode & "','" '사번
            sQry = sQry & CntcName & "','" '성명
            sQry = sQry & Comments & "'" '비고
            
            Call RecordSet01.DoQuery(sQry)
            
            ProgBar01.VALUE = ProgBar01.VALUE + 1
            ProgBar01.Text = ProgBar01.VALUE & "/" & oMat01.VisualRowCount - 1 & "건 저장중..."
            
        End If
    Next
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    MDC_Com.MDC_GF_Message "저장 완료!", "S"
    
    Set RecordSet01 = Nothing
    
    PS_PP982_InsertMatrix01 = True

    Exit Function

PS_PP982_InsertMatrix01_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    PS_PP982_InsertMatrix01 = False
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PS_PP982_InsertMatrix01_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Sub PS_PP982_DeleteMatrix01()
'******************************************************************************
'Function ID : PS_PP982_DeleteMatrix01()
'해당모듈    : PS_PP982
'기능        : 삭제
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP982_DeleteMatrix01_Error

    Dim loopCount As Integer
    Dim sQry      As String
    Dim ErrNum    As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLID As String
    Dim CntcCode As String
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("삭제 중...", 100, False)
    
    BPLID = Trim(oForm01.Items("BPLID").Specific.VALUE)
    
    Call oMat01.FlushToDataSource
    
    For loopCount = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PS_PP982O.GetValue("U_ColReg01", loopCount)) = "Y" Then
            
            CntcCode = Trim(oDS_PS_PP982O.GetValue("U_ColReg02", loopCount))
            
            sQry = "      EXEC [PS_PP982_06] '"
            sQry = sQry & BPLID & "','"
            sQry = sQry & CntcCode & "'"
            
            Call oRecordSet01.DoQuery(sQry)
        End If
    Next
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Call MDC_Com.MDC_GF_Message("삭제 완료!", "S")
    
    Exit Sub

PS_PP982_DeleteMatrix01_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("삭제대상이 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_PP982_DeleteMatrix01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
    
End Sub

Private Function PS_PP982_HeaderSpaceLineDel02() As Boolean
'******************************************************************************
'Function ID : PS_PP982_HeaderSpaceLineDel02()
'해당모듈    : PS_PP982
'기능        : 필수입력사항 체크
'인수        : 없음
'반환값      : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP982_HeaderSpaceLineDel02_Error

    Dim ErrNum As Integer
    ErrNum = 0

    Select Case True
'        Case Trim(oForm01.Items("StdYear").Specific.VALUE) = "" '기준년도
'            ErrNum = 1
'            GoTo PS_PP982_HeaderSpaceLineDel02_Error
    End Select

    PS_PP982_HeaderSpaceLineDel02 = True
Exit Function

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP982_HeaderSpaceLineDel02_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "기준년도는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("FrDate").Click(ct_Regular)
    End If
    PS_PP982_HeaderSpaceLineDel02 = False
End Function

''/ 메트릭스 필수 사항 check
'Private Function PS_PP982_MatrixSpaceLineDel() As Boolean
'On Error GoTo PS_PP982_MatrixSpaceLineDel_Error
'
'    Dim i As Long
'    Dim ErrNum As Integer
'    Dim oRecordSet01 As SAPbobsCOM.Recordset
'    Dim sQry As String
'
'    Set oRecordSet01 = Nothing
'    PS_PP982_MatrixSpaceLineDel = True
'Exit Function
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'PS_PP982_MatrixSpaceLineDel_Error:
'    Set oRecordSet01 = Nothing
'    If ErrNum = 1 Then
'        MDC_Com.MDC_GF_Message "라인 데이터가 없습니다. 확인하세요.", "E"
'    ElseIf ErrNum = 2 Then
'        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 사원코드가 없습니다. 확인하세요.", "E"
'    ElseIf ErrNum = 3 Then
'        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 시간이 없습니다. 확인하세요.", "E"
'    ElseIf ErrNum = 4 Then
'        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 등록일자가 없습니다. 확인하세요.", "E"
'    ElseIf ErrNum = 5 Then
'        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 비가동코드가 없습니다. 확인하세요.", "E"
'    Else
'        MDC_Com.MDC_GF_Message "PS_PP982_MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
'    End If
'    PS_PP982_MatrixSpaceLineDel = False
'End Function

Private Sub PS_PP982_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PS_PP982_FlushToItemValue_Error

    Dim loopCount01 As Integer
    Dim loopCount02 As Integer
    Dim ErrNum      As Integer
    Dim sQry        As String
    Dim BPLID       As String
    Dim CheckBoxID  As String
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
    
        Case "Mat01"
        
            Call oMat01.FlushToDataSource
        
            If oCol = "CntcCode" Then
            
                sQry = "        SELECT      T0.Code AS [CntcCode]," '사번
                sQry = sQry & "             T0.U_FullName AS [CntcName]," '성명
                sQry = sQry & "             T1.U_CodeNm AS [TeamName]," '팀
                sQry = sQry & "             T2.U_CodeNm AS [RspName]," '담당
                sQry = sQry & "             T3.U_CodeNm AS [ClsName]" '반
                sQry = sQry & " FROM        [@PH_PY001A] AS T0"
                sQry = sQry & "             LEFT JOIN"
                sQry = sQry & "             [@PS_HR200L] AS T1"
                sQry = sQry & "                 ON T0.U_TeamCode = T1.U_Code"
                sQry = sQry & "                 AND T1.Code = '1'"
                sQry = sQry & "             LEFT JOIN"
                sQry = sQry & "             [@PS_HR200L] AS T2"
                sQry = sQry & "                 ON T0.U_RspCode = T2.U_Code"
                sQry = sQry & "                 AND T2.Code = '2'"
                sQry = sQry & "             LEFT JOIN"
                sQry = sQry & "             [@PS_HR200L] AS T3"
                sQry = sQry & "                 ON T0.U_ClsCode = T3.U_Code"
                sQry = sQry & "                 AND T3.Code = '9'"
                sQry = sQry & " WHERE       T0.Code = '" & oDS_PS_PP982O.GetValue("U_ColReg02", oRow - 1) & "'"
                
                Call oRecordSet01.DoQuery(sQry)

                Call oDS_PS_PP982O.setValue("U_ColReg01", oRow - 1, "Y") '선택
                Call oDS_PS_PP982O.setValue("U_ColReg02", oRow - 1, oRecordSet01.Fields("CntcCode").VALUE) '사번
                Call oDS_PS_PP982O.setValue("U_ColReg03", oRow - 1, oRecordSet01.Fields("CntcName").VALUE) '성명
                Call oDS_PS_PP982O.setValue("U_ColReg04", oRow - 1, oRecordSet01.Fields("TeamName").VALUE) '팀
                Call oDS_PS_PP982O.setValue("U_ColReg05", oRow - 1, oRecordSet01.Fields("RspName").VALUE) '담당
                Call oDS_PS_PP982O.setValue("U_ColReg06", oRow - 1, oRecordSet01.Fields("ClsName").VALUE) '반
                
                Call oMat01.LoadFromDataSource
                
                If oForm01.Mode = fm_ADD_MODE Then
                    If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 0) And Trim(oMat01.Columns("CntcCode").Cells(oRow).Specific.VALUE) <> "" Then
                        Call PS_PP982_Add_MatrixRow01(oMat01.RowCount, False)
                    End If
                End If
            
            End If
    
            Call oMat01.AutoResizeColumns
    
        Case "BPLID"
            
            BPLID = Trim(oForm01.Items("BPLID").Specific.VALUE)

            '반 체크박스 초기화_S
            For loopCount01 = 1 To CheckBoxCount
            
                If loopCount01 >= 1 And loopCount01 < 10 Then
                    CheckBoxID = "Check0" & loopCount01
                Else
                    CheckBoxID = "Check" & loopCount01
                End If
            
                oForm01.Items(CheckBoxID).Specific.Caption = "정보 없음"
            
            Next
            '반 체크박스 초기화_E

            '반 체크박스 제목 바인딩_S
            sQry = "        SELECT      U_CodeNm + '[' + U_Code +']-' + U_Comment1 + '[' + U_Char1 + ']-' + U_Comment2 + '[' + U_Char2 + ']' AS [FullClsName]"
            sQry = sQry & " FROM        [@PS_HR200L]"
            sQry = sQry & " WHERE       Code = '9'"
            sQry = sQry & "             AND U_Char3 = '" & BPLID & "'"
            sQry = sQry & "             AND U_UseYN = 'Y'"
            sQry = sQry & " ORDER BY    U_Seq"
            
            Call oRecordSet01.DoQuery(sQry)
            
            For loopCount02 = 1 To oRecordSet01.RecordCount
            
                If loopCount02 >= 1 And loopCount02 < 10 Then
                    CheckBoxID = "Check0" & loopCount02
                Else
                    CheckBoxID = "Check" & loopCount02
                End If
            
                oForm01.Items(CheckBoxID).Specific.Caption = oRecordSet01.Fields("FullClsName").VALUE
            
                oRecordSet01.MoveNext
            
            Next
            '반 체크박스 제목 바인딩_E
            
            
        Case "CntcCode"
            
            'oForm01.Items("CntcName").Specific.VALUE = MDC_GetData.Get_ReData("U_FULLNAME", "U_MSTCOD", "[OHEM]", "'" & Trim(oForm01.Items("CntcCode").Specific.VALUE) & "'") '성명
            oForm01.DataSources.UserDataSources("CntcName").VALUE = MDC_GetData.Get_ReData("U_FULLNAME", "U_MSTCOD", "[OHEM]", "'" & Trim(oForm01.Items("CntcCode").Specific.VALUE) & "'") '성명
            
    End Select

    Set oRecordSet01 = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
PS_PP982_FlushToItemValue_Error:

    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
'        Call MDC_Com.MDC_GF_Message("시각은 숫자만 입력이 가능합니다.", "E")
'    ElseIf ErrNum = 2 Then
'        Call MDC_Com.MDC_GF_Message("시각(시)는 24미만의 값만 입력이 가능합니다.", "E")
'    ElseIf ErrNum = 3 Then
'        Call MDC_Com.MDC_GF_Message("시각(분)은 60미만의 값만 입력이 가능합니다.", "E")
    Else
        Call MDC_Com.MDC_GF_Message("PS_PP982_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

'/폼의 아이템 사용지정
Sub PS_PP982_FormItemEnabled()
On Error GoTo PS_PP982_FormItemEnabled_Error
    If (oForm01.Mode = fm_ADD_MODE) Then
'        oMat02.Columns("ItemCode").Cells(1).Click ct_Regular
'        oForm01.Items("ItemCode").Enabled = True

'        Call oForm01.Items("BPLID02").Specific.Select(MDC_PS_Common.User_BPLId, psk_ByValue)
'        Call PS_PP982_FlushToItemValue("BPLID02") '팀, 담당, 반 콤보박스 강제 설정

    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
'        oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP982_FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "PS_PP982_FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/아이템 변경 이벤트
Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_ItemEvent_Error

    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            Call Raise_EVENT_ITEM_PRESSED(FormUID, pval, BubbleEvent)
        Case et_KEY_DOWN: '//2
            Call Raise_EVENT_KEY_DOWN(FormUID, pval, BubbleEvent)
        Case et_COMBO_SELECT: '//5
            Call Raise_EVENT_COMBO_SELECT(FormUID, pval, BubbleEvent)
        Case et_CLICK: '//6
            Call Raise_EVENT_CLICK(FormUID, pval, BubbleEvent)
        Case et_DOUBLE_CLICK: '//7
            Call Raise_EVENT_DOUBLE_CLICK(FormUID, pval, BubbleEvent)
        Case et_MATRIX_LINK_PRESSED '//8
            Call Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, pval, BubbleEvent)
        Case et_VALIDATE: '//10
            Call Raise_EVENT_VALIDATE(FormUID, pval, BubbleEvent)
        Case et_MATRIX_LOAD: '//11
            Call Raise_EVENT_MATRIX_LOAD(FormUID, pval, BubbleEvent)
        Case et_FORM_ACTIVATE: '//18
            '//et_FORM_ACTIVATE
        Case et_FORM_DEACTIVATE: '//19
            '//et_FORM_DEACTIVATE
        Case et_FORM_RESIZE '//20
            Call Raise_EVENT_RESIZE(FormUID, pval, BubbleEvent)
        Case et_CHOOSE_FROM_LIST '//27
            Call Raise_EVENT_CHOOSE_FROM_LIST(FormUID, pval, BubbleEvent)
        Case et_GOT_FOCUS: '//3
            Call Raise_EVENT_GOT_FOCUS(FormUID, pval, BubbleEvent)
        Case et_LOST_FOCUS: '//4
            '//et_LOST_FOCUS
        Case et_FORM_UNLOAD: '//17
            Call Raise_EVENT_FORM_UNLOAD(FormUID, pval, BubbleEvent)
    End Select
    Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Raise_ItemEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_ItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_MenuEvent_Error

    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
            Case "1282": '추가
                '/추가버튼 클릭시 메트릭스 insertrow
                
'                oMat02.Clear
'                oMat02.FlushToDataSource
'                oMat02.LoadFromDataSource
                
'                oForm01.Mode = fm_ADD_MODE
'                BubbleEvent = False
                
                'oForm01.Items("GCode").Click ct_Regular
                
                
                Exit Sub
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
            
            Case "7169": '엑셀 내보내기
            
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
                '//Call PS_PP982_FormItemEnabled '//UDO방식
            Case "1282": '추가
'                oMat02.Clear
'                oDS_PS_PP982H.Clear

'                Call PS_PP982_LoadCaption
'                Call PS_PP982_FormItemEnabled
                '//Call PS_PP982_FormItemEnabled '//UDO방식
                '//Call PS_PP982_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PS_PP982_FormItemEnabled
                
            Case "7169": '엑셀 내보내기
            
        End Select
    End If
    Exit Sub
Raise_MenuEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_MenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_FormDataEvent(ByRef FormUID As String, ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_FormDataEvent_Error
    If (BusinessObjectInfo.BeforeAction = True) Then '//BeforeAction = True
        Select Case BusinessObjectInfo.EventType
            Case et_FORM_DATA_LOAD: '//33
            Case et_FORM_DATA_ADD: '//34
            Case et_FORM_DATA_UPDATE: '//35
            Case et_FORM_DATA_DELETE: '//36
        End Select
    ElseIf (BusinessObjectInfo.BeforeAction = False) Then '//BeforeAction = False
        Select Case BusinessObjectInfo.EventType
            Case et_FORM_DATA_LOAD: '//33
            Case et_FORM_DATA_ADD: '//34
            Case et_FORM_DATA_UPDATE: '//35
            Case et_FORM_DATA_DELETE: '//36
        End Select
    End If
    Exit Sub
Raise_FormDataEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_FormDataEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_RightClickEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_RightClickEvent_Error
    If pval.BeforeAction = True Then
    ElseIf pval.BeforeAction = False Then
    End If
    If pval.ItemUID = "Mat01" Then
        If pval.Row > 0 Then
            oLastItemUID01 = pval.ItemUID
            oLastColUID01 = pval.ColUID
            oLastColRow01 = pval.Row
            
        End If
    Else
        oLastItemUID01 = pval.ItemUID
        oLastColUID01 = ""
        oLastColRow01 = 0
    End If
    Exit Sub
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_ITEM_PRESSED(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_ITEM_PRESSED_Error

    If pval.BeforeAction = True Then
    
        If pval.ItemUID = "PS_PP982" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "BtnSrch01" Then '/조회
            If PS_PP982_HeaderSpaceLineDel02 = False Then
                BubbleEvent = False
                Exit Sub
            Else
                Call PS_PP982_SelectGrid01 '반별자료
                Call PS_PP982_SelectGrid02 '집계자료(주별)
'                Call PS_PP982_SelectGrid03 '집계자료(일별)
            End If
            
        ElseIf pval.ItemUID = "BtnPrt01" Then '출력
            
            Call PS_PP982_Print_Report01
            
        End If
        
        If pval.ItemUID = "BtnSrch04" Then
            
            Call PS_PP982_SelectMatrix01
            
        ElseIf pval.ItemUID = "BtnSave04" Then
        
            Call PS_PP982_InsertMatrix01
            Call PS_PP982_SelectMatrix01
        
        ElseIf pval.ItemUID = "BtnDel04" Then
        
            If Sbo_Application.MessageBox("삭제후 복구는 불가능합니다. 삭제하시겠습니까?", "1", "예", "아니오") = "1" Then
        
                Call PS_PP982_DeleteMatrix01
                Call PS_PP982_SelectMatrix01
            
            End If
        
        End If
        
        If pval.ItemUID = "BtnAll" Then
        
            Call PS_PP982_CheckAll
        
        End If
  
    ElseIf pval.BeforeAction = False Then
    
        '폴더를 사용할 때는 필수 소스_S
        'Folder01이 선택되었을 때
        If pval.ItemUID = "Folder01" Then
            
            oForm01.PaneLevel = 1
            
        End If
        
        'Folder02가 선택되었을 때
        If pval.ItemUID = "Folder02" Then
            
            oForm01.PaneLevel = 2
            
        End If
        
        'Folder03가 선택되었을 때
        If pval.ItemUID = "Folder03" Then
            
            oForm01.PaneLevel = 3
                        
        End If
        
        'Folder04가 선택되었을 때
        If pval.ItemUID = "Folder04" Then
            
            oForm01.PaneLevel = 4
                        
        End If
        '폴더를 사용할 때는 필수 소스_E
    
        If pval.ItemUID = "PS_PP982" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
    End If
    
    Exit Sub
    
Raise_EVENT_ITEM_PRESSED_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ITEM_PRESSED_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_KEY_DOWN(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_KEY_DOWN_Error
    If pval.BeforeAction = True Then
    
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CntcCode", "") '헤더_조회조건-사번
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "CntcCode") '매트릭스-사번
        
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_KEY_DOWN_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_KEY_DOWN_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CLICK(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CLICK_Error

    If pval.BeforeAction = True Then
    
        If pval.ItemUID = "Mat01" Then
            If pval.Row > 0 Then
                
                Call oMat01.SelectRow(pval.Row, True, False)
                
                oLastItemUID01 = pval.ItemUID
                oLastColUID01 = pval.ColUID
                oLastColRow01 = pval.Row
            End If
            
        End If
    
    ElseIf pval.BeforeAction = False Then
    
    End If
    
    Exit Sub
    
Raise_EVENT_CLICK_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_COMBO_SELECT(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_COMBO_SELECT_Error

    If pval.Before_Action = True Then
    
        If pval.ItemUID = "Mat01" Then

            If pval.Row > 0 Then

                Call oMat01.SelectRow(pval.Row, True, False)

            End If

        End If
    
    ElseIf pval.Before_Action = False Then
    
        If pval.ItemChanged = True Then
            Call PS_PP982_FlushToItemValue(pval.ItemUID)
        End If
    
    End If

    Exit Sub
Raise_EVENT_COMBO_SELECT_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_COMBO_SELECT_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_DOUBLE_CLICK(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_DOUBLE_CLICK_Error

    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
    
    End If
    
    Exit Sub
Raise_EVENT_DOUBLE_CLICK_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_DOUBLE_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_MATRIX_LINK_PRESSED(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_MATRIX_LINK_PRESSED_Error
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_MATRIX_LINK_PRESSED_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LINK_PRESSED_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_VALIDATE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_VALIDATE_Error

    Call oForm01.Freeze(True)
    
    If pval.BeforeAction = True Then
    
        If pval.ItemChanged = True Then

            If pval.ItemUID = "Mat01" Then
            
                Call PS_PP982_FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
                
            Else
            
                Call PS_PP982_FlushToItemValue(pval.ItemUID)
            
            End If
                
'            oMat02.LoadFromDataSource
'            oMat02.AutoResizeColumns
'            oForm01.Update
        End If
        
    ElseIf pval.BeforeAction = False Then

    End If
    
    Call oForm01.Freeze(False)
    
    Exit Sub
    
Raise_EVENT_VALIDATE_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_VALIDATE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_MATRIX_LOAD(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_MATRIX_LOAD_Error
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
'        Call PS_PP982_FormItemEnabled
        '//Call PS_PP982_AddMatrixRow(oMat02.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_PP982_FormResize
    End If
    Exit Sub
Raise_EVENT_RESIZE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_RESIZE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CHOOSE_FROM_LIST(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CHOOSE_FROM_LIST_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
'        If (pval.ItemUID = "ItemCode") Then
'            Dim oDataTable01 As SAPbouiCOM.DataTable
'            Set oDataTable01 = pval.SelectedObjects
'            oForm01.DataSources.UserDataSources("ItemCode").Value = oDataTable01.Columns(0).Cells(0).Value
'            Set oDataTable01 = Nothing
'        End If
'        If (pval.ItemUID = "CardCode" Or pval.ItemUID = "CardName") Then
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_PP982H", "U_CardCode,U_CardName")
'        End If
    End If
    Exit Sub
Raise_EVENT_CHOOSE_FROM_LIST_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CHOOSE_FROM_LIST_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_GOT_FOCUS(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_GOT_FOCUS_Error
    If pval.ItemUID = "Mat02" Then
        If pval.Row > 0 Then
            oLastItemUID01 = pval.ItemUID
            oLastColUID01 = pval.ColUID
            oLastColRow01 = pval.Row
        End If
    Else
        oLastItemUID01 = pval.ItemUID
        oLastColUID01 = ""
        oLastColRow01 = 0
    End If
    Exit Sub
Raise_EVENT_GOT_FOCUS_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_GOT_FOCUS_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_FORM_UNLOAD(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_FORM_UNLOAD_Error
    If pval.BeforeAction = True Then
    ElseIf pval.BeforeAction = False Then
        RemoveForms oFormUniqueID01
        Set oForm01 = Nothing
        Set oGrid01 = Nothing
        Set oGrid02 = Nothing
    End If
    Exit Sub
Raise_EVENT_FORM_UNLOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_FORM_UNLOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_ROW_DELETE(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_ROW_DELETE_Error
    Dim i As Long
    If (oLastColRow01 > 0) Then
        If pval.BeforeAction = True Then
'            If (PS_PP982_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_PP982_CreateItems() As Boolean
On Error GoTo PS_PP982_CreateItems_Error

    Call oForm01.Freeze(True)
    
    Set oGrid01 = oForm01.Items("Grid01").Specific
    Set oGrid02 = oForm01.Items("Grid02").Specific
    Set oGrid03 = oForm01.Items("Grid03").Specific
    
    Call oForm01.DataSources.DataTables.Add("PS_PP982L")
    Call oForm01.DataSources.DataTables.Add("PS_PP982M")
    Call oForm01.DataSources.DataTables.Add("PS_PP982N")
    
    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("PS_PP982L")
    oGrid02.DataTable = oForm01.DataSources.DataTables.Item("PS_PP982M")
    oGrid03.DataTable = oForm01.DataSources.DataTables.Item("PS_PP982N")

    Set oDS_PS_PP982L = oForm01.DataSources.DataTables.Item("PS_PP982L")
    Set oDS_PS_PP982M = oForm01.DataSources.DataTables.Item("PS_PP982M")
    Set oDS_PS_PP982N = oForm01.DataSources.DataTables.Item("PS_PP982N")
    
    Set oDS_PS_PP982O = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_Auto
    oMat01.AutoResizeColumns
    
    '사업장
    Call oForm01.DataSources.UserDataSources.Add("BPLID", dt_SHORT_TEXT, 5)
    Call oForm01.Items("BPLID").Specific.DataBind.SetBound(True, "", "BPLID")

    '일자Fr
    Call oForm01.DataSources.UserDataSources.Add("FrDt", dt_DATE)
    Call oForm01.Items("FrDt").Specific.DataBind.SetBound(True, "", "FrDt")
    
    '일자To
    Call oForm01.DataSources.UserDataSources.Add("ToDt", dt_DATE)
    Call oForm01.Items("ToDt").Specific.DataBind.SetBound(True, "", "ToDt")
    
    '사번
    Call oForm01.DataSources.UserDataSources.Add("CntcCode", dt_SHORT_TEXT, 20)
    Call oForm01.Items("CntcCode").Specific.DataBind.SetBound(True, "", "CntcCode")
    
    '성명
    Call oForm01.DataSources.UserDataSources.Add("CntcName", dt_SHORT_TEXT, 50)
    Call oForm01.Items("CntcName").Specific.DataBind.SetBound(True, "", "CntcName")
    
    '반 체크박스
    Call oForm01.DataSources.UserDataSources.Add("Check01", dt_SHORT_TEXT, 5)
    Call oForm01.Items("Check01").Specific.DataBind.SetBound(True, "", "Check01")
    
    Call oForm01.DataSources.UserDataSources.Add("Check02", dt_SHORT_TEXT, 5)
    Call oForm01.Items("Check02").Specific.DataBind.SetBound(True, "", "Check02")
    
    Call oForm01.DataSources.UserDataSources.Add("Check03", dt_SHORT_TEXT, 5)
    Call oForm01.Items("Check03").Specific.DataBind.SetBound(True, "", "Check03")
    
    Call oForm01.DataSources.UserDataSources.Add("Check04", dt_SHORT_TEXT, 5)
    Call oForm01.Items("Check04").Specific.DataBind.SetBound(True, "", "Check04")
    
    Call oForm01.DataSources.UserDataSources.Add("Check05", dt_SHORT_TEXT, 5)
    Call oForm01.Items("Check05").Specific.DataBind.SetBound(True, "", "Check05")
    
    Call oForm01.DataSources.UserDataSources.Add("Check06", dt_SHORT_TEXT, 5)
    Call oForm01.Items("Check06").Specific.DataBind.SetBound(True, "", "Check06")
    
    Call oForm01.DataSources.UserDataSources.Add("Check07", dt_SHORT_TEXT, 5)
    Call oForm01.Items("Check07").Specific.DataBind.SetBound(True, "", "Check07")
    
    Call oForm01.DataSources.UserDataSources.Add("Check08", dt_SHORT_TEXT, 5)
    Call oForm01.Items("Check08").Specific.DataBind.SetBound(True, "", "Check08")
    
    Call oForm01.DataSources.UserDataSources.Add("Check09", dt_SHORT_TEXT, 5)
    Call oForm01.Items("Check09").Specific.DataBind.SetBound(True, "", "Check09")
    
    Call oForm01.DataSources.UserDataSources.Add("Check10", dt_SHORT_TEXT, 5)
    Call oForm01.Items("Check10").Specific.DataBind.SetBound(True, "", "Check10")
    
    Call oForm01.DataSources.UserDataSources.Add("Check11", dt_SHORT_TEXT, 5)
    Call oForm01.Items("Check11").Specific.DataBind.SetBound(True, "", "Check11")
    
    Call oForm01.DataSources.UserDataSources.Add("Check12", dt_SHORT_TEXT, 5)
    Call oForm01.Items("Check12").Specific.DataBind.SetBound(True, "", "Check12")
    
    Call oForm01.Freeze(False)
    
    Exit Function
PS_PP982_CreateItems_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP982_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'/콤보박스 set
Sub PS_PP982_ComboBox_Setting()
On Error GoTo PS_PP982_ComboBox_Setting_Error

    Dim sQry As String
    
    Dim BPLID As String
    BPLID = MDC_PS_Common.User_BPLId
    
    Call oForm01.Freeze(True)
    
    '사업장
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLID").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", BPLID, False, False)
    
    Call oForm01.Freeze(False)
    
    Exit Sub
PS_PP982_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP982_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP982_CF_ChooseFromList()
On Error GoTo PS_PP982_CF_ChooseFromList_Error

    Exit Sub
PS_PP982_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP982_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP982_EnableMenus()
On Error GoTo PS_PP982_EnableMenus_Error
    
    Exit Sub
PS_PP982_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP982_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP982_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_PP982_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_PP982_FormItemEnabled
        '//Call PS_PP982_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_PP982_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_PP982_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP982_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub PS_PP982_FormReset()
'******************************************************************************
'Function ID : PS_PP982_FormReset()
'해당모듈    : PS_PP982
'기능        : 화면 초기화
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP982_FormReset_Error

    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    
    Dim User_BPLId As String
    User_BPLId = MDC_PS_Common.User_BPLId

    '헤더 초기화
    oForm01.DataSources.UserDataSources("FrDt").VALUE = Format(Now, "YYYYMM01") '기준일자
    oForm01.DataSources.UserDataSources("ToDt").VALUE = Format(Now, "YYYYMMDD") '기준일자
    
    '출력버튼 비활성
'    oForm01.Items("BtnPrt01").Visible = False
'    oForm01.Items("BtnPrt02").Visible = False
    
    '라인 초기화
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    Call PS_PP982_Add_MatrixRow01(0, True)
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)

    Exit Sub

PS_PP982_FormReset_Error:
    Call oForm01.Freeze(False)
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PS_PP982_FormReset_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub PS_PP982_CheckAll()
'******************************************************************************
'Function ID : PS_PP982_CheckAll()
'해당모듈    : PS_PP982
'기능        : 체크박스 전체 선택(해제)
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP982_CheckAll_Error

    Dim CheckType As String
    Dim loopCount As Integer
    
    Call oForm01.Freeze(True)
    CheckType = "Y"
    
    Call oMat01.FlushToDataSource
    
    For loopCount = 0 To oMat01.VisualRowCount - 2
        If Trim(oDS_PS_PP982O.GetValue("U_ColReg01", loopCount)) = "N" Then
            CheckType = "N"
            Exit For
        End If
    Next
    
    For loopCount = 0 To oMat01.VisualRowCount - 2
        oDS_PS_PP982O.offset = loopCount
        If CheckType = "N" Then
            Call oDS_PS_PP982O.setValue("U_ColReg01", loopCount, "Y")
        Else
            Call oDS_PS_PP982O.setValue("U_ColReg01", loopCount, "N")
        End If
    Next
    
    Call oMat01.LoadFromDataSource
    
    Call oForm01.Freeze(False)

    Exit Sub
    
PS_PP982_CheckAll_Error:
    Call oForm01.Freeze(False)
    Call Sbo_Application.SetStatusBarMessage("PS_PP982_CheckAll_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)
End Sub

Private Sub PS_PP982_FormResize()
On Error GoTo PS_PP982_FormResize_Error
    
    Call oForm01.Freeze(True)
    
    '그룹박스 크기 동적 할당
    oForm01.Items("GrpBox01").Height = oForm01.Items("Grid01").Height + 25
    oForm01.Items("GrpBox01").Width = oForm01.Items("Grid01").Width + 30
    
    If oGrid01.Columns.Count > 0 Then
        Call oGrid01.AutoResizeColumns
    End If
    
    If oGrid02.Columns.Count > 0 Then
        Call oGrid02.AutoResizeColumns
    End If
    
    If oGrid03.Columns.Count > 0 Then
        Call oGrid03.AutoResizeColumns
    End If
    
    Call oMat01.AutoResizeColumns
    
    Call oForm01.Freeze(False)
    
    Exit Sub
PS_PP982_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP982_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub PS_PP982_GetDetail(ByVal pCpCode As String, ByVal pCpName As String)
On Error GoTo PS_PP982_GetDetail_Error

    Dim loopCount As Integer
    
    Dim BPLID     As String '사업장
    Dim TeamCode  As String '팀
    Dim RspCode   As String '담당
    Dim ClsCode   As String '반
    Dim CardType  As String '거래처구분
    Dim ItemType  As String '품목구분
    Dim WCYN      As String '생산완료여부
    Dim DateStd   As String '일자기준
    Dim FrDt      As String '기간(Fr)
    Dim ToDt      As String '기간(To)
    Dim ItemCode  As String '품목코드(작번)
    Dim CpCode    As String '공정
    Dim CpName    As String '공정명
    Dim OrdGbn    As String '작업구분

    Dim oTempClass As Variant

'    For loopCount = 0 To oGrid01.Rows.Count - 1
'        If oGrid01.Rows.IsSelected(loopCount) = True Then
'
'            '/ Procedure 실행
''             sQry = "EXEC PS_PP361_02 '" & Trim(oGrid01.DataTable.GetValue(0, i)) & "'"
'''            sQry = "EXEC PS_PP361_02 '" & Trim(oGrid01.DataTable.GetValue(0, i)) & "','" & SYYYYMM & "','" & EYYYYMM & "'"
''            oRecordSet.DoQuery sQry
''
''            '// Procedure 실행(Grid 사용)
''            oForm01.DataSources.DataTables.Item("ZTEMP02").ExecuteQuery (sQry)
''            oGrid02.DataTable = oForm01.DataSources.DataTables.Item("ZTEMP02")
'
''            SumPkQty = Trim(oGrid02.DataTable.GetValue(7, i))
''            SumPkWt = Trim(oGrid02.DataTable.GetValue(8, i))
'
'            ItemCode = Trim(oGrid01.DataTable.GetValue(0, loopCount))
'            ItemName = Trim(oGrid01.DataTable.GetValue(1, loopCount))
'
'        End If
'    Next

    BPLID = Trim(oForm01.Items("BPLID02").Specific.VALUE)  '사업장
    TeamCode = Trim(oForm01.Items("TeamCode02").Specific.VALUE)  '팀
    RspCode = Trim(oForm01.Items("RspCode02").Specific.VALUE)  '담당
    ClsCode = Trim(oForm01.Items("ClsCode02").Specific.VALUE)  '반
    CardType = Trim(oForm01.Items("CardType02").Specific.VALUE)  '거래처구분
    ItemType = Trim(oForm01.Items("ItemType02").Specific.VALUE)  '품목구분
    WCYN = Trim(oForm01.Items("WCYN02").Specific.VALUE)  '생산완료여부
    DateStd = Trim(oForm01.Items("DateStd02").Specific.VALUE)  '일자기준
    FrDt = Trim(oForm01.Items("FrDt02").Specific.VALUE)  '기간(Fr)
    ToDt = Trim(oForm01.Items("ToDt02").Specific.VALUE)  '기간(To)
    ItemCode = Trim(oForm01.Items("ItemCode02").Specific.VALUE)  '품목코드(작번)
    CpCode = pCpCode  '공정
    CpName = pCpName '공정명
    OrdGbn = Trim(oForm01.Items("OrdGbn02").Specific.VALUE)  '작업구분

    Set oTempClass = New PS_PP990
    Call oTempClass.LoadForm(BPLID, TeamCode, RspCode, ClsCode, CardType, ItemType, WCYN, DateStd, FrDt, ToDt, ItemCode, CpCode, CpName, OrdGbn)
    
    Exit Sub

PS_PP982_GetDetail_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP982_GetDetail_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP982_Print_Report01()
On Error GoTo Print_Query_Error

    Dim DocNum     As String
    Dim WinTitle   As String
    Dim ReportName As String
    Dim sQry       As String

    Dim i          As Integer
    Dim ErrNum     As Integer
    Dim Sub_sQry   As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)

    Call ConnectODBC
    
    Dim BPLID    As String '사업장
    Dim TeamCode As String '팀
    Dim RspCode  As String '담당
    Dim ClsCode  As String '반
    Dim CardType As String '거래처구분
    Dim ItemType As String '품목구분
    Dim WCYN     As String '생산완료여부
    Dim DateStd  As String '일자기준
    Dim FrDt     As String '기간(Fr)
    Dim ToDt     As String '기간(To)
    Dim ItemCode As String '품목코드(작번)
    Dim CpCode   As String '공정
    Dim OrdGbn   As String '작업구분
    Dim CntcCode As String '조회자 사번
    
    BPLID = Trim(oForm01.Items("BPLID04").Specific.VALUE)  '사업장
    TeamCode = Trim(oForm01.Items("TeamCode04").Specific.VALUE)  '팀
    RspCode = Trim(oForm01.Items("RspCode04").Specific.VALUE)  '담당
    ClsCode = Trim(oForm01.Items("ClsCode04").Specific.VALUE)  '반
    CardType = Trim(oForm01.Items("CardType04").Specific.VALUE)  '거래처구분
    ItemType = Trim(oForm01.Items("ItemType04").Specific.VALUE)  '품목구분
    WCYN = Trim(oForm01.Items("WCYN04").Specific.VALUE)  '생산완료여부
    DateStd = Trim(oForm01.Items("DateStd04").Specific.VALUE)  '일자기준
    FrDt = Trim(oForm01.Items("FrDt04").Specific.VALUE)  '기간(Fr)
    ToDt = Trim(oForm01.Items("ToDt04").Specific.VALUE)  '기간(To)
    ItemCode = Trim(oForm01.Items("ItemCode04").Specific.VALUE)  '품목코드(작번)
    CpCode = Trim(oForm01.Items("CpCode04").Specific.VALUE)  '공정
    OrdGbn = Trim(oForm01.Items("OrdGbn04").Specific.VALUE)  '작업구분
    CntcCode = MDC_PS_Common.User_MSTCOD '조회자 사번
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "[PS_PP982] 레포트"
    ReportName = "PS_PP982_01.rpt"
 
    '// Formula 수식필드
    ReDim gRpt_Formula(2)
    ReDim gRpt_Formula_Value(2)
       
    '// SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    gRpt_SFormula(1, 1) = ""
    gRpt_SFormula_Value(1, 1) = ""
    
    '/ Procedure 실행"
    sQry = "      EXEC [PS_PP982_51] '"
    sQry = sQry & BPLID & "','" '사업장
    sQry = sQry & TeamCode & "','" '팀
    sQry = sQry & RspCode & "','" '담당
    sQry = sQry & ClsCode & "','" '반
    sQry = sQry & CardType & "','" '거래처구분
    sQry = sQry & ItemType & "','" '품목구분
    sQry = sQry & WCYN & "','" '생산완료여부
    sQry = sQry & DateStd & "','" '일자기준
    sQry = sQry & FrDt & "','" '기간(Fr)
    sQry = sQry & ToDt & "','" '기간(To)
    sQry = sQry & ItemCode & "','" '품목코드(작번)
    sQry = sQry & CpCode & "','" '공정
    sQry = sQry & OrdGbn & "','" '작업구분
    sQry = sQry & CntcCode & "'" '조회자사번

'    oRecordSet.DoQuery sQry
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo Print_Query_Error
'    End If

    '/ Action (sub_query가 있을때는 'Y'로...)/
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Query_Error:
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "PS_PP982_Print_Report01_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Private Sub PS_PP982_Print_Report02()
On Error GoTo Print_Query_Error

    Dim DocNum     As String
    Dim WinTitle   As String
    Dim ReportName As String
    Dim sQry       As String

    Dim i          As Integer
    Dim ErrNum     As Integer
    Dim Sub_sQry   As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)

    Call ConnectODBC
    
    Dim BPLID    As String '사업장
    Dim TeamCode As String '팀
    Dim RspCode  As String '담당
    Dim ClsCode  As String '반
    Dim CardType As String '거래처구분
    Dim ItemType As String '품목구분
    Dim WCYN     As String '생산완료여부
    Dim DateStd  As String '일자기준
    Dim FrDt     As String '기간(Fr)
    Dim ToDt     As String '기간(To)
    Dim ItemCode As String '품목코드(작번)
    Dim CpCode   As String '공정
    Dim OrdGbn   As String '작업구분
    Dim CntcCode As String '조회자 사번
    
    BPLID = Trim(oForm01.Items("BPLID02").Specific.VALUE)  '사업장
    TeamCode = Trim(oForm01.Items("TeamCode02").Specific.VALUE)  '팀
    RspCode = Trim(oForm01.Items("RspCode02").Specific.VALUE)  '담당
    ClsCode = Trim(oForm01.Items("ClsCode02").Specific.VALUE)  '반
    CardType = Trim(oForm01.Items("CardType02").Specific.VALUE)  '거래처구분
    ItemType = Trim(oForm01.Items("ItemType02").Specific.VALUE)  '품목구분
    WCYN = Trim(oForm01.Items("WCYN02").Specific.VALUE)  '생산완료여부
    DateStd = Trim(oForm01.Items("DateStd02").Specific.VALUE)  '일자기준
    FrDt = Trim(oForm01.Items("FrDt02").Specific.VALUE)  '기간(Fr)
    ToDt = Trim(oForm01.Items("ToDt02").Specific.VALUE)  '기간(To)
    ItemCode = Trim(oForm01.Items("ItemCode02").Specific.VALUE)  '품목코드(작번)
    CpCode = Trim(oForm01.Items("CpCode02").Specific.VALUE)  '공정
    OrdGbn = Trim(oForm01.Items("OrdGbn02").Specific.VALUE)  '작업구분
    CntcCode = MDC_PS_Common.User_MSTCOD '조회자 사번
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "[PS_PP982] 레포트"
    ReportName = "PS_PP982_01.rpt"
 
    '// Formula 수식필드
    ReDim gRpt_Formula(2)
    ReDim gRpt_Formula_Value(2)
       
    '// SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    gRpt_SFormula(1, 1) = ""
    gRpt_SFormula_Value(1, 1) = ""
    
    '/ Procedure 실행"
    sQry = "      EXEC [PS_PP982_52] '"
    sQry = sQry & BPLID & "','" '사업장
    sQry = sQry & TeamCode & "','" '팀
    sQry = sQry & RspCode & "','" '담당
    sQry = sQry & ClsCode & "','" '반
    sQry = sQry & CardType & "','" '거래처구분
    sQry = sQry & ItemType & "','" '품목구분
    sQry = sQry & WCYN & "','" '생산완료여부
    sQry = sQry & DateStd & "','" '일자기준
    sQry = sQry & FrDt & "','" '기간(Fr)
    sQry = sQry & ToDt & "','" '기간(To)
    sQry = sQry & ItemCode & "','" '품목코드(작번)
    sQry = sQry & CpCode & "','" '공정
    sQry = sQry & OrdGbn & "','" '작업구분
    sQry = sQry & CntcCode & "'" '조회자사번

'    oRecordSet.DoQuery sQry
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo Print_Query_Error
'    End If

    '/ Action (sub_query가 있을때는 'Y'로...)/
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Query_Error:
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "PS_PP982_Print_Report02_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub














