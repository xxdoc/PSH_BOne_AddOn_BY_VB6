VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_GA032"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File               : PS_GA032.cls
'//  Module             : 총무관리 > 예산및비용관리 > 관리담당전용(소모품비)
'//  Desc               : 월별계획등록
'//  FormType           : PS_GA032
'//  Create Date(Start) : 2017.04.26
'//  Create Date(End)   :
'//  Creator            : Song Myoung gyu
'//  Modified Date      :
'//  Modifier           :
'//  Company            : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01  As String
Public oForm01          As SAPbouiCOM.Form

Public oMat01           As SAPbouiCOM.Matrix

Public oDS_PS_GA032A    As SAPbouiCOM.DBDataSource '등록용 Matrix

Private oLastItemUID01  As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01   As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01   As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error

    Dim i         As Long
    Dim oInnerXml As String
    Dim oXmlDoc   As New MSXML2.DOMDocument
    
    oXmlDoc.Load (SubMain.ShareFolderPath & "ScreenPS\PS_GA032.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)

    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_GA032_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_GA032_CreateItems
    Call PS_GA032_ComboBox_Setting
    Call PS_GA032_Initial_Setting
    Call PS_GA032_CF_ChooseFromList
    Call PS_GA032_EnableMenus
    Call PS_GA032_SetDocument(oFromDocEntry01)
    Call PS_GA032_FormResize

'    Call PS_GA032_Add_MatrixRow(0, True)
    Call PS_GA032_LoadCaption
    Call PS_GA032_FormItemEnabled
    
    oForm01.EnableMenu ("1283"), False  '// 삭제
    oForm01.EnableMenu ("1286"), False  '// 닫기
    oForm01.EnableMenu ("1287"), False  '// 복제
    oForm01.EnableMenu ("1285"), False  '// 복원
    oForm01.EnableMenu ("1284"), False  '// 취소
    oForm01.EnableMenu ("1293"), False  '// 행삭제
    oForm01.EnableMenu ("1281"), False
    oForm01.EnableMenu ("1282"), True

    Call PS_GA032_FormReset '폼초기화
    
    oForm01.Update
    oForm01.Freeze False
    
    Call oForm01.Items("StdYear").Click
    
    oForm01.Visible = True
    Set oXmlDoc = Nothing

    Exit Sub
    
LoadForm_Error:
    oForm01.Update
    oForm01.Freeze False
    Set oXmlDoc = Nothing
    Set oForm01 = Nothing
    Sbo_Application.SetStatusBarMessage "Form_Load Error:" & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA032_LoadCaption()
'******************************************************************************
'Function ID : PS_GA032_LoadCaption()
'해당모듈    : PS_GA032
'기능        : Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_GA032_LoadCaption_Error

    Call oForm01.Freeze(True)

    If oForm01.Mode = fm_ADD_MODE Then
        oForm01.Items("BtnAdd").Specific.Caption = "추가"
'        oForm01.Items("BtnDelete").Enabled = False
'    ElseIf oForm01.Mode = fm_OK_MODE Then
'        oForm01.Items("BtnAdd").Specific.Caption = "확인"
    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
        oForm01.Items("BtnAdd").Specific.Caption = "수정"
'        oForm01.Items("BtnDelete").Enabled = True
    End If
    
    Call oForm01.Freeze(False)

    Exit Sub

PS_GA032_LoadCaption_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/메트릭스 Row추가
Sub PS_GA032_Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_GA032_Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_GA032A.InsertRecord (oRow)
    End If
    
    oMat01.AddRow
    oDS_PS_GA032A.offset = oRow
    Call oDS_PS_GA032A.setValue("U_LineNum", oRow, oRow + 1)
    Call oDS_PS_GA032A.setValue("U_ColReg01", oRow, "Y")
    
    oMat01.LoadFromDataSource
    
    Exit Sub

PS_GA032_Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "PS_GA032_Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub PS_GA032_MTX01()
'******************************************************************************
'Function ID : PS_GA032_MTX01()
'해당모듈    : PS_GA032
'기능        : 등록된 데이터 조회
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_GA032_MTX01_Error

    Dim i       As Integer
    Dim sQry    As String
    Dim ErrNum  As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLID As String
    Dim StdYear As String '기준년도
    
    BPLID = Trim(oForm01.Items("BPLID").Specific.VALUE) '사업장
    StdYear = Trim(oForm01.Items("StdYear").Specific.VALUE) '기준년도
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    sQry = "      EXEC [PS_GA032_01] '"
    sQry = sQry & BPLID & "','" '사업장
    sQry = sQry & StdYear & "'" '기준년도
    
    Call oRecordSet01.DoQuery(sQry)
    
    oMat01.Clear
    oDS_PS_GA032A.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource

    If (oRecordSet01.RecordCount = 0) Then
    
        ErrNum = 1
        
        oForm01.Mode = fm_ADD_MODE
        
        Call PS_GA032_Add_MatrixRow(0, True)
        Call PS_GA032_LoadCaption
        
        GoTo PS_GA032_MTX01_Error
        
        Exit Sub
    End If
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_GA032A.Size Then
            oDS_PS_GA032A.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_GA032A.offset = i
        
        Call oDS_PS_GA032A.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_GA032A.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("Select").VALUE)) '선택
        Call oDS_PS_GA032A.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("BudCls").VALUE)) '구분
        Call oDS_PS_GA032A.setValue("U_ColReg03", i, Trim(oRecordSet01.Fields("BudClsNm").VALUE)) '구분명
        Call oDS_PS_GA032A.setValue("U_ColReg04", i, Trim(oRecordSet01.Fields("PlanCls").VALUE)) '계획구분
        Call oDS_PS_GA032A.setValue("U_ColSum01", i, Trim(oRecordSet01.Fields("Month01").VALUE)) '1월
        Call oDS_PS_GA032A.setValue("U_ColSum02", i, Trim(oRecordSet01.Fields("Month02").VALUE)) '2월
        Call oDS_PS_GA032A.setValue("U_ColSum03", i, Trim(oRecordSet01.Fields("Month03").VALUE)) '3월
        Call oDS_PS_GA032A.setValue("U_ColSum04", i, Trim(oRecordSet01.Fields("Month04").VALUE)) '4월
        Call oDS_PS_GA032A.setValue("U_ColSum05", i, Trim(oRecordSet01.Fields("Month05").VALUE)) '5월
        Call oDS_PS_GA032A.setValue("U_ColSum06", i, Trim(oRecordSet01.Fields("Month06").VALUE)) '6월
        Call oDS_PS_GA032A.setValue("U_ColSum07", i, Trim(oRecordSet01.Fields("Month07").VALUE)) '7월
        Call oDS_PS_GA032A.setValue("U_ColSum08", i, Trim(oRecordSet01.Fields("Month08").VALUE)) '8월
        Call oDS_PS_GA032A.setValue("U_ColSum09", i, Trim(oRecordSet01.Fields("Month09").VALUE)) '9월
        Call oDS_PS_GA032A.setValue("U_ColSum10", i, Trim(oRecordSet01.Fields("Month10").VALUE)) '10월
        Call oDS_PS_GA032A.setValue("U_ColSum11", i, Trim(oRecordSet01.Fields("Month11").VALUE)) '11월
        Call oDS_PS_GA032A.setValue("U_ColSum12", i, Trim(oRecordSet01.Fields("Month12").VALUE)) '12월
        Call oDS_PS_GA032A.setValue("U_ColSum13", i, Trim(oRecordSet01.Fields("Total").VALUE)) '계
        
        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
        
    Next
    
    Call PS_GA032_Add_MatrixRow(oMat01.VisualRowCount)
    
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    ProgBar01.Stop
    
    Call oForm01.Freeze(False)
  
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub

PS_GA032_MTX01_Error:
'    ProgBar01.Stop
    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_GA032_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Sub PS_GA032_DeleteData()
'******************************************************************************
'Function ID : PS_GA032_DeleteData()
'해당모듈    : PS_GA032
'기능        : 기본정보 삭제
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_GA032_DeleteData_Error

    Dim loopCount As Integer
    Dim sQry      As String
    Dim ErrNum    As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLID   As String
    Dim StdYear As String
    Dim BudCls  As String
    Dim PlanCls As String
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("삭제 중...", 100, False)
    
    BPLID = oForm01.DataSources.UserDataSources("BPLID").VALUE
    StdYear = oForm01.DataSources.UserDataSources("StdYear").VALUE
    
    Call oMat01.FlushToDataSource
    
    For loopCount = 0 To oMat01.VisualRowCount - 2 '마지막 빈행 제외를 위해 2를 뺌
        If Trim(oDS_PS_GA032A.GetValue("U_ColReg01", loopCount)) = "Y" Then
            
            BudCls = Trim(oDS_PS_GA032A.GetValue("U_ColReg02", loopCount))
            PlanCls = Trim(oDS_PS_GA032A.GetValue("U_ColReg04", loopCount))
            
            sQry = "      EXEC [PS_GA032_03] '"
            sQry = sQry & BPLID & "','" '사업장
            sQry = sQry & StdYear & "','" '기준년도
            sQry = sQry & BudCls & "','" '구분코드
            sQry = sQry & PlanCls & "'" '계획구분
            
            Call oRecordSet01.DoQuery(sQry)
        End If
    Next
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Call MDC_Com.MDC_GF_Message("삭제 완료!", "S")
    
    Exit Sub

PS_GA032_DeleteData_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("삭제대상이 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_GA032_DeleteData_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Function PS_GA032_AddData() As Boolean
'******************************************************************************
'Function ID : PS_GA032_AddData()
'해당모듈    : PS_GA032
'기능        : 데이터 INSERT, UPDATE(기존 데이터가 존재하면 UPDATE, 아니면 INSERT)
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_GA032_AddData_Error

    Dim i    As Integer
    Dim sQry As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLID    As String '사업장
    Dim StdYear  As String '기준년도
    Dim BudCls   As String '구분
    Dim PlanCls  As String '계획구분
    Dim Month01  As Currency '1월
    Dim Month02  As Currency '2월
    Dim Month03  As Currency '3월
    Dim Month04  As Currency '4월
    Dim Month05  As Currency '5월
    Dim Month06  As Currency '6월
    Dim Month07  As Currency '7월
    Dim Month08  As Currency '8월
    Dim Month09  As Currency '9월
    Dim Month10  As Currency '10월
    Dim Month11  As Currency '11월
    Dim Month12  As Currency '12월
    Dim Total  As Currency '계
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("저장 중...", oMat01.VisualRowCount - 1, False)
    
    BPLID = oForm01.DataSources.UserDataSources("BPLID").VALUE
    StdYear = oForm01.DataSources.UserDataSources("StdYear").VALUE
        
    Call oMat01.FlushToDataSource
    For i = 0 To oMat01.VisualRowCount - 2 '마지막 빈행 제외를 위해 2를 뺌
        If Trim(oDS_PS_GA032A.GetValue("U_ColReg01", i)) = "Y" Then
            
            BudCls = Trim(oDS_PS_GA032A.GetValue("U_ColReg02", i)) '구분
            PlanCls = Trim(oDS_PS_GA032A.GetValue("U_ColReg04", i)) '계획구분
            Month01 = Val(Trim(oDS_PS_GA032A.GetValue("U_ColSum01", i))) '1월
            Month02 = Val(Trim(oDS_PS_GA032A.GetValue("U_ColSum02", i))) '2월
            Month03 = Val(Trim(oDS_PS_GA032A.GetValue("U_ColSum03", i))) '3월
            Month04 = Val(Trim(oDS_PS_GA032A.GetValue("U_ColSum04", i))) '4월
            Month05 = Val(Trim(oDS_PS_GA032A.GetValue("U_ColSum05", i))) '5월
            Month06 = Val(Trim(oDS_PS_GA032A.GetValue("U_ColSum06", i))) '6월
            Month07 = Val(Trim(oDS_PS_GA032A.GetValue("U_ColSum07", i))) '7월
            Month08 = Val(Trim(oDS_PS_GA032A.GetValue("U_ColSum08", i))) '8월
            Month09 = Val(Trim(oDS_PS_GA032A.GetValue("U_ColSum09", i))) '9월
            Month10 = Val(Trim(oDS_PS_GA032A.GetValue("U_ColSum10", i))) '10월
            Month11 = Val(Trim(oDS_PS_GA032A.GetValue("U_ColSum11", i))) '11월
            Month12 = Val(Trim(oDS_PS_GA032A.GetValue("U_ColSum12", i))) '12월
            Total = Val(Trim(oDS_PS_GA032A.GetValue("U_ColSum13", i))) '계
            
            sQry = "      EXEC [PS_GA032_02] '"
            sQry = sQry & BPLID & "','" '사업장
            sQry = sQry & StdYear & "','" '기준년도
            sQry = sQry & BudCls & "','" '구분
            sQry = sQry & PlanCls & "','" '계획구분
            sQry = sQry & Month01 & "','" '1월
            sQry = sQry & Month02 & "','" '2월
            sQry = sQry & Month03 & "','" '3월
            sQry = sQry & Month04 & "','" '4월
            sQry = sQry & Month05 & "','" '5월
            sQry = sQry & Month06 & "','" '6월
            sQry = sQry & Month07 & "','" '7월
            sQry = sQry & Month08 & "','" '8월
            sQry = sQry & Month09 & "','" '9월
            sQry = sQry & Month10 & "','" '10월
            sQry = sQry & Month11 & "','" '11월
            sQry = sQry & Month12 & "','" '12월
            sQry = sQry & Total & "'" '계
            
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
    
    PS_GA032_AddData = True

    Exit Function

PS_GA032_AddData_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    PS_GA032_AddData = False
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PS_GA032_AddData_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Private Function PS_GA032_HeaderSpaceLineDel() As Boolean
'******************************************************************************
'Function ID : PS_GA032_HeaderSpaceLineDel()
'해당모듈    : PS_GA032
'기능        : 필수입력사항 체크
'인수        : 없음
'반환값      : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_GA032_HeaderSpaceLineDel_Error

    Dim ErrNum As Integer
    ErrNum = 0

    Select Case True
'        Case Trim(oForm01.Items("DestNo1").Specific.VALUE) = "" '출장번호1
'            ErrNum = 1
'            GoTo PS_GA032_HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("DestNo2").Specific.VALUE) = "" '출장번호2
'            ErrNum = 2
'            GoTo PS_GA032_HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("MSTCOD").Specific.VALUE) = "" '사원번호
'            ErrNum = 3
'            GoTo PS_GA032_HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("FrDate").Specific.VALUE) = "" '시작일자
'            ErrNum = 4
'            GoTo PS_GA032_HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("FrTime").Specific.VALUE) = "" '시작시각
'            ErrNum = 5
'            GoTo PS_GA032_HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("ToDate").Specific.VALUE) = "" '종료일자
'            ErrNum = 6
'            GoTo PS_GA032_HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("ToTime").Specific.VALUE) = "" '종료시각
'            ErrNum = 7
'            GoTo PS_GA032_HeaderSpaceLineDel_Error
    End Select

    PS_GA032_HeaderSpaceLineDel = True
    
    Exit Function

PS_GA032_HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출장번호1은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("DestNo1").Click(ct_Regular)
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "출장번호2는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("DestNo2").Click(ct_Regular)
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "사원번호는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("MSTCOD").Click(ct_Regular)
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "시작일자는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("FrDate").Click(ct_Regular)
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "시작시각은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("FrTime").Click(ct_Regular)
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "종료일자는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("FrDate").Click(ct_Regular)
    ElseIf ErrNum = 7 Then
        MDC_Com.MDC_GF_Message "종료시각은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("FrTime").Click(ct_Regular)
    End If
    PS_GA032_HeaderSpaceLineDel = False
End Function

'/ 메트릭스 필수 사항 check
Private Function PS_GA032_MatrixSpaceLineDel() As Boolean
On Error GoTo PS_GA032_MatrixSpaceLineDel_Error

    Dim i As Long
    Dim ErrNum As Integer
    ErrNum = 0

    Call oMat01.FlushToDataSource
    For i = 0 To oMat01.VisualRowCount - 2 '마직막 빈행 제외를 위해 2를 뺌
        If Trim(oDS_PS_GA032A.GetValue("U_ColReg01", i)) = "Y" Then
            
'            If Trim(oDS_PS_GA032A.GetValue("U_ColReg02", i)) = "" Then  '작번
'                ErrNum = 1
'                GoTo PS_GA032_MatrixSpaceLineDel_Error
'            ElseIf Trim(oDS_PS_GA032A.GetValue("U_ColSum01", i)) = "" Then '금액
'                ErrNum = 2
'                GoTo PS_GA032_MatrixSpaceLineDel_Error
'            ElseIf Trim(oDS_PS_GA032A.GetValue("U_ColReg05", i)) = "" Then '기준년월
'                ErrNum = 3
'                GoTo PS_GA032_MatrixSpaceLineDel_Error
'            End If

        End If
    Next

    PS_GA032_MatrixSpaceLineDel = True
    
    Exit Function

PS_GA032_MatrixSpaceLineDel_Error:

    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 작번이 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 금액이 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 기준년월이 없습니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "PS_GA032_MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    PS_GA032_MatrixSpaceLineDel = False
    
End Function

Private Sub PS_GA032_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PS_GA032_FlushToItemValue_Error

    Dim loopCount As Integer
    Dim ErrNum    As Integer
    Dim sQry      As String
    
    Dim FldValue  As String
    Dim Descr     As String
    Dim Total     As Currency
    
    Dim oCombo As SAPbouiCOM.ComboBox
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID

        Case "Mat01"
        
            Call oMat01.FlushToDataSource
        
            If oCol = "BudCls" Then
            
                '예산구분 조회
                sQry = "        SELECT      T0.FldValue,"
                sQry = sQry & "             T0.Descr"
                sQry = sQry & " FROM        UFD1 AS T0"
                sQry = sQry & " WHERE       T0.TableID = 'JDT1'"
                sQry = sQry & "             AND T0.FieldID = '28'"
                sQry = sQry & "             AND T0.FldValue = '" & oMat01.Columns(oCol).Cells(oRow).Specific.VALUE & "'"
                
                Call oRecordSet01.DoQuery(sQry)
                
                FldValue = oRecordSet01.Fields("FldValue").VALUE
                Descr = oRecordSet01.Fields("Descr").VALUE

                Call oDS_PS_GA032A.setValue("U_ColReg02", oRow - 1, FldValue)  '예산구분코드
                Call oDS_PS_GA032A.setValue("U_ColReg03", oRow - 1, Descr) '예산구분명
'                Call oDS_PS_GA032A.setValue("U_ColSum01", oRow - 1, Amount) '금액
                
                '행 추가
                If oMat01.RowCount = oRow And Trim(oDS_PS_GA032A.GetValue("U_ColReg02", oRow - 1)) <> "" Then
                    Call PS_GA032_Add_MatrixRow(oRow)
                End If
            
'                Call oMat01.Columns("Month01").Cells(oRow).Click(ct_Regular)
                
            ElseIf oCol <> "BudCls" Then
            
                Total = Val(Trim(oDS_PS_GA032A.GetValue("U_ColSum01", oRow - 1))) + _
                        Val(Trim(oDS_PS_GA032A.GetValue("U_ColSum02", oRow - 1))) + _
                        Val(Trim(oDS_PS_GA032A.GetValue("U_ColSum03", oRow - 1))) + _
                        Val(Trim(oDS_PS_GA032A.GetValue("U_ColSum04", oRow - 1))) + _
                        Val(Trim(oDS_PS_GA032A.GetValue("U_ColSum05", oRow - 1))) + _
                        Val(Trim(oDS_PS_GA032A.GetValue("U_ColSum06", oRow - 1))) + _
                        Val(Trim(oDS_PS_GA032A.GetValue("U_ColSum07", oRow - 1))) + _
                        Val(Trim(oDS_PS_GA032A.GetValue("U_ColSum08", oRow - 1))) + _
                        Val(Trim(oDS_PS_GA032A.GetValue("U_ColSum09", oRow - 1))) + _
                        Val(Trim(oDS_PS_GA032A.GetValue("U_ColSum10", oRow - 1))) + _
                        Val(Trim(oDS_PS_GA032A.GetValue("U_ColSum11", oRow - 1))) + _
                        Val(Trim(oDS_PS_GA032A.GetValue("U_ColSum12", oRow - 1)))
                        
                Call oDS_PS_GA032A.setValue("U_ColSum13", oRow - 1, Total)

            End If

            Call oMat01.LoadFromDataSource
'            '강제 포커스 이동_S
            Call oMat01.Columns(oCol).Cells(oRow).Click(ct_Regular)
'            '강제 포커스 이동_E
            Call oMat01.AutoResizeColumns

'        Case "OrdNum01"
'
'            oForm01.Items("ItemName01").Specific.VALUE = MDC_GetData.Get_ReData("FrgnName", "ItemCode", "OITM", "'" & oForm01.Items("OrdNum01").Specific.VALUE & "'") '품명
'            oForm01.Items("ItemSpec01").Specific.VALUE = MDC_GetData.Get_ReData("U_Size", "ItemCode", "OITM", "'" & oForm01.Items("OrdNum01").Specific.VALUE & "'")  '규격

    End Select

    Set oCombo = Nothing
    Set oRecordSet01 = Nothing

    Exit Sub
    
PS_GA032_FlushToItemValue_Error:

    Set oCombo = Nothing
    Set oRecordSet01 = Nothing

    Call MDC_Com.MDC_GF_Message("PS_GA032_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")

End Sub

'/폼의 아이템 사용지정
Sub PS_GA032_FormItemEnabled()
On Error GoTo PS_GA032_FormItemEnabled_Error

    If (oForm01.Mode = fm_ADD_MODE) Then
       
'        oMat01.Columns("ItemCode").Cells(1).Click ct_Regular
'        oForm01.Items("ItemCode").Enabled = True

    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
    
'        oForm01.Items("ItemCode").Enabled = True

    ElseIf (oForm01.Mode = fm_OK_MODE) Then
       
    End If
    
    Exit Sub

PS_GA032_FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "PS_GA032_FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/아이템 변경 이벤트
Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_FormItemEvent_Error

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

Raise_FormItemEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_FormItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_FormMenuEvent_Error

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
            
                oForm01.Mode = fm_ADD_MODE

                Call PS_GA032_FormReset
'                Call PS_GA032_LoadCaption
                
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
            
            Case "7169": '엑셀 내보내기
            
'                '엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
'                Call oForm01.Freeze(True)
'                Call PS_GA032_Add_MatrixRow(oMat01.VisualRowCount)
'                Call oForm01.Freeze(False)
            
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
                '//Call PS_GA032_FormItemEnabled '//UDO방식
            Case "1282": '추가
'                oMat01.Clear
'                oDS_PS_GA032A.Clear

'                Call PS_GA032_LoadCaption
'                Call PS_GA032_FormItemEnabled
                '//Call PS_GA032_FormItemEnabled '//UDO방식
                '//Call PS_GA032_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PS_GA032_FormItemEnabled
                
            Case "7169": '엑셀 내보내기
            
'                '엑셀 내보내기 이후 처리
'                Call oForm01.Freeze(True)
'                Call oDS_PS_GA032B.RemoveRecord(oDS_PS_GA032B.Size - 1)
'                Call oMat01.LoadFromDataSource
'                Call oForm01.Freeze(False)
            
        End Select
    End If
    
    Exit Sub
    
Raise_FormMenuEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_FormMenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
    
        If pval.ItemUID = "PS_GA032" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "BtnAdd" Then '/추가/확인 버튼클릭
        
            If oForm01.Mode = fm_ADD_MODE Then
            
'                If PS_GA032_HeaderSpaceLineDel() = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If

'                If PS_GA032_DataCheck() = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If

                '매트릭스 필수자료 체크
                If PS_GA032_MatrixSpaceLineDel = False Then
                    BubbleEvent = False
                    Exit Sub
                End If

                If PS_GA032_AddData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If

'                Call PS_GA032_FormReset
                oForm01.Mode = fm_ADD_MODE
                
                Call PS_GA032_LoadCaption
                Call PS_GA032_MTX01
                
                oLast_Mode = oForm01.Mode
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            
                If PS_GA032_HeaderSpaceLineDel = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            
'                If PS_GA032_DataCheck() = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
                
                Call PS_GA032_FormReset
                oForm01.Mode = fm_ADD_MODE
                
                Call PS_GA032_LoadCaption
                Call PS_GA032_MTX01
                
'                oForm01.Items("GCode").Click ct_Regular
            End If
            
        ElseIf pval.ItemUID = "BtnSearch" Then '/조회
            
'            Call PS_GA032_FormReset
'            oForm01.Mode = fm_ADD_MODE '/fm_VIEW_MODE
'
'            Call PS_GA032_LoadCaption
            Call PS_GA032_MTX01

        ElseIf pval.ItemUID = "BtnDelete" Then '/삭제
        
            If Sbo_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", "1", "예", "아니오") = "1" Then
        
                Call PS_GA032_DeleteData
'                Call PS_GA032_FormReset
                oForm01.Mode = fm_ADD_MODE '/fm_VIEW_MODE
                
                Call PS_GA032_LoadCaption
                Call PS_GA032_MTX01
                
            Else
            
            End If
        
        ElseIf pval.ItemUID = "BtnSelect" Then '전체선택
        
            Call PS_GA032_CheckAll
        
        ElseIf pval.ItemUID = "BtnPrt01" Then
        
            Call PS_GA032_Print_Report01
        
        End If
  
    ElseIf pval.BeforeAction = False Then
    
        If pval.ItemUID = "PS_GA032" Then
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
        
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "BudCls") 'Matrix-구분
        
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
                
            End If
        End If
    ElseIf pval.BeforeAction = False Then
    
    End If
    
    Exit Sub
    
Raise_EVENT_CLICK_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_COMBO_SELECT(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_COMBO_SELECT_Error

    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
    
        If pval.ItemUID = "Mat01" Then
        
        Else
        
            Call PS_GA032_FlushToItemValue(pval.ItemUID)
        
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
        
            If (pval.ItemUID = "Mat01") Then
            
'                If pval.ColUID = "BudCls" Then
            
                    Call PS_GA032_FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
                
                    'oMat01.AutoResizeColumns
                
'                End If
            
'                oMat01.FlushToDataSource
'
'                If (pval.ColUID = "MSTCOD") Then
'                    Call oDS_PS_GA032B.setValue("U_ColReg02", pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                    Call oDS_PS_GA032B.setValue("U_ColReg03", pval.Row - 1, MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'"))
'                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_GA032B.GetValue("U_ColReg02", pval.Row - 1)) <> "" Then
'                        Call PS_GA032_Add_MatrixRow(pval.Row)
'                    End If
'                Else
''                    Call oDS_PS_GA032B.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                End If
'
'                oMat01.LoadFromDataSource
                oMat01.AutoResizeColumns
                
            Else
                
                Call PS_GA032_FlushToItemValue(pval.ItemUID)
                
'                If pval.ItemUID = "MSTCOD" Then
'
'                    oForm01.Items("MSTNAM").Specific.VALUE = MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm01.Items("MSTCOD").Specific.VALUE & "'") '성명
'
'                End If
                
            End If

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
        Call PS_GA032_FormItemEnabled
        '//Call PS_GA032_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    
    Exit Sub
    
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error

    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_GA032_FormResize
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
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_GA032A", "U_CardCode,U_CardName")
'        End If
    End If
    
    Exit Sub
    
Raise_EVENT_CHOOSE_FROM_LIST_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CHOOSE_FROM_LIST_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub Raise_EVENT_GOT_FOCUS(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_GOT_FOCUS_Error

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
    
Raise_EVENT_GOT_FOCUS_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_GOT_FOCUS_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_FORM_UNLOAD(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_FORM_UNLOAD_Error

    If pval.BeforeAction = True Then
    ElseIf pval.BeforeAction = False Then
        RemoveForms oFormUniqueID01
        Set oForm01 = Nothing
        Set oMat01 = Nothing
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
'            If (PS_GA032_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_GA032A.RemoveRecord(oDS_PS_GA032A.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PS_GA032_Add_MatrixRow(0)
            Else
                If Trim(oDS_PS_GA032A.GetValue("U_ColReg01", oMat01.RowCount - 1)) <> "" Then
                    Call PS_GA032_Add_MatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    
    Exit Sub
    
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_GA032_CreateItems() As Boolean
On Error GoTo PS_GA032_CreateItems_Error

    Call oForm01.Freeze(True)
    
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Set oDS_PS_GA032A = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '사업장
    Call oForm01.DataSources.UserDataSources.Add("BPLID", dt_SHORT_TEXT, 5)
    Call oForm01.Items("BPLID").Specific.DataBind.SetBound(True, "", "BPLID")
    
    '기준년도
    Call oForm01.DataSources.UserDataSources.Add("StdYear", dt_SHORT_TEXT, 4)
    Call oForm01.Items("StdYear").Specific.DataBind.SetBound(True, "", "StdYear")
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    
    Exit Function
    
PS_GA032_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA032_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'/콤보박스 set
Sub PS_GA032_ComboBox_Setting()
On Error GoTo PS_GA032_ComboBox_Setting_Error

    Dim oCombo       As SAPbouiCOM.ComboBox
    Dim sQry         As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    
    '사업장
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLID").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    
    '계획구분
    Call oMat01.Columns("PlanCls").ValidValues.Add("01", "최초계획")
    Call oMat01.Columns("PlanCls").ValidValues.Add("02", "수정계획")
    
    Call oForm01.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
    
PS_GA032_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "PS_GA032_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_GA032_Initial_Setting()
On Error GoTo PS_GA032_Initial_Setting_Error
    '사업장 사용자의 소속 사업장 선택
    Call oForm01.Items("BPLID").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
 
    '날짜 설정
'    oForm01.Items("ToDt").Specific.VALUE = Format(Date, "YYYY") + "1231"
'    oForm01.Items("FrDt").Specific.VALUE = Format(Now, "YYYY") + "0101"
    oForm01.Items("StdYear").Specific.VALUE = Format(Now, "YYYY")

    '년도, 사업장 코드는 Hidden 처리
'    oMat01.Columns("StdYear").Visible = False
'    oMat01.Columns("BPLID").Visible = False

    Exit Sub
PS_GA032_Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA032_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_GA032_CF_ChooseFromList()
On Error GoTo PS_GA032_CF_ChooseFromList_Error

    Exit Sub
    
PS_GA032_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA032_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA032_EnableMenus()
On Error GoTo PS_GA032_EnableMenus_Error
    
    Exit Sub
    
PS_GA032_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA032_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA032_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_GA032_SetDocument_Error

    If (oFromDocEntry01 = "") Then
        Call PS_GA032_FormItemEnabled
        '//Call PS_GA032_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_GA032_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    
    Exit Sub
    
PS_GA032_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA032_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA032_FormResize()
On Error GoTo PS_GA032_FormResize_Error
    
    Call oMat01.AutoResizeColumns
    
    Exit Sub
    
PS_GA032_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA032_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub PS_GA032_FormReset()
'******************************************************************************
'Function ID : PS_GA032_FormReset()
'해당모듈    : PS_GA032
'기능        : 화면 초기화
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_GA032_FormReset_Error

    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    
    Dim User_BPLId As String
    User_BPLId = MDC_PS_Common.User_BPLId

    '헤더 초기화
    oForm01.DataSources.UserDataSources("StdYear").VALUE = Format(Now, "YYYY") '기준년도(등록)
    
    '라인 초기화
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    Call PS_GA032_Add_MatrixRow(0, True)
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)

    Exit Sub

PS_GA032_FormReset_Error:
    Call oForm01.Freeze(False)
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PS_GA032_FormReset_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub PS_GA032_CheckAll()
On Error GoTo PS_GA032_CheckAll_Error

    Dim CheckType As String
    Dim loopCount As Integer
    
    Call oForm01.Freeze(True)
    CheckType = "Y"
    
    Call oMat01.FlushToDataSource
    
    For loopCount = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PS_GA032A.GetValue("U_ColReg01", loopCount)) = "N" Then
            CheckType = "N"
            Exit For
        End If
    Next
    
    For loopCount = 0 To oMat01.VisualRowCount - 1
        oDS_PS_GA032A.offset = loopCount
        If CheckType = "N" Then
            Call oDS_PS_GA032A.setValue("U_ColReg01", loopCount, "Y")
        Else
            Call oDS_PS_GA032A.setValue("U_ColReg01", loopCount, "N")
        End If
    Next
    
    Call oMat01.LoadFromDataSource
    
    Call oForm01.Freeze(False)

    Exit Sub
    
PS_GA032_CheckAll_Error:
    Call oForm01.Freeze(False)
    Call Sbo_Application.SetStatusBarMessage("PS_GA032_CheckAll_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)
End Sub

Private Sub PS_GA032_Print_Report01()
On Error GoTo PS_GA032_Print_Report01_Error
    
    Dim DocNum     As String
    Dim ErrNum     As Integer
    Dim WinTitle   As String
    Dim ReportName As String
    Dim sQry       As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim CLTCOD As String
    Dim DestNo1 As String
    Dim DestNo2 As String
    
     '/ ODBC 연결 체크
    Call ConnectODBC
        
    '//인자 MOVE , Trim 시키기..
    CLTCOD = Trim(oForm01.Items("CLTCOD").Specific.VALUE)
    DestNo1 = Trim(oForm01.Items("DestNo1").Specific.VALUE)
    DestNo2 = Trim(oForm01.Items("DestNo2").Specific.VALUE)
    
    WinTitle = "[PS_GA032] 공용증"
    
    If CLTCOD = "1" Then '창원
        ReportName = "PS_GA032_01.rpt"
    ElseIf CLTCOD = "2" Then '동래
        ReportName = "PS_GA032_02.rpt"
    ElseIf CLTCOD = "3" Then '사상
        ReportName = "PS_GA032_03.rpt"
    End If
    
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
    sQry = "EXEC [PS_GA032_90] '" & CLTCOD & "','" & DestNo1 & "','" & DestNo2 & "'"
    
'    oRecordSet.DoQuery sQry
'
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo PS_GA032_Print_Report01_Error
'    End If
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    Set oRecordSet = Nothing
    
    Exit Sub
    
PS_GA032_Print_Report01_Error:
    If ErrNum = 1 Then
        Set oRecordSet = Nothing
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PS_GA032_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
    
End Sub




