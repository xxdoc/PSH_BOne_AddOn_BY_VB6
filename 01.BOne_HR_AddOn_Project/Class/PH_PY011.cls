VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY011"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PH_PY011.cls
'//  Module : 인사관리>인사마스터관리
'//  Desc : 전문직호칭일괄변경
'//  FormType : PH_PY011
'//  Create Date(Start) : 2013.07.04
'//  Create Date(End) :
'//  Creator : Song Myoung gyu
'//  Modified Date :
'//  Modifier :
'//  Company : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PH_PY011A As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PH_PY011B As SAPbouiCOM.DBDataSource '등록라인

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

'//사용자구조체
Private Type ItemInformations
    ItemCode As String
    LotNo As String
    Quantity As Long
    OPORNo As Long
    POR1No As Long
    check As Boolean
    OPDNNo As Long
    PDN1No As Long
End Type

Private oLast_Mode&

Private ItemInformation() As ItemInformations
Private ItemInformationCount As Long

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml As String
    Dim oXmlDoc As New MSXML2.DOMDocument
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY011.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    
    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PH_PY011_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01, "PH_PY011" '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    '폼 할당
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    '//oForm.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm.Freeze True
    Call PH_PY011_CreateItems
    Call PH_PY011_ComboBox_Setting
    Call PH_PY011_CF_ChooseFromList
    Call PH_PY011_EnableMenus
    Call PH_PY011_SetDocument(oFromDocEntry01)
''    Call PH_PY011_FormResize
    

'    Call PH_PY011_Add_MatrixRow(0, True)
'    Call PH_PY011_FormItemEnabled
    
    oForm.EnableMenu ("1283"), False        '// 삭제
    oForm.EnableMenu ("1286"), False        '// 닫기
    oForm.EnableMenu ("1287"), False        '// 복제
    oForm.EnableMenu ("1285"), False        '// 복원
    oForm.EnableMenu ("1284"), False         '// 취소
    oForm.EnableMenu ("1293"), False         '// 행삭제
    oForm.EnableMenu ("1281"), False
    oForm.EnableMenu ("1282"), True

    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = oCompany.GetBusinessObject(BoRecordset)
        
'    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PH_PY011A]"
'    Call RecordSet01.DoQuery(sQry)
'    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
'        Call oDS_PH_PY011A.setValue("DocEntry", 0, 1)
'    Else
'        Call oDS_PH_PY011A.setValue("DocEntry", 0, Trim(RecordSet01.Fields(0).VALUE) + 1)
'    End If
'
'    Call PH_PY011_FormReset '폼초기화 추가(2013.01.29 송명규)
    
    oForm.Update
    oForm.Freeze False
    
    oForm.Visible = True
    Set oXmlDoc = Nothing

    '기준일자
    oForm.Items("StdDate").Specific.VALUE = Format(Now, "YYYYMM01")

    Exit Sub
LoadForm_Error:
    oForm.Update
    oForm.Freeze False
    Set oXmlDoc = Nothing
    Set oForm = Nothing
    Sbo_Application.SetStatusBarMessage "Form_Load Error:" & Err.Description, bmt_Short, True
End Sub



'/메트릭스 Row추가
Sub PH_PY011_Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PH_PY011_Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PH_PY011B.InsertRecord (oRow)
    End If
    
    oMat01.AddRow
    oDS_PH_PY011B.Offset = oRow
    oDS_PH_PY011B.setValue "U_LineNum", oRow, oRow + 1
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY011_Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "PH_PY011_Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub PH_PY011_MTX01()
'******************************************************************************
'Function ID : PH_PY011_MTX01()
'해당모듈 : PH_PY011
'기능 : 데이터 조회
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY011_MTX01_Error

    Dim i As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)

    Dim CLTCOD As String '사업장
    Dim StdDate As String '기준일자
    Dim ChCnt As Integer '변경인원 수

    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE) '사업장
    StdDate = Trim(oForm.Items("StdDate").Specific.VALUE) '기준일자
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm.Freeze(True)
    
    sQry = "                EXEC [PH_PY011_01] "
    sQry = sQry & "'" & CLTCOD & "'," '사업장
    sQry = sQry & "'" & StdDate & "'"  '기준일자

    Call oRecordSet01.DoQuery(sQry)
    
    oMat01.Clear
    oDS_PH_PY011B.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource

    If (oRecordSet01.RecordCount = 0) Then
    
        ErrNum = 1
        
        oForm.Mode = fm_ADD_MODE
        
'        Call PH_PY011_Add_MatrixRow(0, True)
'        Call PH_PY011_LoadCaption
        
        GoTo PH_PY011_MTX01_Error
        
        Exit Sub
    End If
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PH_PY011B.Size Then
            oDS_PH_PY011B.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PH_PY011B.Offset = i
        
        Call oDS_PH_PY011B.setValue("U_LineNum", i, i + 1)
        Call oDS_PH_PY011B.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("MSTCOD").VALUE)) '사번
        Call oDS_PH_PY011B.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("FullName").VALUE)) '성명
        Call oDS_PH_PY011B.setValue("U_ColReg03", i, Trim(oRecordSet01.Fields("GrpDat").VALUE)) '입사일자
        Call oDS_PH_PY011B.setValue("U_ColReg04", i, Trim(oRecordSet01.Fields("gunsok").VALUE)) '근속년수
        Call oDS_PH_PY011B.setValue("U_ColReg05", i, Trim(oRecordSet01.Fields("Callname").VALUE)) '현재코드
        Call oDS_PH_PY011B.setValue("U_ColReg06", i, Trim(oRecordSet01.Fields("Cname").VALUE)) '현재호칭
        Call oDS_PH_PY011B.setValue("U_ColReg07", i, Trim(oRecordSet01.Fields("ChCallName").VALUE)) '변경후코드
        Call oDS_PH_PY011B.setValue("U_ColReg08", i, Trim(oRecordSet01.Fields("ChCName").VALUE)) '변경후호칭
        Call oDS_PH_PY011B.setValue("U_ColReg09", i, Trim(oRecordSet01.Fields("ChYN").VALUE)) '변경후호칭
        
        If Trim(oRecordSet01.Fields("ChYN").VALUE) = "Y" Then
            ChCnt = ChCnt + 1
        End If
        
        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
        
    Next
    
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    ProgBar01.Stop
    
    oForm.Items("ChCnt").Specific.VALUE = ChCnt
    
    Call oForm.Freeze(False)
  
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY011_MTX01_Error:
'    ProgBar01.Stop
    Call oForm.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PH_PY011_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Function PH_PY011_Update() As Boolean
'******************************************************************************
'Function ID : PH_PY011_Update()
'해당모듈 : PH_PY011
'기능 : 호칭일괄변경
'인수 : 없음
'반환값 : 성공(True), 실패(False)
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY011_Update_Error

    Dim i As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)


    For i = 1 To oMat01.RowCount
        
        If oMat01.Columns("ChYN").Cells(i).Specific.VALUE = "Y" Then
    
            sQry = "                EXEC [PH_PY011_02] "
            sQry = sQry & "'" & oMat01.Columns("MSTCOD").Cells(i).Specific.VALUE & "'," '사번
            sQry = sQry & "'" & oMat01.Columns("ChCallName").Cells(i).Specific.VALUE & "'"  '변경후호칭코드
            
            Call oRecordSet01.DoQuery(sQry)
            
        End If
    
    Next

    Set oRecordSet01 = Nothing
    PH_PY011_Update = True
    Exit Function

PH_PY011_Update_Error:

    Set oRecordSet01 = Nothing
    PH_PY011_Update = True
    
    Call MDC_Com.MDC_GF_Message("PH_PY011_Update_Error:" & Err.Number & " - " & Err.Description, "E")

End Function




Private Function PH_PY011_HeaderSpaceLineDel() As Boolean
'******************************************************************************
'Function ID : PH_PY011_HeaderSpaceLineDel()
'해당모듈 : PH_PY011
'기능 : 필수입력사항 체크
'인수 : 없음
'반환값 : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY011_HeaderSpaceLineDel_Error

    Dim ErrNum As Integer
    ErrNum = 0

    Select Case True
        Case Trim(oForm.Items("DestNo1").Specific.VALUE) = "" '출장번호1
            ErrNum = 1
            GoTo PH_PY011_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("DestNo2").Specific.VALUE) = "" '출장번호2
            ErrNum = 2
            GoTo PH_PY011_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("MSTCOD").Specific.VALUE) = "" '사원번호
            ErrNum = 3
            GoTo PH_PY011_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("FrDate").Specific.VALUE) = "" '시작일자
            ErrNum = 4
            GoTo PH_PY011_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("FrTime").Specific.VALUE) = "" '시작시각
            ErrNum = 5
            GoTo PH_PY011_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("ToDate").Specific.VALUE) = "" '종료일자
            ErrNum = 6
            GoTo PH_PY011_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("ToTime").Specific.VALUE) = "" '종료시각
            ErrNum = 7
            GoTo PH_PY011_HeaderSpaceLineDel_Error
    End Select

    PH_PY011_HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY011_HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출장번호1은 필수사항입니다. 확인하세요.", "E"
        Call oForm.Items("DestNo1").CLICK(ct_Regular)
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "출장번호2는 필수사항입니다. 확인하세요.", "E"
        Call oForm.Items("DestNo2").CLICK(ct_Regular)
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "사원번호는 필수사항입니다. 확인하세요.", "E"
        Call oForm.Items("MSTCOD").CLICK(ct_Regular)
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "시작일자는 필수사항입니다. 확인하세요.", "E"
        Call oForm.Items("FrDate").CLICK(ct_Regular)
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "시작시각은 필수사항입니다. 확인하세요.", "E"
        Call oForm.Items("FrTime").CLICK(ct_Regular)
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "종료일자는 필수사항입니다. 확인하세요.", "E"
        Call oForm.Items("FrDate").CLICK(ct_Regular)
    ElseIf ErrNum = 7 Then
        MDC_Com.MDC_GF_Message "종료시각은 필수사항입니다. 확인하세요.", "E"
        Call oForm.Items("FrTime").CLICK(ct_Regular)
    End If
    PH_PY011_HeaderSpaceLineDel = False
End Function

'/ 메트릭스 필수 사항 check
Private Function PH_PY011_MatrixSpaceLineDel() As Boolean
On Error GoTo PH_PY011_MatrixSpaceLineDel_Error

    Dim i As Long
    Dim ErrNum As Integer
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim sQry As String

    Set oRecordSet01 = Nothing
    PH_PY011_MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY011_MatrixSpaceLineDel_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "라인 데이터가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 사원코드가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 시간이 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 등록일자가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 비가동코드가 없습니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "PH_PY011_MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    PH_PY011_MatrixSpaceLineDel = False
End Function

Private Sub PH_PY011_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PH_PY011_FlushToItemValue_Error

    Dim i As Integer
    Dim ErrNum As Integer
    Dim sQry As String
    Dim ItemCode As String
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim CLTCOD As String
    Dim TeamCode As String
    Dim RspCode As String
    
    Select Case oUID
        
        Case "CLTCOD"
            
            CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
            
            If oForm.Items("TeamCode").Specific.ValidValues.Count > 0 Then
                For i = oForm.Items("TeamCode").Specific.ValidValues.Count - 1 To 0 Step -1
                    Call oForm.Items("TeamCode").Specific.ValidValues.Remove(i, psk_Index)
                Next i
            End If
            
            '부서콤보세팅
            Call oForm.Items("TeamCode").Specific.ValidValues.Add("%", "전체")
            sQry = "            SELECT      U_Code AS [Code],"
            sQry = sQry & "                 U_CodeNm As [Name]"
            sQry = sQry & "  FROM       [@PS_HR200L]"
            sQry = sQry & "  WHERE      Code = '1'"
            sQry = sQry & "                 AND U_UseYN = 'Y'"
            sQry = sQry & "                 AND U_Char2 = '" & CLTCOD & "'"
            sQry = sQry & "  ORDER BY  U_Seq"
            Call MDC_SetMod.Set_ComboList(oForm.Items("TeamCode").Specific, sQry, "", False, False)
            Call oForm.Items("TeamCode").Specific.Select(0, psk_Index)
            
        Case "TeamCode"
            
            TeamCode = Trim(oForm.Items("TeamCode").Specific.VALUE)
            
            If oForm.Items("RspCode").Specific.ValidValues.Count > 0 Then
                For i = oForm.Items("RspCode").Specific.ValidValues.Count - 1 To 0 Step -1
                    Call oForm.Items("RspCode").Specific.ValidValues.Remove(i, psk_Index)
                Next i
            End If
            
            '담당콤보세팅
            Call oForm.Items("RspCode").Specific.ValidValues.Add("%", "전체")
            sQry = "            SELECT      U_Code AS [Code],"
            sQry = sQry & "                 U_CodeNm As [Name]"
            sQry = sQry & "  FROM       [@PS_HR200L]"
            sQry = sQry & "  WHERE      Code = '2'"
            sQry = sQry & "                 AND U_UseYN = 'Y'"
            sQry = sQry & "                 AND U_Char1 = '" & TeamCode & "'"
            sQry = sQry & "  ORDER BY  U_Seq"
            Call MDC_SetMod.Set_ComboList(oForm.Items("RspCode").Specific, sQry, "", False, False)
            Call oForm.Items("RspCode").Specific.Select(0, psk_Index)
        
        Case "RspCode"
            
            TeamCode = Trim(oForm.Items("TeamCode").Specific.VALUE)
            RspCode = Trim(oForm.Items("RspCode").Specific.VALUE)
            
            If oForm.Items("ClsCode").Specific.ValidValues.Count > 0 Then
                For i = oForm.Items("ClsCode").Specific.ValidValues.Count - 1 To 0 Step -1
                    Call oForm.Items("ClsCode").Specific.ValidValues.Remove(i, psk_Index)
                Next i
            End If
            
            '반콤보세팅
            Call oForm.Items("ClsCode").Specific.ValidValues.Add("%", "전체")
            sQry = "            SELECT      U_Code AS [Code],"
            sQry = sQry & "                 U_CodeNm As [Name]"
            sQry = sQry & "  FROM       [@PS_HR200L]"
            sQry = sQry & "  WHERE      Code = '9'"
            sQry = sQry & "                 AND U_UseYN = 'Y'"
            sQry = sQry & "                 AND U_Char1 = '" & RspCode & "'"
            sQry = sQry & "                 AND U_Char2 = '" & TeamCode & "'"
            sQry = sQry & "  ORDER BY  U_Seq"
            Call MDC_SetMod.Set_ComboList(oForm.Items("ClsCode").Specific, sQry, "", False, False)
            Call oForm.Items("ClsCode").Specific.Select(0, psk_Index)
        
    End Select

   Set oRecordSet01 = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
PH_PY011_FlushToItemValue_Error:

    Set oRecordSet01 = Nothing

    Call MDC_Com.MDC_GF_Message("PH_PY011_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")

End Sub

'/폼의 아이템 사용지정
Sub PH_PY011_FormItemEnabled()
On Error GoTo PH_PY011_FormItemEnabled_Error
    If (oForm.Mode = fm_ADD_MODE) Then
    
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
'        Call CLTCOD_Select(oForm, "SCLTCOD")
        
'        oMat01.Columns("ItemCode").Cells(1).Click ct_Regular
'        oForm.Items("ItemCode").Enabled = True

    ElseIf (oForm.Mode = fm_FIND_MODE) Then
    
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
'        Call CLTCOD_Select(oForm, "SCLTCOD")
        
'        oForm.Items("ItemCode").Enabled = True

    ElseIf (oForm.Mode = fm_OK_MODE) Then
    
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
'        Call CLTCOD_Select(oForm, "SCLTCOD")
        
    End If
    
    Exit Sub
    
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY011_FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "PH_PY011_FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/아이템 변경 이벤트
Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Raise_FormItemEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_FormItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_FormMenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_FormMenuEvent_Error

    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = oCompany.GetBusinessObject(BoRecordset)

    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
            Case "1282": '추가
                '/추가버튼 클릭시 메트릭스 insertrow
                
'                oMat01.Clear
'                oMat01.FlushToDataSource
'                oMat01.LoadFromDataSource
                
'                oForm.Mode = fm_ADD_MODE
'                BubbleEvent = False
'                Call PH_PY011_LoadCaption
                
                'oForm.Items("GCode").Click ct_Regular
                
                
                Exit Sub
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
            
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
                Call PH_PY011_Add_MatrixRow(oMat01.VisualRowCount)
            
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
                '//Call PH_PY011_FormItemEnabled '//UDO방식
            Case "1282": '추가
'                oMat01.Clear
'                oDS_PH_PY011A.Clear

'                Call PH_PY011_LoadCaption
'                Call PH_PY011_FormItemEnabled
                '//Call PH_PY011_FormItemEnabled '//UDO방식
                '//Call PH_PY011_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PH_PY011_FormItemEnabled
                
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 이후 처리
                Call oForm.Freeze(True)
                Call oDS_PH_PY011B.RemoveRecord(oDS_PH_PY011B.Size - 1)
                Call oMat01.LoadFromDataSource
                Call oForm.Freeze(False)
            
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
    
        If pval.ItemUID = "PH_PY011" Then
            If oForm.Mode = fm_ADD_MODE Then
            ElseIf oForm.Mode = fm_UPDATE_MODE Then
            ElseIf oForm.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "BtnUpdate" Then '수정 버튼
        
'            If oForm.Mode = fm_ADD_MODE Then
            
'                If PH_PY011_HeaderSpaceLineDel() = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
'
''                If PH_PY011_DataCheck() = False Then
''                    BubbleEvent = False
''                    Exit Sub
''                End If
'
''                If PH_PY011_AddData() = False Then
''                    BubbleEvent = False
''                    Exit Sub
''                End If
'
''                Call PH_PY011_FormReset
'                oForm.Mode = fm_ADD_MODE
'
''                Call PH_PY011_LoadCaption
'                Call PH_PY011_MTX01
'
'                oLast_Mode = oForm.Mode
                
'            ElseIf oForm.Mode = fm_UPDATE_MODE Then
'
'                If PH_PY011_HeaderSpaceLineDel = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
            
'                If PH_PY011_DataCheck() = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
            
'                If PH_PY011_UpdateData() = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
'
'                Call PH_PY011_FormReset
'                oForm.Mode = fm_ADD_MODE
'
'                Call PH_PY011_LoadCaption
'                Call PH_PY011_MTX01
                
'                oForm.Items("GCode").Click ct_Regular
'            End If
            
            If Sbo_Application.MessageBox("수정하시겠습니까?", "1", "예", "아니오") = "1" Then
            
                If PH_PY011_Update = True Then
                
                    Call MDC_GF_Message("수정을 완료하였습니다.", "S")
                    
                Else
                
                    Call MDC_GF_Message("수정하는 동안 오류가 발생했습니다. 관리자에게 문의하세요.", "E")
                    
                End If
                
            
            Else
            
'                BubblEvent = False
'                Exit Sub
            
            End If
            
        ElseIf pval.ItemUID = "BtnSearch" Then '/조회
            
'            Call PH_PY011_FormReset
'            oForm.Mode = fm_ADD_MODE '/fm_VIEW_MODE
            
'            Call PH_PY011_LoadCaption
            Call PH_PY011_MTX01
            
'        ElseIf pval.ItemUID = "BtnDelete" Then '/삭제
'
'            If Sbo_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", "1", "예", "아니오") = "1" Then
'
'                Call PH_PY011_DeleteData
'                Call PH_PY011_FormReset
'                oForm.Mode = fm_ADD_MODE '/fm_VIEW_MODE
'
'                Call PH_PY011_LoadCaption
'                Call PH_PY011_MTX01
'
'            Else
'
'            End If
        
        ElseIf pval.ItemUID = "BtnPrint" Then
        
            Call PH_PY011_Print_Report01
        
        End If
  
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PH_PY011" Then
            If oForm.Mode = fm_ADD_MODE Then
            ElseIf oForm.Mode = fm_UPDATE_MODE Then
            ElseIf oForm.Mode = fm_OK_MODE Then
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
    
        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pval, BubbleEvent, "MSTCOD", "") '사번
        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pval, BubbleEvent, "ShiftDatCd", "") '근무형태
        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pval, BubbleEvent, "GNMUJOCd", "") '근무조
        
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
                
'                Call oForm.Freeze(True)
'
'                'DataSource를 이용하여 각 컨트롤에 값을 출력
'                Call oDS_PH_PY011A.setValue("DocEntry", 0, oMat01.Columns("DocEntry").Cells(pval.Row).Specific.VALUE) '관리번호
'                Call oDS_PH_PY011A.setValue("U_CLTCOD", 0, oMat01.Columns("CLTCOD").Cells(pval.Row).Specific.VALUE) '사업장
'                Call oDS_PH_PY011A.setValue("U_DestNo1", 0, oMat01.Columns("DestNo1").Cells(pval.Row).Specific.VALUE) '출장번호1
'                Call oDS_PH_PY011A.setValue("U_DestNo2", 0, oMat01.Columns("DestNo2").Cells(pval.Row).Specific.VALUE) '출장번호2
'                Call oDS_PH_PY011A.setValue("U_MSTCOD", 0, oMat01.Columns("MSTCOD").Cells(pval.Row).Specific.VALUE) '사원번호
'                Call oDS_PH_PY011A.setValue("U_MSTNAM", 0, oMat01.Columns("MSTNAM").Cells(pval.Row).Specific.VALUE) '사원성명
'                Call oDS_PH_PY011A.setValue("U_Destinat", 0, oMat01.Columns("Destinat").Cells(pval.Row).Specific.VALUE) '출장지
'                Call oDS_PH_PY011A.setValue("U_Dest2", 0, oMat01.Columns("Dest2").Cells(pval.Row).Specific.VALUE) '출장지상세
'                Call oDS_PH_PY011A.setValue("U_CoCode", 0, oMat01.Columns("CoCode").Cells(pval.Row).Specific.VALUE) '작번
'                Call oDS_PH_PY011A.setValue("U_FrDate", 0, Replace(oMat01.Columns("FrDate").Cells(pval.Row).Specific.VALUE, ".", "")) '시작일자
'                Call oDS_PH_PY011A.setValue("U_FrTime", 0, oMat01.Columns("FrTime").Cells(pval.Row).Specific.VALUE) '시작시각
'                Call oDS_PH_PY011A.setValue("U_ToDate", 0, Replace(oMat01.Columns("ToDate").Cells(pval.Row).Specific.VALUE, ".", "")) '종료일자
'                Call oDS_PH_PY011A.setValue("U_ToTime", 0, oMat01.Columns("ToTime").Cells(pval.Row).Specific.VALUE) '종료시각
'                Call oDS_PH_PY011A.setValue("U_Object", 0, oMat01.Columns("Object").Cells(pval.Row).Specific.VALUE) '목적
'                Call oDS_PH_PY011A.setValue("U_Comments", 0, oMat01.Columns("Comments").Cells(pval.Row).Specific.VALUE) '비고
'                Call oDS_PH_PY011A.setValue("U_RegCls", 0, oMat01.Columns("RegCls").Cells(pval.Row).Specific.VALUE) '등록구분
'                Call oDS_PH_PY011A.setValue("U_ObjCls", 0, oMat01.Columns("ObjCls").Cells(pval.Row).Specific.VALUE) '목적구분
'                Call oDS_PH_PY011A.setValue("U_DestCode", 0, oMat01.Columns("DestCode").Cells(pval.Row).Specific.VALUE) '출장지역
'                Call oDS_PH_PY011A.setValue("U_DestDiv", 0, oMat01.Columns("DestDiv").Cells(pval.Row).Specific.VALUE) '출장구분
'                Call oDS_PH_PY011A.setValue("U_Vehicle", 0, oMat01.Columns("Vehicle").Cells(pval.Row).Specific.VALUE) '차량구분
'                Call oDS_PH_PY011A.setValue("U_FuelPrc", 0, oMat01.Columns("FuelPrc").Cells(pval.Row).Specific.VALUE) '1L단가
'                Call oDS_PH_PY011A.setValue("U_FuelType", 0, oMat01.Columns("FuelType").Cells(pval.Row).Specific.VALUE) '유류
'                Call oDS_PH_PY011A.setValue("U_Distance", 0, oMat01.Columns("Distance").Cells(pval.Row).Specific.VALUE) '거리
'                Call oDS_PH_PY011A.setValue("U_TransExp", 0, oMat01.Columns("TransExp").Cells(pval.Row).Specific.VALUE) '교통비
'                Call oDS_PH_PY011A.setValue("U_DayExp", 0, oMat01.Columns("DayExp").Cells(pval.Row).Specific.VALUE) '일비
'                Call oDS_PH_PY011A.setValue("U_FoodNum", 0, oMat01.Columns("FoodNum").Cells(pval.Row).Specific.VALUE) '식수
'                Call oDS_PH_PY011A.setValue("U_FoodExp", 0, oMat01.Columns("FoodExp").Cells(pval.Row).Specific.VALUE) '식비
'                Call oDS_PH_PY011A.setValue("U_ParkExp", 0, oMat01.Columns("ParkExp").Cells(pval.Row).Specific.VALUE) '주차비
'                Call oDS_PH_PY011A.setValue("U_TollExp", 0, oMat01.Columns("TollExp").Cells(pval.Row).Specific.VALUE) '도로비
'                Call oDS_PH_PY011A.setValue("U_TotalExp", 0, oMat01.Columns("TotalExp").Cells(pval.Row).Specific.VALUE) '합계
'
'                oForm.Mode = fm_UPDATE_MODE
'                Call PH_PY011_LoadCaption
'
'                Call oForm.Freeze(False)
                
            End If
        End If
    ElseIf pval.BeforeAction = False Then
    
    End If
    
    Exit Sub
    
Raise_EVENT_CLICK_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_COMBO_SELECT(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_COMBO_SELECT_Error

    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
    
        Call PH_PY011_FlushToItemValue(pval.ItemUID)
        
    End If

    Exit Sub
Raise_EVENT_COMBO_SELECT_Error:
    Call oForm.Freeze(False)
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

    Call oForm.Freeze(True)
    
    If pval.BeforeAction = True Then
    
        If pval.ItemChanged = True Then
        
            If (pval.ItemUID = "Mat01") Then
'                If (pval.ColUID = "ItemCode") Then
'                    '//기타작업
'                    Call oDS_PH_PY011B.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                    If oMat01.RowCount = pval.Row And Trim(oDS_PH_PY011B.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
'                        PH_PY011_AddMatrixRow (pval.Row)
'                    End If
'                Else
'                    Call oDS_PH_PY011B.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                End If
            Else
                
                Call PH_PY011_FlushToItemValue(pval.ItemUID)
                
                If pval.ItemUID = "MSTCOD" Then
                
                    oForm.Items("MSTNAM").Specific.VALUE = MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm.Items("MSTCOD").Specific.VALUE & "'") '성명
                    
                ElseIf pval.ItemUID = "ShiftDatCd" Then
                
                    oForm.Items("ShiftDatNm").Specific.VALUE = MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L] AS T0", "'" & oForm.Items("ShiftDatCd").Specific.VALUE & "'", " AND T0.Code = 'P154' AND T0.U_UseYN = 'Y'") '근무형태
                    
                ElseIf pval.ItemUID = "GNMUJOCd" Then
                
                    oForm.Items("GNMUJONm").Specific.VALUE = MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L] AS T0", "'" & oForm.Items("GNMUJOCd").Specific.VALUE & "'", " AND T0.Code = 'P155' AND T0.U_UseYN = 'Y'") '근무조
                
                End If
                
            End If
'            oMat01.LoadFromDataSource
'            oMat01.AutoResizeColumns
'            oForm.Update
        End If
        
    ElseIf pval.BeforeAction = False Then

    End If
    
    Call oForm.Freeze(False)
    
    Exit Sub
    
Raise_EVENT_VALIDATE_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_VALIDATE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_MATRIX_LOAD(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_MATRIX_LOAD_Error
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        Call PH_PY011_FormItemEnabled
        '//Call PH_PY011_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PH_PY011_FormResize
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
'            oForm.DataSources.UserDataSources("ItemCode").Value = oDataTable01.Columns(0).Cells(0).Value
'            Set oDataTable01 = Nothing
'        End If
'        If (pval.ItemUID = "CardCode" Or pval.ItemUID = "CardName") Then
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY011A", "U_CardCode,U_CardName")
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
        Set oForm = Nothing
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
'            If (PH_PY011_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PH_PY011A.RemoveRecord(oDS_PH_PY011A.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PH_PY011_Add_MatrixRow(0)
            Else
                If Trim(oDS_PH_PY011A.GetValue("U_CntcCode", oMat01.RowCount - 1)) <> "" Then
                    Call PH_PY011_Add_MatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PH_PY011_CreateItems() As Boolean
On Error GoTo PH_PY011_CreateItems_Error

    Call oForm.Freeze(True)
    
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
'    Set oDS_PH_PY011A = oForm.DataSources.DBDataSources("@PH_PY011A")
    Set oDS_PH_PY011B = oForm.DataSources.DBDataSources("@PS_USERDS01")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '사업장
    Call oForm.DataSources.UserDataSources.Add("CLTCOD", dt_SHORT_TEXT, 10)
    Call oForm.Items("CLTCOD").Specific.DataBind.SetBound(True, "", "CLTCOD")
    
    '기준일자
    Call oForm.DataSources.UserDataSources.Add("StdDate", dt_DATE)
    Call oForm.Items("StdDate").Specific.DataBind.SetBound(True, "", "StdDate")
    
    '변경인원 수
    Call oForm.DataSources.UserDataSources.Add("ChCnt", dt_SHORT_TEXT, 10)
    Call oForm.Items("ChCnt").Specific.DataBind.SetBound(True, "", "ChCnt")
    
    Set oRecordSet01 = Nothing
    Call oForm.Freeze(False)
    
    Exit Function
PH_PY011_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY011_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'/콤보박스 set
Sub PH_PY011_ComboBox_Setting()
On Error GoTo PH_PY011_ComboBox_Setting_Error

'    Dim oCombo As SAPbouiCOM.ComboBox
'    Dim sQry As String
'    Dim oRecordSet01 As SAPbobsCOM.Recordset
'
'    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
'
'    Call oForm.Freeze(True)
'
'    Call oForm.Freeze(False)
'    Set oCombo = Nothing
'    Set oRecordSet01 = Nothing
    
    Exit Sub
PH_PY011_ComboBox_Setting_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY011_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY011_CF_ChooseFromList()
On Error GoTo PH_PY011_CF_ChooseFromList_Error

    Exit Sub
PH_PY011_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY011_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY011_EnableMenus()
On Error GoTo PH_PY011_EnableMenus_Error
    
    Exit Sub
PH_PY011_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY011_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY011_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PH_PY011_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PH_PY011_FormItemEnabled
        '//Call PH_PY011_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm.Mode = fm_FIND_MODE
'        Call PH_PY011_FormItemEnabled
'        oForm.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm.Items("1").Click ct_Regular
    End If
    Exit Sub
PH_PY011_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY011_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY011_FormResize()
On Error GoTo PH_PY011_FormResize_Error
    
    Call oMat01.AutoResizeColumns
    
    Exit Sub
PH_PY011_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY011_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PH_PY011_Print_Report01()

    Dim DocNum As String
    Dim ErrNum As Integer
    Dim WinTitle As String
    Dim ReportName As String
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY011_Print_Report01_Error
    
    Dim CLTCOD As String
    Dim DestNo1 As String
    Dim DestNo2 As String
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
     '/ ODBC 연결 체크
    If ConnectODBC = False Then
        GoTo PH_PY011_Print_Report01_Error
    End If
    
    '//인자 MOVE , Trim 시키기..
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    DestNo1 = Trim(oForm.Items("DestNo1").Specific.VALUE)
    DestNo2 = Trim(oForm.Items("DestNo2").Specific.VALUE)
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

    WinTitle = "[PH_PY011] 공용증"
    
    If CLTCOD = "1" Then '창원
        ReportName = "PH_PY011_01.rpt"
    ElseIf CLTCOD = "2" Then '동래
        ReportName = "PH_PY011_02.rpt"
    ElseIf CLTCOD = "3" Then '사상
        ReportName = "PH_PY011_03.rpt"
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
    sQry = "EXEC [PH_PY011_90] '" & CLTCOD & "','" & DestNo1 & "','" & DestNo2 & "'"
    
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo PH_PY011_Print_Report01_Error
    End If
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    Set oRecordSet = Nothing
    Exit Sub
    
PH_PY011_Print_Report01_Error:
    If ErrNum = 1 Then
        Set oRecordSet = Nothing
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY011_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
    
End Sub







