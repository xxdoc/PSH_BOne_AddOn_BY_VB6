VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_SD057"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_SD057.cls
'//  Module : 판매관리 > 애드온 > 기성매출관리
'//  Desc : 작번별미청구금액등록
'//  FormType : PS_SD057
'//  Create Date(Start) : 2014.06.10
'//  Create Date(End) : 2014.06.10
'//  Creator : Song Myoung gyu
'//  Modified Date :
'//  Modifier :
'//  Company : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_SD057H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_SD057L As SAPbouiCOM.DBDataSource '등록라인

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01 As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_SD057.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    
    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_SD057_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
     
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    Call oForm01.Freeze(True)
    
    Call PS_SD057_CreateItems
    Call PS_SD057_ComboBox_Setting
    Call PS_SD057_CF_ChooseFromList
    Call PS_SD057_EnableMenus
    Call PS_SD057_SetDocument(oFromDocEntry01)
    Call PS_SD057_FormResize

'    Call PS_SD057_Add_MatrixRow(0, True)
    Call PS_SD057_LoadCaption
    Call PS_SD057_FormItemEnabled
    Call PS_SD057_Initial_Setting
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1285"), False        '// 복원
    oForm01.EnableMenu ("1284"), False         '// 취소
    oForm01.EnableMenu ("1293"), False         '// 행삭제
    oForm01.EnableMenu ("1281"), False
    oForm01.EnableMenu ("1282"), True
    
    Call oForm01.Update
    Call oForm01.Freeze(False)
    
    oForm01.Visible = True
    Set oXmlDoc01 = Nothing
 
    Exit Sub
LoadForm_Error:
    oForm01.Update
    oForm01.Freeze False
    Set oXmlDoc01 = Nothing
    Set oForm01 = Nothing
    Sbo_Application.SetStatusBarMessage "Form_Load Error:" & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD057_LoadCaption()
'******************************************************************************
'Function ID : PS_SD057_LoadCaption()
'해당모듈 : PS_SD057
'기능 : Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_SD057_LoadCaption_Error

    Call oForm01.Freeze(True)

    If oForm01.Mode = fm_ADD_MODE Then
    
        oForm01.Items("BtnAdd").Specific.Caption = "추가"

'    ElseIf oForm01.Mode = fm_OK_MODE Then
'        oForm01.Items("BtnAdd").Specific.Caption = "확인"

    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
    
        oForm01.Items("BtnAdd").Specific.Caption = "수정"
        
    End If
    
    Call oForm01.Freeze(False)

    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_SD057_LoadCaption_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/메트릭스 Row추가
Sub PS_SD057_Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_SD057_Add_MatrixRow_Error

    If RowIserted = False Then '//행추가여부
        oDS_PS_SD057L.InsertRecord (oRow)
    End If
    
    oMat01.AddRow
    oDS_PS_SD057L.Offset = oRow
    Call oDS_PS_SD057L.setValue("U_LineNum", oRow, oRow + 1)
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_SD057_Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "PS_SD057_Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub PS_SD057_MTX01(ByVal pItemUID As String)
'******************************************************************************
'Function ID : PS_SD057_MTX01()
'해당모듈 : PS_SD057
'기능 : 데이터 조회
'인수 : pItemUID(아이템UID)
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_SD057_MTX01_Error

    Dim i As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim OrdNum As String '작번
    Dim StdYM As String '기준년월
    Dim StdCnt As String '기준회차
    Dim CntcCode As String '사용자 사번
        
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    If pItemUID = "BtnSearch1" Then
    
        OrdNum = Trim(oForm01.Items("OrdNum").Specific.VALUE) '작번
        StdYM = Trim(oForm01.Items("StdYM").Specific.VALUE) '기준년월
        StdCnt = Trim(oForm01.Items("StdCnt").Specific.VALUE) '기준회차
        CntcCode = MDC_PS_Common.User_MSTCOD '사용자사번
        
        sQry = "         EXEC [PS_SD057_01] '"
        sQry = sQry & OrdNum & "','"
        sQry = sQry & StdYM & "','"
        sQry = sQry & StdCnt & "','"
        sQry = sQry & CntcCode & "'"
        
        Call oRecordSet01.DoQuery(sQry)
        
        oMat01.Clear
        oDS_PS_SD057L.Clear
        oMat01.FlushToDataSource
        oMat01.LoadFromDataSource
    
        If (oRecordSet01.RecordCount = 0) Then
        
            ErrNum = 1
            
            oForm01.Mode = fm_ADD_MODE
            
    '        Call PS_SD057_Add_MatrixRow(0, True)
            Call PS_SD057_LoadCaption
            
            GoTo PS_SD057_MTX01_Error
            
            Exit Sub
        End If
        
        For i = 0 To oRecordSet01.RecordCount - 1
            If i + 1 > oDS_PS_SD057L.Size Then
                oDS_PS_SD057L.InsertRecord (i)
            End If
            
            oMat01.AddRow
            oDS_PS_SD057L.Offset = i
            
            Call oDS_PS_SD057L.setValue("U_LineNum", i, i + 1)
            Call oDS_PS_SD057L.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("Check").VALUE)) '선택
            Call oDS_PS_SD057L.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("StdYM").VALUE)) '기준년월
            Call oDS_PS_SD057L.setValue("U_ColReg03", i, Trim(oRecordSet01.Fields("StdCnt").VALUE)) '기준회차
            Call oDS_PS_SD057L.setValue("U_ColReg04", i, Trim(oRecordSet01.Fields("OrdNum").VALUE)) '품목코드(작번)
            Call oDS_PS_SD057L.setValue("U_ColReg05", i, Trim(oRecordSet01.Fields("FrgnName").VALUE)) '품목명(작번)
            Call oDS_PS_SD057L.setValue("U_ColReg06", i, Trim(oRecordSet01.Fields("Spec").VALUE)) '규격(작번)
            Call oDS_PS_SD057L.setValue("U_ColReg07", i, Trim(oRecordSet01.Fields("InOut").VALUE)) '자체/외주
            Call oDS_PS_SD057L.setValue("U_ColSum01", i, Trim(oRecordSet01.Fields("MatAmt").VALUE)) '견적자재비
            Call oDS_PS_SD057L.setValue("U_ColSum02", i, Trim(oRecordSet01.Fields("TotAmt").VALUE)) '누계자재비
            Call oDS_PS_SD057L.setValue("U_ColSum04", i, Trim(oRecordSet01.Fields("PreEA").VALUE)) '직전미청구금액
            Call oDS_PS_SD057L.setValue("U_ColSum03", i, Trim(oRecordSet01.Fields("ExpAmt").VALUE)) '미청구금액
            Call oDS_PS_SD057L.setValue("U_ColTxt01", i, Trim(oRecordSet01.Fields("Comment").VALUE)) '비고
            Call oDS_PS_SD057L.setValue("U_ColReg10", i, Trim(oRecordSet01.Fields("CreateUser").VALUE)) '등록자(사번)
            Call oDS_PS_SD057L.setValue("U_ColDt01", i, Format(Trim(oRecordSet01.Fields("CreateDate").VALUE), "YYYYMMDD")) '등록일자
            Call oDS_PS_SD057L.setValue("U_ColReg12", i, Trim(oRecordSet01.Fields("UpdateUser").VALUE)) '수정자(사번)
            Call oDS_PS_SD057L.setValue("U_ColDt02", i, Format(Trim(oRecordSet01.Fields("UpdateDate").VALUE), "YYYYMMDD")) '수정일자
            
            oRecordSet01.MoveNext
            ProgBar01.VALUE = ProgBar01.VALUE + 1
            ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
            
        Next
        
        oMat01.LoadFromDataSource
        oMat01.AutoResizeColumns
        ProgBar01.Stop
    
    ElseIf pItemUID = "BtnSearch2" Then
    
        OrdNum = Trim(oForm01.Items("OrdNumS").Specific.VALUE) '작번
        StdYM = Trim(oForm01.Items("StdYMS").Specific.VALUE) '기준년월
        StdCnt = Trim(oForm01.Items("StdCntS").Specific.VALUE) '기준회차
        CntcCode = MDC_PS_Common.User_MSTCOD '사용자사번
    
        sQry = "         EXEC [PS_SD057_02] '"
        sQry = sQry & OrdNum & "','"
        sQry = sQry & StdYM & "','"
        sQry = sQry & StdCnt & "','"
        sQry = sQry & CntcCode & "'"
        
        Call oRecordSet01.DoQuery(sQry)
        
        oMat01.Clear
        oDS_PS_SD057L.Clear
        oMat01.FlushToDataSource
        oMat01.LoadFromDataSource
    
        If (oRecordSet01.RecordCount = 0) Then
        
            ErrNum = 1
            
            oForm01.Mode = fm_ADD_MODE
            
    '        Call PS_SD057_Add_MatrixRow(0, True)
            Call PS_SD057_LoadCaption
            
            GoTo PS_SD057_MTX01_Error
            
            Exit Sub
        End If
        
        For i = 0 To oRecordSet01.RecordCount - 1
            If i + 1 > oDS_PS_SD057L.Size Then
                oDS_PS_SD057L.InsertRecord (i)
            End If
            
            oMat01.AddRow
            oDS_PS_SD057L.Offset = i
            
            Call oDS_PS_SD057L.setValue("U_LineNum", i, i + 1)
            Call oDS_PS_SD057L.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("Check").VALUE)) '선택
            Call oDS_PS_SD057L.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("StdYM").VALUE)) '기준년월
            Call oDS_PS_SD057L.setValue("U_ColReg03", i, Trim(oRecordSet01.Fields("StdCnt").VALUE)) '기준회차
            Call oDS_PS_SD057L.setValue("U_ColReg04", i, Trim(oRecordSet01.Fields("OrdNum").VALUE)) '품목코드(작번)
            Call oDS_PS_SD057L.setValue("U_ColReg05", i, Trim(oRecordSet01.Fields("FrgnName").VALUE)) '품목명(작번)
            Call oDS_PS_SD057L.setValue("U_ColReg06", i, Trim(oRecordSet01.Fields("Spec").VALUE)) '규격(작번)
            Call oDS_PS_SD057L.setValue("U_ColReg07", i, Trim(oRecordSet01.Fields("InOut").VALUE)) '자체/외주
            Call oDS_PS_SD057L.setValue("U_ColSum01", i, Trim(oRecordSet01.Fields("MatAmt").VALUE)) '견적자재비
            Call oDS_PS_SD057L.setValue("U_ColSum02", i, Trim(oRecordSet01.Fields("TotAmt").VALUE)) '누계자재비
            Call oDS_PS_SD057L.setValue("U_ColSum04", i, Trim(oRecordSet01.Fields("PreEA").VALUE)) '직전미청구금액
            Call oDS_PS_SD057L.setValue("U_ColSum03", i, Trim(oRecordSet01.Fields("ExpAmt").VALUE)) '미청구금액
            Call oDS_PS_SD057L.setValue("U_ColTxt01", i, Trim(oRecordSet01.Fields("Comment").VALUE)) '비고
            Call oDS_PS_SD057L.setValue("U_ColReg10", i, Trim(oRecordSet01.Fields("CreateUser").VALUE)) '등록자(사번)
            Call oDS_PS_SD057L.setValue("U_ColDt01", i, Format(Trim(oRecordSet01.Fields("CreateDate").VALUE), "YYYYMMDD")) '등록일자
            Call oDS_PS_SD057L.setValue("U_ColReg12", i, Trim(oRecordSet01.Fields("UpdateUser").VALUE)) '수정자(사번)
            Call oDS_PS_SD057L.setValue("U_ColDt02", i, Format(Trim(oRecordSet01.Fields("UpdateDate").VALUE), "YYYYMMDD")) '수정일자

            oRecordSet01.MoveNext
            ProgBar01.VALUE = ProgBar01.VALUE + 1
            ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
            
        Next
        
        oMat01.LoadFromDataSource
        oMat01.AutoResizeColumns
        ProgBar01.Stop
    
    End If
    
    Call oForm01.Freeze(False)
  
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_SD057_MTX01_Error:
'    ProgBar01.Stop
    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_SD057_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Function PS_SD057_AddData() As Boolean
'******************************************************************************
'Function ID : PS_SD057_AddData()
'해당모듈 : PS_SD057
'기능 : 데이터 INSERT
'인수 : 없음
'반환값 : 성공여부
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_SD057_AddData_Error

    Dim loopCount As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim StdYM As String '기준년월
    Dim StdCnt As String '기준회차
    Dim OrdNum As String '품목코드(작번)
    Dim FrgnName As String '품목명(작번)
    Dim SPEC As String '규격(작번)
    Dim InOut As String '자체/외주
    Dim MatAmt As Double '견적자재비
    Dim TotAmt As Double '누계자재비
    Dim PreEA As Double '직전미청구금액
    Dim ExpAmt As Double '미청구금액
    Dim Comment As String '비고
    Dim Amount As Double '금액
    Dim CntcCode As String '등록자 및 수정자
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("저장 중...", 100, False)
    
    CntcCode = MDC_PS_Common.User_MSTCOD '사용자사번
    
    Call oMat01.FlushToDataSource
    For loopCount = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PS_SD057L.GetValue("U_ColReg01", loopCount)) = "Y" Then
            
            StdYM = Trim(oDS_PS_SD057L.GetValue("U_ColReg02", loopCount)) '기준년월
            StdCnt = Trim(oDS_PS_SD057L.GetValue("U_ColReg03", loopCount)) '기준회차
            OrdNum = Trim(oDS_PS_SD057L.GetValue("U_ColReg04", loopCount)) '품목코드(작번)
            FrgnName = Trim(oDS_PS_SD057L.GetValue("U_ColReg05", loopCount)) '품목명(작번)
            SPEC = Trim(oDS_PS_SD057L.GetValue("U_ColReg06", loopCount)) '규격(작번)
            InOut = Trim(oDS_PS_SD057L.GetValue("U_ColReg07", loopCount)) '자체/외주
            MatAmt = Trim(oDS_PS_SD057L.GetValue("U_ColSum01", loopCount)) '견적자재비
            TotAmt = Trim(oDS_PS_SD057L.GetValue("U_ColSum02", loopCount)) '누계자재비
            PreEA = Trim(oDS_PS_SD057L.GetValue("U_ColSum04", loopCount)) '직전미청구금액
            ExpAmt = Trim(oDS_PS_SD057L.GetValue("U_ColSum03", loopCount)) '미청구금액
            Comment = Trim(oDS_PS_SD057L.GetValue("U_ColTxt01", loopCount)) '비고

            sQry = "                EXEC [PS_SD057_03] "
            sQry = sQry & "'" & StdYM & "'," '기준년월
            sQry = sQry & "'" & StdCnt & "',"  '기준회차
            sQry = sQry & "'" & OrdNum & "',"  '품목코드(작번)
            sQry = sQry & "'" & FrgnName & "',"  '품목명(작번)
            sQry = sQry & "'" & SPEC & "',"  '규격(작번)
            sQry = sQry & "'" & InOut & "',"  '자체/외주
            sQry = sQry & "'" & MatAmt & "',"  '견적자재비
            sQry = sQry & "'" & TotAmt & "',"  '누계자재비
            sQry = sQry & "'" & PreEA & "'," '직전미청구금액
            sQry = sQry & "'" & ExpAmt & "',"  '미청구금액
            sQry = sQry & "'" & Comment & "'," '비고
            sQry = sQry & "'" & CntcCode & "'"  '등록자 및 수정자

            Call RecordSet01.DoQuery(sQry)
    
        End If
        
    Next

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set RecordSet01 = Nothing
    
    Call MDC_Com.MDC_GF_Message("등록 완료!", "S")
    
    PS_SD057_AddData = True

    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_SD057_AddData_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set RecordSet01 = Nothing

    Call MDC_Com.MDC_GF_Message("PS_SD057_AddData_Error:" & Err.Number & " - " & Err.Description, "E")
    
    PS_SD057_AddData = False
    
End Function

Public Sub PS_SD057_DeleteData()
'******************************************************************************
'Function ID : PS_SD057_DeleteData()
'해당모듈 : PS_SD057
'기능 : 기본정보 삭제
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_SD057_DeleteData_Error

    Dim loopCount As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim StdYM As String '기준년월
    Dim StdCnt As String '기준회차
    Dim OrdNum As String '품목코드
    Dim CpCode As String '공정코드
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If oMat01.VisualRowCount = 0 Then
    
        ErrNum = 1
        GoTo PS_SD057_DeleteData_Error
    
    End If
    
    For loopCount = 0 To oMat01.VisualRowCount - 1
    
        StdYM = Trim(oDS_PS_SD057L.GetValue("U_ColReg02", loopCount)) '기준년월
        StdCnt = Trim(oDS_PS_SD057L.GetValue("U_ColReg03", loopCount)) '기준회차
        OrdNum = Trim(oDS_PS_SD057L.GetValue("U_ColReg04", loopCount)) '품목코드
        
        sQry = "                EXEC [PS_SD057_04] "
        sQry = sQry & "'" & StdYM & "'," '기준년월
        sQry = sQry & "'" & StdCnt & "'," '기준회차
        sQry = sQry & "'" & OrdNum & "'" '품목코드(작번)
        
        Call oRecordSet01.DoQuery(sQry)
    
    Next
    
    Call MDC_Com.MDC_GF_Message("삭제 완료!", "S")
    
    Set oRecordSet01 = Nothing
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_SD057_DeleteData_Error:

    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("삭제대상이 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_SD057_DeleteData_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Private Function PS_SD057_HeaderSpaceLineDel() As Boolean
'******************************************************************************
'Function ID : PS_SD057_HeaderSpaceLineDel()
'해당모듈 : PS_SD057
'기능 : 필수입력사항 체크
'인수 : 없음
'반환값 : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_SD057_HeaderSpaceLineDel_Error

    Dim ErrNum As Integer
    ErrNum = 0

    Select Case True
        Case Trim(oForm01.Items("BPLID").Specific.VALUE) = "%" '사업장
            ErrNum = 1
            GoTo PS_SD057_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZDPTCD").Specific.VALUE) = "%" '관련자팀
            ErrNum = 2
            GoTo PS_SD057_HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("ZSECCD").Specific.VALUE) = "%" '관련자담당
'            ErrNum = 3
'            GoTo PS_SD057_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZEMPIE").Specific.VALUE) = "" '관련자사번
            ErrNum = 4
            GoTo PS_SD057_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZGODAY").Specific.VALUE) = "" '기준일자
            ErrNum = 5
            GoTo PS_SD057_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZGOTIME").Specific.VALUE) = "" '출입시간(Fr)
            ErrNum = 6
            GoTo PS_SD057_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZOUTIME").Specific.VALUE) = "" '출입시간(To)
            ErrNum = 7
            GoTo PS_SD057_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZUSRNN").Specific.VALUE) = "" '방문자성명
            ErrNum = 8
            GoTo PS_SD057_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZBUSNM").Specific.VALUE) = "" '업체명
            ErrNum = 9
            GoTo PS_SD057_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZRESID").Specific.VALUE) = "" '주소
            ErrNum = 10
            GoTo PS_SD057_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZCARJO").Specific.VALUE) = "%" '차량종류
            ErrNum = 11
            GoTo PS_SD057_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZMOKJU").Specific.VALUE) = "%" '출입목적
            ErrNum = 12
            GoTo PS_SD057_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZPOINT").Specific.VALUE) = "%" '장소
            ErrNum = 13
            GoTo PS_SD057_HeaderSpaceLineDel_Error
    End Select

    PS_SD057_HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_SD057_HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "사업장은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("BPLID").Click(ct_Regular)
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "관련자팀은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZDPTCD").Click(ct_Regular)
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "관련자담당은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZSECCD").Click(ct_Regular)
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "관련자사번은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZEMPIE").Click(ct_Regular)
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "기준일자는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZGODAY").Click(ct_Regular)
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "출입시간(시작)은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZGOTIME").Click(ct_Regular)
    ElseIf ErrNum = 7 Then
        MDC_Com.MDC_GF_Message "출입시간(종료)은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZOUTIME").Click(ct_Regular)
    ElseIf ErrNum = 8 Then
        MDC_Com.MDC_GF_Message "방문자성명은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZUSRNN").Click(ct_Regular)
    ElseIf ErrNum = 9 Then
        MDC_Com.MDC_GF_Message "업체명은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZBUSNM").Click(ct_Regular)
    ElseIf ErrNum = 10 Then
        MDC_Com.MDC_GF_Message "주소는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZRESID").Click(ct_Regular)
    ElseIf ErrNum = 11 Then
        MDC_Com.MDC_GF_Message "차량종류는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZCARJO").Click(ct_Regular)
    ElseIf ErrNum = 12 Then
        MDC_Com.MDC_GF_Message "출입목적은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZMOKJU").Click(ct_Regular)
    ElseIf ErrNum = 13 Then
        MDC_Com.MDC_GF_Message "장소는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZPOINT").Click(ct_Regular)
    End If
    PS_SD057_HeaderSpaceLineDel = False
End Function

Private Function PS_SD057_CheckBeforeSearch(ByVal pItemUID As String) As Boolean
'******************************************************************************
'Function ID : PS_SD057_CheckBeforeSearch()
'해당모듈 : PS_SD057
'기능 : 필수입력사항 체크
'인수 : pItemUID(아이템 UID)
'반환값 : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_SD057_CheckBeforeSearch_Error

    Dim ErrNum As Integer
    ErrNum = 0

    If pItemUID = "BtnSearch1" Then

        Select Case True
'            Case Trim(oForm01.Items("OrdNum").Specific.VALUE) = "" '작번
'                ErrNum = 1
'                GoTo PS_SD057_CheckBeforeSearch_Error '작번은 제외
            Case Trim(oForm01.Items("StdYM").Specific.VALUE) = "" '기준년월
                ErrNum = 2
                GoTo PS_SD057_CheckBeforeSearch_Error
            Case Trim(oForm01.Items("StdCnt").Specific.VALUE) = "%" '기준회차
                ErrNum = 3
                GoTo PS_SD057_CheckBeforeSearch_Error
        End Select
    
    ElseIf pItemUID = "BtnSearch2" Then
        
        Select Case True
'            Case Trim(oForm01.Items("OrdNumS").Specific.VALUE) = "" '작번
'                ErrNum = 4
'                GoTo PS_SD057_CheckBeforeSearch_Error '작번은 제외
            Case Trim(oForm01.Items("StdYMS").Specific.VALUE) = "" '기준년월
                ErrNum = 5
                GoTo PS_SD057_CheckBeforeSearch_Error
            Case Trim(oForm01.Items("StdCntS").Specific.VALUE) = "%" '기준회차
                ErrNum = 6
                GoTo PS_SD057_CheckBeforeSearch_Error
        End Select
    
    End If

    PS_SD057_CheckBeforeSearch = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_SD057_CheckBeforeSearch_Error:
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("입력정보의 작번은 필수사항입니다. 확인하세요.", "E")
        Call oForm01.Items("OrdNum").Click(ct_Regular)
    ElseIf ErrNum = 2 Then
        Call MDC_Com.MDC_GF_Message("입력정보의 기준년월은 필수사항입니다. 확인하세요.", "E")
        Call oForm01.Items("StdYM").Click(ct_Regular)
    ElseIf ErrNum = 3 Then
        Call MDC_Com.MDC_GF_Message("입력정보의 기준회차는 필수사항입니다. 확인하세요.", "E")
        Call oForm01.Items("StdCnt").Click(ct_Regular)
    ElseIf ErrNum = 4 Then
        Call MDC_Com.MDC_GF_Message("조회정보의 작번은 필수사항입니다. 확인하세요.", "E")
        Call oForm01.Items("OrdNumS").Click(ct_Regular)
    ElseIf ErrNum = 5 Then
        Call MDC_Com.MDC_GF_Message("조회정보의 기준년월은 필수사항입니다. 확인하세요.", "E")
        Call oForm01.Items("StdYMS").Click(ct_Regular)
    ElseIf ErrNum = 6 Then
        Call MDC_Com.MDC_GF_Message("조회정보의 기준회차는 필수사항입니다. 확인하세요.", "E")
        Call oForm01.Items("StdCntS").Click(ct_Regular)
    End If
    PS_SD057_CheckBeforeSearch = False
End Function


'/ 메트릭스 필수 사항 check
Private Function PS_SD057_MatrixSpaceLineDel() As Boolean
On Error GoTo PS_SD057_MatrixSpaceLineDel_Error

    Dim i As Long
    Dim ErrNum As Integer
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim sQry As String

    Set oRecordSet01 = Nothing
    PS_SD057_MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_SD057_MatrixSpaceLineDel_Error:
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
        MDC_Com.MDC_GF_Message "PS_SD057_MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    PS_SD057_MatrixSpaceLineDel = False
End Function

Private Sub PS_SD057_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PS_SD057_FlushToItemValue_Error

    Dim Price As Double '단가
    Dim RWorkTime As Double '예상추가공수
    Dim Amount As Double '금액
    Dim loopCount As Integer
    Dim totalAmt As Double '금액 합계

    Select Case oUID

        Case "Mat01"
        
            If oCol = "RWorkTime" Then
            
                Call oMat01.FlushToDataSource
                
                RWorkTime = oDS_PS_SD057L.GetValue("U_ColQty01", oRow - 1)
                Price = oDS_PS_SD057L.GetValue("U_ColSum01", oRow - 1)
                Amount = RWorkTime * Price
                Call oDS_PS_SD057L.setValue("U_ColSum02", oRow - 1, Amount)
            
                For loopCount = 0 To oMat01.VisualRowCount - 1
                
                    totalAmt = totalAmt + oDS_PS_SD057L.GetValue("U_ColSum02", loopCount)
                
                Next
            
                oForm01.Items("Total").Specific.VALUE = totalAmt
            
                Call oMat01.LoadFromDataSource
                
            End If
            
            Call oMat01.AutoResizeColumns

        Case "OrdNum"
        
            oForm01.Items("FrgnName").Specific.VALUE = MDC_GetData.Get_ReData("FrgnName", "ItemCode", "[OITM]", "'" & oForm01.Items(oUID).Specific.VALUE & "'")
            
        Case "OrdNumS"
        
            oForm01.Items("FrgnNameS").Specific.VALUE = MDC_GetData.Get_ReData("FrgnName", "ItemCode", "[OITM]", "'" & oForm01.Items(oUID).Specific.VALUE & "'")
        
    End Select

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
PS_SD057_FlushToItemValue_Error:

    Call MDC_Com.MDC_GF_Message("PS_SD057_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")

End Sub

'/폼의 아이템 사용지정
Sub PS_SD057_FormItemEnabled()
On Error GoTo PS_SD057_FormItemEnabled_Error
    If (oForm01.Mode = fm_ADD_MODE) Then
'        oMat01.Columns("ItemCode").Cells(1).Click ct_Regular
'        oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
'        oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_SD057_FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "PS_SD057_FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
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
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
            Case "1282": '추가
                '/추가버튼 클릭시 메트릭스 insertrow
                
'                oMat01.Clear
'                oMat01.FlushToDataSource
'                oMat01.LoadFromDataSource
                
                oForm01.Mode = fm_ADD_MODE
                BubbleEvent = False
                Call PS_SD057_LoadCaption
                
                'oForm01.Items("GCode").Click ct_Regular
                
                
                Exit Sub
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
                '//Call PS_SD057_FormItemEnabled '//UDO방식
            Case "1282": '추가
'                oMat01.Clear
'                oDS_PS_SD057H.Clear

'                Call PS_SD057_LoadCaption
'                Call PS_SD057_FormItemEnabled
                '//Call PS_SD057_FormItemEnabled '//UDO방식
                '//Call PS_SD057_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PS_SD057_FormItemEnabled
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
    
        If pval.ItemUID = "PS_SD057" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
            
        If pval.ItemUID = "BtnAdd" Then '추가(수정) 버튼클릭
        
            If oForm01.Mode = fm_ADD_MODE Then
            
'                If PS_SD057_HeaderSpaceLineDel = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
                
                If PS_SD057_AddData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If

                oForm01.Mode = fm_ADD_MODE
                
                Call PS_SD057_LoadCaption
                
                oLast_Mode = oForm01.Mode
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            
'                If PS_SD057_HeaderSpaceLineDel = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
            
                '쿼리에서 로직 처리
                '행이 존재하면 UPDATE, 존재하지 않으면 INSERT
                If PS_SD057_AddData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
'                oForm01.Mode = fm_ADD_MODE
                
                Call PS_SD057_MTX01("BtnSearch2")
                Call PS_SD057_LoadCaption

            End If
            
        ElseIf pval.ItemUID = "BtnSearch1" Then '/조회

            If PS_SD057_CheckBeforeSearch(pval.ItemUID) = False Then
                BubbleEvent = False
                Exit Sub
            End If

            oForm01.Mode = fm_ADD_MODE '/fm_VIEW_MODE
            
            Call PS_SD057_LoadCaption
            Call PS_SD057_MTX01(pval.ItemUID)
            
        ElseIf pval.ItemUID = "BtnSearch2" Then '조회2
        
            If PS_SD057_CheckBeforeSearch(pval.ItemUID) = False Then
                BubbleEvent = False
                Exit Sub
            End If

            oForm01.Mode = fm_UPDATE_MODE '/fm_VIEW_MODE
            
            Call PS_SD057_LoadCaption
            Call PS_SD057_MTX01(pval.ItemUID)
        
        ElseIf pval.ItemUID = "BtnDelete" Then '삭제
        
            If Sbo_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", "1", "예", "아니오") = "1" Then
        
                Call PS_SD057_DeleteData
                
                oForm01.Mode = fm_ADD_MODE '/fm_VIEW_MODE
                
                Call PS_SD057_LoadCaption
'                Call PS_GA020_MTX01
                
            Else
            
            End If
        
        ElseIf pval.ItemUID = "BtnChk" Then '전체선택(해제)
        
            Call PS_SD057_CheckAll
            
        End If
  
    ElseIf pval.BeforeAction = False Then
        
        If pval.ItemUID = "BtnAdd" Or pval.ItemUID = "BtnDelete" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
    
        If pval.ItemUID = "PS_SD057" Then
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
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "OrdNum", "") '//사용자값활성
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "OrdNumS", "") '//사용자값활성
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "ItemCode") '//사용자값활성
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_KEY_DOWN_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_KEY_DOWN_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CLICK(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CLICK_Error

    If pval.BeforeAction = True Then
    
        
        
    ElseIf pval.BeforeAction = False Then
    
        If pval.ItemUID = "Mat01" Then
        
            If pval.Row > 0 Then
            
                Call oMat01.SelectRow(pval.Row, True, False)
                
            End If
        End If
    
    End If
    
    Exit Sub
    
Raise_EVENT_CLICK_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_COMBO_SELECT(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_COMBO_SELECT_Error

    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
    
        Call PS_SD057_FlushToItemValue(pval.ItemUID)
    
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
            
                Call PS_SD057_FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
        
            Else
        
                Call PS_SD057_FlushToItemValue(pval.ItemUID)
            
            End If
                
        End If
        
'            oMat01.LoadFromDataSource
'            oMat01.AutoResizeColumns
'            oForm01.Update
        
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
        Call PS_SD057_FormItemEnabled
        '//Call PS_SD057_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_SD057_FormResize
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
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_SD057H", "U_CardCode,U_CardName")
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
        
            
        ElseIf pval.BeforeAction = False Then
            
'            If oMat01.RowCount = 0 Then
'                Call PS_SD057_Add_MatrixRow(0)
'            Else
'                If Trim(oDS_PS_SD057L.GetValue("U_ColReg01", oMat01.RowCount - 1)) <> "" Then
'                    Call PS_SD057_Add_MatrixRow(oMat01.RowCount)
'                End If
'            End If

        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_SD057_CreateItems() As Boolean
On Error GoTo PS_SD057_CreateItems_Error

    Call oForm01.Freeze(True)
    
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    Set oDS_PS_SD057H = oForm01.DataSources.DBDataSources("@PS_SD057H")
    Set oDS_PS_SD057L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '/////입력정보/////
    '작번
    Call oForm01.DataSources.UserDataSources.Add("OrdNum", dt_SHORT_TEXT, 20)
    Call oForm01.Items("OrdNum").Specific.DataBind.SetBound(True, "", "OrdNum")
    
    '품명
    Call oForm01.DataSources.UserDataSources.Add("FrgnName", dt_SHORT_TEXT, 100)
    Call oForm01.Items("FrgnName").Specific.DataBind.SetBound(True, "", "FrgnName")
    
    '기준년월
    Call oForm01.DataSources.UserDataSources.Add("StdYM", dt_SHORT_TEXT, 6)
    Call oForm01.Items("StdYM").Specific.DataBind.SetBound(True, "", "StdYM")
    
    '기준회차
    Call oForm01.DataSources.UserDataSources.Add("StdCnt", dt_SHORT_TEXT, 5)
    Call oForm01.Items("StdCnt").Specific.DataBind.SetBound(True, "", "StdCnt")
    '/////입력정보/////
    
    '/////조회정보/////
    '작번
    Call oForm01.DataSources.UserDataSources.Add("OrdNumS", dt_SHORT_TEXT, 20)
    Call oForm01.Items("OrdNumS").Specific.DataBind.SetBound(True, "", "OrdNumS")
    
    '품명
    Call oForm01.DataSources.UserDataSources.Add("FrgnNameS", dt_SHORT_TEXT, 100)
    Call oForm01.Items("FrgnNameS").Specific.DataBind.SetBound(True, "", "FrgnNameS")
    
    '기준년월
    Call oForm01.DataSources.UserDataSources.Add("StdYMS", dt_SHORT_TEXT, 6)
    Call oForm01.Items("StdYMS").Specific.DataBind.SetBound(True, "", "StdYMS")
    
    '기준회차
    Call oForm01.DataSources.UserDataSources.Add("StdCntS", dt_SHORT_TEXT, 5)
    Call oForm01.Items("StdCntS").Specific.DataBind.SetBound(True, "", "StdCntS")
    '/////조회정보/////

    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    
    Exit Function
PS_SD057_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD057_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'/콤보박스 set
Sub PS_SD057_ComboBox_Setting()
On Error GoTo PS_SD057_ComboBox_Setting_Error

    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)

    '//////////입력정보//////////S
    '기준회차
    Call oForm01.Items("StdCnt").Specific.ValidValues.Add("%", "선택")
    sQry = "           SELECT      U_Minor AS [Code],"
    sQry = sQry & "                U_CdName AS [Name]"
    sQry = sQry & " FROM       [@PS_SY001L]"
    sQry = sQry & " WHERE      Code = 'S008'"
    sQry = sQry & "                AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY  U_Seq"
    Call MDC_SetMod.Set_ComboList(oForm01.Items("StdCnt").Specific, sQry, "", False, False)
    Call oForm01.Items("StdCnt").Specific.Select("%", psk_ByValue)
    '//////////입력정보//////////E
    
    '//////////조회정보//////////S
    '기준회차
    Call oForm01.Items("StdCntS").Specific.ValidValues.Add("%", "전체")
    sQry = "           SELECT      U_Minor AS [Code],"
    sQry = sQry & "                U_CdName AS [Name]"
    sQry = sQry & " FROM       [@PS_SY001L]"
    sQry = sQry & " WHERE      Code = 'S008'"
    sQry = sQry & "                AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY  U_Seq"
    Call MDC_SetMod.Set_ComboList(oForm01.Items("StdCntS").Specific, sQry, "", False, False)
    Call oForm01.Items("StdCntS").Specific.Select("%", psk_ByValue)
    '//////////조회정보//////////E
    
    '//////////매트릭스//////////S
    '기준회차
    sQry = "           SELECT      U_Minor AS [Code],"
    sQry = sQry & "                U_CdName AS [Name]"
    sQry = sQry & " FROM       [@PS_SY001L]"
    sQry = sQry & " WHERE      Code = 'S008'"
    sQry = sQry & "                AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY  U_Seq"
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("StdCnt"), sQry)

'    '품의구분
'    sQry = "           SELECT       Code AS [Code],"
'    sQry = sQry & "                 Name AS [Name]"
'    sQry = sQry & " FROM        [@PSH_ORDTYP]"
'    sQry = sQry & " WHERE       Code IN ('10','20','30','40')" '4개 품의대해서만 조회
'    sQry = sQry & " ORDER BY   Code"
'    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("POType"), sQry)
    '//////////매트릭스//////////E

    Call oForm01.Freeze(False)
    Set oRecordSet01 = Nothing
    
    Exit Sub
PS_SD057_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Set oRecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "PS_SD057_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD057_CF_ChooseFromList()
On Error GoTo PS_SD057_CF_ChooseFromList_Error

    Exit Sub
PS_SD057_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD057_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD057_EnableMenus()
On Error GoTo PS_SD057_EnableMenus_Error
    
    Exit Sub
PS_SD057_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD057_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD057_Initial_Setting()
On Error GoTo PS_SD057_Initial_Setting_Error
    
    '컬럼 Hidden 처리
    oMat01.Columns("StdYM").Visible = False '기준년월
    oMat01.Columns("StdCnt").Visible = False '기준회차
    
    oForm01.Items("StdYM").Specific.VALUE = Format(Now, "YYYYMM")
    oForm01.Items("StdYMS").Specific.VALUE = Format(Now, "YYYYMM")
    
    Call oForm01.Items("OrdNum").Click
    
    Exit Sub
PS_SD057_Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD057_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD057_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_SD057_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_SD057_FormItemEnabled
        '//Call PS_SD057_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_SD057_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_SD057_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD057_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD057_FormResize()
On Error GoTo PS_SD057_FormResize_Error
    
    Call oMat01.AutoResizeColumns
    
    Exit Sub
PS_SD057_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD057_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD057_CheckAll()
On Error GoTo PS_SD057_CheckAll_Error

    Dim CheckType As String
    Dim loopCount As Integer
    
    Call oForm01.Freeze(True)
    CheckType = "Y"
    
    Call oMat01.FlushToDataSource
    
    For loopCount = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PS_SD057L.GetValue("U_ColReg01", loopCount)) = "N" Then
            CheckType = "N"
            Exit For
        End If
    Next
    
    For loopCount = 0 To oMat01.VisualRowCount - 1
        oDS_PS_SD057L.Offset = loopCount
        If CheckType = "N" Then
            Call oDS_PS_SD057L.setValue("U_ColReg01", loopCount, "Y")
        Else
            Call oDS_PS_SD057L.setValue("U_ColReg01", loopCount, "N")
        End If
    Next
    
    Call oMat01.LoadFromDataSource
    
    Call oForm01.Freeze(False)

    Exit Sub
    
PS_SD057_CheckAll_Error:
    Call oForm01.Freeze(False)
    Call Sbo_Application.SetStatusBarMessage("PS_SD057_CheckAll_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)
End Sub





