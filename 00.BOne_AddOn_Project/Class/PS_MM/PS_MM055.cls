VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_MM055"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_MM055.cls
'//  Module : 구매관리 > 애드온 > 구매관리
'//  Desc : 미입고품 입고예정일 관리
'//  FormType : PS_MM055
'//  Create Date(Start) : 2014.11.14
'//  Create Date(End) :
'//  Creator : Song Myoung gyu
'//  Modified Date :
'//  Modifier :
'//  Company : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_MM055H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_MM055L As SAPbouiCOM.DBDataSource '등록라인

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
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_MM055.srf")
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
    
    oFormUniqueID01 = "PS_MM055_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
     
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    Call oForm01.Freeze(True)
    
    Call PS_MM055_CreateItems
    Call PS_MM055_ComboBox_Setting
    Call PS_MM055_CF_ChooseFromList
    Call PS_MM055_EnableMenus
    Call PS_MM055_SetDocument(oFromDocEntry01)
    Call PS_MM055_FormResize

'    Call PS_MM055_Add_MatrixRow(0, True)
    Call PS_MM055_LoadCaption
    Call PS_MM055_FormItemEnabled
    Call PS_MM055_Initial_Setting
    
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

Private Sub PS_MM055_LoadCaption()
'******************************************************************************
'Function ID : PS_MM055_LoadCaption()
'해당모듈 : PS_MM055
'기능 : Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_MM055_LoadCaption_Error

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
PS_MM055_LoadCaption_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/메트릭스 Row추가
Sub PS_MM055_Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_MM055_Add_MatrixRow_Error

    If RowIserted = False Then '//행추가여부
        oDS_PS_MM055L.InsertRecord (oRow)
    End If
    
    oMat01.AddRow
    oDS_PS_MM055L.Offset = oRow
    Call oDS_PS_MM055L.setValue("U_LineNum", oRow, oRow + 1)
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_MM055_Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "PS_MM055_Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub PS_MM055_MTX01(ByVal pItemUID As String)
'******************************************************************************
'Function ID : PS_MM055_MTX01()
'해당모듈 : PS_MM055
'기능 : 데이터 조회
'인수 : pItemUID(아이템UID)
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_MM055_MTX01_Error

    Dim i As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLId As String '작번
    Dim FrDt As String '기간(시작)
    Dim ToDt As String '기간(종료)
    Dim CardType As String '거래처구분
    Dim ItemType As String '품목구분
    Dim Amount As Currency '품의금액
    Dim OrdCls As String '품의구분
    Dim StdDt As String '기준일자
    Dim PhsCode As String '구매업체
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    If pItemUID = "BtnSearch1" Then
    
        BPLId = Trim(oForm01.Items("BPLID").Specific.VALUE) '사업장
        FrDt = Trim(oForm01.Items("FrDt").Specific.VALUE) '기간(시작)
        ToDt = Trim(oForm01.Items("ToDt").Specific.VALUE) '기간(종료)
        CardType = Trim(oForm01.Items("CardType").Specific.VALUE) '거래처구분
        ItemType = Trim(oForm01.Items("ItemType").Specific.VALUE) '품목구분
        Amount = Trim(oForm01.Items("Amount").Specific.VALUE) '품의금액
        OrdCls = Trim(oForm01.Items("OrdCls").Specific.VALUE) '품의구분
        StdDt = Trim(oForm01.Items("StdDt").Specific.VALUE) '기준일자
        PhsCode = Trim(oForm01.Items("PhsCode").Specific.VALUE) '구매업체
        
        sQry = "         EXEC [PS_MM055_01] '"
        sQry = sQry & BPLId & "','"
        sQry = sQry & FrDt & "','"
        sQry = sQry & ToDt & "','"
        sQry = sQry & CardType & "','"
        sQry = sQry & ItemType & "','"
        sQry = sQry & Amount & "','"
        sQry = sQry & OrdCls & "','"
        sQry = sQry & StdDt & "','"
        sQry = sQry & PhsCode & "'"
                
        Call oRecordSet01.DoQuery(sQry)
        
        oMat01.Clear
        oDS_PS_MM055L.Clear
        oMat01.FlushToDataSource
        oMat01.LoadFromDataSource
    
        If (oRecordSet01.RecordCount = 0) Then
        
            ErrNum = 1
            
            oForm01.Mode = fm_ADD_MODE
            
    '        Call PS_MM055_Add_MatrixRow(0, True)
            Call PS_MM055_LoadCaption
            
            GoTo PS_MM055_MTX01_Error
            
            Exit Sub
        End If
        
        For i = 0 To oRecordSet01.RecordCount - 1
            If i + 1 > oDS_PS_MM055L.Size Then
                oDS_PS_MM055L.InsertRecord (i)
            End If
            
            oMat01.AddRow
            oDS_PS_MM055L.Offset = i
            
            Call oDS_PS_MM055L.setValue("U_LineNum", i, i + 1)
            Call oDS_PS_MM055L.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("Check").VALUE)) '선택
            Call oDS_PS_MM055L.setValue("U_ColDt01", i, Format(Trim(oRecordSet01.Fields("InDueDt").VALUE), "YYYYMMDD")) '입고예정일
            Call oDS_PS_MM055L.setValue("U_ColTxt01", i, Trim(oRecordSet01.Fields("Comment").VALUE)) '비고
            Call oDS_PS_MM055L.setValue("U_ColReg03", i, Format(Trim(oRecordSet01.Fields("OrdDueDt").VALUE), "YYYY.MM.DD")) '납기일
            Call oDS_PS_MM055L.setValue("U_ColReg04", i, Format(Trim(oRecordSet01.Fields("OrdDocDt").VALUE), "YYYY.MM.DD")) '품의일
            Call oDS_PS_MM055L.setValue("U_ColReg05", i, Trim(oRecordSet01.Fields("MatItemCd").VALUE)) '자재품목코드
            Call oDS_PS_MM055L.setValue("U_ColReg06", i, Trim(oRecordSet01.Fields("MatItemNm").VALUE)) '자재품명
            Call oDS_PS_MM055L.setValue("U_ColReg07", i, Trim(oRecordSet01.Fields("MatItemSpc").VALUE)) '자재규격
            Call oDS_PS_MM055L.setValue("U_ColReg11", i, Trim(oRecordSet01.Fields("OrdType").VALUE)) '자체/통합
            Call oDS_PS_MM055L.setValue("U_ColReg08", i, Trim(oRecordSet01.Fields("OrdCls").VALUE)) '품의구분
            Call oDS_PS_MM055L.setValue("U_ColReg09", i, Trim(oRecordSet01.Fields("OrdCardNm").VALUE)) '업체명
            Call oDS_PS_MM055L.setValue("U_ColReg10", i, Trim(oRecordSet01.Fields("OrdUnit").VALUE)) '단위
            Call oDS_PS_MM055L.setValue("U_ColQty01", i, Trim(oRecordSet01.Fields("OrdQty").VALUE)) '수량
            Call oDS_PS_MM055L.setValue("U_ColSum01", i, Trim(oRecordSet01.Fields("OrdAmt").VALUE)) '품의금액
            Call oDS_PS_MM055L.setValue("U_ColSum02", i, Trim(oRecordSet01.Fields("TOrdAmt").VALUE)) '누적품의금액
            Call oDS_PS_MM055L.setValue("U_ColSum03", i, Trim(oRecordSet01.Fields("SaleAmt").VALUE)) '수주금액
            Call oDS_PS_MM055L.setValue("U_ColReg13", i, Trim(oRecordSet01.Fields("ItemCode").VALUE)) '작번
            Call oDS_PS_MM055L.setValue("U_ColReg14", i, Trim(oRecordSet01.Fields("ItemName").VALUE)) '품명
            Call oDS_PS_MM055L.setValue("U_ColReg15", i, Trim(oRecordSet01.Fields("ItemSpec").VALUE)) '규격
            Call oDS_PS_MM055L.setValue("U_ColReg19", i, Trim(oRecordSet01.Fields("BPLID").VALUE)) '사업장코드
            Call oDS_PS_MM055L.setValue("U_ColReg16", i, Format(Trim(oRecordSet01.Fields("StdDt").VALUE), "YYYY.MM.DD")) '기준일자
            Call oDS_PS_MM055L.setValue("U_ColReg17", i, Trim(oRecordSet01.Fields("OrdEntry").VALUE)) '품의문서번호
            Call oDS_PS_MM055L.setValue("U_ColReg18", i, Trim(oRecordSet01.Fields("OrdLineID").VALUE)) '품의라인번호
            
            oRecordSet01.MoveNext
            ProgBar01.VALUE = ProgBar01.VALUE + 1
            ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
            
        Next
        
        oMat01.LoadFromDataSource
        oMat01.AutoResizeColumns
        ProgBar01.Stop
    
    ElseIf pItemUID = "BtnSearch2" Then
    
        BPLId = Trim(oForm01.Items("S_BPLID").Specific.VALUE) '사업장
        OrdCls = Trim(oForm01.Items("S_OrdCls").Specific.VALUE) '품의구분
        StdDt = Trim(oForm01.Items("S_StdDt").Specific.VALUE) '기준일자

        sQry = "         EXEC [PS_MM055_02] '"
        sQry = sQry & BPLId & "','"
        sQry = sQry & OrdCls & "','"
        sQry = sQry & StdDt & "'"

        Call oRecordSet01.DoQuery(sQry)

        oMat01.Clear
        oDS_PS_MM055L.Clear
        oMat01.FlushToDataSource
        oMat01.LoadFromDataSource

        If (oRecordSet01.RecordCount = 0) Then

            ErrNum = 1

            oForm01.Mode = fm_ADD_MODE

    '        Call PS_MM055_Add_MatrixRow(0, True)
            Call PS_MM055_LoadCaption

            GoTo PS_MM055_MTX01_Error

            Exit Sub
        End If

        For i = 0 To oRecordSet01.RecordCount - 1
            If i + 1 > oDS_PS_MM055L.Size Then
                oDS_PS_MM055L.InsertRecord (i)
            End If

            oMat01.AddRow
            oDS_PS_MM055L.Offset = i

            Call oDS_PS_MM055L.setValue("U_LineNum", i, i + 1)
            Call oDS_PS_MM055L.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("Check").VALUE)) '선택
            Call oDS_PS_MM055L.setValue("U_ColDt01", i, Format(Trim(oRecordSet01.Fields("InDueDt").VALUE), "YYYYMMDD")) '입고예정일
            Call oDS_PS_MM055L.setValue("U_ColTxt01", i, Trim(oRecordSet01.Fields("Comment").VALUE)) '비고
            Call oDS_PS_MM055L.setValue("U_ColReg03", i, Format(Trim(oRecordSet01.Fields("OrdDueDt").VALUE), "YYYY.MM.DD")) '납기일
            Call oDS_PS_MM055L.setValue("U_ColReg04", i, Format(Trim(oRecordSet01.Fields("OrdDocDt").VALUE), "YYYY.MM.DD")) '품의일
            Call oDS_PS_MM055L.setValue("U_ColReg05", i, Trim(oRecordSet01.Fields("MatItemCd").VALUE)) '자재품목코드
            Call oDS_PS_MM055L.setValue("U_ColReg06", i, Trim(oRecordSet01.Fields("MatItemNm").VALUE)) '자재품명
            Call oDS_PS_MM055L.setValue("U_ColReg07", i, Trim(oRecordSet01.Fields("MatItemSpc").VALUE)) '자재규격
            Call oDS_PS_MM055L.setValue("U_ColReg11", i, Trim(oRecordSet01.Fields("OrdType").VALUE)) '자체/통합
            Call oDS_PS_MM055L.setValue("U_ColReg08", i, Trim(oRecordSet01.Fields("OrdCls").VALUE)) '품의구분
            Call oDS_PS_MM055L.setValue("U_ColReg09", i, Trim(oRecordSet01.Fields("OrdCardNm").VALUE)) '업체명
            Call oDS_PS_MM055L.setValue("U_ColReg10", i, Trim(oRecordSet01.Fields("OrdUnit").VALUE)) '단위
            Call oDS_PS_MM055L.setValue("U_ColQty01", i, Trim(oRecordSet01.Fields("OrdQty").VALUE)) '수량
            Call oDS_PS_MM055L.setValue("U_ColSum01", i, Trim(oRecordSet01.Fields("OrdAmt").VALUE)) '품의금액
            Call oDS_PS_MM055L.setValue("U_ColSum02", i, Trim(oRecordSet01.Fields("TOrdAmt").VALUE)) '누적품의금액
            Call oDS_PS_MM055L.setValue("U_ColSum03", i, Trim(oRecordSet01.Fields("SaleAmt").VALUE)) '수주금액
            Call oDS_PS_MM055L.setValue("U_ColReg13", i, Trim(oRecordSet01.Fields("ItemCode").VALUE)) '작번
            Call oDS_PS_MM055L.setValue("U_ColReg14", i, Trim(oRecordSet01.Fields("ItemName").VALUE)) '품명
            Call oDS_PS_MM055L.setValue("U_ColReg15", i, Trim(oRecordSet01.Fields("ItemSpec").VALUE)) '규격
            Call oDS_PS_MM055L.setValue("U_ColReg19", i, Trim(oRecordSet01.Fields("BPLID").VALUE)) '사업장코드
            Call oDS_PS_MM055L.setValue("U_ColReg16", i, Format(Trim(oRecordSet01.Fields("StdDt").VALUE), "YYYY.MM.DD")) '기준일자
            Call oDS_PS_MM055L.setValue("U_ColReg17", i, Trim(oRecordSet01.Fields("OrdEntry").VALUE)) '품의문서번호
            Call oDS_PS_MM055L.setValue("U_ColReg18", i, Trim(oRecordSet01.Fields("OrdLineID").VALUE)) '품의라인번호

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
PS_MM055_MTX01_Error:
'    ProgBar01.Stop
    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_MM055_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Function PS_MM055_AddData() As Boolean
'******************************************************************************
'Function ID : PS_MM055_AddData()
'해당모듈 : PS_MM055
'기능 : 데이터 INSERT
'인수 : 없음
'반환값 : 성공여부
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_MM055_AddData_Error

    Dim loopCount As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim InDueDt As String '입고예정일
    Dim Comment As String '비고
    Dim OrdDueDt As String '납기일
    Dim OrdDocDt As String '품의일
    Dim MatItemCd As String '자재품목코드
    Dim MatItemNm As String '자재품명
    Dim MatItemSpc As String '자재규격
    Dim OrdType As String '자체/통합
    Dim OrdCls As String '품의구분
    Dim OrdCardNm As String '업체명
    Dim OrdUnit As String '단위
    Dim OrdQty As Double '수량
    Dim OrdAmt As Currency '품의금액
    Dim TOrdAmt As Currency '누적품의금액
    Dim SaleAmt As Currency '수주금액
    Dim ItemCode As String '작번
    Dim ItemName As String '품명
    Dim BPLId As String '사업장코드
    Dim ItemSpec As String '규격
    Dim StdDt As String '기준일자
    Dim OrdEntry As String '품의문서번호
    Dim OrdLineID As String '품의라인번호
    
    Dim InDueDt2 As String 'xxxx.xx.xx 형태로 변경하기 위한 변수
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("저장 중...", 100, False)
    
    
    Call oMat01.FlushToDataSource
    For loopCount = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PS_MM055L.GetValue("U_ColReg01", loopCount)) = "Y" Then
            
            InDueDt2 = Trim(oDS_PS_MM055L.GetValue("U_ColDt01", loopCount))
            
            InDueDt = Left(InDueDt2, 4) & "." & Right(Left(InDueDt2, 6), 2) & "." & Right(InDueDt2, 2) '입고예정일
            Comment = Trim(oDS_PS_MM055L.GetValue("U_ColTxt01", loopCount)) '비고
            OrdDueDt = Trim(oDS_PS_MM055L.GetValue("U_ColReg03", loopCount)) '납기일
            OrdDocDt = Trim(oDS_PS_MM055L.GetValue("U_ColReg04", loopCount)) '품의일
            MatItemCd = Trim(oDS_PS_MM055L.GetValue("U_ColReg05", loopCount)) '자재품목코드
            MatItemNm = Replace(Trim(oDS_PS_MM055L.GetValue("U_ColReg06", loopCount)), "'", "`") '자재품명(자재품명 중에 '를 `로 변경)
            MatItemSpc = Replace(Trim(oDS_PS_MM055L.GetValue("U_ColReg07", loopCount)), "'", "`") '자재규격(자재규격 중에 '를 `로 변경)
            OrdType = Trim(oDS_PS_MM055L.GetValue("U_ColReg11", loopCount)) '자체/통합
            OrdCls = Trim(oDS_PS_MM055L.GetValue("U_ColReg08", loopCount)) '품의구분
            OrdCardNm = Trim(oDS_PS_MM055L.GetValue("U_ColReg09", loopCount)) '업체명
            OrdUnit = Trim(oDS_PS_MM055L.GetValue("U_ColReg10", loopCount)) '단위
            OrdQty = Trim(oDS_PS_MM055L.GetValue("U_ColQty01", loopCount)) '수량
            OrdAmt = Trim(oDS_PS_MM055L.GetValue("U_ColSum01", loopCount)) '품의금액
            TOrdAmt = Trim(oDS_PS_MM055L.GetValue("U_ColSum02", loopCount)) '누적품의금액
            SaleAmt = Trim(oDS_PS_MM055L.GetValue("U_ColSum03", loopCount)) '수주금액
            ItemCode = Trim(oDS_PS_MM055L.GetValue("U_ColReg13", loopCount)) '작번
            ItemName = Replace(Trim(oDS_PS_MM055L.GetValue("U_ColReg14", loopCount)), "'", "`") '품명(품명 중에 '를 `로 변경)
            ItemSpec = Replace(Trim(oDS_PS_MM055L.GetValue("U_ColReg15", loopCount)), "'", "`") '규격(규격 중에 '를 `로 변경
            BPLId = Trim(oDS_PS_MM055L.GetValue("U_ColReg19", loopCount)) '사업장코드
            StdDt = Trim(oDS_PS_MM055L.GetValue("U_ColReg16", loopCount)) '기준일자
            OrdEntry = Trim(oDS_PS_MM055L.GetValue("U_ColReg17", loopCount)) '품의문서번호
            OrdLineID = Trim(oDS_PS_MM055L.GetValue("U_ColReg18", loopCount)) '품의라인번호
            
            sQry = "                EXEC [PS_MM055_03] "
            sQry = sQry & "'" & InDueDt & "'," '입고예정일
            sQry = sQry & "'" & Comment & "'," '비고
            sQry = sQry & "'" & OrdDueDt & "',"  '납기일
            sQry = sQry & "'" & OrdDocDt & "',"  '품의일
            sQry = sQry & "'" & MatItemCd & "',"  '자재품목코드
            sQry = sQry & "'" & MatItemNm & "'," '자재품명
            sQry = sQry & "'" & MatItemSpc & "'," '자재규격
            sQry = sQry & "'" & OrdType & "'," '자체/통합
            sQry = sQry & "'" & OrdCls & "'," '품의구분
            sQry = sQry & "'" & OrdCardNm & "',"  '업체명
            sQry = sQry & "'" & OrdUnit & "',"  '단위
            sQry = sQry & "'" & OrdQty & "',"  '수량
            sQry = sQry & "'" & OrdAmt & "',"  '품의금액
            sQry = sQry & "'" & TOrdAmt & "',"  '누적품의금액
            sQry = sQry & "'" & SaleAmt & "',"  '수주금액
            sQry = sQry & "'" & ItemCode & "',"  '작번
            sQry = sQry & "'" & ItemName & "',"  '품명
            sQry = sQry & "'" & ItemSpec & "',"  '규격
            sQry = sQry & "'" & BPLId & "',"  '사업장코드
            sQry = sQry & "'" & StdDt & "',"  '기준일자
            sQry = sQry & "'" & OrdEntry & "',"  '품의문서번호
            sQry = sQry & "'" & OrdLineID & "'"  '품의라인번호

            Call RecordSet01.DoQuery(sQry)
    
        End If
        
    Next

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set RecordSet01 = Nothing
    
    Call MDC_Com.MDC_GF_Message("등록 완료!", "S")
    
    PS_MM055_AddData = True

    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_MM055_AddData_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set RecordSet01 = Nothing

    Call MDC_Com.MDC_GF_Message("PS_MM055_AddData_Error:" & Err.Number & " - " & Err.Description, "E")
    
    PS_MM055_AddData = False
    
End Function

Public Sub PS_MM055_DeleteData()
'******************************************************************************
'Function ID : PS_MM055_DeleteData()
'해당모듈 : PS_MM055
'기능 : 기본정보 삭제
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_MM055_DeleteData_Error

    Dim loopCount As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim StdYM As String '기준년월
    Dim StdCnt As String '기준회차
    Dim ReqNo As String '요청번호
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If oMat01.VisualRowCount = 0 Then
    
        ErrNum = 1
        GoTo PS_MM055_DeleteData_Error
    
    End If
    
    For loopCount = 0 To oMat01.VisualRowCount - 1
    
        StdYM = Trim(oDS_PS_MM055L.GetValue("U_ColReg02", loopCount)) '기준년월
        StdCnt = Trim(oDS_PS_MM055L.GetValue("U_ColReg03", loopCount)) '기준회차
        ReqNo = Trim(oDS_PS_MM055L.GetValue("U_ColReg04", loopCount)) '요청번호
        
        sQry = "                EXEC [PS_MM055_04] "
        sQry = sQry & "'" & StdYM & "'," '기준년월
        sQry = sQry & "'" & StdCnt & "',"  '기준회차
        sQry = sQry & "'" & ReqNo & "'"  '요청번호
        
        Call oRecordSet01.DoQuery(sQry)
    
    Next
    
    Call MDC_Com.MDC_GF_Message("삭제 완료!", "S")
    
    Set oRecordSet01 = Nothing
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_MM055_DeleteData_Error:

    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("삭제대상이 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_MM055_DeleteData_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Private Function PS_MM055_HeaderSpaceLineDel() As Boolean
'******************************************************************************
'Function ID : PS_MM055_HeaderSpaceLineDel()
'해당모듈 : PS_MM055
'기능 : 필수입력사항 체크
'인수 : 없음
'반환값 : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_MM055_HeaderSpaceLineDel_Error

    Dim ErrNum As Integer
    ErrNum = 0

    Select Case True
        Case Trim(oForm01.Items("BPLID").Specific.VALUE) = "%" '사업장
            ErrNum = 1
            GoTo PS_MM055_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZDPTCD").Specific.VALUE) = "%" '관련자팀
            ErrNum = 2
            GoTo PS_MM055_HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("ZSECCD").Specific.VALUE) = "%" '관련자담당
'            ErrNum = 3
'            GoTo PS_MM055_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZEMPIE").Specific.VALUE) = "" '관련자사번
            ErrNum = 4
            GoTo PS_MM055_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZGODAY").Specific.VALUE) = "" '기준일자
            ErrNum = 5
            GoTo PS_MM055_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZGOTIME").Specific.VALUE) = "" '출입시간(Fr)
            ErrNum = 6
            GoTo PS_MM055_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZOUTIME").Specific.VALUE) = "" '출입시간(To)
            ErrNum = 7
            GoTo PS_MM055_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZUSRNN").Specific.VALUE) = "" '방문자성명
            ErrNum = 8
            GoTo PS_MM055_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZBUSNM").Specific.VALUE) = "" '업체명
            ErrNum = 9
            GoTo PS_MM055_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZRESID").Specific.VALUE) = "" '주소
            ErrNum = 10
            GoTo PS_MM055_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZCARJO").Specific.VALUE) = "%" '차량종류
            ErrNum = 11
            GoTo PS_MM055_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZMOKJU").Specific.VALUE) = "%" '출입목적
            ErrNum = 12
            GoTo PS_MM055_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZPOINT").Specific.VALUE) = "%" '장소
            ErrNum = 13
            GoTo PS_MM055_HeaderSpaceLineDel_Error
    End Select

    PS_MM055_HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_MM055_HeaderSpaceLineDel_Error:
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
    PS_MM055_HeaderSpaceLineDel = False
End Function

Private Function PS_MM055_CheckBeforeSearch(ByVal pItemUID As String) As Boolean
'******************************************************************************
'Function ID : PS_MM055_CheckBeforeSearch()
'해당모듈 : PS_MM055
'기능 : 필수입력사항 체크
'인수 : pItemUID(아이템 UID)
'반환값 : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_MM055_CheckBeforeSearch_Error

    Dim ErrNum As Integer
    ErrNum = 0

    If pItemUID = "BtnSearch1" Then

        Select Case True
'            Case Trim(oForm01.Items("OrdNum").Specific.VALUE) = "" '작번
'                ErrNum = 1
'                GoTo PS_MM055_CheckBeforeSearch_Error
'            Case Trim(oForm01.Items("StdYM").Specific.VALUE) = "" '기준년월
'                ErrNum = 2
'                GoTo PS_MM055_CheckBeforeSearch_Error
'            Case Trim(oForm01.Items("StdCnt").Specific.VALUE) = "%" '기준회차
'                ErrNum = 3
'                GoTo PS_MM055_CheckBeforeSearch_Error
        End Select
    
    ElseIf pItemUID = "BtnSearch2" Then
        
        Select Case True
'            Case Trim(oForm01.Items("OrdNumS").Specific.VALUE) = "" '작번
'                ErrNum = 4
'                GoTo PS_MM055_CheckBeforeSearch_Error
'            Case Trim(oForm01.Items("StdYMS").Specific.VALUE) = "" '기준년월
'                ErrNum = 5
'                GoTo PS_MM055_CheckBeforeSearch_Error
'            Case Trim(oForm01.Items("StdCntS").Specific.VALUE) = "%" '기준회차
'                ErrNum = 6
'                GoTo PS_MM055_CheckBeforeSearch_Error
        End Select
    
    End If

    PS_MM055_CheckBeforeSearch = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_MM055_CheckBeforeSearch_Error:
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
    PS_MM055_CheckBeforeSearch = False
End Function


'/ 메트릭스 필수 사항 check
Private Function PS_MM055_MatrixSpaceLineDel() As Boolean
On Error GoTo PS_MM055_MatrixSpaceLineDel_Error

    Dim i As Long
    Dim ErrNum As Integer
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim sQry As String

    Set oRecordSet01 = Nothing
    PS_MM055_MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_MM055_MatrixSpaceLineDel_Error:
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
        MDC_Com.MDC_GF_Message "PS_MM055_MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    PS_MM055_MatrixSpaceLineDel = False
End Function

Private Sub PS_MM055_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PS_MM055_FlushToItemValue_Error

    Dim loopCount As Integer
    Dim totalAmt As Double

    Select Case oUID

        Case "Mat01"

            If oCol = "Amount" Then

                Call oMat01.FlushToDataSource

                For loopCount = 0 To oMat01.VisualRowCount - 1

                    totalAmt = totalAmt + oDS_PS_MM055L.GetValue("U_ColSum03", loopCount)

                Next

                oForm01.Items("Total").Specific.VALUE = totalAmt

                Call oMat01.LoadFromDataSource

            End If

            Call oMat01.AutoResizeColumns

        Case "PhsCode"

            oForm01.Items("PhsName").Specific.VALUE = MDC_GetData.Get_ReData("CardName", "CardCode", "[OCRD]", "'" & oForm01.Items(oUID).Specific.VALUE & "'")

'        Case "OrdNumS"
'
'            oForm01.Items("FrgnNameS").Specific.VALUE = MDC_GetData.Get_ReData("FrgnName", "ItemCode", "[OITM]", "'" & oForm01.Items(oUID).Specific.VALUE & "'")
        
    End Select

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
PS_MM055_FlushToItemValue_Error:

    Call MDC_Com.MDC_GF_Message("PS_MM055_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")

End Sub

'/폼의 아이템 사용지정
Sub PS_MM055_FormItemEnabled()
On Error GoTo PS_MM055_FormItemEnabled_Error
    If (oForm01.Mode = fm_ADD_MODE) Then
'        oMat01.Columns("ItemCode").Cells(1).Click ct_Regular
'        oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
'        oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_MM055_FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "PS_MM055_FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
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
                Call PS_MM055_LoadCaption
                
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
                '//Call PS_MM055_FormItemEnabled '//UDO방식
            Case "1282": '추가
'                oMat01.Clear
'                oDS_PS_MM055H.Clear

'                Call PS_MM055_LoadCaption
'                Call PS_MM055_FormItemEnabled
                '//Call PS_MM055_FormItemEnabled '//UDO방식
                '//Call PS_MM055_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PS_MM055_FormItemEnabled
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
    
        If pval.ItemUID = "PS_MM055" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
            
        If pval.ItemUID = "BtnAdd" Then '추가(수정) 버튼클릭
        
            If oForm01.Mode = fm_ADD_MODE Then
            
'                If PS_MM055_HeaderSpaceLineDel = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
                
                If PS_MM055_AddData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If

                oForm01.Mode = fm_ADD_MODE
                
                Call PS_MM055_LoadCaption
                
                oLast_Mode = oForm01.Mode
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            
'                If PS_MM055_HeaderSpaceLineDel = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
            
                '쿼리에서 로직 처리
                '행이 존재하면 UPDATE, 존재하지 않으면 INSERT
                If PS_MM055_AddData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
'                oForm01.Mode = fm_ADD_MODE
                
                Call PS_MM055_MTX01("BtnSearch2")
                Call PS_MM055_LoadCaption

            End If
            
        ElseIf pval.ItemUID = "BtnSearch1" Then '/조회

            If PS_MM055_CheckBeforeSearch(pval.ItemUID) = False Then
                BubbleEvent = False
                Exit Sub
            End If

            oForm01.Mode = fm_ADD_MODE '/fm_VIEW_MODE
            
            Call PS_MM055_LoadCaption
            Call PS_MM055_MTX01(pval.ItemUID)
            
        ElseIf pval.ItemUID = "BtnSearch2" Then '조회2
        
            If PS_MM055_CheckBeforeSearch(pval.ItemUID) = False Then
                BubbleEvent = False
                Exit Sub
            End If

            oForm01.Mode = fm_UPDATE_MODE '/fm_VIEW_MODE
            
            Call PS_MM055_LoadCaption
            Call PS_MM055_MTX01(pval.ItemUID)
        
        ElseIf pval.ItemUID = "BtnDelete" Then '삭제
        
            If Sbo_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", "1", "예", "아니오") = "1" Then
        
                Call PS_MM055_DeleteData
                
                oForm01.Mode = fm_ADD_MODE '/fm_VIEW_MODE
                
                Call PS_MM055_LoadCaption
'                Call PS_GA020_MTX01
                
            Else
            
            End If
        
        ElseIf pval.ItemUID = "BtnChk" Then '전체선택(해제)
        
            Call PS_MM055_CheckAll
            
        ElseIf pval.ItemUID = "BtnPrint" Then '출력
        
            Call PS_MM055_Print_Report01
            
        End If
  
    ElseIf pval.BeforeAction = False Then
        
        If pval.ItemUID = "BtnAdd" Or pval.ItemUID = "BtnDelete" Then
            If oForm01.Mode = fm_ADD_MODE Then
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "PS_MM055" Then
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
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "PhsCode", "") '구매업체코드
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "OrdNumS", "") '//사용자값활성
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
    
        Call PS_MM055_FlushToItemValue(pval.ItemUID)
    
    End If

    Exit Sub
Raise_EVENT_COMBO_SELECT_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_COMBO_SELECT_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_DOUBLE_CLICK(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_DOUBLE_CLICK_Error
    If pval.BeforeAction = True Then
        
        If pval.ItemUID = "Mat01" Then
        
            If pval.Row = 0 Then
            
                oMat01.Columns(pval.ColUID).TitleObject.Sortable = True
                oMat01.FlushToDataSource
            
            End If
        
        End If
    
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
            
                Call PS_MM055_FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
        
            Else
        
                Call PS_MM055_FlushToItemValue(pval.ItemUID)
            
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
        Call PS_MM055_FormItemEnabled
        '//Call PS_MM055_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_MM055_FormResize
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
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_MM055H", "U_CardCode,U_CardName")
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
'                Call PS_MM055_Add_MatrixRow(0)
'            Else
'                If Trim(oDS_PS_MM055L.GetValue("U_ColReg01", oMat01.RowCount - 1)) <> "" Then
'                    Call PS_MM055_Add_MatrixRow(oMat01.RowCount)
'                End If
'            End If

        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_MM055_CreateItems() As Boolean
On Error GoTo PS_MM055_CreateItems_Error

    Call oForm01.Freeze(True)
    
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    Set oDS_PS_MM055H = oForm01.DataSources.DBDataSources("@PS_MM055H")
    Set oDS_PS_MM055L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '/////입력정보조회/////
    '사업장
    Call oForm01.DataSources.UserDataSources.Add("BPLID", dt_SHORT_TEXT, 5)
    Call oForm01.Items("BPLID").Specific.DataBind.SetBound(True, "", "BPLID")
    
    '기간(품의)시작
    Call oForm01.DataSources.UserDataSources.Add("FrDt", dt_DATE)
    Call oForm01.Items("FrDt").Specific.DataBind.SetBound(True, "", "FrDt")
    
    '기간(품의)종료
    Call oForm01.DataSources.UserDataSources.Add("ToDt", dt_DATE)
    Call oForm01.Items("ToDt").Specific.DataBind.SetBound(True, "", "ToDt")
    
    '거래처구분
    Call oForm01.DataSources.UserDataSources.Add("CardType", dt_SHORT_TEXT, 5)
    Call oForm01.Items("CardType").Specific.DataBind.SetBound(True, "", "CardType")
    
    '품목구분
    Call oForm01.DataSources.UserDataSources.Add("ItemType", dt_SHORT_TEXT, 5)
    Call oForm01.Items("ItemType").Specific.DataBind.SetBound(True, "", "ItemType")
    
    '품의구분
    Call oForm01.DataSources.UserDataSources.Add("OrdCls", dt_SHORT_TEXT, 5)
    Call oForm01.Items("OrdCls").Specific.DataBind.SetBound(True, "", "OrdCls")
    
    '품의금액
    Call oForm01.DataSources.UserDataSources.Add("Amount", dt_SUM)
    Call oForm01.Items("Amount").Specific.DataBind.SetBound(True, "", "Amount")
    
    '기준일자
    Call oForm01.DataSources.UserDataSources.Add("StdDt", dt_DATE)
    Call oForm01.Items("StdDt").Specific.DataBind.SetBound(True, "", "StdDt")
    
    '구매업체코드
    Call oForm01.DataSources.UserDataSources.Add("PhsCode", dt_SHORT_TEXT, 20)
    Call oForm01.Items("PhsCode").Specific.DataBind.SetBound(True, "", "PhsCode")
    
    '구매업체명
    Call oForm01.DataSources.UserDataSources.Add("PhsName", dt_SHORT_TEXT, 50)
    Call oForm01.Items("PhsName").Specific.DataBind.SetBound(True, "", "PhsName")
    '/////입력정보조회/////
    
    '/////출력정보조회/////
    '사업장
    Call oForm01.DataSources.UserDataSources.Add("S_BPLID", dt_SHORT_TEXT, 5)
    Call oForm01.Items("S_BPLID").Specific.DataBind.SetBound(True, "", "S_BPLID")
    
    '품의구분
    Call oForm01.DataSources.UserDataSources.Add("S_OrdCls", dt_SHORT_TEXT, 5)
    Call oForm01.Items("S_OrdCls").Specific.DataBind.SetBound(True, "", "S_OrdCls")
    
    '기준일자
    Call oForm01.DataSources.UserDataSources.Add("S_StdDt", dt_DATE)
    Call oForm01.Items("S_StdDt").Specific.DataBind.SetBound(True, "", "S_StdDt")
    '/////출력정보조회/////

    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    
    Exit Function
PS_MM055_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_MM055_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'/콤보박스 set
Sub PS_MM055_ComboBox_Setting()
On Error GoTo PS_MM055_ComboBox_Setting_Error

    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)

    '//////////입력정보조회//////////S
    '사업장
'    Call oForm01.Items("BPLID").Specific.ValidValues.Add("%", "선택")
    sQry = "           SELECT      BPLId AS [Code],"
    sQry = sQry & "                BPLName AS [Name]"
    sQry = sQry & " FROM       [OBPL]"
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLID").Specific, sQry, "", False, False)
    Call oForm01.Items("BPLID").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
    
    '거래처구분
    sQry = "           SELECT     U_Minor AS [Code], "
    sQry = sQry & "               U_CdName AS [Name]"
    sQry = sQry & " FROM      [@PS_SY001L]"
    sQry = sQry & " WHERE     Code = 'C100'"
    Call oForm01.Items("CardType").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("CardType").Specific, sQry, "", False, False)
    Call oForm01.Items("CardType").Specific.Select("%", psk_ByValue)
    
    '품목구분
    sQry = "           SELECT     U_Minor AS [Code], "
    sQry = sQry & "               U_CdName AS [Name]"
    sQry = sQry & " FROM      [@PS_SY001L]"
    sQry = sQry & " WHERE     Code = 'S002'"
    Call oForm01.Items("ItemType").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ItemType").Specific, sQry, "", False, False)
    Call oForm01.Items("ItemType").Specific.Select("%", psk_ByValue)
    
    '품의구분
    sQry = "           SELECT       Code AS [Code],"
    sQry = sQry & "                 Name AS [Name]"
    sQry = sQry & " FROM        [@PSH_ORDTYP]"
    sQry = sQry & " WHERE       Code IN ('10','20','30','40')" '4개 품의대해서만 조회
    sQry = sQry & " ORDER BY   Code"
    Call oForm01.Items("OrdCls").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("OrdCls").Specific, sQry, "", False, False)
    Call oForm01.Items("OrdCls").Specific.Select("%", psk_ByValue)
    '//////////입력정보조회//////////E
    
    '//////////출력정보조회//////////S
    '사업장
'    Call oForm01.Items("BPLID").Specific.ValidValues.Add("%", "선택")
    sQry = "           SELECT      BPLId AS [Code],"
    sQry = sQry & "                BPLName AS [Name]"
    sQry = sQry & " FROM       [OBPL]"
    Call MDC_SetMod.Set_ComboList(oForm01.Items("S_BPLID").Specific, sQry, "", False, False)
    Call oForm01.Items("S_BPLID").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
    
    '품의구분
    sQry = "           SELECT       Code AS [Code],"
    sQry = sQry & "                 Name AS [Name]"
    sQry = sQry & " FROM        [@PSH_ORDTYP]"
    sQry = sQry & " WHERE       Code IN ('10','20','30','40')" '4개 품의대해서만 조회
    sQry = sQry & " ORDER BY   Code"
    Call oForm01.Items("S_OrdCls").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("S_OrdCls").Specific, sQry, "", False, False)
    Call oForm01.Items("S_OrdCls").Specific.Select("%", psk_ByValue)
    '//////////출력정보조회//////////E

    '//////////매트릭스//////////S
    '품의구분
    sQry = "           SELECT       Code AS [Code],"
    sQry = sQry & "                 Name AS [Name]"
    sQry = sQry & " FROM        [@PSH_ORDTYP]"
    sQry = sQry & " WHERE       Code IN ('10','20','30','40')" '4개 품의대해서만 조회
    sQry = sQry & " ORDER BY   Code"
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("OrdCls"), sQry)
    
    '자체/통합
    sQry = "           SELECT      Code,"
    sQry = sQry & "                Name "
    sQry = sQry & " FROM       [@PSH_RETYPE]"
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("OrdType"), sQry)
    '//////////매트릭스//////////E

    Call oForm01.Freeze(False)
    Set oRecordSet01 = Nothing
    
    Exit Sub
PS_MM055_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Set oRecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "PS_MM055_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_MM055_CF_ChooseFromList()
On Error GoTo PS_MM055_CF_ChooseFromList_Error

    Exit Sub
PS_MM055_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM055_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_MM055_EnableMenus()
On Error GoTo PS_MM055_EnableMenus_Error
    
    Exit Sub
PS_MM055_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM055_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_MM055_Initial_Setting()
On Error GoTo PS_MM055_Initial_Setting_Error
    
    '컬럼 Hidden 처리
'    oMat01.Columns("StdDt").Visible = False '기준일자
'    oMat01.Columns("OrdEntry").Visible = False '품의문서번호
'    oMat01.Columns("OrdLineID").Visible = False '품의라인번호
    
    oForm01.Items("StdDt").Enabled = False '입력정보 기준일자 수정 불가처리
    
    oForm01.Items("StdDt").Specific.VALUE = Format(Now, "YYYYMMDD")
    oForm01.Items("S_StdDt").Specific.VALUE = Format(Now, "YYYYMMDD")
    
    Call oForm01.Items("Amount").Click
    
    Exit Sub
PS_MM055_Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM055_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_MM055_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_MM055_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_MM055_FormItemEnabled
        '//Call PS_MM055_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_MM055_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_MM055_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM055_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_MM055_FormResize()
On Error GoTo PS_MM055_FormResize_Error
    
    Call oMat01.AutoResizeColumns
    
    Exit Sub
PS_MM055_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM055_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_MM055_CheckAll()
On Error GoTo PS_MM055_CheckAll_Error

    Dim CheckType As String
    Dim loopCount As Integer
    
    Call oForm01.Freeze(True)
    CheckType = "Y"
    
    Call oMat01.FlushToDataSource
    
    For loopCount = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PS_MM055L.GetValue("U_ColReg01", loopCount)) = "N" Then
            CheckType = "N"
            Exit For
        End If
    Next
    
    For loopCount = 0 To oMat01.VisualRowCount - 1
        oDS_PS_MM055L.Offset = loopCount
        If CheckType = "N" Then
            Call oDS_PS_MM055L.setValue("U_ColReg01", loopCount, "Y")
        Else
            Call oDS_PS_MM055L.setValue("U_ColReg01", loopCount, "N")
        End If
    Next
    
    Call oMat01.LoadFromDataSource
    
    Call oForm01.Freeze(False)

    Exit Sub
    
PS_MM055_CheckAll_Error:
    Call oForm01.Freeze(False)
    Call Sbo_Application.SetStatusBarMessage("PS_MM055_CheckAll_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)
End Sub

Private Sub PS_MM055_Print_Report01()
On Error GoTo Print_Query_Error

    Dim DocNum As String
    Dim WinTitle As String
    Dim ReportName As String
    Dim sQry As String

    Dim i As Integer
    Dim ErrNum As Integer
    Dim Sub_sQry As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
            
    Call ConnectODBC
    
    Dim BPLId As String '사업장
    Dim OrdCls As String '품의구분
    Dim StdDt As String '기준일자
    
    BPLId = Trim(oForm01.Items("S_BPLID").Specific.VALUE) '사업장
    OrdCls = Trim(oForm01.Items("S_OrdCls").Specific.VALUE) '품의구분
    StdDt = Trim(oForm01.Items("S_StdDt").Specific.VALUE) '기준일자

    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
 
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
    
    WinTitle = "[PS_MM055] 미입고품 입고예정일 조회"
    
    ReportName = "PS_MM055_01.rpt"
    
    sQry = "         EXEC [PS_MM055_02] '"
    sQry = sQry & BPLId & "','"
    sQry = sQry & OrdCls & "','"
    sQry = sQry & StdDt & "'"
    
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
        MDC_Com.MDC_GF_Message "PS_MM055_Print_Report01_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub


