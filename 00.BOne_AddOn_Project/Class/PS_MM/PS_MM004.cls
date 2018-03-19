VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_MM004"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_MM004.cls
'//  Module         : 자재관리 > 운송확인증 출력
'//  Desc           : 기본
'//  FormType       : PS_MM004
'//  Create Date    : 2014.03.25 N.G.Y
'//  Copyright  (c) Morning Data
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_MM004H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_MM004L As SAPbouiCOM.DBDataSource '등록라인

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oFormName As String '참조Object
Private oFormDocNo As String '참조문서번호

'//사용자구조체
Private Type ItemInformations
    ItemCode As String
    LotNo As String
    Quantity As Long
    OPORNo As Long
    POR1No As Long
    Check As Boolean
    OPDNNo As Long
    PDN1No As Long
End Type

Private oLast_Mode&

Private ItemInformation() As ItemInformations
Private ItemInformationCount As Long

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromName01 As String, Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim DocDate As String
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_MM004.srf")
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
    
    oFormUniqueID01 = "PS_MM004_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
     
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oFormName = oFromName01
    oFormDocNo = oFromDocEntry01
        
    oForm01.Freeze True
    Call PS_MM004_CreateItems
    Call PS_MM004_FormReset '화면초기화
    Call PS_MM004_ComboBox_Setting
    Call PS_MM004_CF_ChooseFromList
    Call PS_MM004_EnableMenus
    Call PS_MM004_FormResize
    
'    Call Add_MatrixRow(0, True)
    Call LoadCaption
    Call FormItemEnabled
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1285"), False        '// 복원
    oForm01.EnableMenu ("1284"), False         '// 취소
    oForm01.EnableMenu ("1293"), False         '// 행삭제
    oForm01.EnableMenu ("1281"), False
    oForm01.EnableMenu ("1282"), True
        
    
    oMat01.Columns("Check").Visible = False '선택 체크박스 Visible = False
    
    
    
    If Trim(oFormDocNo) <> "" Then
        Dim sQry As String
        
        Dim RecordSet01 As SAPbobsCOM.Recordset
        Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
        
        sQry = "Select Cnt = Count(*) From [@PS_MM004H] Where U_ObjType = '" + oFormName + "' And U_DocNo = '" + oFormDocNo + "'"
        Call RecordSet01.DoQuery(sQry)
        
        If RecordSet01.Fields(0).VALUE > 0 Then
            sQry = "Select DocEntry, U_DocDate From [@PS_MM004H] Where U_ObjType = '" + oFormName + "' And U_DocNo = '" + oFormDocNo + "'"
            Call RecordSet01.DoQuery(sQry)
            
            oForm01.Items("SDocDateF").Specific.VALUE = Format(RecordSet01.Fields(1).VALUE, "YYYYMMDD")
            oForm01.Items("SDocDateT").Specific.VALUE = Format(RecordSet01.Fields(1).VALUE, "YYYYMMDD")
            
            oForm01.Items("SObjType").Specific.VALUE = oFormName
            oForm01.Items("SDocNo").Specific.VALUE = oFormDocNo
            
            Call oForm01.Items("BtnSearch").Click
            
            Call PS_MM004_RET(Trim(RecordSet01.Fields(0).VALUE))
        
        Else
        
        
            oForm01.Items("ObjType").Specific.VALUE = oFormName
            oForm01.Items("DocNo").Specific.VALUE = oFormDocNo
            
            oForm01.Items("SectionF").Specific.VALUE = "창원"
            
            If oFormName = "PS_SD040" Then
                sQry = "Select Distinct b.U_ItmBsort From [@PS_SD040L] a Inner join OITM b On a.U_ItemCode = b.ItemCode and a.DocEntry = '" + oFormDocNo + "'"
                Call RecordSet01.DoQuery(sQry)
                
                If RecordSet01.RecordCount > 0 Then
                    oForm01.Items("ItmBsort").Specific.VALUE = Trim(RecordSet01.Fields(0).VALUE)
                    If Trim(RecordSet01.Fields(0).VALUE) = "102" Then
                        oForm01.Items("SectionT").Specific.VALUE = "안강"
                    End If
                End If
                
                
                sQry = "Select U_TranCard, U_TranCode, U_Tonnage, U_Destin, U_DocDate From [@PS_SD040H] Where DocEntry = '" + oFormDocNo + "'"
                Call RecordSet01.DoQuery(sQry)
                
                If RecordSet01.RecordCount > 0 Then
                    oForm01.Items("TranCard").Specific.VALUE = Trim(RecordSet01.Fields(0).VALUE)
                    oForm01.Items("TranCode").Specific.VALUE = Trim(RecordSet01.Fields(1).VALUE)
                    oForm01.Items("Tonnage").Specific.VALUE = Trim(RecordSet01.Fields(2).VALUE)
                    oForm01.Items("Destin").Specific.VALUE = Trim(RecordSet01.Fields(3).VALUE)
                    
                    oForm01.Items("DocDate").Specific.VALUE = Format(RecordSet01.Fields(4).VALUE, "YYYYMMDD")
                    oForm01.Items("SDocDateF").Specific.VALUE = Format(RecordSet01.Fields(4).VALUE, "YYYYMMDD")
                    
                End If
                oForm01.Items("LocCode").Specific.VALUE = "07"
                
            ElseIf oFormName = "PS_PP095" Then
                sQry = "Select U_Destin, U_Section, U_Tonnage, U_TranCard, U_TranCode, U_TWeight, U_DocDate From [@PS_PP095H] Where DocEntry = '" + oFormDocNo + "'"
                Call RecordSet01.DoQuery(sQry)
                
                If RecordSet01.RecordCount > 0 Then
                    oForm01.Items("Destin").Specific.VALUE = Trim(RecordSet01.Fields(0).VALUE)
                    oForm01.Items("SectionT").Specific.VALUE = Trim(RecordSet01.Fields(1).VALUE)
                    oForm01.Items("Tonnage").Specific.VALUE = Trim(RecordSet01.Fields(2).VALUE)
                    oForm01.Items("TranCard").Specific.VALUE = Trim(RecordSet01.Fields(3).VALUE)
                    oForm01.Items("TranCode").Specific.VALUE = Trim(RecordSet01.Fields(4).VALUE)
                    oForm01.Items("TWeight").Specific.VALUE = Trim(RecordSet01.Fields(5).VALUE)
                    oForm01.Items("DocDate").Specific.VALUE = Format(RecordSet01.Fields(6).VALUE, "YYYYMMDD")
                    oForm01.Items("SDocDateF").Specific.VALUE = Format(RecordSet01.Fields(6).VALUE, "YYYYMMDD")
                    
                End If
                
                oForm01.Items("ItmBsort").Specific.VALUE = "104"
                
                sQry = "Select DocDate = Convert(Char(8),Max(U_DocDate),112) From [@PS_MM003H] Where U_BPLId = '" & oForm01.Items("BPLId").Specific.VALUE & "'"
                sQry = sQry & " And U_DocDate <= '" & oForm01.Items("DocDate").Specific.VALUE & "'"
                Call RecordSet01.DoQuery(sQry)
                
                DocDate = Trim(RecordSet01.Fields(0).VALUE)
                
                sQry = " Select b.U_LocCode From [@PS_MM003H] a Inner Join [@PS_MM003L] b On a.Code = b.Code "
                sQry = sQry + " Where a.U_BPLId = '" + oForm01.Items("BPLId").Specific.VALUE + "'"
                sQry = sQry + " And a.U_Docdate = '" + DocDate + "'"
                sQry = sQry + " And b.U_LocName Like '%" & oForm01.Items("SectionT").Specific.VALUE & "%'"
                Call RecordSet01.DoQuery(sQry)
                
                oForm01.Items("LocCode").Specific.VALUE = Trim(RecordSet01.Fields(0).VALUE)
                
                '//실중량
                sQry = "Select Weight = Sum(U_Weight) From [@PS_PP095L] Where DocEntry = '" + oFormDocNo + "'"
                Call RecordSet01.DoQuery(sQry)
                If RecordSet01.RecordCount > 0 Then
                    oForm01.Items("Weight").Specific.VALUE = Trim(RecordSet01.Fields(0).VALUE)
                End If
                
            End If
        
        End If
    Else
        '//일자 Set
        oForm01.Items("DocDate").Specific.VALUE = Format(Date, "YYYYMMDD")
        oForm01.Items("SDocDateF").Specific.VALUE = Format(Date, "YYYYMMDD")
    End If
    
    oForm01.Items("SDocDateT").Specific.VALUE = Format(Date, "YYYYMMDD")
    oForm01.Update
    oForm01.Freeze False
    
    oForm01.Visible = True
    Set oXmlDoc01 = Nothing
    
    
'     oMat01.Columns("LineNum").Cells(1).Specific.Value = 1

     Call oForm01.Items("DocDate").Click
     
    Exit Sub
LoadForm_Error:
    oForm01.Update
    oForm01.Freeze False
    Set oXmlDoc01 = Nothing
    Set oForm01 = Nothing
    Sbo_Application.SetStatusBarMessage "Form_Load Error:" & Err.Description, bmt_Short, True
End Sub

'/ 추가 확인 갱신 버튼 이름 변경
Private Sub LoadCaption()
On Error GoTo LoadCaption_Error
    Call oForm01.Freeze(True)

    If oForm01.Mode = fm_ADD_MODE Then
        oForm01.Items("BtnAdd").Specific.Caption = "추가"
        oForm01.Items("BtnDelete").Enabled = False
'    ElseIf oForm01.Mode = fm_OK_MODE Then
'        oForm01.Items("BtnAdd").Specific.Caption = "확인"
    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
        oForm01.Items("BtnAdd").Specific.Caption = "수정"
        oForm01.Items("BtnDelete").Enabled = True
    End If
    
    Call oForm01.Freeze(False)

    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadCaption_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub
'/메트릭스 Row추가
Sub Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_MM004H.InsertRecord (oRow)
    End If
    
    oMat01.AddRow
    oDS_PS_MM004H.Offset = oRow
    oDS_PS_MM004H.setValue "U_LineNum", oRow, oRow + 1
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub PS_MM004_MTX01()
'******************************************************************************
'Function ID : PS_MM004_MTX01()
'해당모듈 : PS_MM004
'기능 : 데이터 조회
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_MM004_MTX01_Error

    Dim i As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim SBPLID As String '사업장
    Dim SDocDateF As String '조회기간1
    Dim SDocDateT As String '조회기간2
    Dim SItmBsort As String '폼목분류
    Dim SLocCode As String '운송지역
    Dim SWay As String '편도/회로
    Dim STranCard As String '운송업체
    Dim STonnage As String '차종(톤수)
    Dim SObjType As String '참조Object
    Dim SDocNo  As String '참조문서번호
    

    SBPLID = Trim(oForm01.Items("SBPLId").Specific.VALUE) '사업장
    SDocDateF = Trim(oForm01.Items("SDocDateF").Specific.VALUE) '조회시작일
    SDocDateT = Trim(oForm01.Items("SDocDateT").Specific.VALUE) '조회종료일
    SItmBsort = Trim(oForm01.Items("SItmBsort").Specific.VALUE) '품목분류
    SLocCode = Trim(oForm01.Items("SLocCode").Specific.VALUE) '운송지역
    SWay = Trim(oForm01.Items("SWay").Specific.VALUE) '편도/회로
    STranCard = Trim(oForm01.Items("STranCard").Specific.VALUE) '운송업체
    STonnage = Trim(oForm01.Items("STonnage").Specific.VALUE) '차종(톤수)
    SObjType = Trim(oForm01.Items("SObjType").Specific.VALUE) '참조Object
    SDocNo = Trim(oForm01.Items("SDocNo").Specific.VALUE) '참조문서번호
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    sQry = "                EXEC [PS_MM004_01] "
    sQry = sQry & "'" & SBPLID & "',"  '사업장
    sQry = sQry & "'" & SDocDateF & "',"  '조회시작일
    sQry = sQry & "'" & SDocDateT & "',"  '조회종료일
    sQry = sQry & "'" & SItmBsort & "',"  '품목분류
    sQry = sQry & "'" & SLocCode & "',"  '운송지역
    sQry = sQry & "'" & SWay & "',"  '편도/회로
    sQry = sQry & "'" & STranCard & "',"  '운송업체
    sQry = sQry & "'" & STonnage & "',"  '차종(톤수)
    sQry = sQry & "'" & SObjType & "',"  '참조Object
    sQry = sQry & "'" & SDocNo & "'"  '참조문서번호
    
    Call oRecordSet01.DoQuery(sQry)
    
    oMat01.Clear
    oDS_PS_MM004L.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource

    If (oRecordSet01.RecordCount = 0) Then
    
        ErrNum = 1
        
        oForm01.Mode = fm_ADD_MODE
        
'        Call PS_MM004_Add_MatrixRow(0, True)
        Call LoadCaption
        
        GoTo PS_MM004_MTX01_Error
        
        Exit Sub
    End If
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_MM004L.Size Then
            oDS_PS_MM004L.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_MM004L.Offset = i
        
        Call oDS_PS_MM004L.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_MM004L.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("DocEntry").VALUE)) '관리번호
        Call oDS_PS_MM004L.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("BPLId").VALUE)) '사업장
        Call oDS_PS_MM004L.setValue("U_ColReg03", i, Trim(oRecordSet01.Fields("DocDate").VALUE)) '일자
        Call oDS_PS_MM004L.setValue("U_ColReg04", i, Trim(oRecordSet01.Fields("ItmBName").VALUE)) '품목분류명
        Call oDS_PS_MM004L.setValue("U_ColReg05", i, Trim(oRecordSet01.Fields("LocName").VALUE)) '운송지역
        Call oDS_PS_MM004L.setValue("U_ColReg06", i, Trim(oRecordSet01.Fields("Way").VALUE)) '편도/회로
        Call oDS_PS_MM004L.setValue("U_ColReg07", i, Trim(oRecordSet01.Fields("SectionF").VALUE)) '운행구간(출발)
        Call oDS_PS_MM004L.setValue("U_ColReg08", i, Trim(oRecordSet01.Fields("SectionT").VALUE)) '운행구간(도착)
        Call oDS_PS_MM004L.setValue("U_ColReg09", i, Trim(oRecordSet01.Fields("Destin").VALUE)) '도착장소
        Call oDS_PS_MM004L.setValue("U_ColReg10", i, Trim(oRecordSet01.Fields("TranCard").VALUE)) '운송업체
        Call oDS_PS_MM004L.setValue("U_ColReg11", i, Trim(oRecordSet01.Fields("TranCode").VALUE)) '차량번호
        Call oDS_PS_MM004L.setValue("U_ColReg12", i, Trim(oRecordSet01.Fields("Tonnage").VALUE)) '차종(톤수)
        Call oDS_PS_MM004L.setValue("U_ColReg13", i, Trim(oRecordSet01.Fields("ItemName").VALUE)) '품명
        Call oDS_PS_MM004L.setValue("U_ColNum01", i, Trim(oRecordSet01.Fields("Weight").VALUE)) '운송량
        Call oDS_PS_MM004L.setValue("U_ColSum01", i, Trim(oRecordSet01.Fields("TotAmt").VALUE)) '금액
        Call oDS_PS_MM004L.setValue("U_ColReg15", i, Trim(oRecordSet01.Fields("PassYN").VALUE)) '비고
        Call oDS_PS_MM004L.setValue("U_ColReg16", i, Trim(oRecordSet01.Fields("Comments").VALUE)) '비고
        
        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
        
    Next
    
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    ProgBar01.Stop
    Call oForm01.Freeze(False)
  
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_MM004_MTX01_Error:
'    ProgBar01.Stop
    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_MM004_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Sub PS_MM004_DeleteData()
'******************************************************************************
'Function ID : PS_MM004_DeleteData()
'해당모듈 : PS_MM004
'기능 : 기본정보 삭제
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_MM004_DeleteData_Error

    Dim i As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As String
    
    If oForm01.Mode = fm_UPDATE_MODE Then
    
        DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
        
        sQry = "SELECT COUNT(*) FROM [@PS_MM004H] WHERE DocEntry = '" & DocEntry & "'"
        Call oRecordSet01.DoQuery(sQry)
        
        If (oRecordSet01.RecordCount = 0) Then
            ErrNum = 1
            oForm01.Mode = fm_ADD_MODE
            GoTo PS_MM004_DeleteData_Error
        Else
            sQry = "EXEC PS_MM004_04 '" + DocEntry + "'"
            Call oRecordSet01.DoQuery(sQry)
        End If
    End If
    
    Call MDC_Com.MDC_GF_Message("삭제 완료!", "S")
    
'    Call PS_MM004_FormReset
    
'    oForm01.Mode = fm_ADD_MODE
    
'    Call oForm01.Items("BtnSearch").Click(ct_Regular)
    
'    oMat01.Clear
'    oMat01.FlushToDataSource
'    oMat01.LoadFromDataSource
'    Call PS_MM004_Add_MatrixRow(0, True)
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_MM004_DeleteData_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("삭제대상이 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_MM004_DeleteData_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Function PS_MM004_UpdateData() As Boolean
'******************************************************************************
'Function ID : PS_MM004_UpdateData()
'해당모듈 : PS_MM004
'기능 : 기본정보를 수정
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_MM004_UpdateData_Error

    Dim i As Integer
    Dim j As Integer
    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As Integer '관리번호
    Dim BPLID As String '사업장
    Dim DocDate As String '일자
    Dim ItmBsort As String '분류코드
    Dim ItmBName As String '분류명
    Dim LocCode As String '운송지역코드
    Dim LocName As String '운송지역명
    Dim Way As String '편도/회로
    Dim SectionF As String '운행구간출발
    Dim SectionT As String '운행구간도착
    Dim Destin As String '도착장소
    Dim TranCard As String '운송업체
    Dim TranCode As String '차량번호
    Dim Tonnage As Double '차종(톤수)
    Dim ItemName As String '품명
    Dim Weight As Double '실중량
    Dim TWeight As Double '총중량
    Dim PassYN As String
    Dim Amt As Double '기본운임
    Dim PassAmt As Double '경유운임
    Dim TotAmt As Double '총운임
    Dim Comments As String '비고

    DocEntry = oForm01.Items("DocEntry").Specific.VALUE '관리번호
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE) '사업장
    DocDate = Trim(oForm01.Items("DocDate").Specific.VALUE) '일자
    ItmBsort = Trim(oForm01.Items("ItmBsort").Specific.VALUE) '품목분류코드
    ItmBName = Trim(oForm01.Items("ItmBName").Specific.VALUE) '품목분류명
    LocCode = Trim(oForm01.Items("LocCode").Specific.VALUE) '운송지역코드
    LocName = Trim(oForm01.Items("LocName").Specific.VALUE) '운송지역명
    Way = Trim(oForm01.Items("Way").Specific.VALUE) '편도/회로
    SectionF = Trim(oForm01.Items("SectionF").Specific.VALUE) '운행구간출발
    SectionT = Trim(oForm01.Items("SectionT").Specific.VALUE) '운행구간도착
    Destin = Trim(oForm01.Items("Destin").Specific.VALUE) '도착장소
    TranCard = Trim(oForm01.Items("TranCard").Specific.VALUE) '운송업체
    TranCode = Trim(oForm01.Items("TranCode").Specific.VALUE) '차량번호
    Tonnage = oForm01.Items("Tonnage").Specific.VALUE '차종(톤수)
    ItemName = oForm01.Items("ItemName").Specific.VALUE '품명
    Weight = oForm01.Items("Weight").Specific.VALUE '실중량
    TWeight = oForm01.Items("TWeight").Specific.VALUE '계근중량
    PassYN = oForm01.Items("PassYN").Specific.VALUE '경유지YN
    Amt = oForm01.Items("Amt").Specific.VALUE '기본운임
    PassAmt = oForm01.Items("PassAmt").Specific.VALUE '경유지운임
    TotAmt = oForm01.Items("TotAmt").Specific.VALUE '총운임
    Comments = Trim(oForm01.Items("Comments").Specific.VALUE) '비고
    
    If Trim(DocEntry) = "" Then
        MDC_Com.MDC_GF_Message "수정할 항목이 없습니다. 수정하실려면 항목을 선택하세요!", "E"
        PS_MM004_UpdateData = False
        Exit Function
    End If
    
    sQry = "                EXEC [PS_MM004_03] "
    sQry = sQry & DocEntry & "," '관리번호
    sQry = sQry & "'" & BPLID & "',"  '사업장
    sQry = sQry & "'" & DocDate & "',"  '출장번호1
    sQry = sQry & "'" & ItmBsort & "',"  '출장번호2
    sQry = sQry & "'" & ItmBName & "',"  '사원번호
    sQry = sQry & "'" & LocCode & "',"  '사원성명
    sQry = sQry & "'" & LocName & "',"  '출장지
    sQry = sQry & "'" & Way & "',"  '출장지상세
    sQry = sQry & "'" & SectionF & "',"  '작번
    sQry = sQry & "'" & SectionT & "',"  '시작일자
    sQry = sQry & "'" & Destin & "',"  '시작시각
    sQry = sQry & "'" & TranCard & "',"  '종료일자
    sQry = sQry & "'" & TranCode & "',"  '종료시각
    sQry = sQry & Tonnage & ","  '목적
    sQry = sQry & ItemName & ","  '품명
    sQry = sQry & Weight & ","  '실중량
    sQry = sQry & TWeight & ","  '총중량
    sQry = sQry & "'" & PassYN & "',"  '경유지YN
    sQry = sQry & Amt & ","  '기본운임
    sQry = sQry & PassAmt & ","  '경유지운임
    sQry = sQry & TotAmt & ","  '총운임
    sQry = sQry & "'" & Comments & "'"  '비고
    
    Call RecordSet01.DoQuery(sQry)
    
    MDC_Com.MDC_GF_Message "수정 완료!", "S"
    
    Set RecordSet01 = Nothing
    PS_MM004_UpdateData = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_MM004_UpdateData_Error:
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PS_MM004_UpdateData_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Function PS_MM004_AddData() As Boolean
'******************************************************************************
'Function ID : PS_MM004_AddData()
'해당모듈 : PS_MM004
'기능 : 데이터 INSERT
'인수 : 없음
'반환값 : 성공여부
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_MM004_AddData_Error

    Dim i As Integer
    Dim sQry As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim RecordSet02 As SAPbobsCOM.Recordset
    Set RecordSet02 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As Long
    Dim ObjType As String '참조문서타입
    Dim DocNo As Long
    Dim BPLID As String '사업장
    Dim DocDate As String '일자
    Dim ItmBsort As String '분류코드
    Dim ItmBName As String '분류명
    Dim LocCode As String '운송지역코드
    Dim LocName As String '운송지역명
    Dim Way As String '편도/회로
    Dim SectionF As String '운행구간출발
    Dim SectionT As String '운행구간도착
    Dim Destin As String '도착장소
    Dim TranCard As String '운송업체
    Dim TranCode As String '차량번호
    Dim Tonnage As Double '차종(톤수)
    Dim ItemName As String '품명
    
    Dim Weight As Double '실중량
    Dim TWeight As Double '총중량
    
    Dim PassYN As String '경유지 YN
    Dim Amt As Double '금액
    Dim PassAmt As Double '경유지금액
    Dim TotAmt As Double '총금액
    Dim Comments As String '비고
    
    ObjType = Trim(oForm01.Items("ObjType").Specific.VALUE) '참조문서Object
    DocNo = Val(oForm01.Items("DocNo").Specific.VALUE) '참조문서번호
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE) '사업장
    DocDate = Trim(oForm01.Items("DocDate").Specific.VALUE) '일자
    ItmBsort = Trim(oForm01.Items("ItmBsort").Specific.VALUE) '품목분류
    ItmBName = Trim(oForm01.Items("ItmBName").Specific.VALUE) '품목분류명
    LocCode = Trim(oForm01.Items("LocCode").Specific.VALUE) '운송지역
    LocName = Trim(oForm01.Items("LocName").Specific.VALUE) '운송지역명
    Way = Trim(oForm01.Items("Way").Specific.VALUE) '편도/회로
    SectionF = Trim(oForm01.Items("SectionF").Specific.VALUE) '운행구간(출발)
    SectionT = Trim(oForm01.Items("SectionT").Specific.VALUE) '운행구간(도착)
    Destin = Trim(oForm01.Items("Destin").Specific.VALUE) '도착장소
    TranCard = Trim(oForm01.Items("TranCard").Specific.VALUE) '운송업체
    TranCode = Trim(oForm01.Items("TranCode").Specific.VALUE) '차량번호
    Tonnage = oForm01.Items("Tonnage").Specific.VALUE '차종(톤수)
    ItemName = oForm01.Items("ItemName").Specific.VALUE '품명
    
    Weight = oForm01.Items("Weight").Specific.VALUE '실중량
    TWeight = oForm01.Items("TWeight").Specific.VALUE '계근중량
    PassYN = oForm01.Items("PassYN").Specific.VALUE '경유지YN
    Amt = oForm01.Items("Amt").Specific.VALUE '금액
    PassAmt = oForm01.Items("PassAmt").Specific.VALUE '경유지금액
    TotAmt = oForm01.Items("TotAmt").Specific.VALUE '총금액
    Comments = Trim(oForm01.Items("Comments").Specific.VALUE) '비고
    
    'DocEntry는 화면상의 DocEntry가 아닌 입력 시점의 최종 DocEntry를 조회한 후 +1하여 INSERT를 해줘야 함
    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM[@PS_MM004H]"
    Call RecordSet01.DoQuery(sQry)
    
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        DocEntry = 1
    Else
        DocEntry = Trim(RecordSet01.Fields(0).VALUE) + 1
    End If
    
    sQry = "                EXEC [PS_MM004_02] "
    sQry = sQry & DocEntry & "," '관리번호
    sQry = sQry & "'" & ObjType & "',"  '참조문서Object
    sQry = sQry & DocNo & ","  '참조문서번호
    sQry = sQry & "'" & BPLID & "',"  '사업장
    sQry = sQry & "'" & DocDate & "',"  '일자
    sQry = sQry & "'" & ItmBsort & "',"  '분류코드
    sQry = sQry & "'" & ItmBName & "',"  '분류명
    sQry = sQry & "'" & LocCode & "',"  '운송지역
    sQry = sQry & "'" & LocName & "',"  '운송지역명
    sQry = sQry & "'" & Way & "',"  '편도/회로
    sQry = sQry & "'" & SectionF & "',"  '운행구간출발
    sQry = sQry & "'" & SectionT & "',"  '운행구간도착
    sQry = sQry & "'" & Destin & "',"  '도착장소
    sQry = sQry & "'" & TranCard & "',"  '운송업체
    sQry = sQry & "'" & TranCode & "',"  '차량번호
    sQry = sQry & Tonnage & ","  '차종(톤수)
    sQry = sQry & "'" & ItemName & "',"  '품명
    sQry = sQry & Weight & ","  '실중량
    sQry = sQry & TWeight & ","  '총중량
    sQry = sQry & "'" & PassYN & "',"  '경유지YN
    sQry = sQry & Amt & ","  '금액
    sQry = sQry & PassAmt & ","  '경유지금액
    sQry = sQry & TotAmt & ","  '금액
    sQry = sQry & "'" & Comments & "'"  '비고
    
    Call RecordSet02.DoQuery(sQry)
    
    MDC_Com.MDC_GF_Message "등록 완료!", "S"
    
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    PS_MM004_AddData = True

    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_MM004_AddData_Error:
    PS_MM004_AddData = False
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    MDC_Com.MDC_GF_Message "PS_MM004_AddData_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Private Function PS_MM004_HeaderSpaceLineDel() As Boolean
'******************************************************************************
'Function ID : PS_MM004_HeaderSpaceLineDel()
'해당모듈 : PS_MM004
'기능 : 필수입력사항 체크
'인수 : 없음
'반환값 : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_MM004_HeaderSpaceLineDel_Error

    Dim ErrNum As Integer
    ErrNum = 0

    Select Case True
        Case Trim(oForm01.Items("BPLId").Specific.VALUE) = "" '사업장
            ErrNum = 1
            GoTo PS_MM004_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ItmBsort").Specific.VALUE) = "" '품목분류
            ErrNum = 2
            GoTo PS_MM004_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("LocCode").Specific.VALUE) = "" '운송지역코드
            ErrNum = 3
            GoTo PS_MM004_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("Way").Specific.VALUE) = "" '편도/회로
            ErrNum = 4
            GoTo PS_MM004_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("TranCard").Specific.VALUE) = "" '운송업체
            ErrNum = 5
            GoTo PS_MM004_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("TranCode").Specific.VALUE) = "" '차량번호
            ErrNum = 6
            GoTo PS_MM004_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("Tonnage").Specific.VALUE) = "" '차종(톤수)
            ErrNum = 7
            GoTo PS_MM004_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ItemName").Specific.VALUE) = "" '품명
            ErrNum = 8
            GoTo PS_MM004_HeaderSpaceLineDel_Error
    End Select

    PS_MM004_HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_MM004_HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "사업장은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("BPLId").Click(ct_Regular)
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "품목분류는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ItmBsort").Click(ct_Regular)
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "운송지역코드는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("LocCode").Click(ct_Regular)
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "편도/회로는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("Way").Click(ct_Regular)
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "운송업체는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("TranCard").Click(ct_Regular)
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "차량번호는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("TranCode").Click(ct_Regular)
    ElseIf ErrNum = 7 Then
        MDC_Com.MDC_GF_Message "차종(톤수)은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("Tonnage").Click(ct_Regular)
    ElseIf ErrNum = 8 Then
        MDC_Com.MDC_GF_Message "품명은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("Tonnage").Click(ct_Regular)
    End If
    PS_MM004_HeaderSpaceLineDel = False
End Function
'/코드의 코드명 보여주기
Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error
    Dim i&
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim TWeight         As Double
    Dim oCombo          As SAPbouiCOM.ComboBox
    
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    Dim BoxWeight As Double
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    
    Select Case oUID
        Case "DocDate"
            oForm01.Items("SDocDateF").Specific.VALUE = oForm01.Items("DocDate").Specific.VALUE
            oForm01.Items("SDocDateT").Specific.VALUE = oForm01.Items("DocDate").Specific.VALUE
            
        Case "ItmBsort" '/품목분류
            sQry = "Select Name From [@PSH_ITMBSORT] Where Code = '" & Trim(oForm01.Items("ItmBsort").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            oForm01.Items("ItmBName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            
            If Trim(oForm01.Items("ObjType").Specific.VALUE) <> "" Then
                oForm01.Items("ItemName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            End If
            
            Set oCombo = oForm01.Items("SItmBsort").Specific
            oCombo.Select Trim(oForm01.Items("ItmBsort").Specific.VALUE), psk_ByValue
            
        Case "LocCode" '/운송지역코드
            sQry = "Select b.U_LocName, b.U_PassAmt From [@PS_MM003H] a Inner Join [@PS_MM003L] b On a.Code = b.Code"
            sQry = sQry + " Where a.U_BPLId =  '" & Trim(oForm01.Items("BPLId").Specific.VALUE) & "'"
            sQry = sQry + "   And b.U_LocCode = '" & Trim(oForm01.Items("LocCode").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            oForm01.Items("LocName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            
            
            If oForm01.Items("PassYN").Specific.VALUE = "Y" Then
                oForm01.Items("PassAmt").Specific.VALUE = oRecordSet01.Fields(1).VALUE
            Else
                oForm01.Items("PassAmt").Specific.VALUE = 0
            End If
            
            
            TWeight = CDbl(oForm01.Items("TWeight").Specific.VALUE)
            '//총계근중량으로 운송비 다시계산
            If TWeight > 0 Then
                sQry = "Select Amt = Isnull(b.U_Amt,0) From [@PS_MM003H] a Inner Join [@PS_MM003L] b On a.Code = b.Code"
                sQry = sQry + " Where a.U_BPLId =  '" & Trim(oForm01.Items("BPLId").Specific.VALUE) & "'"
                sQry = sQry + "   And b.U_LocCode = '" & Trim(oForm01.Items("LocCode").Specific.VALUE) & "'"
                sQry = sQry + "   And " & TWeight & " Between b.U_KgF And b.U_KgT "
                oRecordSet01.DoQuery sQry
                
                If oForm01.Items("Way").Specific.VALUE = "10" Then
                    oForm01.Items("Amt").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                    
                ElseIf oForm01.Items("Way").Specific.VALUE = "20" Then
                    oForm01.Items("Amt").Specific.VALUE = (oRecordSet01.Fields(0).VALUE * 0.8)
                End If
            End If
            
'            oForm01.Items("TotAmt").Specific.VALUE = oForm01.Items("Amt").Specific.VALUE + CDbl(oForm01.Items("PassAmt").Specific.VALUE)
            
'        Case "Tonnage" '/차종(톤수)
'            If CStr(oForm01.Items("Tonnage").Specific.VALUE) <> "" Then
'                sQry = "Select b.U_Amt From [@PS_MM003H] a Inner Join [@PS_MM003L] b On a.Code = b.Code"
'                sQry = sQry + " Where a.U_BPLId =  '" & Trim(oForm01.Items("BPLId").Specific.VALUE) & "'"
'                sQry = sQry + "   And b.U_LocCode = '" & Trim(oForm01.Items("LocCode").Specific.VALUE) & "'"
'                sQry = sQry + "   And b.U_Tonnage = '" & Trim(oForm01.Items("Tonnage").Specific.VALUE) & "'"
'                oRecordSet01.DoQuery sQry
'
'                If oForm01.Items("Way").Specific.VALUE = "10" Then
'                    oForm01.Items("Amt").Specific.VALUE = oRecordSet01.Fields(0).VALUE
'                ElseIf oForm01.Items("Way").Specific.VALUE = "20" Then
'                    oForm01.Items("Amt").Specific.VALUE = (oRecordSet01.Fields(0).VALUE * 0.8)
'                End If
'            Else
'                oForm01.Items("Amt").Specific.VALUE = 0
'            End If
            
            
        Case "Amt"
            oForm01.Items("TotAmt").Specific.VALUE = oForm01.Items("Amt").Specific.VALUE + CDbl(oForm01.Items("PassAmt").Specific.VALUE)
        Case "PassAmt"
            oForm01.Items("TotAmt").Specific.VALUE = oForm01.Items("Amt").Specific.VALUE + CDbl(oForm01.Items("PassAmt").Specific.VALUE)
        Case "TWeight"
            '//중량입력시 화물운송료 계산
            '//MM095 총중량 구간으로으로 금액 Select
            TWeight = CDbl(oForm01.Items("TWeight").Specific.VALUE)
            '//총계근중량으로 운송비 다시계산
            If TWeight > 0 Then
                sQry = "Select Amt = Isnull(b.U_Amt,0) From [@PS_MM003H] a Inner Join [@PS_MM003L] b On a.Code = b.Code"
                sQry = sQry + " Where a.U_BPLId =  '" & Trim(oForm01.Items("BPLId").Specific.VALUE) & "'"
                sQry = sQry + "   And b.U_LocCode = '" & Trim(oForm01.Items("LocCode").Specific.VALUE) & "'"
                sQry = sQry + "   And " & TWeight & " Between b.U_KgF And b.U_KgT "
                oRecordSet01.DoQuery sQry
                
                If Trim(oForm01.Items("Way").Specific.VALUE) = "10" Then
                    oForm01.Items("Amt").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                    
                ElseIf Trim(oForm01.Items("Way").Specific.VALUE) = "20" Then
                    oForm01.Items("Amt").Specific.VALUE = (oRecordSet01.Fields(0).VALUE * 0.8)
                End If
            End If
        End Select
    
   Set oRecordSet01 = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
FlushToItemValue_Error:
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub
'/폼의 아이템 사용지정
Sub FormItemEnabled()
On Error GoTo FormItemEnabled_Error
    If (oForm01.Mode = fm_ADD_MODE) Then
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/아이템 변경 이벤트
Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_ItemEvent_Error


    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            Call Raise_EVENT_ITEM_PRESSED(FormUID, pval, BubbleEvent)
        Case et_KEY_DOWN: '//2
            If pval.CharPressed = 9 Then
                If pval.ItemUID = "ItmBsort" Then
                    If oForm01.Items("ItmBsort").Specific.VALUE = "" Then
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                 ElseIf pval.ItemUID = "LocCode" Then '/운송지역
                    If oForm01.Items("LocCode").Specific.VALUE = "" Then
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If

                ElseIf pval.ItemUID = "Tonnage" Then '차종(톤수)
                    If oForm01.Items("Tonnage").Specific.VALUE = "" Then
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
            End If
           
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
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "ItmBsort" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "LocCode" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "Tonnage" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "Amt" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "PassAmt" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "TWeight" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "DocDate" Then
                        FlushToItemValue pval.ItemUID
                    End If
                End If
           ' Call Raise_EVENT_VALIDATE(FormUID, pval, BubbleEvent)
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
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                
            Case "1281": '찾기
            Case "1282": '추가
                '/추가버튼 클릭시 메트릭스 insertrow
                Call PS_MM004_FormReset
                
'                oMat01.Clear
'                oMat01.FlushToDataSource
'                oMat01.LoadFromDataSource
                
                oForm01.Mode = fm_ADD_MODE
                BubbleEvent = False
                Call LoadCaption
                Exit Sub
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                
            Case "1281": '찾기
                '//Call PS_MM004_FormItemEnabled '//UDO방식
            Case "1282": '추가
'                oMat01.Clear
'                oDS_PS_MM004H.Clear

'                Call LoadCaption
'                Call FormItemEnabled
                '//Call PS_MM004_FormItemEnabled '//UDO방식
                '//Call PS_MM004_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PS_MM004_FormItemEnabled
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
        If pval.ItemUID = "PS_MM004" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "BtnAdd" Then '/추가/확인 버튼클릭
            If oForm01.Mode = fm_ADD_MODE Then
                If PS_MM004_HeaderSpaceLineDel = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                If PS_MM004_AddData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If

                Call PS_MM004_FormReset
                oForm01.Mode = fm_ADD_MODE
                
                Call LoadCaption
                Call PS_MM004_MTX01
                
                oLast_Mode = oForm01.Mode
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                If PS_MM004_HeaderSpaceLineDel = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                If PS_MM004_UpdateData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                Call PS_MM004_FormReset
                oForm01.Mode = fm_ADD_MODE
                
                Call LoadCaption
                Call PS_MM004_MTX01
            End If
            
        ElseIf pval.ItemUID = "BtnSearch" Then '/조회
'            Call PS_MM004_FormReset
            oForm01.Mode = fm_ADD_MODE '/fm_VIEW_MODE
            
            Call LoadCaption
            Call PS_MM004_MTX01
            
        ElseIf pval.ItemUID = "BtnDelete" Then '/삭제
            If Sbo_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", "1", "예", "아니오") = "1" Then
        
                Call PS_MM004_DeleteData
                Call PS_MM004_FormReset
                oForm01.Mode = fm_ADD_MODE '/fm_VIEW_MODE
                
                Call LoadCaption
                Call PS_MM004_MTX01
                
            Else
            
            End If
        ElseIf pval.ItemUID = "BtnPrint" Then
        
            Call PS_MM004_Print_Report01
        
        
        End If
        
  
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_MM004" Then
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
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '//사용자값활성
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "FixCode") '자산번호 포맷서치
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_KEY_DOWN_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_KEY_DOWN_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub Raise_EVENT_CLICK(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CLICK_Error
    Dim DocEntry As String
    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If pval.BeforeAction = True Then
        If pval.ItemUID = "Mat01" Then
            If pval.Row > 0 Then
            
                Call oMat01.SelectRow(pval.Row, True, False)
                
'                Call oForm01.Freeze(True)
                
                DocEntry = oMat01.Columns("DocEntry").Cells(pval.Row).Specific.VALUE
                
                Call PS_MM004_RET(DocEntry)
                
'                sQry = "Select DocEntry, U_ObjType, U_DocNo, U_BPLId, U_DocDate, U_ItmBsort, U_ItmBName, U_LocCode, U_LocName, "
'                sQry = sQry + " U_Way, U_SectionF, U_SectionT, U_Destin, U_TranCard, U_TranCode, U_Tonnage, U_Weight, U_Amt, U_Comments "
'                sQry = sQry + " From [@PS_MM004H] Where DocEntry = '" + DocEntry + "'"
'
'                oRecordSet01.DoQuery sQry
'
'                'DataSource를 이용하여 각 컨트롤에 값을 출력
'                Call oDS_PS_MM004H.setValue("DocEntry", 0, oRecordSet01.Fields("DocEntry").VALUE) '관리번호
'                Call oDS_PS_MM004H.setValue("U_ObjType", 0, Trim(oRecordSet01.Fields("U_ObjType").VALUE)) '참조Object
'                Call oDS_PS_MM004H.setValue("U_DocNo", 0, oRecordSet01.Fields("U_DocNo").VALUE) '참조문서번호
'                Call oDS_PS_MM004H.setValue("U_BPLId", 0, Trim(oRecordSet01.Fields("U_BPLId").VALUE)) '사업장
'                Call oDS_PS_MM004H.setValue("U_DocDate", 0, Format(oRecordSet01.Fields("U_DocDate").VALUE, "YYYYMMDD")) '일자
'                Call oDS_PS_MM004H.setValue("U_ItmBsort", 0, Trim(oRecordSet01.Fields("U_ItmBsort").VALUE)) '분류코드
'                Call oDS_PS_MM004H.setValue("U_ItmBName", 0, Trim(oRecordSet01.Fields("U_ItmBName").VALUE)) '분류명
'                Call oDS_PS_MM004H.setValue("U_LocCode", 0, Trim(oRecordSet01.Fields("U_LocCode").VALUE)) '운송지역코드
'                Call oDS_PS_MM004H.setValue("U_LocName", 0, Trim(oRecordSet01.Fields("U_LocName").VALUE)) '운송지역명
'                Call oDS_PS_MM004H.setValue("U_Way", 0, Trim(oRecordSet01.Fields("U_Way").VALUE)) '편도/회로
'                Call oDS_PS_MM004H.setValue("U_SectionF", 0, Trim(oRecordSet01.Fields("U_SectionF").VALUE)) '운행구간시작
'                Call oDS_PS_MM004H.setValue("U_SectionT", 0, Trim(oRecordSet01.Fields("U_SectionT").VALUE)) '운행구간종료
'                Call oDS_PS_MM004H.setValue("U_Destin", 0, Trim(oRecordSet01.Fields("U_Destin").VALUE)) '도착장소
'                Call oDS_PS_MM004H.setValue("U_TranCard", 0, Trim(oRecordSet01.Fields("U_TranCard").VALUE)) '운송업체
'                Call oDS_PS_MM004H.setValue("U_TranCode", 0, Trim(oRecordSet01.Fields("U_TranCode").VALUE)) '차량번호
'                Call oDS_PS_MM004H.setValue("U_Tonnage", 0, oRecordSet01.Fields("U_Tonnage").VALUE) '차종(톤수)
'                Call oDS_PS_MM004H.setValue("U_Weight", 0, oRecordSet01.Fields("U_Weight").VALUE) '운송량
'                Call oDS_PS_MM004H.setValue("U_Amt", 0, oRecordSet01.Fields("U_Amt").VALUE) '금액
'                Call oDS_PS_MM004H.setValue("U_Comments", 0, Trim(oRecordSet01.Fields("U_Comments").VALUE)) '비고
'
'                oForm01.Mode = fm_UPDATE_MODE
'                Call LoadCaption
'
'                Call oForm01.Freeze(False)
                
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

Dim i&
Dim sQry            As String

Dim sCount&, sSeq&
Dim sCode As String
Dim SCpCode As String
Dim TWeight As Double

Dim oRecordSet01      As SAPbobsCOM.Recordset

Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
    
            If pval.ItemUID = "Way" Then
                If Trim(oForm01.Items("LocCode").Specific.VALUE) <> "" And oForm01.Items("TWeight").Specific.VALUE > 0 Then
                    TWeight = oForm01.Items("TWeight").Specific.VALUE
                    
                    sQry = "Select Amt = Isnull(b.U_Amt,0) From [@PS_MM003H] a Inner Join [@PS_MM003L] b On a.Code = b.Code"
                    sQry = sQry + " Where a.U_BPLId =  '" & Trim(oForm01.Items("BPLId").Specific.VALUE) & "'"
                    sQry = sQry + "   And b.U_LocCode = '" & Trim(oForm01.Items("LocCode").Specific.VALUE) & "'"
                    sQry = sQry + "   And " & TWeight & " Between b.U_KgF And b.U_KgT "
                    oRecordSet01.DoQuery sQry
                    
                    If oForm01.Items("Way").Specific.VALUE = "10" Then
                        oForm01.Items("Amt").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
'                        oForm01.Items("TotAmt").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE) + oForm01.Items("PassAmt").Specific.VALUE
                    ElseIf oForm01.Items("Way").Specific.VALUE = "20" Then
                        oForm01.Items("Amt").Specific.VALUE = (oRecordSet01.Fields(0).VALUE * 0.8)
'                        oForm01.Items("TotAmt").Specific.VALUE = (oRecordSet01.Fields(0).VALUE * 0.8) + oForm01.Items("PassAmt").Specific.VALUE
                    End If
                    
                End If
                
                If oForm01.Items("Way").Specific.VALUE = "10" Then
                    oForm01.Items("SectionF").Specific.VALUE = ""
                    oForm01.Items("SectionT").Specific.VALUE = ""
                    oForm01.Items("ItemName").Specific.VALUE = ""
                    oForm01.Items("Destin").Specific.VALUE = ""
                ElseIf oForm01.Items("Way").Specific.VALUE = "20" Then
                    oForm01.Items("ItmBsort").Specific.VALUE = "102"
                    oForm01.Items("LocCode").Specific.VALUE = "07"
                    oForm01.Items("SectionF").Specific.VALUE = "안강"
                    oForm01.Items("SectionT").Specific.VALUE = "창원"
                    oForm01.Items("ItemName").Specific.VALUE = "부품소재"
                    oForm01.Items("Destin").Specific.VALUE = "창원"
                End If
                
            ElseIf pval.ItemUID = "PassYN" Then
                If oForm01.Items("PassYN").Specific.VALUE = "Y" Then
                    '경유지(혼적) Y이면
                    sQry = "Select PassAmt = Isnull(U_PassAmt,0) From [@PS_MM003H] a Inner Join [@PS_MM003L] b On a.Code = b.Code"
                    sQry = sQry + " Where a.U_BPLId =  '" & Trim(oForm01.Items("BPLId").Specific.VALUE) & "'"
                    sQry = sQry + "   And b.U_LocCode = '" & Trim(oForm01.Items("LocCode").Specific.VALUE) & "'"
                    oRecordSet01.DoQuery sQry
                    
                    oForm01.Items("PassAmt").Specific.VALUE = oRecordSet01.Fields(0).VALUE
'                    oForm01.Items("TotAmt").Specific.VALUE = oForm01.Items("Amt").Specific.VALUE + oRecordSet01.Fields(0).VALUE
                                        
                Else
                    oForm01.Items("PassAmt").Specific.VALUE = 0
'                    oForm01.Items("TotAmt").Specific.VALUE = oForm01.Items("Amt").Specific.VALUE
                End If
            End If
            
'
'
'                sCount = oForm01.Items("CpCode").Specific.ValidValues.Count
'                sSeq = sCount
'                For i = 1 To sCount
'                    oForm01.Items("CpCode").Specific.ValidValues.Remove sSeq - 1, psk_Index
'                    sSeq = sSeq - 1
'                Next i
'
'                '//공정구분에 따른 공정코드변경
'
'                Select Case oForm01.Items("CpGbn").Specific.VALUE
'                    Case "10"
'                        '//멀티게이지-금형래핑
'                        Call oForm01.Items("WorkGbn").Specific.Select("104", psk_ByValue)
'                        oForm01.Items("ItmBsort").Enabled = True
'                        oMat01.Columns("ItmBsort").Editable = True
'                    Case "20"
'                        '//멀티게이지-외경연삭
'                        Call oForm01.Items("WorkGbn").Specific.Select("104", psk_ByValue)
'                        oForm01.Items("ItmBsort").Enabled = False
'                        oMat01.Columns("ItmBsort").Editable = False
'                    Case "30"
'                        '//멀티게이지-포장
'                        Call oForm01.Items("WorkGbn").Specific.Select("104", psk_ByValue)
'                        oForm01.Items("ItmBsort").Enabled = False
'                        oMat01.Columns("ItmBsort").Editable = False
'                    Case "40"
'                        '//휘팅-바렐
'                        Call oForm01.Items("WorkGbn").Specific.Select("101", psk_ByValue)
'                        oForm01.Items("ItmBsort").Enabled = False
'                        oMat01.Columns("ItmBsort").Editable = False
'                    Case "50"
'                        '//휘팅-포장
'                        Call oForm01.Items("WorkGbn").Specific.Select("101", psk_ByValue)
'                        oForm01.Items("ItmBsort").Enabled = False
'                        oMat01.Columns("ItmBsort").Editable = False
'                    Case "60", "70"
'                        '//검사공수 입력
'                        Call oForm01.Items("WorkGbn").Specific.Select("105", psk_ByValue)
'                    Case Else
'
'                End Select
'
'            ElseIf pval.ItemUID = "OrdType" Then
'
'                If oForm01.Items("OrdType").Specific.VALUE = "10" Then
'                    '// 실동입력
'                    oForm01.Items("CpGbn").Enabled = True
''                    oForm01.Items("CntcCode").Enabled = True '사번 TextBox 활성, 비활성 기능 제외(2013.05.03 송명규 추가)
'                    oForm01.Items("CpCode").Enabled = True
'                    oForm01.Items("ItmBsort").Enabled = True
'                    oMat01.Columns("NCode").Editable = False
'                Else
'                    '// 비가동입력
'                    oForm01.Items("CpGbn").Enabled = False
'
'                    Call oForm01.Items("CpGbn").Specific.Select("", psk_ByValue)
''                    oForm01.Items("CntcCode").Enabled = False '사번 TextBox 활성, 비활성 기능 제외(2013.05.03 송명규 추가)
'
'                    oForm01.Items("CpCode").Enabled = False
'                    Call oForm01.Items("CpCode").Specific.Select("", psk_ByValue)
'                    oForm01.Items("ItmBsort").Enabled = False
'                    oMat01.Columns("NCode").Editable = True
'
'
'                End If
'            ElseIf pval.ItemUID = "CpCode" Then '// 공정코드
'                sQry = "SELECT U_CpName From [@PS_PP001L] Where U_CpCode = '" & Trim(oForm01.Items("CpCode").Specific.VALUE) & "'"
'                oRecordSet01.DoQuery sQry
'                oForm01.Items("CpName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
'            End If



    End If
    Call oForm01.Freeze(False)
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
'/현재 사용안함
Private Sub Raise_EVENT_VALIDATE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_VALIDATE_Error
    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
'        If pval.ItemChanged = True Then
'            If (pval.ItemUID = "Mat01") Then
'                If (pval.ColUID = "ItemCode") Then
'                    '//기타작업
'                    Call oDS_PS_MM004L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Value)
'                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_MM004L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
'                        PS_MM004_AddMatrixRow (pval.Row)
'                    End If
'                Else
'                    Call oDS_PS_MM004L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Value)
'                End If
'            Else
'                If (pval.ItemUID = "DocEntry") Then
'                    Call oDS_PS_MM004H.setValue(pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Value)
'                ElseIf (pval.ItemUID = "CardCode") Then
'                    Call oDS_PS_MM004H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Value)
'                    Call oDS_PS_MM004H.setValue("U_CardName", 0, MDC_GetData.Get_ReData("CardName", "CardCode", "[OCRD]", "'" & oForm01.Items(pval.ItemUID).Specific.Value & "'"))
'                Else
'                    Call oDS_PS_MM004H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Value)
'                End If
'            End If
'            oMat01.LoadFromDataSource
'            oMat01.AutoResizeColumns
'            oForm01.Update
'        End If
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
        Call PS_MM004_FormItemEnabled
        '//Call PS_MM004_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_MM004_FormResize
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
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_MM004H", "U_CardCode,U_CardName")
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

Private Function PS_MM004_CreateItems() As Boolean
On Error GoTo PS_MM004_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
         
    
    
    Set oDS_PS_MM004H = oForm01.DataSources.DBDataSources("@PS_MM004H")
    Set oDS_PS_MM004L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '/유저데이타 속성 선언 날짜형식
    Call oForm01.DataSources.UserDataSources.Add("SDocDateF", dt_DATE, 8)
    Call oForm01.Items("SDocDateF").Specific.DataBind.SetBound(True, "", "SDocDateF")
    
    Call oForm01.DataSources.UserDataSources.Add("SDocDateT", dt_DATE, 8)
    Call oForm01.Items("SDocDateT").Specific.DataBind.SetBound(True, "", "SDocDateT")
    
    Call oForm01.DataSources.UserDataSources.Add("SBPLId", dt_SHORT_TEXT, 5)
    Call oForm01.Items("SBPLId").Specific.DataBind.SetBound(True, "", "SBPLId")
    
    Call oForm01.DataSources.UserDataSources.Add("SItmBsort", dt_SHORT_TEXT, 10)
    Call oForm01.Items("SItmBsort").Specific.DataBind.SetBound(True, "", "SItmBsort")
    
    Call oForm01.DataSources.UserDataSources.Add("SLocCode", dt_SHORT_TEXT, 10)
    Call oForm01.Items("SLocCode").Specific.DataBind.SetBound(True, "", "SLocCode")
    
    Call oForm01.DataSources.UserDataSources.Add("SWay", dt_SHORT_TEXT, 10)
    Call oForm01.Items("SWay").Specific.DataBind.SetBound(True, "", "SWay")
    
    Call oForm01.DataSources.UserDataSources.Add("STranCard", dt_SHORT_TEXT, 10)
    Call oForm01.Items("STranCard").Specific.DataBind.SetBound(True, "", "STranCard")
    
    Call oForm01.DataSources.UserDataSources.Add("STonnage", dt_SHORT_TEXT, 10)
    Call oForm01.Items("STonnage").Specific.DataBind.SetBound(True, "", "STonnage")
    
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_MM004_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_MM004_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function
'/콤보박스 set
Sub PS_MM004_ComboBox_Setting()
On Error GoTo PS_MM004_ComboBox_Setting_Error

    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim sQry            As String
    
    Dim oRecordSet01      As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    '//콤보에 기본값설정
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM004", "CpGbn", "ItemCode", "01", "완제품")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM004", "Mat01", "ItemCode", "02", "반제품")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("Column"), "PS_MM004", "Mat01", "ItemCode")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM004", "CpGbn", "", "10", "멀티게이지-금형래핑")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM004", "CpGbn", "", "20", "멀티게이지-외경연삭")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("CpGbn").Specific, "10", "멀티게이지")
        
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    Set oCombo = oForm01.Items("BPLId").Specific
    Call oCombo.Select(MDC_PS_Common.User_BPLId, psk_ByValue)
    
    Call MDC_SetMod.Set_ComboList(oForm01.Items("SBPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    Set oCombo = oForm01.Items("SBPLId").Specific
    Call oCombo.Select(MDC_PS_Common.User_BPLId, psk_ByValue)
    
    If MDC_PS_Common.User_BPLId = "1" Then
        oForm01.Items("SectionF").Specific.VALUE = "창원"
    End If
    
    
    oForm01.Items("SItmBsort").Specific.ValidValues.Add "", ""
    Call MDC_SetMod.Set_ComboList(oForm01.Items("SItmBsort").Specific, "select Code, Name from [@PSH_ITMBSORT] Where U_PudYN = 'Y' order by Code", "101", False, False)
    Set oCombo = oForm01.Items("SItmBsort").Specific
    Call oCombo.Select("", psk_ByValue)
    
    'Call MDC_SetMod.Set_ComboList(oMat01.Columns("NCode").Cells( Items("WorkGbn").Specific, "select Code, Name from [@PSH_ITMBSORT] Where U_PudYN = 'Y' order by Code", "101", False, False)
    
    Set oCombo = oForm01.Items("Way").Specific
    oCombo.ValidValues.Add "10", "편도"
    oCombo.ValidValues.Add "20", "회로"
    oCombo.Select "10", psk_ByValue
    
    Set oCombo = oForm01.Items("PassYN").Specific
    oCombo.ValidValues.Add "N", "일반"
    oCombo.ValidValues.Add "Y", "경유지"
    oCombo.Select "N", psk_ByValue
    
    oForm01.Items("SLocCode").Specific.ValidValues.Add "", ""
    Call MDC_SetMod.Set_ComboList(oForm01.Items("SLocCode").Specific, "select Distinct b.U_LocCode, b.U_LocName from [@PS_MM003H] a Inner Join [@PS_MM003L] b On a.Code = b.Code Where a.U_BPLId = '" + MDC_PS_Common.User_BPLId + "' order by b.U_LocCode", "", False, False)
    Set oCombo = oForm01.Items("SLocCode").Specific
    Call oCombo.Select("", psk_ByValue)
    
    oForm01.Items("STonnage").Specific.ValidValues.Add "", ""
    Call MDC_SetMod.Set_ComboList(oForm01.Items("STonnage").Specific, "select Distinct b.U_Tonnage, b.U_TonNm from [@PS_MM003H] a Inner Join [@PS_MM003L] b On a.Code = b.Code Where a.U_BPLId = '" + MDC_PS_Common.User_BPLId + "' order by b.U_Tonnage", "", False, False)
    Set oCombo = oForm01.Items("STonnage").Specific
    Call oCombo.Select("", psk_ByValue)
    
'    oForm01.Items("STranCard").Specific.ValidValues.Add "", ""
'    Call MDC_SetMod.Set_ComboList(oForm01.Items("STranCard").Specific, "select Distinct a.U_TranCard, a.U_TranCard from [@PS_MM004H] a Where a.U_BPLId = '" + MDC_PS_Common.User_BPLId + "' order by a.U_TranCard", "", False, False)
'    Set oCombo = oForm01.Items("STranCard").Specific
'    Call oCombo.Select("", psk_ByValue)
    
    oForm01.Items("SWay").Specific.ValidValues.Add "", ""
    oForm01.Items("SWay").Specific.ValidValues.Add "10", "편도"
    oForm01.Items("SWay").Specific.ValidValues.Add "20", "회로"
    
        
    Call oMat01.Columns("Way").ValidValues.Add("10", "편도")
    Call oMat01.Columns("Way").ValidValues.Add("20", "회로")
            
    sQry = "SELECT BPLId, BPLName FROM OBPL order by BPLId"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oMat01.Columns("BPLId").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
'    Call MDC_SetMod.Set_ComboList(oForm01.Items("CpCode").Specific, "select U_CpCode, U_CpName from [@PS_PP001L] Where Code = 'CP301' ", "CP30112", False, False)
    
    Call oForm01.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
PS_MM004_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_MM004_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_MM004_CF_ChooseFromList()
On Error GoTo PS_MM004_CF_ChooseFromList_Error

    Exit Sub
PS_MM004_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM004_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_MM004_FormItemEnabled()
On Error GoTo PS_MM004_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
        '//Call PS_MM004_FormClear '//UDO방식
'        Call oForm01.EnableMenu("1281", True) '//찾기
'        Call oForm01.EnableMenu("1282", False) '//추가
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
'        Call oForm01.EnableMenu("1281", False) '//찾기
'        Call oForm01.EnableMenu("1282", True) '//추가
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_MM004_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_MM004_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_MM004_EnableMenus()
On Error GoTo PS_MM004_EnableMenus_Error
    
    Exit Sub
PS_MM004_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM004_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PS_MM004_FormResize()
On Error GoTo PS_MM004_FormResize_Error
    
    Exit Sub
PS_MM004_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM004_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub PS_MM004_FormReset()
'******************************************************************************
'Function ID : PS_MM004_FormReset()
'해당모듈 : PS_MM004
'기능 : 화면 초기화
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_MM004_FormReset_Error

    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    
    '관리번호
    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PS_MM004H]"
    Call RecordSet01.DoQuery(sQry)
    
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        Call oDS_PS_MM004H.setValue("DocEntry", 0, 1)
    Else
        Call oDS_PS_MM004H.setValue("DocEntry", 0, Trim(RecordSet01.Fields(0).VALUE) + 1)
    End If
    
    Dim User_BPLId As String
'    User_BPLID = MDC_PS_Common.User_BPLID
    
    '//////////기준정보//////////
'    Call oDS_PS_MM004H.setValue("U_CLTCOD", 0, User_BPLID) '사업장

    Call oDS_PS_MM004H.setValue("U_DocDate", 0, Format(Now, "yyyyMMdd")) '일자
    Call oDS_PS_MM004H.setValue("U_ItmBsort", 0, "") '품목분류
    Call oDS_PS_MM004H.setValue("U_ItmBName", 0, "") '품목분류명
    Call oDS_PS_MM004H.setValue("U_ObjType", 0, "") '참조Object
    Call oDS_PS_MM004H.setValue("U_DocNo", 0, "") '참조문서번호
    Call oDS_PS_MM004H.setValue("U_LocCode", 0, "") '운송지역코드
    Call oDS_PS_MM004H.setValue("U_LocName", 0, "") '운송지역명
    Call oDS_PS_MM004H.setValue("U_Way", 0, "10") '편도/회로
  
    Call oDS_PS_MM004H.setValue("U_SectionF", 0, "") '운행구간1
    Call oDS_PS_MM004H.setValue("U_SectionT", 0, "") '운행구간2
    Call oDS_PS_MM004H.setValue("U_Destin", 0, "") '도착장소
    Call oDS_PS_MM004H.setValue("U_TranCard", 0, "") '운송업체
    Call oDS_PS_MM004H.setValue("U_TranCode", 0, "") '차량번호
    Call oDS_PS_MM004H.setValue("U_Tonnage", 0, "") '차종(톤수)
    Call oDS_PS_MM004H.setValue("U_ItemName", 0, "") '품명
    Call oDS_PS_MM004H.setValue("U_Weight", 0, "") '운송량
    Call oDS_PS_MM004H.setValue("U_TWeight", 0, "") '계근량
    Call oDS_PS_MM004H.setValue("U_PassYN", 0, "N") '경유Y/N
    Call oDS_PS_MM004H.setValue("U_Amt", 0, "") '기본운임
    Call oDS_PS_MM004H.setValue("U_PassAmt", 0, "") '경유금액
    Call oDS_PS_MM004H.setValue("U_TotAmt", 0, "") '총금액
    Call oDS_PS_MM004H.setValue("U_Comments", 0, "") '비고
    
    '//////////조회정보//////////
    Call oForm01.Items("SLocCode").Specific.Select("", psk_ByValue) '운송지역
    Call oForm01.Items("SWay").Specific.Select("", psk_ByValue) '편도/회로
    'Call oForm01.Items("STranCard").Specific.Select("", psk_ByValue) '운송업체
    Call oForm01.Items("STonnage").Specific.Select("", psk_ByValue) '차종(톤수)
    oForm01.Items("SObjType").Specific.VALUE = ""
    oForm01.Items("SDocNo").Specific.VALUE = ""

'    oForm01.Items("SFrDate").Specific.VALUE = Format(Now, "YYYY.MM")
'    oForm01.Items("SToDate").Specific.VALUE = Format(Now, "YYYY.MM")
    
    Call oForm01.Items("ItmBsort").Click
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)

Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_MM004_FormReset_Error:
    Call oForm01.Freeze(False)
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PS_MM004_FormReset_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub


Private Sub PS_MM004_Print_Report01()

    Dim DocNum As String
    Dim ErrNum As Integer
    Dim WinTitle As String
    Dim ReportName As String
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PS_MM004_Print_Report01_Error
    
    Dim DocEntry As Long
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
     '/ ODBC 연결 체크
    Call ConnectODBC
    
    '//인자 MOVE , Trim 시키기..
    DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

    WinTitle = "[PS_MM004] 운송확인증"
    
    ReportName = "PS_MM004_10.rpt"
    
    
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
    sQry = "EXEC [PS_MM004_10] '" & DocEntry & "'"
    
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo PS_MM004_Print_Report01_Error
    End If
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    Set oRecordSet = Nothing
    Exit Sub
    
PS_MM004_Print_Report01_Error:
    If ErrNum = 1 Then
        Set oRecordSet = Nothing
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PS_MM004_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
    
End Sub

Private Sub PS_MM004_RET(ByVal oDocEntry As String)
On Error GoTo PS_MM004_RET_Error
    Dim DocEntry As String
    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
                
    sQry = "Select DocEntry, U_ObjType, U_DocNo, U_BPLId, U_DocDate, U_ItmBsort, U_ItmBName, U_LocCode, U_LocName, "
    sQry = sQry + " U_Way, U_SectionF, U_SectionT, U_Destin, U_TranCard, U_TranCode, U_Tonnage, U_ItemName, "
    sQry = sQry + " U_Weight, U_TWeight, U_PassYN, U_Amt, U_PassAmt, U_TotAmt, U_Comments "
    sQry = sQry + " From [@PS_MM004H] Where DocEntry = '" + oDocEntry + "'"
    
    oRecordSet01.DoQuery sQry
                    
    'DataSource를 이용하여 각 컨트롤에 값을 출력
    Call oDS_PS_MM004H.setValue("DocEntry", 0, oRecordSet01.Fields("DocEntry").VALUE) '관리번호
    Call oDS_PS_MM004H.setValue("U_ObjType", 0, Trim(oRecordSet01.Fields("U_ObjType").VALUE)) '참조Object
    Call oDS_PS_MM004H.setValue("U_DocNo", 0, oRecordSet01.Fields("U_DocNo").VALUE) '참조문서번호
    Call oDS_PS_MM004H.setValue("U_BPLId", 0, Trim(oRecordSet01.Fields("U_BPLId").VALUE)) '사업장
    Call oDS_PS_MM004H.setValue("U_DocDate", 0, Format(oRecordSet01.Fields("U_DocDate").VALUE, "YYYYMMDD")) '일자
    Call oDS_PS_MM004H.setValue("U_ItmBsort", 0, Trim(oRecordSet01.Fields("U_ItmBsort").VALUE)) '분류코드
    Call oDS_PS_MM004H.setValue("U_ItmBName", 0, Trim(oRecordSet01.Fields("U_ItmBName").VALUE)) '분류명
    Call oDS_PS_MM004H.setValue("U_LocCode", 0, Trim(oRecordSet01.Fields("U_LocCode").VALUE)) '운송지역코드
    Call oDS_PS_MM004H.setValue("U_LocName", 0, Trim(oRecordSet01.Fields("U_LocName").VALUE)) '운송지역명
    Call oDS_PS_MM004H.setValue("U_Way", 0, Trim(oRecordSet01.Fields("U_Way").VALUE)) '편도/회로
    Call oDS_PS_MM004H.setValue("U_SectionF", 0, Trim(oRecordSet01.Fields("U_SectionF").VALUE)) '운행구간시작
    Call oDS_PS_MM004H.setValue("U_SectionT", 0, Trim(oRecordSet01.Fields("U_SectionT").VALUE)) '운행구간종료
    Call oDS_PS_MM004H.setValue("U_Destin", 0, Trim(oRecordSet01.Fields("U_Destin").VALUE)) '도착장소
    Call oDS_PS_MM004H.setValue("U_TranCard", 0, Trim(oRecordSet01.Fields("U_TranCard").VALUE)) '운송업체
    Call oDS_PS_MM004H.setValue("U_TranCode", 0, Trim(oRecordSet01.Fields("U_TranCode").VALUE)) '차량번호
    Call oDS_PS_MM004H.setValue("U_Tonnage", 0, oRecordSet01.Fields("U_Tonnage").VALUE) '차종(톤수)
    Call oDS_PS_MM004H.setValue("U_ItemName", 0, oRecordSet01.Fields("U_ItemName").VALUE) '품명
    Call oDS_PS_MM004H.setValue("U_Weight", 0, oRecordSet01.Fields("U_Weight").VALUE) '운송량
    Call oDS_PS_MM004H.setValue("U_TWeight", 0, oRecordSet01.Fields("U_TWeight").VALUE) '계근중량
    Call oDS_PS_MM004H.setValue("U_PassYN", 0, oRecordSet01.Fields("U_PassYN").VALUE) '경유여부
    Call oDS_PS_MM004H.setValue("U_Amt", 0, oRecordSet01.Fields("U_Amt").VALUE) '기본운임
    Call oDS_PS_MM004H.setValue("U_PassAmt", 0, oRecordSet01.Fields("U_PassAmt").VALUE) '경유지운임
    Call oDS_PS_MM004H.setValue("U_TotAmt", 0, oRecordSet01.Fields("U_TotAmt").VALUE) '총금액
    Call oDS_PS_MM004H.setValue("U_Comments", 0, Trim(oRecordSet01.Fields("U_Comments").VALUE)) '비고
    
    oForm01.Mode = fm_UPDATE_MODE
    Call LoadCaption
    
    Call oForm01.Freeze(False)
    
    Exit Sub
PS_MM004_RET_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM004_RET_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
