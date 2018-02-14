VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_QM240"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_QM240.cls
'//  Module         : QM
'//  Desc           : 계측기 등록
'//  FormType       : PS_QM240
'//  Create Date    : 2012.05.30 N.G.Y
'//  Copyright  (c) Morning Data
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_QM240H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_QM240L As SAPbouiCOM.DBDataSource '등록라인

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
    Check As Boolean
    OPDNNo As Long
    PDN1No As Long
End Type

Private oLast_Mode&

Private ItemInformation() As ItemInformations
Private ItemInformationCount As Long
'/ 추가 확인 갱신 버튼 이름 변경
Private Sub LoadCaption()
On Error GoTo LoadCaption_Error
    If oForm01.Mode = fm_ADD_MODE Then
        oForm01.Items("Btn_save").Specific.Caption = "추가"
        oForm01.Items("Btn_del").Enabled = False
'    ElseIf oForm01.Mode = fm_OK_MODE Then
'        oForm01.Items("Btn_save").Specific.Caption = "확인"
    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
        oForm01.Items("Btn_save").Specific.Caption = "수정"
        oForm01.Items("Btn_del").Enabled = True
    End If

    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadCaption_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'//조회데이타 가져오기
Public Sub LoadData()
On Error GoTo LoadData_Error
    Dim i As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim sBPLId$, SGCode$
    
    sBPLId = Trim(oForm01.Items("SBPLId").Specific.VALUE)
    SGCode = Trim(oForm01.Items("SGCode").Specific.VALUE)
    
    If SGCode = "" Then SGCode = "%"
 
        
    sQry = "EXEC [PS_QM240_01] '" & sBPLId & "','" & SGCode & "'"
    oRecordSet01.DoQuery sQry
    
    oMat01.Clear
    oDS_PS_QM240L.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
                
    If (oRecordSet01.RecordCount = 0) Then
        MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
        Set oRecordSet01 = Nothing
        oForm01.Mode = fm_ADD_MODE
        
        Call Add_MatrixRow(0, True)
        Call LoadCaption
        Exit Sub
    End If
    
    oForm01.Freeze True
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_QM240L.Size Then
            oDS_PS_QM240L.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_QM240L.Offset = i
        
        oDS_PS_QM240L.setValue "U_LineNum", i, i + 1
        oDS_PS_QM240L.setValue "U_ColNum01", i, Trim(oRecordSet01.Fields(0).VALUE)
        oDS_PS_QM240L.setValue "U_ColReg01", i, Trim(oRecordSet01.Fields(1).VALUE)
        oDS_PS_QM240L.setValue "U_ColReg02", i, Trim(oRecordSet01.Fields(2).VALUE)
        oDS_PS_QM240L.setValue "U_ColReg03", i, Trim(oRecordSet01.Fields(3).VALUE)
        oDS_PS_QM240L.setValue "U_ColReg04", i, Trim(oRecordSet01.Fields(4).VALUE)
        oDS_PS_QM240L.setValue "U_ColReg05", i, Trim(oRecordSet01.Fields(5).VALUE)
        oDS_PS_QM240L.setValue "U_ColReg06", i, Trim(oRecordSet01.Fields(6).VALUE)
        oDS_PS_QM240L.setValue "U_ColReg07", i, Trim(oRecordSet01.Fields(7).VALUE)
        oDS_PS_QM240L.setValue "U_ColReg08", i, Trim(oRecordSet01.Fields(8).VALUE)
        oDS_PS_QM240L.setValue "U_ColSum01", i, Trim(oRecordSet01.Fields(9).VALUE)
        oDS_PS_QM240L.setValue "U_ColReg09", i, Trim(oRecordSet01.Fields(10).VALUE)
        oDS_PS_QM240L.setValue "U_ColDt01", i, Format(Trim(oRecordSet01.Fields(11).VALUE), "YYYYMMDD")
        
        
        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
    Next
    
    
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    ProgBar01.Stop
    oForm01.Freeze False
  
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadData_Error:
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "LoadData_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub
Public Sub DeleteData()
On Error GoTo DeleteData_Error
    Dim i As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry$, ItemCode$
    Dim Check As String
    
    If oForm01.Mode = fm_UPDATE_MODE Then
        DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
        
        sQry = "Select Count(*) From [@PS_QM240H] where DocEntry = '" & DocEntry & "'"
        oRecordSet01.DoQuery sQry
        
        If (oRecordSet01.RecordCount = 0) Then
            MDC_Com.MDC_GF_Message "삭제대상이 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
            Set oRecordSet01 = Nothing
            oForm01.Mode = fm_ADD_MODE
        Else
            sQry = "Delete From [@PS_QM240H] where DocEntry = '" & DocEntry & "'"
            oRecordSet01.DoQuery sQry
        End If
    End If
    
    Call FormReset
    
    oForm01.Mode = fm_ADD_MODE
    
    oForm01.Items("Btn_ret").Click ct_Regular
    

'    oMat01.Clear
'    oMat01.FlushToDataSource
'    oMat01.LoadFromDataSource
'    Call Add_MatrixRow(0, True)
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
DeleteData_Error:
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "DeleteData_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub
'/메트릭스 Row추가
Sub Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_QM240L.InsertRecord (oRow)
    End If
    
    oMat01.AddRow
    oDS_PS_QM240L.Offset = oRow
    oDS_PS_QM240L.setValue "U_LineNum", oRow, oRow + 1
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub
'// 데이타 insert
Public Function Updatedata(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
On Error GoTo Updatedata_Error
    Dim i As Integer
    Dim j As Integer
    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry$, GCode$, Code$, BPLId$, CdName$, Size$, Maker$
    Dim ProCd$, Sect$, Remark$
    Dim DocDate$
    Dim Cycle As Integer
    
    DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
    BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
    DocDate = Trim(oForm01.Items("DocDate").Specific.VALUE)
    GCode = Trim(oForm01.Items("GCode").Specific.VALUE)
    Code = Trim(oForm01.Items("Code").Specific.VALUE)
    CdName = Trim(oForm01.Items("CdName").Specific.VALUE)
    Size = Trim(oForm01.Items("Size").Specific.VALUE)
    Maker = Trim(oForm01.Items("Maker").Specific.VALUE)
    ProCd = Trim(oForm01.Items("ProCd").Specific.VALUE)
    Sect = Trim(oForm01.Items("Sect").Specific.VALUE)
    Cycle = Trim(oForm01.Items("Cycle").Specific.VALUE)
    Remark = Trim(oForm01.Items("Remark").Specific.VALUE)
    
    
    If Trim(DocEntry) = "" Then
        MDC_Com.MDC_GF_Message "수정할 항목이 없습니다. 수정하실려면 항목을 선택을 하세요!", "E"
        Updatedata = False
        Exit Function
    End If
        
    
    sQry = "Update [@PS_QM240H]"
    sQry = sQry & " set "
    sQry = sQry & " U_BPLId = '" & BPLId & "',"
    sQry = sQry & " U_GCode = '" & GCode & "',"
    sQry = sQry & " U_Code = '" & Code & "',"
    sQry = sQry & " U_CdName = '" & CdName & "',"
    sQry = sQry & " U_Size = '" & Size & "',"
    sQry = sQry & " U_Maker = '" & Maker & "',"
    sQry = sQry & " U_ProCd  = '" & ProCd & "',"
    sQry = sQry & " U_Sect  = '" & Sect & "',"
    sQry = sQry & " U_Cycle  = '" & Cycle & "',"
    sQry = sQry & " U_Remark = '" & Remark & "'"
            
    sQry = sQry & " Where DocEntry = '" & DocEntry & "'"
    RecordSet01.DoQuery sQry
    
    
    MDC_Com.MDC_GF_Message "수정 완료!", "S"
    
    Set RecordSet01 = Nothing
    Updatedata = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Updatedata_Error:
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "Updatedata_Error:" & Err.Number & " - " & Err.Description, "E"
End Function
'// 데이타 insert
Public Function Add_PurchaseDemand(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
On Error GoTo Add_PurchaseDemand_Error
    Dim i As Integer
    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim RecordSet02 As SAPbobsCOM.Recordset
    Set RecordSet02 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry$, BPLId$, GCode$, Code$, CdName$, Size$, Maker$
    Dim ProCd$, Sect$, Remark$
    Dim DocDate$
    Dim Cycle As Integer
    
    DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
    BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
    DocDate = Trim(oForm01.Items("DocDate").Specific.VALUE)
    GCode = Trim(oForm01.Items("GCode").Specific.VALUE)
    Code = Trim(oForm01.Items("Code").Specific.VALUE)
    CdName = Trim(oForm01.Items("CdName").Specific.VALUE)
    Size = Trim(oForm01.Items("Size").Specific.VALUE)
    Maker = Trim(oForm01.Items("Maker").Specific.VALUE)
    ProCd = Trim(oForm01.Items("ProCd").Specific.VALUE)
    Sect = Trim(oForm01.Items("Sect").Specific.VALUE)
    Cycle = Trim(oForm01.Items("Cycle").Specific.VALUE)
    Remark = Trim(oForm01.Items("Remark").Specific.VALUE)
    
    
    sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_QM240H]"
    RecordSet01.DoQuery sQry
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        DocEntry = 1
    Else
        DocEntry = Trim(RecordSet01.Fields(0).VALUE) + 1
    End If
        
    sQry = "Select right('00' + rtrim(Convert(Char(3),Convert(Integer,right(Max(U_Code),3)) + 1)),3) From [@PS_QM240H] Where U_BPLId = '" & BPLId & "' And U_GCode = '" & GCode & "'"
    RecordSet01.DoQuery sQry
    If Trim(RecordSet01.Fields(0).VALUE) = "" Then
        Code = BPLId + GCode + "001"
    Else
        Code = BPLId + GCode + Trim(RecordSet01.Fields(0).VALUE)
    End If
    
    
    sQry = "INSERT INTO [@PS_QM240H]"
    sQry = sQry & " ("
    sQry = sQry & " DocEntry,"
    sQry = sQry & " DocNum,"
    sQry = sQry & " U_BPLId,"
    sQry = sQry & " U_DocDate,"
    sQry = sQry & " U_GCode,"
    sQry = sQry & " U_Code,"
    sQry = sQry & " U_CdName,"
    sQry = sQry & " U_Size,"
    sQry = sQry & " U_Maker,"
    sQry = sQry & " U_ProCd,"
    sQry = sQry & " U_Sect,"
    sQry = sQry & " U_Cycle,"
    sQry = sQry & " U_Remark"
    sQry = sQry & " ) "
    sQry = sQry & "VALUES("
    sQry = sQry & DocEntry & ","
    sQry = sQry & DocEntry & ","
    sQry = sQry & "'" & BPLId & "',"
    sQry = sQry & "'" & DocDate & "',"
    sQry = sQry & "'" & GCode & "',"
    sQry = sQry & "'" & Code & "',"
    sQry = sQry & "'" & CdName & "',"
    sQry = sQry & "'" & Size & "',"
    sQry = sQry & "'" & Maker & "',"
    sQry = sQry & "'" & ProCd & "',"
    sQry = sQry & "'" & Sect & "',"
    sQry = sQry & "'" & Cycle & "',"
    sQry = sQry & "'" & Remark & "'"
    sQry = sQry & ")"
    RecordSet02.DoQuery sQry
    
    
    MDC_Com.MDC_GF_Message "등록 완료!", "S"
    
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    Add_PurchaseDemand = True
    
    
    
    
    
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_PurchaseDemand_Error:
    Add_PurchaseDemand = False
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "Add_PurchaseDemand_Error:" & Err.Number & " - " & Err.Description, "E"
End Function
'/ 입력필수 사항 Check
Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum          As Integer
    
    ErrNum = 0

    '// Check
    Select Case True
        Case Trim(oForm01.Items("GCode").Specific.VALUE) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("Cycle").Specific.VALUE) = 0
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "분류는 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "검사주기는 필수사항입니다. 확인하세요.", "E"
    End If
    HeaderSpaceLineDel = False
End Function
'/ 메트릭스 필수 사항 check
Private Function MatrixSpaceLineDel() As Boolean
On Error GoTo MatrixSpaceLineDel_Error
    Dim i               As Long
    Dim ErrNum          As Integer
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    Dim sQry            As String


    Set oRecordSet01 = Nothing
    MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
MatrixSpaceLineDel_Error:
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
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function
'/코드의 코드명 보여주기
Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error
    Dim i&
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim ItemCode$, Qty&
    Dim Calculate_Weight As Double
    Dim oRecordSet01    As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
        Case "ItemCode" '/사번
            sQry = "Select ItemName From OITM Where ItemCode = '" & Trim(oForm01.Items("ItemCode").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            oForm01.Items("ItemName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            
            sQry = "Select a.CardCode, a.CardName, a.DocEntry, b.LineNum From ORDR a, RDR1 b Where a.DocEntry = b.DocEntry and b.ItemCode = '" & Trim(oForm01.Items("ItemCode").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            oForm01.Items("CardCode").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            oForm01.Items("CardName").Specific.VALUE = Trim(oRecordSet01.Fields(1).VALUE)
            oForm01.Items("ORDRNum").Specific.VALUE = Trim(oRecordSet01.Fields(2).VALUE)
            oForm01.Items("RDR1Num").Specific.VALUE = Trim(oRecordSet01.Fields(3).VALUE)
            
    End Select

'   Set oRecordSet = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
FlushToItemValue_Error:
    MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub
'/폼의 아이템 사용지정
Sub FormItemEnabled()
On Error GoTo FormItemEnabled_Error
    If (oForm01.Mode = fm_ADD_MODE) Then
'        oMat01.Columns("ItemCode").Cells(1).Click ct_Regular
'        oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
'        oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub


'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_QM240.srf")
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
    
    oFormUniqueID01 = "PS_QM240_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
     
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_QM240_CreateItems
    Call PS_QM240_ComboBox_Setting
    Call PS_QM240_CF_ChooseFromList
    Call PS_QM240_EnableMenus
    Call PS_QM240_SetDocument(oFromDocEntry01)
    Call PS_QM240_FormResize
    
    
    Call Add_MatrixRow(0, True)
    Call LoadCaption
    Call FormItemEnabled
    
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1285"), False        '// 복원
    oForm01.EnableMenu ("1284"), True         '// 취소
    oForm01.EnableMenu ("1293"), True         '// 행삭제
    oForm01.EnableMenu ("1281"), False
    oForm01.EnableMenu ("1282"), True
        
    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
        
    sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_QM240H]"
    RecordSet01.DoQuery sQry
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = Trim(RecordSet01.Fields(0).VALUE) + 1
    End If
        
    
        
    oForm01.Update
    oForm01.Freeze False
    
    oForm01.Visible = True
    Set oXmlDoc01 = Nothing
    
    '//일자 Set
     oForm01.Items("DocDate").Specific.VALUE = Format(Date, "YYYYMMDD")

     'oForm01.Items("GCode").Click ct_Regular
    Exit Sub
LoadForm_Error:
    oForm01.Update
    oForm01.Freeze False
    Set oXmlDoc01 = Nothing
    Set oForm01 = Nothing
    Sbo_Application.SetStatusBarMessage "Form_Load Error:" & Err.Description, bmt_Short, True
End Sub
'/아이템 변경 이벤트
Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_ItemEvent_Error


    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            Call Raise_EVENT_ITEM_PRESSED(FormUID, pval, BubbleEvent)
        Case et_KEY_DOWN: '//2
            If pval.CharPressed = 9 Then
'                If pval.ItemUID = "ItemCode" Then
'                    If oForm01.Items("ItemCode").Specific.VALUE = "" Then
'                        Sbo_Application.ActivateMenuItem ("7425")
'                        BubbleEvent = False
'                    End If
'
'                End If
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
'                    If pval.ItemUID = "ItemCode" Then
'                        FlushToItemValue pval.ItemUID
'                    End If
                    If pval.ItemUID = "DocDate" Then
                    
'                        Dim sQry As String
'                        Dim DocDate As String
'                        Dim RecordSet01 As SAPbobsCOM.Recordset
'                        Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
'
'                        DocDate = Trim(oForm01.Items("DocDate").Specific.VALUE)
'
'                        sQry = "Select left('" & DocDate & "',6) + '01', Convert(char(8),Dateadd(dd, -1, left(convert(char(8),Dateadd(mm, 1, '" & DocDate & "'),112), 6) + '01'),112)"
'                        RecordSet01.DoQuery sQry
'                        oForm01.Items("DocDateF").Specific.VALUE = Trim(RecordSet01.Fields(0).VALUE)
'                        oForm01.Items("DocDateT").Specific.VALUE = Trim(RecordSet01.Fields(1).VALUE)
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
                
                Call FormReset
                
'                oMat01.Clear
'                oMat01.FlushToDataSource
'                oMat01.LoadFromDataSource
                
                oForm01.Mode = fm_ADD_MODE
                BubbleEvent = False
                Call LoadCaption
                
                'oForm01.Items("GCode").Click ct_Regular
                
                
                Exit Sub
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
                '//Call PS_QM240_FormItemEnabled '//UDO방식
            Case "1282": '추가
'                oMat01.Clear
'                oDS_PS_QM240H.Clear

'                Call LoadCaption
'                Call FormItemEnabled
                '//Call PS_QM240_FormItemEnabled '//UDO방식
                '//Call PS_QM240_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PS_QM240_FormItemEnabled
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
        If pval.ItemUID = "PS_QM240" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "Btn_save" Then '/추가/확인 버튼클릭
            If oForm01.Mode = fm_ADD_MODE Then
                If HeaderSpaceLineDel = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                If Add_PurchaseDemand(pval) = False Then
                    BubbleEvent = False
                    Exit Sub
                End If

                Call FormReset
                oForm01.Mode = fm_ADD_MODE
                
                Call LoadCaption
                Call LoadData
                
                oLast_Mode = oForm01.Mode
                
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                If Updatedata(pval) = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                Call FormReset
                oForm01.Mode = fm_ADD_MODE
                Call LoadCaption
                Call LoadData
                
                
'                oForm01.Items("GCode").Click ct_Regular
            End If
            
        ElseIf pval.ItemUID = "Btn_ret" Then '/조회
                        
            Call LoadData
        ElseIf pval.ItemUID = "Btn_del" Then '/삭제
'            If DeleteBefor = False Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            
            Call DeleteData
            Call LoadData
            oForm01.Mode = fm_OK_MODE '/fm_VIEW_MODE
        End If
  
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_QM240" Then
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
        If pval.ItemUID = "Mat01" Then
            If pval.Row > 0 Then
                Call oMat01.SelectRow(pval.Row, True, False)
                
                oForm01.Items("DocEntry").Specific.VALUE = oMat01.Columns("DocEntry").Cells(pval.Row).Specific.VALUE
                Call oForm01.Items("BPLId").Specific.Select(oMat01.Columns("BPLId").Cells(pval.Row).Specific.VALUE, psk_ByValue)

                oForm01.Items("DocDate").Specific.VALUE = oMat01.Columns("DocDate").Cells(pval.Row).Specific.VALUE
                Call oForm01.Items("GCode").Specific.Select(oMat01.Columns("GCode").Cells(pval.Row).Specific.VALUE, psk_ByValue)
                Call oForm01.Items("SGCode").Specific.Select(oMat01.Columns("GCode").Cells(pval.Row).Specific.VALUE, psk_ByValue)
                oForm01.Items("Code").Specific.VALUE = oMat01.Columns("Code").Cells(pval.Row).Specific.VALUE
                oForm01.Items("CdName").Specific.VALUE = oMat01.Columns("CdName").Cells(pval.Row).Specific.VALUE
                oForm01.Items("Size").Specific.VALUE = oMat01.Columns("Size").Cells(pval.Row).Specific.VALUE
                oForm01.Items("Maker").Specific.VALUE = oMat01.Columns("Maker").Cells(pval.Row).Specific.VALUE
                oForm01.Items("ProCd").Specific.VALUE = oMat01.Columns("ProCd").Cells(pval.Row).Specific.VALUE
                Call oForm01.Items("Sect").Specific.Select(oMat01.Columns("Sect").Cells(pval.Row).Specific.VALUE, psk_ByValue)
                oForm01.Items("Cycle").Specific.VALUE = oMat01.Columns("Cycle").Cells(pval.Row).Specific.VALUE
                oForm01.Items("Remark").Specific.VALUE = oMat01.Columns("Remark").Cells(pval.Row).Specific.VALUE
                oForm01.Mode = fm_UPDATE_MODE
                Call LoadCaption
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
Dim GCode As String
Dim BPLId As String

Dim oRecordSet01      As SAPbobsCOM.Recordset

Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "GCode" Or pval.ItemUID = "BPLId" Then
            BPLId = oForm01.Items("BPLId").Specific.VALUE
            GCode = oForm01.Items("GCode").Specific.VALUE 'Selected.
            sQry = "Select right('00' + rtrim(Convert(Char(3),Convert(Integer,right(Max(U_Code),3)) + 1)),3) From [@PS_QM240H] Where U_BPLId = '" & BPLId & "' And U_GCode = '" & GCode & "'"
            oRecordSet01.DoQuery sQry
            
            If Trim(BPLId) <> "" And Trim(GCode) <> "%" Then
                If Trim(oRecordSet01.Fields(0).VALUE) = "" Then
                    oForm01.Items("Code").Specific.String = BPLId + GCode + "001"
                Else
                    oForm01.Items("Code").Specific.String = BPLId + GCode + Trim(oRecordSet01.Fields(0).VALUE)
                End If
            End If
        End If
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
'                    Call oDS_PS_QM240L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Value)
'                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_QM240L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
'                        PS_QM240_AddMatrixRow (pval.Row)
'                    End If
'                Else
'                    Call oDS_PS_QM240L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Value)
'                End If
'            Else
'                If (pval.ItemUID = "DocEntry") Then
'                    Call oDS_PS_QM240H.setValue(pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Value)
'                ElseIf (pval.ItemUID = "CardCode") Then
'                    Call oDS_PS_QM240H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Value)
'                    Call oDS_PS_QM240H.setValue("U_CardName", 0, MDC_GetData.Get_ReData("CardName", "CardCode", "[OCRD]", "'" & oForm01.Items(pval.ItemUID).Specific.Value & "'"))
'                Else
'                    Call oDS_PS_QM240H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Value)
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
        Call PS_QM240_FormItemEnabled
        '//Call PS_QM240_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_QM240_FormResize
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
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_QM240H", "U_CardCode,U_CardName")
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
'            If (PS_QM240_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_QM240H.RemoveRecord(oDS_PS_QM240H.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call Add_MatrixRow(0)
            Else
                If Trim(oDS_PS_QM240H.GetValue("U_CntcCode", oMat01.RowCount - 1)) <> "" Then
                    Call Add_MatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_QM240_CreateItems() As Boolean
On Error GoTo PS_QM240_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
         
    
    
    Set oDS_PS_QM240H = oForm01.DataSources.DBDataSources("@PS_QM240H")
    Set oDS_PS_QM240L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_QM240_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM240_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function
'/콤보박스 set
Sub PS_QM240_ComboBox_Setting()
On Error GoTo PS_QM240_ComboBox_Setting_Error

    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim sQry            As String
    
    Dim oRecordSet01      As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
        
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    Call MDC_SetMod.Set_ComboList(oForm01.Items("SBPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    
    
    Set oCombo = oForm01.Items("BPLId").Specific
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
    Call MDC_SetMod.Set_ComboList(oForm01.Items("GCode").Specific, "Select U_Minor = '%', U_CdName = '선택' Union All SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a Inner Join [@PS_SY001L] b On a.Code = b.Code And a.Code = 'Q007' order by U_Minor", "", False, False)
    Call MDC_SetMod.Set_ComboList(oForm01.Items("SGCode").Specific, "Select U_Minor = '%', U_CdName = '전체' Union All SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a Inner Join [@PS_SY001L] b On a.Code = b.Code And a.Code = 'Q007' order by U_Minor", "", False, False)
    
    Call oForm01.Items("GCode").Specific.Select("%", psk_ByValue)
    Call oForm01.Items("SGCode").Specific.Select("%", psk_ByValue)
    
    Call MDC_SetMod.Set_ComboList(oForm01.Items("Sect").Specific, "Select U_Minor = '%', U_CdName = '선택' Union All SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a Inner Join [@PS_SY001L] b On a.Code = b.Code And a.Code = 'Q010' order by U_Minor", "", False, False)
    Call oForm01.Items("Sect").Specific.Select("%", psk_ByValue)
    
    sQry = "SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a Inner Join [@PS_SY001L] b On a.Code = b.Code And a.Code = 'Q007' order by b.U_Minor"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oMat01.Columns("GCode").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    sQry = "SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a Inner Join [@PS_SY001L] b On a.Code = b.Code And a.Code = 'Q010' order by b.U_Minor"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oMat01.Columns("Sect").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    Call oForm01.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
PS_QM240_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM240_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM240_CF_ChooseFromList()
On Error GoTo PS_QM240_CF_ChooseFromList_Error

    Exit Sub
PS_QM240_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM240_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM240_FormItemEnabled()
On Error GoTo PS_QM240_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
        '//Call PS_QM240_FormClear '//UDO방식
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
PS_QM240_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM240_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM240_EnableMenus()
On Error GoTo PS_QM240_EnableMenus_Error
    
    Exit Sub
PS_QM240_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM240_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM240_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_QM240_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_QM240_FormItemEnabled
        '//Call PS_QM240_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_QM240_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_QM240_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM240_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM240_FormResize()
On Error GoTo PS_QM240_FormResize_Error
    
    Exit Sub
PS_QM240_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM240_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub FormReset()
On Error GoTo FormReset_Error

Dim sQry As String
Dim RecordSet01 As SAPbobsCOM.Recordset
Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
   

Call oForm01.Items("GCode").Specific.Select("%", psk_ByValue)
Call oForm01.Items("SGCode").Specific.Select("%", psk_ByValue)
oForm01.Items("Code").Specific.VALUE = ""
oForm01.Items("CdName").Specific.VALUE = ""
oForm01.Items("Size").Specific.VALUE = ""
oForm01.Items("Maker").Specific.VALUE = ""
oForm01.Items("ProCd").Specific.VALUE = ""
Call oForm01.Items("Sect").Specific.Select("%", psk_ByValue)
oForm01.Items("Cycle").Specific.VALUE = 0
oForm01.Items("Remark").Specific.VALUE = ""

sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_QM240H]"
RecordSet01.DoQuery sQry
If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
    oForm01.Items("DocEntry").Specific.VALUE = 1
Else
    oForm01.Items("DocEntry").Specific.VALUE = Trim(RecordSet01.Fields(0).VALUE) + 1
End If

oForm01.Mode = fm_ADD_MODE

Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormReset_Error:
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "FormReset_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub


