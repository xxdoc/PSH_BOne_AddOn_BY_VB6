VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP568"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File               : PS_PP568.cls
'//  Module             : 생산 관리 > 애드온 > 장비이익률관리(기계)
'//  Desc               : 장비별 공정진행률 등록
'//  FormType           : PS_PP568
'//  Create Date(Start) : 2017.06.16
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

Public oDS_PS_PP568A    As SAPbouiCOM.DBDataSource '등록용 Matrix

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
    
    oXmlDoc.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP568.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)

    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_PP568_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_PP568_CreateItems
    Call PS_PP568_ComboBox_Setting
    Call PS_PP568_Initial_Setting
'    Call PS_PP568_CF_ChooseFromList
'    Call PS_PP568_EnableMenus
'    Call PS_PP568_SetDocument(oFromDocEntry01)
'    Call PS_PP568_FormResize

    Call PS_PP568_LoadCaption
    Call PS_PP568_FormItemEnabled
    
    oForm01.EnableMenu ("1283"), False  '// 삭제
    oForm01.EnableMenu ("1286"), False  '// 닫기
    oForm01.EnableMenu ("1287"), False  '// 복제
    oForm01.EnableMenu ("1285"), False  '// 복원
    oForm01.EnableMenu ("1284"), False  '// 취소
    oForm01.EnableMenu ("1293"), False  '// 행삭제
    oForm01.EnableMenu ("1281"), False
    oForm01.EnableMenu ("1282"), True

'    Call PS_PP568_FormReset '폼초기화
    
    oForm01.Update
    oForm01.Freeze False
    
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

Private Sub PS_PP568_LoadCaption()
'******************************************************************************
'Function ID : PS_PP568_LoadCaption()
'해당모듈    : PS_PP568
'기능        : Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP568_LoadCaption_Error

    Call oForm01.Freeze(True)

    If oForm01.Mode = fm_ADD_MODE Then
        oForm01.Items("BtnAdd").Specific.Caption = "저장"
'        oForm01.Items("BtnDelete").Enabled = False
'    ElseIf oForm01.Mode = fm_OK_MODE Then
'        oForm01.Items("BtnAdd").Specific.Caption = "확인"
    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
        oForm01.Items("BtnAdd").Specific.Caption = "수정"
'        oForm01.Items("BtnDelete").Enabled = True
    End If
    
    Call oForm01.Freeze(False)

    Exit Sub

PS_PP568_LoadCaption_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/메트릭스 Row추가
Sub PS_PP568_Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_PP568_Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_PP568A.InsertRecord (oRow)
    End If
    
    oMat01.AddRow
    oDS_PS_PP568A.offset = oRow
    Call oDS_PS_PP568A.setValue("U_LineNum", oRow, oRow + 1)
    Call oDS_PS_PP568A.setValue("U_ColReg01", oRow, "Y")
    
    oMat01.LoadFromDataSource
    
    Exit Sub

PS_PP568_Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "PS_PP568_Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub PS_PP568_MTX01()
'******************************************************************************
'Function ID : PS_PP568_MTX01()
'해당모듈    : PS_PP568
'기능        : 등록된 데이터 조회
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP568_MTX01_Error

    Dim i       As Integer
    Dim Query01 As String
    Dim ErrNum  As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim FrYM     As String '기준년월(시작)
    Dim ToYM     As String '기준년월(종료)
    Dim CardType As String '거래처구분
    Dim CardCode As String '수주처
    Dim ItemCode As String '작번
    Dim InOut    As String '자체/외주
    Dim OrdAmt   As Double '수주금액
    Dim CmltYN   As String '생산완료
    
    FrYM = Trim(oForm01.Items("FrYM").Specific.VALUE) '기준년월(시작)
    ToYM = Trim(oForm01.Items("ToYM").Specific.VALUE) '기준년월(종료)
    CardType = Trim(oForm01.Items("CardType").Specific.Selected.VALUE) '거래처구분
    CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE) '수주처
    ItemCode = Trim(oForm01.Items("ItemCode").Specific.VALUE) '작번
    InOut = Trim(oForm01.Items("InOut").Specific.Selected.VALUE) '자체/외주
    OrdAmt = Trim(oForm01.Items("OrdAmt").Specific.VALUE) '수주금액
    If oForm01.Items("CmltYN").Specific.Checked = True Then
    
        CmltYN = "Y"
    
    Else
    
        CmltYN = "N"
    
    End If
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    Query01 = "         EXEC [PS_PP568_01] '"
    Query01 = Query01 & FrYM & "','"
    Query01 = Query01 & ToYM & "','"
    Query01 = Query01 & CardType & "','"
    Query01 = Query01 & CardCode & "','"
    Query01 = Query01 & ItemCode & "','"
    Query01 = Query01 & InOut & "','"
    Query01 = Query01 & OrdAmt & "','"
    Query01 = Query01 & CmltYN & "'"
    
    Call oRecordSet01.DoQuery(Query01)
    
    oMat01.Clear
    oDS_PS_PP568A.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource

    If (oRecordSet01.RecordCount = 0) Then
    
        ErrNum = 1
        
        oForm01.Mode = fm_ADD_MODE
        
        Call PS_PP568_Add_MatrixRow(0, True)
        Call PS_PP568_LoadCaption
        
        GoTo PS_PP568_MTX01_Error
        
        Exit Sub
    End If
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_PP568A.Size Then
            oDS_PS_PP568A.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_PP568A.offset = i
        
        Call oDS_PS_PP568A.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_PP568A.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("Select").VALUE)) '선택
        Call oDS_PS_PP568A.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("ItemCode").VALUE)) '작번
        Call oDS_PS_PP568A.setValue("U_ColReg03", i, Trim(oRecordSet01.Fields("ItemName").VALUE)) '작명
        Call oDS_PS_PP568A.setValue("U_ColReg04", i, Trim(oRecordSet01.Fields("ItemSpec").VALUE)) '규격
        Call oDS_PS_PP568A.setValue("U_ColReg05", i, Trim(oRecordSet01.Fields("PreRegDt").VALUE)) '직전등록일
        Call oDS_PS_PP568A.setValue("U_ColQty01", i, Trim(oRecordSet01.Fields("PreRate").VALUE)) '직전공정률
        Call oDS_PS_PP568A.setValue("U_ColQty02", i, Trim(oRecordSet01.Fields("CurRate").VALUE)) '현재공정률
        
        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
        
    Next
    
'    Call PS_PP568_Add_MatrixRow(oMat01.VisualRowCount)
    
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    ProgBar01.Stop
    
    Call oForm01.Freeze(False)
  
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub

PS_PP568_MTX01_Error:
'    ProgBar01.Stop
    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_PP568_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Sub PS_PP568_DeleteData()
'******************************************************************************
'Function ID : PS_PP568_DeleteData()
'해당모듈    : PS_PP568
'기능        : 기본정보 삭제
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP568_DeleteData_Error

    Dim loopCount As Integer
    Dim sQry      As String
    Dim ErrNum    As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ItemCode  As String
    Dim RegDt     As String
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("삭제 중...", 100, False)
    
    Call oMat01.FlushToDataSource
    
    For loopCount = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PS_PP568A.GetValue("U_ColReg01", loopCount)) = "Y" Then
            
            ItemCode = Trim(oDS_PS_PP568A.GetValue("U_ColReg02", loopCount))
            RegDt = Trim(oDS_PS_PP568A.GetValue("U_ColReg05", loopCount))
            
            sQry = "      EXEC [PS_PP568_03] '"
            sQry = sQry & ItemCode & "','" '작번
            sQry = sQry & RegDt & "'" '등록일
            
            Call oRecordSet01.DoQuery(sQry)
            
        End If
    Next
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Call MDC_Com.MDC_GF_Message("삭제 완료!", "S")
    
    Exit Sub

PS_PP568_DeleteData_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("삭제대상이 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_PP568_DeleteData_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Function PS_PP568_AddData() As Boolean
'******************************************************************************
'Function ID : PS_PP568_AddData()
'해당모듈    : PS_PP568
'기능        : 데이터 INSERT, UPDATE(기존 데이터가 존재하면 UPDATE, 아니면 INSERT)
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP568_AddData_Error

    Dim i    As Integer
    Dim sQry As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ItemCode  As String '작번
    Dim RegDt     As String '등록일
    Dim CpRate    As Double '공정진도율
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("저장 중...", oMat01.VisualRowCount - 1, False)
    
    RegDt = oForm01.DataSources.UserDataSources("RegDt").VALUE
        
    Call oMat01.FlushToDataSource
    For i = 0 To oMat01.VisualRowCount - 1 '마지막 빈행 제외를 위해 2를 뺌
        If Trim(oDS_PS_PP568A.GetValue("U_ColReg01", i)) = "Y" Then
            
            ItemCode = Trim(oDS_PS_PP568A.GetValue("U_ColReg02", i)) '작번
            CpRate = Trim(oDS_PS_PP568A.GetValue("U_ColQty02", i)) '현재공정진도율
            
            sQry = "      EXEC [PS_PP568_02] '"
            sQry = sQry & ItemCode & "','" '작번
            sQry = sQry & RegDt & "','" '등록일
            sQry = sQry & CpRate & "'" '현재공정진도율
            
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
    
    PS_PP568_AddData = True

    Exit Function

PS_PP568_AddData_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    PS_PP568_AddData = False
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PS_PP568_AddData_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Private Function PS_PP568_HeaderSpaceLineDel() As Boolean
'******************************************************************************
'Function ID : PS_PP568_HeaderSpaceLineDel()
'해당모듈    : PS_PP568
'기능        : 필수입력사항 체크
'인수        : 없음
'반환값      : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP568_HeaderSpaceLineDel_Error

    Dim ErrNum As Integer
    ErrNum = 0

    Select Case True
'        Case Trim(oForm01.Items("DestNo1").Specific.VALUE) = "" '출장번호1
'            ErrNum = 1
'            GoTo PS_PP568_HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("DestNo2").Specific.VALUE) = "" '출장번호2
'            ErrNum = 2
'            GoTo PS_PP568_HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("MSTCOD").Specific.VALUE) = "" '사원번호
'            ErrNum = 3
'            GoTo PS_PP568_HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("FrDate").Specific.VALUE) = "" '시작일자
'            ErrNum = 4
'            GoTo PS_PP568_HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("FrTime").Specific.VALUE) = "" '시작시각
'            ErrNum = 5
'            GoTo PS_PP568_HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("ToDate").Specific.VALUE) = "" '종료일자
'            ErrNum = 6
'            GoTo PS_PP568_HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("ToTime").Specific.VALUE) = "" '종료시각
'            ErrNum = 7
'            GoTo PS_PP568_HeaderSpaceLineDel_Error
    End Select

    PS_PP568_HeaderSpaceLineDel = True
    
    Exit Function

PS_PP568_HeaderSpaceLineDel_Error:
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
    PS_PP568_HeaderSpaceLineDel = False
End Function

'/ 메트릭스 필수 사항 check
Private Function PS_PP568_MatrixSpaceLineDel() As Boolean
On Error GoTo PS_PP568_MatrixSpaceLineDel_Error

    Dim i As Long
    Dim ErrNum As Integer
    ErrNum = 0

    Call oMat01.FlushToDataSource
    For i = 0 To oMat01.VisualRowCount - 2 '마직막 빈행 제외를 위해 2를 뺌
        If Trim(oDS_PS_PP568A.GetValue("U_ColReg01", i)) = "Y" Then
            
'            If Trim(oDS_PS_PP568A.GetValue("U_ColReg02", i)) = "" Then  '작번
'                ErrNum = 1
'                GoTo PS_PP568_MatrixSpaceLineDel_Error
'            ElseIf Trim(oDS_PS_PP568A.GetValue("U_ColSum01", i)) = "" Then '금액
'                ErrNum = 2
'                GoTo PS_PP568_MatrixSpaceLineDel_Error
'            ElseIf Trim(oDS_PS_PP568A.GetValue("U_ColReg05", i)) = "" Then '기준년월
'                ErrNum = 3
'                GoTo PS_PP568_MatrixSpaceLineDel_Error
'            End If

        End If
    Next

    PS_PP568_MatrixSpaceLineDel = True
    
    Exit Function

PS_PP568_MatrixSpaceLineDel_Error:

    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 작번이 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 금액이 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 기준년월이 없습니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "PS_PP568_MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    PS_PP568_MatrixSpaceLineDel = False
    
End Function

Private Sub PS_PP568_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PS_PP568_FlushToItemValue_Error

    Dim loopCount As Integer
    Dim ErrNum    As Integer
    Dim sQry      As String
    
    Dim BPLID     As String
    
    Dim oCombo As SAPbouiCOM.ComboBox
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID

'        Case "Mat01"
'
'            Call oForm01.Freeze(True)
'
'            Call oMat01.FlushToDataSource
'
'            If oCol = "Location" Then
'
'                If oMat01.Columns(oCol).Cells(oRow).Specific.VALUE = "01" Then '자택
'
'                    '자택 주소 조회
'                    sQry = "        SELECT      T0.U_Address AS [PosAddress]"
'                    sQry = sQry & " FROM        [@PH_PY001A] AS T0"
'                    sQry = sQry & " WHERE       T0.U_CLTCOD = '" & oForm01.Items("BPLID").Specific.VALUE & "'"
'                    sQry = sQry & "             AND T0.Code = '" & oMat01.Columns("CntcCode").Cells(oRow).Specific.VALUE & "'"
'
'                Else '회사
'
'                    '회사 주소 조회
'                    sQry = "        SELECT      T0.U_PosAdd AS [PosAddress]"
'                    sQry = sQry & " FROM        [@PH_PY005A] AS T0"
'                    sQry = sQry & " WHERE       T0.Code = '" & oForm01.Items("BPLID").Specific.VALUE & "'"
'
'                End If
'
'                Call oRecordSet01.DoQuery(sQry)
'
'                Call oDS_PS_PP568A.setValue("U_ColReg12", oRow - 1, oRecordSet01.Fields("PosAddress").VALUE)  '주소
'                Call oMat01.LoadFromDataSource
'
'                Call oMat01.Columns("Hospital").Cells(oRow).Click(ct_Regular)
'
'            Else
'
'                '강제 포커스 이동_S
'                Call oMat01.Columns(oCol).Cells(oRow).Click(ct_Regular)
'                '강제 포커스 이동_E
'
'            End If
'
''            Call oMat01.LoadFromDataSource
'            Call oMat01.AutoResizeColumns
'
'            Call oForm01.Freeze(False)

'        Case "OrdNum01"
'
'            oForm01.Items("ItemName01").Specific.VALUE = MDC_GetData.Get_ReData("FrgnName", "ItemCode", "OITM", "'" & oForm01.Items("OrdNum01").Specific.VALUE & "'") '품명
'            oForm01.Items("ItemSpec01").Specific.VALUE = MDC_GetData.Get_ReData("U_Size", "ItemCode", "OITM", "'" & oForm01.Items("OrdNum01").Specific.VALUE & "'")  '규격

'        Case "BPLID"
'
'            BPLID = Trim(oForm01.Items("BPLID").Specific.VALUE)
'
'            If oForm01.Items("TeamCode").Specific.ValidValues.Count > 0 Then
'                For loopCount = oForm01.Items("TeamCode").Specific.ValidValues.Count - 1 To 0 Step -1
'                    Call oForm01.Items("TeamCode").Specific.ValidValues.Remove(loopCount, psk_Index)
'                Next
'            End If
'
'            '부서콤보세팅
'            Call oForm01.Items("TeamCode").Specific.ValidValues.Add("%", "전체")
'            sQry = "        SELECT      U_Code AS [Code],"
'            sQry = sQry & "             U_CodeNm As [Name]"
'            sQry = sQry & " FROM        [@PS_HR200L]"
'            sQry = sQry & " WHERE       Code = '1'"
'            sQry = sQry & "             AND U_UseYN = 'Y'"
'            sQry = sQry & "             AND U_Char2 = '" & BPLID & "'"
'            sQry = sQry & " ORDER BY    U_Seq"
'            Call MDC_SetMod.Set_ComboList(oForm01.Items("TeamCode").Specific, sQry, "", False, False)
'            Call oForm01.Items("TeamCode").Specific.Select(0, psk_Index)
            
'        Case "CntcCode"
'
'            oForm01.DataSources.UserDataSources("CntcName").VALUE = MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm01.Items("CntcCode").Specific.VALUE & "'") '성명

        Case "CardCode"
        
            oForm01.Items("CardName").Specific.VALUE = MDC_GetData.Get_ReData("CardName", "CardCode", "[OCRD]", "'" & Trim(oForm01.Items("CardCode").Specific.VALUE) & "'") '거래처
        
        Case "ItemCode"
        
            oForm01.Items("ItemName").Specific.VALUE = MDC_GetData.Get_ReData("FrgnName", "ItemCode", "[OITM]", "'" & Trim(oForm01.Items("ItemCode").Specific.VALUE) & "'") '작번
            oForm01.Items("ItemSpec").Specific.VALUE = MDC_GetData.Get_ReData("U_Size", "ItemCode", "[OITM]", "'" & Trim(oForm01.Items("ItemCode").Specific.VALUE) & "'") '규격
            
    End Select

    Set oCombo = Nothing
    Set oRecordSet01 = Nothing

    Exit Sub
    
PS_PP568_FlushToItemValue_Error:

    Set oCombo = Nothing
    Set oRecordSet01 = Nothing

    Call MDC_Com.MDC_GF_Message("PS_PP568_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")

End Sub

'/폼의 아이템 사용지정
Sub PS_PP568_FormItemEnabled()
On Error GoTo PS_PP568_FormItemEnabled_Error

    If (oForm01.Mode = fm_ADD_MODE) Then
       
'        oMat01.Columns("ItemCode").Cells(1).Click ct_Regular
'        oForm01.Items("ItemCode").Enabled = True

    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
    
'        oForm01.Items("ItemCode").Enabled = True

    ElseIf (oForm01.Mode = fm_OK_MODE) Then
       
    End If
    
    Exit Sub

PS_PP568_FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "PS_PP568_FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
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

                Call PS_PP568_FormReset
'                Call PS_PP568_LoadCaption
                
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
            
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
                Call oForm01.Freeze(True)
                Call PS_PP568_Add_MatrixRow(oMat01.VisualRowCount)
                Call oForm01.Freeze(False)
            
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
                '//Call PS_PP568_FormItemEnabled '//UDO방식
            Case "1282": '추가
'                oMat01.Clear
'                oDS_PS_PP568A.Clear

'                Call PS_PP568_LoadCaption
'                Call PS_PP568_FormItemEnabled
                '//Call PS_PP568_FormItemEnabled '//UDO방식
                '//Call PS_PP568_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PS_PP568_FormItemEnabled
                
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 이후 처리
                Call oForm01.Freeze(True)
                Call oDS_PS_PP568A.RemoveRecord(oDS_PS_PP568A.Size - 1)
                Call oMat01.LoadFromDataSource
                Call oForm01.Freeze(False)
            
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
    
        If pval.ItemUID = "PS_PP568" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "BtnAdd" Then '/추가/확인 버튼클릭
        
            If oForm01.Mode = fm_ADD_MODE Then
            
'                If PS_PP568_HeaderSpaceLineDel() = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If

'                If PS_PP568_DataCheck() = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If

                '매트릭스 필수자료 체크
                If PS_PP568_MatrixSpaceLineDel = False Then
                    BubbleEvent = False
                    Exit Sub
                End If

                If PS_PP568_AddData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If

'                Call PS_PP568_FormReset
                oForm01.Mode = fm_ADD_MODE
                
                Call PS_PP568_LoadCaption
                Call PS_PP568_MTX01
                
                oLast_Mode = oForm01.Mode
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            
                If PS_PP568_HeaderSpaceLineDel = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            
'                If PS_PP568_DataCheck() = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
                
                Call PS_PP568_FormReset
                oForm01.Mode = fm_ADD_MODE
                
                Call PS_PP568_LoadCaption
                Call PS_PP568_MTX01
                
'                oForm01.Items("GCode").Click ct_Regular
            End If
            
        ElseIf pval.ItemUID = "BtnSearch" Then '/조회
            
'            Call PS_PP568_FormReset
'            oForm01.Mode = fm_ADD_MODE '/fm_VIEW_MODE
'
'            Call PS_PP568_LoadCaption
            Call PS_PP568_MTX01

        ElseIf pval.ItemUID = "BtnDelete" Then '/삭제
        
            If Sbo_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", "1", "예", "아니오") = "1" Then
        
                Call PS_PP568_DeleteData
'                Call PS_PP568_FormReset
                oForm01.Mode = fm_ADD_MODE '/fm_VIEW_MODE
                
                Call PS_PP568_LoadCaption
                Call PS_PP568_MTX01
                
            Else
            
            End If
        
        ElseIf pval.ItemUID = "BtnSelect" Then '전체선택
        
            Call PS_PP568_CheckAll
        
        ElseIf pval.ItemUID = "BtnPrt01" Then
        
            Call PS_PP568_Print_Report01
        
        End If
  
    ElseIf pval.BeforeAction = False Then
    
        If pval.ItemUID = "PS_PP568" Then
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
        
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CardCode", "") '거래처
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '작번
        
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
        
            Call PS_PP568_FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
        
        Else
        
            Call PS_PP568_FlushToItemValue(pval.ItemUID)
        
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
    
    ElseIf pval.BeforeAction = False Then

        If pval.ItemChanged = True Then
        
            If (pval.ItemUID = "Mat01") Then
            
                Call PS_PP568_FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
                
            Else
                
                Call PS_PP568_FlushToItemValue(pval.ItemUID)
                
            End If

        End If
        
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
        Call PS_PP568_FormItemEnabled
        '//Call PS_PP568_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    
    Exit Sub
    
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error

    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_PP568_FormResize
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
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_PP568A", "U_CardCode,U_CardName")
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
'            If (PS_PP568_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_PP568A.RemoveRecord(oDS_PS_PP568A.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PS_PP568_Add_MatrixRow(0)
            Else
                If Trim(oDS_PS_PP568A.GetValue("U_ColReg01", oMat01.RowCount - 1)) <> "" Then
                    Call PS_PP568_Add_MatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    
    Exit Sub
    
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_PP568_CreateItems() As Boolean
On Error GoTo PS_PP568_CreateItems_Error

    Call oForm01.Freeze(True)
    
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Set oDS_PS_PP568A = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '수주년월(시작)
    Call oForm01.DataSources.UserDataSources.Add("FrYM", dt_SHORT_TEXT, 6)
    Call oForm01.Items("FrYM").Specific.DataBind.SetBound(True, "", "FrYM")
    
    '수주년월(종료)
    Call oForm01.DataSources.UserDataSources.Add("ToYM", dt_SHORT_TEXT, 6)
    Call oForm01.Items("ToYM").Specific.DataBind.SetBound(True, "", "ToYM")

    '거래처구분
    Call oForm01.DataSources.UserDataSources.Add("CardType", dt_SHORT_TEXT, 5)
    Call oForm01.Items("CardType").Specific.DataBind.SetBound(True, "", "CardType")

    '수주처
    Call oForm01.DataSources.UserDataSources.Add("CardCode", dt_SHORT_TEXT, 20)
    Call oForm01.Items("CardCode").Specific.DataBind.SetBound(True, "", "CardCode")

    '수주처명
    Call oForm01.DataSources.UserDataSources.Add("CardName", dt_SHORT_TEXT, 200)
    Call oForm01.Items("CardName").Specific.DataBind.SetBound(True, "", "CardName")

    '작번
    Call oForm01.DataSources.UserDataSources.Add("ItemCode", dt_SHORT_TEXT, 20)
    Call oForm01.Items("ItemCode").Specific.DataBind.SetBound(True, "", "ItemCode")

    '품명
    Call oForm01.DataSources.UserDataSources.Add("ItemName", dt_SHORT_TEXT, 200)
    Call oForm01.Items("ItemName").Specific.DataBind.SetBound(True, "", "ItemName")

    '규격
    Call oForm01.DataSources.UserDataSources.Add("ItemSpec", dt_SHORT_TEXT, 200)
    Call oForm01.Items("ItemSpec").Specific.DataBind.SetBound(True, "", "ItemSpec")
    
    '자체/외주
    Call oForm01.DataSources.UserDataSources.Add("InOut", dt_SHORT_TEXT, 5)
    Call oForm01.Items("InOut").Specific.DataBind.SetBound(True, "", "InOut")
    
    '수주금액
    Call oForm01.DataSources.UserDataSources.Add("OrdAmt", dt_SUM)
    Call oForm01.Items("OrdAmt").Specific.DataBind.SetBound(True, "", "OrdAmt")
    
    '생산완료포함CheckBox
    Call oForm01.DataSources.UserDataSources.Add("CmltYN", dt_SHORT_TEXT, 5)
    Call oForm01.Items("CmltYN").Specific.DataBind.SetBound(True, "", "CmltYN")
    
    '등록일
    Call oForm01.DataSources.UserDataSources.Add("RegDt", dt_DATE)
    Call oForm01.Items("RegDt").Specific.DataBind.SetBound(True, "", "RegDt")
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    
    Exit Function
    
PS_PP568_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP568_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'/콤보박스 set
Sub PS_PP568_ComboBox_Setting()
On Error GoTo PS_PP568_ComboBox_Setting_Error

    Dim oCombo       As SAPbouiCOM.ComboBox
    Dim sQry         As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    '//콤보에 기본값설정
    
    '사업장
'    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLID").Specific, "SELECT BPLID, BPLName FROM OBPL order by BPLID", User_BPLId, False, False)
    
    '거래처구분
    sQry = "        SELECT      U_Minor,"
    sQry = sQry & "             U_CdName"
    sQry = sQry & " FROM        [@PS_SY001L]"
    sQry = sQry & " WHERE       Code = 'C100'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    U_Seq"
    
    Call oForm01.Items("CardType").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("CardType").Specific, sQry, "", False, False)
    Call oForm01.Items("CardType").Specific.Select(0, psk_Index)
    '거래처구분

    '자체/외주
    Call oForm01.Items("InOut").Specific.ValidValues.Add("%", "전체")
    Call oForm01.Items("InOut").Specific.ValidValues.Add("IN", "자체")
    Call oForm01.Items("InOut").Specific.ValidValues.Add("OUT", "외주")
    Call oForm01.Items("InOut").Specific.Select(0, psk_Index)

    Call oForm01.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
    
PS_PP568_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "PS_PP568_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP568_Initial_Setting()
On Error GoTo PS_PP568_Initial_Setting_Error

    oForm01.Items("FrYM").Specific.VALUE = Format(DateAdd("m", -6, Date), "YYYYMM")
    oForm01.Items("ToYM").Specific.VALUE = Format(Now, "YYYYMM")
    oForm01.Items("OrdAmt").Specific.VALUE = "100000000"
    oForm01.Items("RegDt").Specific.VALUE = Format(Now, "YYYYMMDD")
    
    Call oForm01.Items("CardCode").Click
    
    Exit Sub
PS_PP568_Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP568_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP568_CF_ChooseFromList()
On Error GoTo PS_PP568_CF_ChooseFromList_Error

    Exit Sub
    
PS_PP568_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP568_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP568_EnableMenus()
On Error GoTo PS_PP568_EnableMenus_Error
    
    Exit Sub
    
PS_PP568_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP568_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP568_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_PP568_SetDocument_Error

    If (oFromDocEntry01 = "") Then
        Call PS_PP568_FormItemEnabled
        '//Call PS_PP568_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_PP568_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    
    Exit Sub
    
PS_PP568_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP568_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP568_FormResize()
On Error GoTo PS_PP568_FormResize_Error
    
    Call oMat01.AutoResizeColumns
    
    Exit Sub
    
PS_PP568_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP568_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub PS_PP568_FormReset()
'******************************************************************************
'Function ID : PS_PP568_FormReset()
'해당모듈    : PS_PP568
'기능        : 화면 초기화
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP568_FormReset_Error

    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    
    Dim User_BPLId As String
    User_BPLId = MDC_PS_Common.User_BPLId

    '헤더 초기화
    oForm01.DataSources.UserDataSources("StdDate").VALUE = Format(Now, "YYYYMMDD") '기준일자
    
    '라인 초기화
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    Call PS_PP568_Add_MatrixRow(0, True)
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)

    Exit Sub

PS_PP568_FormReset_Error:
    Call oForm01.Freeze(False)
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PS_PP568_FormReset_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub PS_PP568_CheckAll()
On Error GoTo PS_PP568_CheckAll_Error

    Dim CheckType As String
    Dim loopCount As Integer
    
    Call oForm01.Freeze(True)
    CheckType = "Y"
    
    Call oMat01.FlushToDataSource
    
    For loopCount = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PS_PP568A.GetValue("U_ColReg01", loopCount)) = "N" Then
            CheckType = "N"
            Exit For
        End If
    Next
    
    For loopCount = 0 To oMat01.VisualRowCount - 1
        oDS_PS_PP568A.offset = loopCount
        If CheckType = "N" Then
            Call oDS_PS_PP568A.setValue("U_ColReg01", loopCount, "Y")
        Else
            Call oDS_PS_PP568A.setValue("U_ColReg01", loopCount, "N")
        End If
    Next
    
    Call oMat01.LoadFromDataSource
    
    Call oForm01.Freeze(False)

    Exit Sub
    
PS_PP568_CheckAll_Error:
    Call oForm01.Freeze(False)
    Call Sbo_Application.SetStatusBarMessage("PS_PP568_CheckAll_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)
End Sub

Private Sub PS_PP568_Print_Report01()
On Error GoTo PS_PP568_Print_Report01_Error
    
    Dim DocNum     As String
    Dim ErrNum     As Integer
    Dim WinTitle   As String
    Dim ReportName As String
    Dim sQry       As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLID    As String '사업장
    Dim TeamCode As String '부서
    Dim StdDate  As String '기준일자
    Dim CntcCode As String '사번
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
     '/ ODBC 연결 체크
    Call ConnectODBC
    
    BPLID = Trim(oForm01.Items("BPLID").Specific.VALUE) '사업장
    TeamCode = Trim(oForm01.Items("TeamCode").Specific.VALUE) '부서
    StdDate = Trim(oForm01.Items("StdDate").Specific.VALUE) '기준일자
    CntcCode = Trim(oForm01.Items("CntcCode").Specific.VALUE) '사번
    
    WinTitle = "[PS_PP568] 종합건강검진 희망병원"
    
    ReportName = "PS_PP568_01.rpt"
    
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
    sQry = "      EXEC [PS_PP568_90] '"
    sQry = sQry & BPLID & "','" '사업장
    sQry = sQry & TeamCode & "','" '부서
    sQry = sQry & StdDate & "','" '기준일자
    sQry = sQry & CntcCode & "'" '사번
    
'    oRecordSet.DoQuery sQry
'
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo PS_PP568_Print_Report01_Error
'    End If
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    
    Exit Sub
    
PS_PP568_Print_Report01_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet = Nothing

    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
        Sbo_Application.SetStatusBarMessage "PS_PP568_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
    
End Sub







