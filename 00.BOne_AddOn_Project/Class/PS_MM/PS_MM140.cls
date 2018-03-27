VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_MM140"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_MM140.cls
'//  Module         : MM
'//  Description    : 입고확인등록
'//  FormType       : PS_MM140
'//  Create Date    : 2010.10.10
'//  Modified Date  :
'//  Creator        : Ryu Yung Jo
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01             As SAPbouiCOM.Form
Public oMat01              As SAPbouiCOM.Matrix
Private oDS_PS_MM140H As SAPbouiCOM.DBDataSource    '등록헤더

Private oLast_Item_UID      As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID       As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row       As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

'DI를 위한 정보를 가지는 구조체
Private Type StockInfos
    ItemCode As String            '품목코드
    FromWarehouseCode As String   '창고코드
    ToWarehouseCode As String     '창고코드
    Weight As Currency            '중량
    Qty   As Long                 '수량
    OutDoc As String              '반출문서번호
    Chk As String
    MatrixRow As Long
    StockTransDocEntry  As String '재고이전문서번호
    StockTransLineNum   As String '재고이전라인번호
    DocDate  As String            '전기일
    LineNum As Long               '문서라인번호
End Type

Private StockInfo() As StockInfos

Private oLast_Mode&

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_MM140.srf")
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
    
    oFormUniqueID01 = "PS_MM140_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting
    Call Initialization
    Call FormItemEnabled
    
    oForm01.EnableMenu ("1281"), False        '// 찾기
    oForm01.EnableMenu ("1282"), False        '// 추가
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1285"), False        '// 복원
    oForm01.EnableMenu ("1284"), False        '// 취소
    oForm01.EnableMenu ("1293"), True         '// 행삭제
        
    oForm01.Update
    oForm01.Freeze False
    oForm01.Visible = True
    
    Set oXmlDoc01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadForm_Error:
    oForm01.Update
    oForm01.Freeze False
    Set oXmlDoc01 = Nothing
    If (oForm01 Is Nothing) = False Then
        Set oForm01 = Nothing
    End If
    MDC_Com.MDC_GF_Message "LoadForm_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'****************************************************************************************************************
'// ItemEventHander
'****************************************************************************************************************
Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_ItemEvent_Error
    Dim i&
    Dim ErrNum&
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    
    Dim ItemCode$, ItemName$, Size$, Qty&, Weight As Currency, Unit$, RequestDate$, DueDate$, ItemType$, RequestNo$
    Dim RFC_Sender$
    Dim Seq&
    Dim Check$, OutDoc$
    
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "Btn01" Then
                    If Trim(oForm01.Items("Status").Specific.Value) = "N" Then '미입고
                        If HeaderSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        If MatrixSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        If Add_oStockTransfer() = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        oMat01.Clear
                        oDS_PS_MM140H.Clear
                        oMat01.FlushToDataSource
                        oMat01.LoadFromDataSource
                    End If
                ElseIf pval.ItemUID = "Btn02" Then
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    Call LoadData
                    Call FormItemEnabled
                End If
'et_KEY_DOWN ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "CardCode" Then
                        If oForm01.Items("CardCode").Specific.Value = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "OutDoc" Then
                        If oForm01.Items("OutDoc").Specific.Value = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "OutWhCd" Then
                            If oMat01.Columns("OutWhCd").Cells(pval.Row).Specific.Value = "" Then
                                Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                            End If
                        ElseIf pval.ColUID = "InWhCd" Then
                            If oMat01.Columns("InWhCd").Cells(pval.Row).Specific.Value = "" Then
                                Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                            End If
                        End If
                    End If
                End If
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
            Case et_MATRIX_LOAD: '//11
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
            Case et_FORM_RESIZE '//20
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
            Case et_LOST_FOCUS: '//4
            Case et_FORM_UNLOAD: '//17
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
            Case et_KEY_DOWN: '//2
'et_COMBO_SELECT ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_COMBO_SELECT: '//5
                If pval.ItemUID = "BPLId" Or pval.ItemUID = "OUtGbn" Or pval.ItemUID = "Status" Then
                    oForm01.Freeze True
                    oMat01.Clear
                    oDS_PS_MM140H.Clear
                    oMat01.FlushToDataSource
                    oMat01.LoadFromDataSource
                    oForm01.Freeze False
                End If
'et_CLICK ///////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_CLICK: '//6
                If pval.ItemUID = "Mat01" And oLast_Col_Row > 0 And pval.ColUID = "Check" Then
                    oForm01.Freeze True
                    oMat01.FlushToDataSource
                    Check = Trim(oDS_PS_MM140H.GetValue("U_Check", pval.Row - 1))
                    OutDoc = Trim(oDS_PS_MM140H.GetValue("U_OutDoc", pval.Row - 1))
                    For i = 0 To oMat01.VisualRowCount - 1
                        If Trim(oDS_PS_MM140H.GetValue("U_OutDoc", i)) = OutDoc Then
                            oDS_PS_MM140H.setValue "U_Check", i, Check
                        End If
                    Next i
                    oMat01.LoadFromDataSource
                    oForm01.Freeze False
                End If
'et_DOUBLE_CLICK ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_DOUBLE_CLICK: '//7
                If pval.ItemUID = "Mat01" And pval.Row = 0 And pval.ColUID = "Check" Then
                    oForm01.Freeze True
                    If oMat01.Columns("Check").Cells(1).Specific.Checked = False Then
                        Check = "True"
                    ElseIf oMat01.Columns("Check").Cells(1).Specific.Checked = True Then
                        Check = "False"
                    End If
                    For i = 0 To oMat01.VisualRowCount - 1
                        oMat01.Columns("Check").Cells(i + 1).Specific.Checked = Check
                    Next i
                    oForm01.Freeze False
                End If
            Case et_MATRIX_LINK_PRESSED '//8
'et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "CardCode" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "OutWhCd" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        ElseIf pval.ColUID = "InWhCd" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        End If
                    End If
                End If
            Case et_MATRIX_LOAD: '//11
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
            Case et_FORM_RESIZE '//20
            Case et_CHOOSE_FROM_LIST '//27
'et_GOT_FOCUS ///////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_GOT_FOCUS: '//3
                oLast_Col_Row = pval.Row
            Case et_LOST_FOCUS: '//4
'et_FORM_UNLOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_FORM_UNLOAD: '//17
                RemoveForms oFormUniqueID01
                Set oForm01 = Nothing
                Set oMat01 = Nothing
                Set oDS_PS_MM140H = Nothing
        End Select
    End If
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
    Set ProgressBar01 = Nothing
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "Raise_ItemEvent_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_MenuEvent_Error
    Dim i&
    Dim ReturnValue&
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
'[1293:행삭제] //////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1293": '행삭제
                If oMat01.RowCount <> oMat01.VisualRowCount Then
                    oForm01.Freeze True
                    For i = 0 To oMat01.VisualRowCount - 1
                        oMat01.Columns("DocNum").Cells(i + 1).Specific.Value = i + 1
                    Next i

                    oMat01.FlushToDataSource
                    oDS_PS_MM140H.RemoveRecord oDS_PS_MM140H.Size - 1       '// Mat01에 마지막라인(빈라인) 삭제
                    oMat01.Clear
                    oMat01.LoadFromDataSource
                    oForm01.Freeze False
                End If
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    End If
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_MenuEvent_Error:
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "Raise_MenuEvent_Error:" & Err.Number & " - " & Err.Description, "E"
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
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_FormDataEvent_Error:
    MDC_Com.MDC_GF_Message "Raise_FormDataEvent_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub Raise_RightClickEvent(ByRef FormUID As String, ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_RightClickEvent_Error
    If (eventInfo.BeforeAction = True) Then
        If eventInfo.Row > 0 Then
        End If
    ElseIf (eventInfo.BeforeAction = False) Then
        '//작업
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub CreateItems()
On Error GoTo CreateItems_Error
    '//디비데이터 소스 개체 할당
    Set oDS_PS_MM140H = oForm01.DataSources.DBDataSources("@PS_MM140H")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    
    Call oForm01.DataSources.UserDataSources.Add("DocDateFr", dt_DATE, 8)
    oForm01.Items("DocDateFr").Specific.DataBind.SetBound True, "", "DocDateFr"
'    oForm01.DataSources.UserDataSources.Item("DocDateFr").Value = Format(Now, "YYYY-MM") & "-01"
    
    Call oForm01.DataSources.UserDataSources.Add("DocDateTo", dt_DATE, 8)
    oForm01.Items("DocDateTo").Specific.DataBind.SetBound True, "", "DocDateTo"
'    oForm01.DataSources.UserDataSources.Item("DocDateTo").Value = Date

    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
CreateItems_Error:
    MDC_Com.MDC_GF_Message "CreateItems_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub ComboBox_Setting()
On Error GoTo ComboBox_Setting_Error
    '//콤보에 기본값설정
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    '// 사업장
    Set oCombo = oForm01.Items("BPLId").Specific
    sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).Value), Trim(oRecordSet01.Fields(1).Value)
        oRecordSet01.MoveNext
    Loop
    
    
    '//작업구분
    Set oCombo = oForm01.Items("OutGbn").Specific
    oCombo.ValidValues.Add "10", "원재료"
    oCombo.ValidValues.Add "20", "재공"
    oCombo.Select 0, psk_Index
    
    oMat01.Columns("OutGbn").ValidValues.Add "10", "원재료"
    oMat01.Columns("OutGbn").ValidValues.Add "20", "재공"
    
    '//상태
    Set oCombo = oForm01.Items("Status").Specific
    oCombo.ValidValues.Add "N", "미입고"
    oCombo.ValidValues.Add "Y", "입고"
    oCombo.Select 0, psk_Index
    
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
ComboBox_Setting_Error:
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "ComboBox_Setting_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub Initialization()
On Error GoTo Initialization_Error
    Dim oCombo          As SAPbouiCOM.ComboBox
    
    '//아이디별 사업장 세팅
    Set oCombo = oForm01.Items("BPLId").Specific
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
    '//아이디별 사번 세팅
'    oForm01.Items("CntcCode").Specific.Value = MDC_PS_Common.User_MSTCOD
    
    '//아이디별 부서 세팅
'    Set oCombo = oForm01.Items("DeptCode").Specific
'    oCombo.Select MDC_PS_Common.User_DeptCode, psk_ByValue
    Set oCombo = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Initialization_Error:
    Set oCombo = Nothing
    MDC_Com.MDC_GF_Message "Initialization_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub CF_ChooseFromList()
On Error GoTo CF_ChooseFromList_Error
    '//ChooseFromList 설정
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
CF_ChooseFromList_Error:
    MDC_Com.MDC_GF_Message "CF_ChooseFromList_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub FormItemEnabled()
On Error GoTo FormItemEnabled_Error
    If Trim(oForm01.Items("Status").Specific.Selected.Value) = "Y" Then
        oForm01.Items("Mat01").Enabled = False
    ElseIf Trim(oForm01.Items("Status").Specific.Selected.Value) = "N" Then
        oForm01.Items("Mat01").Enabled = True
        oMat01.Columns("Check").Editable = True
'        oMat01.Columns("InQty").Editable = True
'        oMat01.Columns("InWt").Editable = True
        oMat01.Columns("DocDate").Editable = True
    End If
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_MM140H.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_MM140H.Offset = oRow
    oDS_PS_MM140H.setValue "DocNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

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
        Case "CardCode"
            sQry = "Select CardName From OCRD Where CardCode = '" & Trim(oForm01.Items("CardCode").Specific.Value) & "'"
            oRecordSet01.DoQuery sQry
            oForm01.Items("CardName").Specific.Value = Trim(oRecordSet01.Fields(0).Value)
        
    End Select

    Set oRecordSet01 = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
FlushToItemValue_Error:
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum          As Integer
    
    ErrNum = 0

    '// Check
    Select Case True
        Case Trim(oForm01.Items("BPLId").Specific.Value) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("CntcCode").Specific.Value) = ""
'            ErrNum = 3
'            GoTo HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("DeptCode").Specific.Value) = ""
'            ErrNum = 4
'            GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "사업장은 필수사항입니다. 확인하세요.", "E"
'    ElseIf ErrNum = 2 Then
'        MDC_Com.MDC_GF_Message "사업장은 필수사항입니다. 확인하세요.", "E"
'    ElseIf ErrNum = 3 Then
'        MDC_Com.MDC_GF_Message "청구인은 필수사항입니다. 확인하세요.", "E"
'    ElseIf ErrNum = 4 Then
'        MDC_Com.MDC_GF_Message "청구부서는 필수사항입니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function

Private Function MatrixSpaceLineDel() As Boolean
On Error GoTo MatrixSpaceLineDel_Error
    Dim i               As Long
    Dim ErrNum          As Integer
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    Dim sQry            As String

    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

'    ErrNum = 0
'
'    oMat01.FlushToDataSource
'
'    '// 라인
'    If oMat01.VisualRowCount = 0 Then
'        ErrNum = 1
'        GoTo MatrixSpaceLineDel_Error
'    End If
'
'    For i = 0 To oMat01.VisualRowCount - 2
'        Select Case True
'            Case oDS_PS_MM140H.GetValue("U_ItemCode", i) = ""
'                ErrNum = 2
'                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_MM140H.GetValue("U_Qty", i) = ""
'                ErrNum = 3
'                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_MM140H.GetValue("U_DocDate", i) = ""
'                ErrNum = 4
'                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_MM140H.GetValue("U_DueDate", i) = ""
'                ErrNum = 5
'                GoTo MatrixSpaceLineDel_Error
'        End Select
'    Next
'    oMat01.LoadFromDataSource

    Set oRecordSet01 = Nothing
    MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
MatrixSpaceLineDel_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "라인 데이터가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 품목코드가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 수량이 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 청구일자가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 납기일자가 없습니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function

'Private Sub LoadCaption()
'On Error GoTo LoadCaption_Error
'    If oForm01.Mode = fm_ADD_MODE Then
'        oForm01.Items("Btn01").Specific.Caption = "추가"
'    ElseIf oForm01.Mode = fm_OK_MODE Then
'        oForm01.Items("Btn01").Specific.Caption = "확인"
'    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'        oForm01.Items("Btn01").Specific.Caption = "갱신"
'    End If
'
'    Exit Sub
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'LoadCaption_Error:
'    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
'End Sub

'Public Function Add_MM140H(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
'On Error GoTo Add_MM140H_Error
'    Dim i As Integer
'    Dim sQry As String
'    Dim RecordSet01 As SAPbobsCOM.Recordset
'    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
'
'    Dim DocEntry$, BPLId$, OutDoc$, OutLine$, ItemCode$, ItemName$, Size$, Mark$, OutItmCd$, OutItmNm$, OutGbn$
'    Dim OutQty As Currency, OutWt As Currency, InQty As Currency, InWt As Currency, JakNum$, DocDate$
'
'    oMat01.FlushToDataSource
'
'    BPLId = Trim(oForm01.Items("BPLId").Specific.Value)
'    For i = 0 To oMat01.RowCount - 2
'        OutDoc = Trim(oDS_PS_MM140H.GetValue("U_OutDoc", i))
'        OutLine = Trim(oDS_PS_MM140H.GetValue("U_OutLine", i))
'        ItemCode = Trim(oDS_PS_MM140H.GetValue("U_ItemCode", i))
'        ItemName = Trim(oDS_PS_MM140H.GetValue("U_ItemName", i))
'        Size = Trim(oDS_PS_MM140H.GetValue("U_Size", i))
'        Mark = Trim(oDS_PS_MM140H.GetValue("U_Mark", i))
'        OutItmCd = Trim(oDS_PS_MM140H.GetValue("U_OutItmCd", i))
'        OutItmNm = Trim(oDS_PS_MM140H.GetValue("U_OutItmNm", i))
'        OutGbn = Trim(oDS_PS_MM140H.GetValue("U_OutGbn", i))
'        OutQty = Trim(oDS_PS_MM140H.GetValue("U_OutQty", i))
'        OutWt = Trim(oDS_PS_MM140H.GetValue("U_OutWt", i))
'        InQty = Trim(oDS_PS_MM140H.GetValue("U_InQty", i))
'        InWt = Trim(oDS_PS_MM140H.GetValue("U_InWt", i))
'        JakNum = Trim(oDS_PS_MM140H.GetValue("U_JakNum", i))
'        DocDate = Trim(oDS_PS_MM140H.GetValue("U_DocDate", i))
'
''        If Trim(oDS_PS_MM140H.GetValue("U_SjQty", i)) = "" Then
''            SjQty = 0
''        Else
''            SjQty = Trim(oDS_PS_MM140H.GetValue("U_SjQty", i))
''        End If
'
'        '// DocEntry
'        sQry = "Select IsNULL(Max(DocEntry), 0) From [@PS_MM140H] "
'        RecordSet01.DoQuery sQry
'        If Trim(RecordSet01.Fields(0).Value) = 0 Then
'            DocEntry = 1
'        Else
'            DocEntry = Trim(RecordSet01.Fields(0).Value) + 1
'        End If
'
'        sQry = "INSERT INTO [@PS_MM140H]"
'        sQry = sQry & " ("
'        sQry = sQry & " DocEntry,"
'        sQry = sQry & " DocNum,"
'        sQry = sQry & " U_BPLId,"
'        sQry = sQry & " U_OutDoc,"
'        sQry = sQry & " U_OutLine,"
'        sQry = sQry & " U_ItemCode,"
'        sQry = sQry & " U_ItemName,"
'        sQry = sQry & " U_Size,"
'        sQry = sQry & " U_Mark,"
'        sQry = sQry & " U_OutItmCd,"
'        sQry = sQry & " U_OutItmNm,"
'        sQry = sQry & " U_OutGbn,"
'        sQry = sQry & " U_OutQty,"
'        sQry = sQry & " U_OutWt,"
'        sQry = sQry & " U_InQty,"
'        sQry = sQry & " U_InWt,"
'        sQry = sQry & " U_JakNum,"
'        sQry = sQry & " U_DocDate"
'        sQry = sQry & " ) "
'        sQry = sQry & "VALUES("
'        sQry = sQry & DocEntry & ","
'        sQry = sQry & DocEntry & ","
'        sQry = sQry & "'" & BPLId & "',"
'        sQry = sQry & "'" & OutDoc & "',"
'        sQry = sQry & "'" & OutLine & "',"
'        sQry = sQry & "'" & ItemCode & "',"
'        sQry = sQry & "'" & ItemName & "',"
'        sQry = sQry & "'" & Size & "',"
'        sQry = sQry & "'" & Mark & "',"
'        sQry = sQry & "'" & OutItmCd & "',"
'        sQry = sQry & "'" & OutItmNm & "',"
'        sQry = sQry & "'" & OutGbn & "',"
'        sQry = sQry & "'" & OutQty & "',"
'        sQry = sQry & "'" & OutWt & "',"
'        sQry = sQry & "'" & InQty & "',"
'        sQry = sQry & "'" & InWt & "',"
'        sQry = sQry & "'" & JakNum & "',"
'        sQry = sQry & "'" & DocDate & "'"
'        sQry = sQry & ")"
'        RecordSet01.DoQuery sQry
'    Next
'
'    MDC_Com.MDC_GF_Message "입고확인 완료!", "S"
'
'    Set RecordSet01 = Nothing
'    Add_MM140H = True
'    Exit Function
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Add_MM140H_Error:
'    Add_MM140Ht = False
'    Set RecordSet01 = Nothing
'    MDC_Com.MDC_GF_Message "Add_MM140H_Error:" & Err.Number & " - " & Err.Description, "E"
'End Function

Public Sub LoadData()
On Error GoTo LoadData_Error
    Dim i As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLId$, OutDoc$, OutGbn$, CardCode$, DocDateFr$, DocDateTo$, Status$
        
    BPLId = Trim(oForm01.Items("BPLId").Specific.Value)
    OutDoc = Trim(oForm01.Items("OutDoc").Specific.Value)
    OutGbn = Trim(oForm01.Items("OutGbn").Specific.Value)
    CardCode = Trim(oForm01.Items("CardCode").Specific.Value)
    DocDateFr = Trim(oForm01.Items("DocDateFr").Specific.Value)
    DocDateTo = Trim(oForm01.Items("DocDateTo").Specific.Value)
    Status = Trim(oForm01.Items("Status").Specific.Value)
    
    If BPLId = "" Then BPLId = "%"
    If OutDoc = "" Then OutDoc = "%"
    If OutGbn = "" Then OutGbn = "%"
    If CardCode = "" Then CardCode = "%"
    If DocDateFr = "" Then DocDateFr = "19000101"
    If DocDateTo = "" Then DocDateTo = "20991231"
    
    sQry = "EXEC [PS_MM140_01] '" & BPLId & "','" & OutDoc & "','" & OutGbn & "','" & CardCode & "','" & DocDateFr & "','" & DocDateTo & "','" & Status & "'"
    oRecordSet01.DoQuery sQry
    
    oMat01.Clear
    oDS_PS_MM140H.Clear
    
    If (oRecordSet01.RecordCount = 0) Then
        MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
        Set oRecordSet01 = Nothing
        Exit Sub
    End If
    
    oForm01.Freeze True
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_MM140H.Size Then
            oDS_PS_MM140H.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_MM140H.Offset = i
        oDS_PS_MM140H.setValue "DocNum", i, i + 1
        oDS_PS_MM140H.setValue "U_OutDoc", i, Trim(oRecordSet01.Fields("U_OutDoc").Value)
        oDS_PS_MM140H.setValue "U_OutLine", i, Trim(oRecordSet01.Fields("U_LineNum").Value)
        oDS_PS_MM140H.setValue "U_ItemCode", i, Trim(oRecordSet01.Fields("U_ItemCode").Value)
        oDS_PS_MM140H.setValue "U_ItemName", i, Trim(oRecordSet01.Fields("U_ItemName").Value)
        oDS_PS_MM140H.setValue "U_Size", i, Trim(oRecordSet01.Fields("U_Size").Value)
        oDS_PS_MM140H.setValue "U_Mark", i, Trim(oRecordSet01.Fields("U_Mark").Value)
        oDS_PS_MM140H.setValue "U_OutItmCd", i, Trim(oRecordSet01.Fields("U_OutItmCd").Value)
        oDS_PS_MM140H.setValue "U_OutItmNm", i, Trim(oRecordSet01.Fields("U_OutItmNm").Value)
        oDS_PS_MM140H.setValue "U_OutGbn", i, Trim(oRecordSet01.Fields("U_OutGbn").Value)
        oDS_PS_MM140H.setValue "U_OutWhCd", i, Trim(oRecordSet01.Fields("U_OutWhCd").Value)
        oDS_PS_MM140H.setValue "U_OutWhNm", i, Trim(oRecordSet01.Fields("U_OutWhNm").Value)
        oDS_PS_MM140H.setValue "U_InWhCd", i, Trim(oRecordSet01.Fields("U_InWhCd").Value)
        oDS_PS_MM140H.setValue "U_InWhNm", i, Trim(oRecordSet01.Fields("U_InWhNm").Value)
        oDS_PS_MM140H.setValue "U_OutQty", i, Trim(oRecordSet01.Fields("U_OutQty").Value)
        oDS_PS_MM140H.setValue "U_OutWt", i, Trim(oRecordSet01.Fields("U_OutWt").Value)
        oDS_PS_MM140H.setValue "U_OutDate", i, Format(Trim(oRecordSet01.Fields("U_OutDate").Value), "YYYYMMDD")
        oDS_PS_MM140H.setValue "U_JakNum", i, Trim(oRecordSet01.Fields("U_OrdNum").Value)
        oDS_PS_MM140H.setValue "U_CardCode", i, Trim(oRecordSet01.Fields("U_CardCode").Value)
        oDS_PS_MM140H.setValue "U_CardName", i, Trim(oRecordSet01.Fields("U_CardName").Value)
        
        If Trim(oForm01.Items("Status").Specific.Selected.Value) = "Y" Then
            oDS_PS_MM140H.setValue "U_DocDate", i, Format(Trim(oRecordSet01.Fields("U_DocDate").Value), "YYYYMMDD")
            oDS_PS_MM140H.setValue "U_InQty", i, Trim(oRecordSet01.Fields("U_InQty").Value)
            oDS_PS_MM140H.setValue "U_InWt", i, Trim(oRecordSet01.Fields("U_InWt").Value)
        ElseIf Trim(oForm01.Items("Status").Specific.Selected.Value) = "N" Then
            oDS_PS_MM140H.setValue "U_DocDate", i, Format(Date, "YYYYMMDD")
            oDS_PS_MM140H.setValue "U_InQty", i, Trim(oRecordSet01.Fields("U_OutQty").Value)
            oDS_PS_MM140H.setValue "U_InWt", i, Trim(oRecordSet01.Fields("U_OutWt").Value)
        End If

        oRecordSet01.MoveNext
        ProgBar01.Value = ProgBar01.Value + 1
        ProgBar01.Text = ProgBar01.Value & "/" & oRecordSet01.RecordCount & "건 조회중...!"
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

Private Function Add_oStockTransfer()
On Error GoTo Add_oStockTransfer_Error
    Dim RetVal As Long
    Dim ErrNum As String
    Dim ErrCode As Long
    Dim ErrMsg As String
    Dim lTypeCount As Integer   '전체 StockInfo 구조체배열의 RowCount
    Dim i, j, k, Q, z, r  As Long
    Dim DocCnt        As Long
    Dim Chk1_Val As String
    Dim sCur_OutDoc$, sNxt_OutDoc$
    Dim sCur_ItemCode As String
    Dim sNxt_ItemCode As String
    Dim sCur_OutWhs As String
    Dim sNxt_OutWhs As String
    Dim sCur_InWhs As String
    Dim sNxt_InWhs As String
    Dim RtnDocNum As String
    Dim DI_oStockTransfer As SAPbobsCOM.StockTransfer
    Dim oPrgBar As SAPbouiCOM.ProgressBar
    Dim StockTransLineCounter As Long
    
    Dim DocEntry$, BPLId$, OutDoc$, OutLine$, ItemCode$, ItemName$, Size$, Mark$, OutItmCd$, OutItmNm$, OutGbn$, CardCode, CardName$, OutDate$
    Dim OutQty As Currency, OutWt As Currency, InQty As Currency, InWt As Currency, JakNum$, DocDate$, OutWhCd$, OutWhNm$, InWhCd$, InWhNm$, STDocNum$
    
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Add_oStockTransfer = True
    
    oMat01.FlushToDataSource
    For i = 0 To oMat01.RowCount - 1
        'DI API
        If Trim(oDS_PS_MM140H.GetValue("U_Check", i)) = "Y" Then
            ReDim Preserve StockInfo(lTypeCount)
            
            StockInfo(lTypeCount).ItemCode = Trim(oDS_PS_MM140H.GetValue("U_OutItmCd", i))
            StockInfo(lTypeCount).FromWarehouseCode = Trim(oDS_PS_MM140H.GetValue("U_OutWhCd", i))
            StockInfo(lTypeCount).ToWarehouseCode = Trim(oDS_PS_MM140H.GetValue("U_InWhCd", i))
            StockInfo(lTypeCount).Weight = Round(Trim(oDS_PS_MM140H.GetValue("U_InWt", i)), 2)
            StockInfo(lTypeCount).Qty = Val(Trim(oDS_PS_MM140H.GetValue("U_InQty", i)))
            
            StockInfo(lTypeCount).OutDoc = Trim(oDS_PS_MM140H.GetValue("U_OutDoc", i))
            StockInfo(lTypeCount).Chk = "N"
            StockInfo(lTypeCount).MatrixRow = (i + 1)
            StockInfo(lTypeCount).DocDate = Trim(oDS_PS_MM140H.GetValue("U_DocDate", i))
            StockInfo(lTypeCount).LineNum = Trim(oDS_PS_MM140H.GetValue("DocNum", i)) - 1
            lTypeCount = lTypeCount + 1
        End If
    Next
    
    For i = 0 To (UBound(StockInfo))
        StockInfo(i).StockTransDocEntry = ""
    Next
    
    Sbo_Company.StartTransaction
    
    If Trim(oForm01.Items("OutGbn").Specific.Value) = "10" Then
        For i = 0 To (UBound(StockInfo))
            If StockInfo(i).Chk <> "N" Then GoTo Continue_First
            
            sCur_OutDoc = StockInfo(i).OutDoc
            sCur_OutWhs = StockInfo(i).FromWarehouseCode
            
            Set DI_oStockTransfer = Nothing
            Set DI_oStockTransfer = Sbo_Company.GetBusinessObject(oStockTransfer)
            
            DI_oStockTransfer.DocDate = Format(StockInfo(i).DocDate, "0000-00-00")
            DI_oStockTransfer.FromWarehouse = sCur_OutWhs
            DI_oStockTransfer.Comments = "[PS_MM140]외주업체 입고확인."
            
            StockTransLineCounter = -1
            For k = i To (UBound(StockInfo))
                If StockInfo(k).Chk <> "N" Then GoTo Continue_Second
                
                sNxt_OutDoc = StockInfo(k).OutDoc
                sNxt_OutWhs = StockInfo(k).FromWarehouseCode
                
                If sCur_OutWhs <> sNxt_OutWhs Or sCur_OutDoc <> sNxt_OutDoc Then
                    GoTo Continue_Second
                End If
                
                If (i <> k) Then
                    DI_oStockTransfer.Lines.Add
                End If
                StockTransLineCounter = StockTransLineCounter + 1
                '--< Line > --
                With DI_oStockTransfer.Lines
                    .ItemCode = StockInfo(k).ItemCode
                    .UserFields("U_Qty").Value = Trim(StockInfo(k).Qty)
                    .Quantity = Round(StockInfo(k).Weight, 2)
                    .WarehouseCode = StockInfo(k).ToWarehouseCode
                    StockInfo(k).Chk = "Y"   '/ 적용한 라인에 대한 표시
                    StockInfo(k).StockTransDocEntry = "Checked"
                    StockInfo(k).StockTransLineNum = StockTransLineCounter
                End With
Continue_Second:
            Next
            
            RetVal = DI_oStockTransfer.Add
            If RetVal = 0 Then
                DocCnt = DocCnt + 1
                Call Sbo_Company.GetNewObjectCode(RtnDocNum) '//재고이전문서번호
                For r = 0 To UBound(StockInfo)
                    If StockInfo(r).StockTransDocEntry = "Checked" Then
                        StockInfo(r).StockTransDocEntry = RtnDocNum
                    End If
                Next
             Else
                Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
                GoTo Add_oStockTransfer_Error
            End If
            
Continue_First:
        Next
    End If
    
    '//문서저장
    BPLId = Trim(oForm01.Items("BPLId").Specific.Value)
    For i = 0 To (UBound(StockInfo))
'        If StockInfo(i).Chk = "Y" Then
            OutDoc = Trim(oDS_PS_MM140H.GetValue("U_OutDoc", StockInfo(i).LineNum))
            OutLine = Trim(oDS_PS_MM140H.GetValue("U_OutLine", StockInfo(i).LineNum))
            ItemCode = Trim(oDS_PS_MM140H.GetValue("U_ItemCode", StockInfo(i).LineNum))
            ItemName = Trim(oDS_PS_MM140H.GetValue("U_ItemName", StockInfo(i).LineNum))
            Size = Trim(oDS_PS_MM140H.GetValue("U_Size", StockInfo(i).LineNum))
            Mark = Trim(oDS_PS_MM140H.GetValue("U_Mark", StockInfo(i).LineNum))
            OutItmCd = Trim(oDS_PS_MM140H.GetValue("U_OutItmCd", StockInfo(i).LineNum))
            OutItmNm = Trim(oDS_PS_MM140H.GetValue("U_OutItmNm", StockInfo(i).LineNum))
            OutGbn = Trim(oDS_PS_MM140H.GetValue("U_OutGbn", StockInfo(i).LineNum))
            OutWhCd = Trim(oDS_PS_MM140H.GetValue("U_OutWhCd", StockInfo(i).LineNum))
            OutWhNm = Trim(oDS_PS_MM140H.GetValue("U_OutWhNm", StockInfo(i).LineNum))
            InWhCd = Trim(oDS_PS_MM140H.GetValue("U_InWhCd", StockInfo(i).LineNum))
            InWhNm = Trim(oDS_PS_MM140H.GetValue("U_InWhNm", StockInfo(i).LineNum))
            OutQty = Trim(oDS_PS_MM140H.GetValue("U_OutQty", StockInfo(i).LineNum))
            OutWt = Trim(oDS_PS_MM140H.GetValue("U_OutWt", StockInfo(i).LineNum))
            InQty = Trim(oDS_PS_MM140H.GetValue("U_InQty", StockInfo(i).LineNum))
            InWt = Trim(oDS_PS_MM140H.GetValue("U_InWt", StockInfo(i).LineNum))
            JakNum = Trim(oDS_PS_MM140H.GetValue("U_JakNum", StockInfo(i).LineNum))
            OutDate = Trim(oDS_PS_MM140H.GetValue("U_OutDate", StockInfo(i).LineNum))
            DocDate = Trim(oDS_PS_MM140H.GetValue("U_DocDate", StockInfo(i).LineNum))
            CardCode = Trim(oDS_PS_MM140H.GetValue("U_CardCode", StockInfo(i).LineNum))
            CardName = Trim(oDS_PS_MM140H.GetValue("U_CardName", StockInfo(i).LineNum))
            STDocNum = StockInfo(i).StockTransDocEntry
    
    '        If Trim(oDS_PS_MM140H.GetValue("U_SjQty", i)) = "" Then
    '            SjQty = 0
    '        Else
    '            SjQty = Trim(oDS_PS_MM140H.GetValue("U_SjQty", i))
    '        End If
            
            '// DocEntry
            sQry = "Select IsNULL(Max(DocEntry), 0) From [@PS_MM140H] "
            oRecordSet01.DoQuery sQry
            If Trim(oRecordSet01.Fields(0).Value) = 0 Then
                DocEntry = 1
            Else
                DocEntry = Trim(oRecordSet01.Fields(0).Value) + 1
            End If
            
            sQry = "INSERT INTO [@PS_MM140H]"
            sQry = sQry & " ("
            sQry = sQry & " DocEntry,"
            sQry = sQry & " DocNum,"
            sQry = sQry & " U_BPLId,"
            sQry = sQry & " U_OutDoc,"
            sQry = sQry & " U_OutLine,"
            sQry = sQry & " U_ItemCode,"
            sQry = sQry & " U_ItemName,"
            sQry = sQry & " U_Size,"
            sQry = sQry & " U_Mark,"
            sQry = sQry & " U_OutItmCd,"
            sQry = sQry & " U_OutItmNm,"
            sQry = sQry & " U_OutGbn,"
            sQry = sQry & " U_OutWhCd,"
            sQry = sQry & " U_OutWhNm,"
            sQry = sQry & " U_InWhCd,"
            sQry = sQry & " U_InWhNm,"
            sQry = sQry & " U_OutQty,"
            sQry = sQry & " U_OutWt,"
            sQry = sQry & " U_InQty,"
            sQry = sQry & " U_InWt,"
            sQry = sQry & " U_JakNum,"
            sQry = sQry & " U_OutDate,"
            sQry = sQry & " U_DocDate,"
            sQry = sQry & " U_CardCode,"
            sQry = sQry & " U_CardName,"
            sQry = sQry & " U_STDocNum"
            sQry = sQry & " ) "
            sQry = sQry & "VALUES("
            sQry = sQry & DocEntry & ","
            sQry = sQry & DocEntry & ","
            sQry = sQry & "'" & BPLId & "',"
            sQry = sQry & "'" & OutDoc & "',"
            sQry = sQry & "'" & OutLine & "',"
            sQry = sQry & "'" & ItemCode & "',"
            sQry = sQry & "'" & ItemName & "',"
            sQry = sQry & "'" & Size & "',"
            sQry = sQry & "'" & Mark & "',"
            sQry = sQry & "'" & OutItmCd & "',"
            sQry = sQry & "'" & OutItmNm & "',"
            sQry = sQry & "'" & OutGbn & "',"
            sQry = sQry & "'" & OutWhCd & "',"
            sQry = sQry & "'" & OutWhNm & "',"
            sQry = sQry & "'" & InWhCd & "',"
            sQry = sQry & "'" & InWhNm & "',"
            sQry = sQry & "'" & OutQty & "',"
            sQry = sQry & "'" & OutWt & "',"
            sQry = sQry & "'" & InQty & "',"
            sQry = sQry & "'" & InWt & "',"
            sQry = sQry & "'" & JakNum & "',"
            sQry = sQry & "'" & OutDate & "',"
            sQry = sQry & "'" & DocDate & "',"
            sQry = sQry & "'" & CardCode & "',"
            sQry = sQry & "'" & CardName & "',"
            sQry = sQry & "'" & STDocNum & "'"
            sQry = sQry & ")"
            oRecordSet01.DoQuery sQry
'        End If
    Next
    
    If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_Commit
    MDC_Com.MDC_GF_Message DocCnt & " 개의 재고이전 문서가 발행되었고 선택한 입고확인문서가 등록되었습니다.", "S"
    Set DI_oStockTransfer = Nothing
    Set oRecordSet01 = Nothing
    Add_oStockTransfer = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_oStockTransfer_Error:
    If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_RollBack
    Set DI_oStockTransfer = Nothing
    Set oRecordSet01 = Nothing
    Add_oStockTransfer = False
    MDC_Com.MDC_GF_Message "Add_oStockTransfer_Error:" & ErrCode & " - " & ErrMsg, "E"
End Function
