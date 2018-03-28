VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_SD920"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_SD920.cls
'//  Module         : CO
'//  Description    : A/R송장 미처리 납품 처리
'//  FormType       : PS_SD920
'//  Create Date    : 2011.06.30
'//  Modified Date  :
'//  Creator        : Ryu Yung Jo
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01             As SAPbouiCOM.Form
Public oMat01              As SAPbouiCOM.Matrix
Private oDS_PS_SD920H As SAPbouiCOM.DBDataSource    '등록헤더
Private oDS_PS_SD920L As SAPbouiCOM.DBDataSource    '등록라인

Private oLast_Item_UID      As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID       As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row       As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

'DI를 위한 정보를 가지는 구조체
Private Type oS_PS_SD920Ls
    LineNum         As Long
    ODLNDocB        As String
    DLN1LinB        As String
    CardCode        As String
    CardName        As String
    ItemCode        As String
    ItemName        As String
    WhsCode         As String
    Quantity        As Currency
    OpenQty         As Currency
    Price           As Currency
    LineTotal       As Currency
    Qty             As Long
    BaseType        As String
    BaseEntry       As String
    BaseLine        As String
    ULineNum        As String
    ORDRDoc         As Long
    RDR1Line        As Long
    ORDNDoc         As String
    RDN1Line        As String
    ODLNDoc         As String
    DLN1Line        As String
    Check           As Boolean
End Type

Private oS_PS_SD920L() As oS_PS_SD920Ls

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_SD920.srf")
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
    
    oFormUniqueID01 = "PS_SD920_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
  
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    '화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
    oForm01.DataBrowser.BrowseBy = "DocEntry"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting
    Call Initialization
    Call FormClear
'    Call Add_MatrixRow(0, True)
    Call FormItemEnabled
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1284"), False        '// 취소
    oForm01.EnableMenu ("1293"), False        '// 행삭제
        
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
    Dim i&, j&



                    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE Or oForm01.Mode = fm_UPDATE_MODE Then
                        If HeaderSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        If MatrixSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        If oForm01.Mode = fm_ADD_MODE Then
                            If Add_oReturns_oDeliveryNotes(1) = False Then
                                BubbleEvent = False
                                Exit Sub
                            Else
'                                oMat01.FlushToDataSource
'                                j = 1
'                                For i = 0 To oMat01.VisualRowCount - 1
'                                    If Trim(oDS_PS_SD920L.GetValue("U_Check", i)) = "Y" Then
'                                        Call oDS_PS_SD920L.setValue("U_LineNum", i, j)
'                                    Else
'                                        oDS_PS_SD920L.RemoveRecord (i)
'                                    End If
'                                oMat01.LoadFromDataSource
'                                Next i
                            End If
                        End If
                    End If
                End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            
                If pval.ItemUID = "Mat01" Then

                    If pval.Row > 0 Then
        
                        Call oMat01.SelectRow(pval.Row, True, False)
        
                    End If
        
                End If
            
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
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE And pval.Action_Success = True Then
                        oForm01.Mode = fm_OK_MODE
                        Call Sbo_Application.ActivateMenuItem("1282")
                    End If
                ElseIf pval.ItemUID = "Btn01" Then
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    Call LoadData
                End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
'et_CLICK ///////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_CLICK: '//6
                If pval.ItemUID = "Mat01" And pval.Row <> 0 And pval.ColUID = "Check" Then
                    oForm01.Freeze True
                    oMat01.FlushToDataSource
                    For i = 0 To oMat01.VisualRowCount - 1
                        If oDS_PS_SD920L.GetValue("U_ODLNDocB", i) = oDS_PS_SD920L.GetValue("U_ODLNDocB", pval.Row - 1) Then
                            Call oDS_PS_SD920L.setValue("U_Check", i, Trim(oDS_PS_SD920L.GetValue("U_Check", pval.Row - 1)))
                        End If
                    Next i
                    oMat01.LoadFromDataSource
'                    oMat01.Columns("Check").Cells(pval.Row).Click
'                    For i = 1 To oMat01.VisualRowCount
'                        If oMat01.Columns("ODLNDocB").Cells(i).Specific.VALUE = oMat01.Columns("ODLNDocB").Cells(pval.Row).Specific.VALUE Then
'                            If oMat01.Columns("Check").Cells(pval.Row).Specific.Checked = True Then
'                                oMat01.Columns("Check").Cells(i).Specific.Checked = True
'                            Else
'                                oMat01.Columns("Check").Cells(i).Specific.Checked = False
'                            End If
'
'                        End If
'                    Next i
                    oForm01.Freeze False
                End If
'et_DOUBLE_CLICK ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_DOUBLE_CLICK: '//7
                If pval.ItemUID = "Mat01" And pval.Row = 0 And pval.ColUID = "Check" Then
                    Dim Check$
                    oForm01.Freeze True
                    oMat01.FlushToDataSource
                    If Trim(oDS_PS_SD920L.GetValue("U_Check", 0)) = "" Or Trim(oDS_PS_SD920L.GetValue("U_Check", 0)) = "N" Then
                        Check = "Y"
                    ElseIf Trim(oDS_PS_SD920L.GetValue("U_Check", 0)) = "Y" Then
                        Check = "N"
                    End If
                    For i = 0 To oMat01.VisualRowCount - 1
                        oDS_PS_SD920L.setValue "U_Check", i, Check
                    Next i
                    oMat01.LoadFromDataSource
                    oForm01.Freeze False
                End If
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
'et_MATRIX_LOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_MATRIX_LOAD: '//11
                Call Add_MatrixRow(oMat01.RowCount, False)
                Call oMat01.AutoResizeColumns
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
            Case et_FORM_RESIZE '//20
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
            Case et_LOST_FOCUS: '//4
'et_FORM_UNLOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_FORM_UNLOAD: '//17
                RemoveForms oFormUniqueID01
                Set oForm01 = Nothing
                Set oMat01 = Nothing
                Set oDS_PS_SD920H = Nothing
                Set oDS_PS_SD920L = Nothing
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
    oForm01.Freeze False
    MDC_Com.MDC_GF_Message "Raise_ItemEvent_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_MenuEvent_Error
    Dim i&
    
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
            Case "1293": '행삭제
                oForm01.Freeze True
                If oMat01.RowCount <> oMat01.VisualRowCount Then
                    For i = 0 To oMat01.VisualRowCount - 1
                        oMat01.Columns("LineNum").Cells(i + 1).Specific.VALUE = i + 1
                    Next i
                    
                    oMat01.FlushToDataSource
                    oDS_PS_SD920L.RemoveRecord oDS_PS_SD920L.Size - 1       '// Mat01에 마지막라인(빈라인) 삭제
                    oMat01.Clear
                    oMat01.LoadFromDataSource
                    
                    If oMat01.Columns("OrdMgNum").Cells(oMat01.RowCount).Specific.VALUE <> "" Then
                        Call Add_MatrixRow(oMat01.RowCount, False)
                    End If
                End If
                oForm01.Freeze False
            Case "1281": '찾기
                oForm01.Freeze True
                Call FormItemEnabled
'                oForm01.Items("CycleCod").Click ct_Regular
                oForm01.Freeze False
            Case "1282": '추가
                oForm01.Freeze True
                Call FormItemEnabled
                Call Initialization
                Call FormClear
                oForm01.Freeze False
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                oForm01.Freeze True
                Call FormItemEnabled
'                If oMat01.VisualRowCount > 0 Then
'                    If oMat01.Columns("CycleCod").Cells(oMat01.VisualRowCount).Specific.Value <> "" Then
'                        Add_MatrixRow oMat01.RowCount, False
'                    End If
'                End If
                oForm01.Freeze False
            Case "1287": '// 복제
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_MenuEvent_Error:
    oForm01.Freeze False
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
        '//작업
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
    Set oDS_PS_SD920H = oForm01.DataSources.DBDataSources("@PS_SD920H")
    Set oDS_PS_SD920L = oForm01.DataSources.DBDataSources("@PS_SD920L")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_Auto
    oMat01.AutoResizeColumns
    
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
    
    '// 마감년월
'    Set oCombo = oForm01.Items("ClsPrd").Specific
'    sQry = "SELECT Code, Name From [OFPR]"
'    oRecordSet01.DoQuery sQry
'    Do Until oRecordSet01.EOF
'        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
'        oRecordSet01.MoveNext
'    Loop
    
    '// 사업장
    Set oCombo = oForm01.Items("BPLId").Specific
    sQry = "SELECT BPLId, BPLName From [OBPL] order by 1"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
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
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    '//아이디별 사업장 세팅
    Set oCombo = oForm01.Items("BPLId").Specific
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
  
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Initialization_Error:
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
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
    If (oForm01.Mode = fm_ADD_MODE) Then
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("YYYYMM").Enabled = True
        oMat01.Columns("Check").Editable = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("YYYYMM").Enabled = True
        oMat01.Columns("Check").Editable = False
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        oForm01.Items("BPLId").Enabled = False
        oForm01.Items("YYYYMM").Enabled = False
        oMat01.Columns("Check").Editable = False
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub FormClear()
On Error GoTo FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_SD920'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormClear_Error:
    MDC_Com.MDC_GF_Message "FormClear_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_SD920L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_SD920L.Offset = oRow
    oDS_PS_SD920L.setValue "U_LineNum", oRow, oRow + 1
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
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    Dim sRow As Long
    Dim sSeq$
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sRow = oRow
    
    Select Case oUID
        Case "Mat01"
'            If oCol = "CycleCod" Then
'                oForm01.Freeze True
'                oMat01.FlushToDataSource
'
'                If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 0) And Trim(oMat01.Columns("CycleCod").Cells(oRow).Specific.Value) <> "" Then
'                    oMat01.FlushToDataSource
'                    Call Add_MatrixRow(oMat01.RowCount, False)
'                    oMat01.Columns("CycleCod").Cells(oRow).Click ct_Regular
'                End If
'
''                sQry = "Select ItemName, FrgnName From OITM Where ItemCode = '" & Trim(oMat01.Columns("ItemCode").Cells(oRow).Specific.Value) & "'"
''                oRecordSet01.DoQuery sQry
''                oMat01.Columns("ItemName").Cells(oRow).Specific.Value = Trim(oRecordSet01.Fields(0).Value)
''                oMat01.Columns("FrgnName").Cells(oRow).Specific.Value = Trim(oRecordSet01.Fields(1).Value)
''
''                oMat01.Columns("ItemCode").Cells(oRow).Click ct_Regular
'                oForm01.Freeze False
'            End If
    End Select

    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FlushToItemValue_Error:
    Set oRecordSet01 = Nothing
    oForm01.Freeze False
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "구매견적문서가 취소되었거나 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
    Else
        MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum          As Integer
    Dim DocNum          As String

    ErrNum = 0

    '// Check
    Select Case True
        Case oDS_PS_SD920H.GetValue("U_YYYYMM", 0) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_SD920H.GetValue("U_BPLId", 0) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "전기년월은 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "사업장은 필수입력사항입니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function

Private Function MatrixSpaceLineDel() As Boolean
On Error GoTo MatrixSpaceLineDel_Error
    Dim i               As Long
    Dim ErrNum          As Integer
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String

    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    ErrNum = 0
    
    oMat01.FlushToDataSource

    '// 라인
    If oMat01.VisualRowCount = 0 Then
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    ElseIf oMat01.VisualRowCount = 1 Then
'        If oDS_PS_SD920L.GetValue("U_CycleCod", 0) = "" Then
'            ErrNum = 2
'            GoTo MatrixSpaceLineDel_Error
'        End If
    End If
    
'    For i = 0 To oMat01.VisualRowCount - 2
'        Select Case True
'            Case oDS_PS_SD920L.GetValue("U_ItemCode", i) = ""
'                ErrNum = 2
'                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_SD920L.GetValue("U_Qty", i) = "" Or oDS_PS_SD920L.GetValue("U_Qty", i) = 0
'                ErrNum = 3
'                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_SD920L.GetValue("U_Weight", i) = ""
'                ErrNum = 4
'                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_SD920L.GetValue("U_Price", i) = 0
'                ErrNum = 5
'                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_SD920L.GetValue("U_LinTotal", i) = 0
'                ErrNum = 6
'                GoTo MatrixSpaceLineDel_Error
'        End Select
'    Next i
    oMat01.LoadFromDataSource

    Set oRecordSet = Nothing
    MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
MatrixSpaceLineDel_Error:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "라인 데이터가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "첫라인에 배부사이클 코드가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "수량은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "중량은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "단가는 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "금액은 필수사항입니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function

Sub Delete_EmptyRow()
On Error GoTo Delete_EmptyRow_Error
    Dim i&
    
    oMat01.FlushToDataSource
    
    For i = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PS_SD920L.GetValue("U_CycleCod", i)) = "" Then
            oDS_PS_SD920L.RemoveRecord i   '// Mat01에 마지막라인(빈라인) 삭제
        End If
    Next i
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Delete_EmptyRow_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub LoadData()
On Error GoTo LoadData_Error
    Dim i As Long
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim YYYYMM$, BPLID$
    
    YYYYMM = Trim(oForm01.Items("YYYYMM").Specific.VALUE)
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
           
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
           
    sQry = "EXEC [PS_SD920_01] '" & BPLID & "','" & YYYYMM & "'"
    oRecordSet01.DoQuery sQry
    
    oMat01.Clear
    oDS_PS_SD920L.Clear
    
    If (oRecordSet01.RecordCount = 0) Then
        MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
        Set oRecordSet01 = Nothing
        Exit Sub
    End If
    
    oForm01.Freeze True
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_SD920L.Size Then
            oDS_PS_SD920L.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_SD920L.Offset = i
        oDS_PS_SD920L.setValue "U_LineNum", i, i + 1
        oDS_PS_SD920L.setValue "U_ODLNDocB", i, Trim(oRecordSet01.Fields("DocEntry").VALUE)
        oDS_PS_SD920L.setValue "U_DLN1LinB", i, Trim(oRecordSet01.Fields("LineNum").VALUE)
        oDS_PS_SD920L.setValue "U_CardCode", i, Trim(oRecordSet01.Fields("CardCode").VALUE)
        oDS_PS_SD920L.setValue "U_CardName", i, Trim(oRecordSet01.Fields("CardName").VALUE)
        oDS_PS_SD920L.setValue "U_ItemCode", i, Trim(oRecordSet01.Fields("ItemCode").VALUE)
        oDS_PS_SD920L.setValue "U_ItemName", i, Trim(oRecordSet01.Fields("Dscription").VALUE)
        oDS_PS_SD920L.setValue "U_WhsCode", i, Trim(oRecordSet01.Fields("WhsCode").VALUE)
        oDS_PS_SD920L.setValue "U_Quantity", i, Trim(oRecordSet01.Fields("Quantity").VALUE)
        oDS_PS_SD920L.setValue "U_OpenQty", i, Trim(oRecordSet01.Fields("OpenQty").VALUE)
        oDS_PS_SD920L.setValue "U_Price", i, Trim(oRecordSet01.Fields("Price").VALUE)
        oDS_PS_SD920L.setValue "U_LinTotal", i, Trim(oRecordSet01.Fields("LineTotal").VALUE)
        oDS_PS_SD920L.setValue "U_Qty", i, Trim(oRecordSet01.Fields("U_Qty").VALUE)
        oDS_PS_SD920L.setValue "U_BaseType", i, Trim(oRecordSet01.Fields("U_BaseType").VALUE)
        oDS_PS_SD920L.setValue "U_BaseDoc", i, Trim(oRecordSet01.Fields("U_BaseEntry").VALUE)
        oDS_PS_SD920L.setValue "U_BaseLine", i, Trim(oRecordSet01.Fields("U_BaseLine").VALUE)
        oDS_PS_SD920L.setValue "U_ULineNum", i, Trim(oRecordSet01.Fields("U_LineNum").VALUE)
        oDS_PS_SD920L.setValue "U_ORDRDoc", i, Trim(oRecordSet01.Fields("BaseEntry").VALUE)
        oDS_PS_SD920L.setValue "U_RDR1Line", i, Trim(oRecordSet01.Fields("BaseLine").VALUE)

        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
    Next
    
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    oForm01.Freeze False

    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadData_Error:

    oForm01.Freeze False
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "LoadData_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'Private Function Add_oReturns_oDeliveryNotes(ChkType As Integer) As Boolean
'On Error GoTo Add_oReturns_oDeliveryNotes_Error
'
'    Dim DI_oReturns       As SAPbobsCOM.Documents '반품 문서객체
'    Dim DI_oDeliveryNotes As SAPbobsCOM.Documents '납품 문서객체
'
'    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
'
'    Dim ProgBar01 As SAPbouiCOM.ProgressBar
'    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("반품 납품 생성!", lTypeCount, False)
'
'    If Sbo_Company.InTransaction = True Then
'        Sbo_Company.EndTransaction wf_RollBack
'    End If
'    Sbo_Company.StartTransaction
'
'
'    oMat01.FlushToDataSource
'
'    '전기일 처리(2012.01.06 송명규 추가)
'    ODLNDocDate = CDate(DateAdd("M", 1, Left(oDS_PS_SD920H.GetValue("U_YYYYMM", 0), 4) & "-" & Right(oDS_PS_SD920H.GetValue("U_YYYYMM", 0), 2) & "-01")) 'Format(TempI, "0000-00-00")
'    ORDNDocDate = DateAdd("D", -1, ODLNDocDate)
'
'        For i = 0 To oMat01.RowCount - 1
'        'DI API
'        If Trim(oDS_PS_SD920L.GetValue("U_Check", i)) = "Y" Then
'            ReDim Preserve oS_PS_SD920L(lTypeCount)
'            oS_PS_SD920L(lTypeCount).LineNum = Trim(oDS_PS_SD920L.GetValue("U_LineNum", i))
'            oS_PS_SD920L(lTypeCount).ODLNDocB = Trim(oDS_PS_SD920L.GetValue("U_ODLNDocB", i))
'            oS_PS_SD920L(lTypeCount).DLN1LinB = Trim(oDS_PS_SD920L.GetValue("U_DLN1LinB", i))
'            oS_PS_SD920L(lTypeCount).CardCode = Trim(oDS_PS_SD920L.GetValue("U_CardCode", i))
'            oS_PS_SD920L(lTypeCount).CardName = Trim(oDS_PS_SD920L.GetValue("U_CardName", i))
'            oS_PS_SD920L(lTypeCount).ItemCode = Trim(oDS_PS_SD920L.GetValue("U_ItemCode", i))
'            oS_PS_SD920L(lTypeCount).ItemName = Trim(oDS_PS_SD920L.GetValue("U_ItemName", i))
'            oS_PS_SD920L(lTypeCount).WhsCode = Trim(oDS_PS_SD920L.GetValue("U_WhsCode", i))
'            oS_PS_SD920L(lTypeCount).Quantity = Val(Trim(oDS_PS_SD920L.GetValue("U_OpenQty", i))) 'Val(Trim(oDS_PS_SD920L.GetValue("U_Quantity", i)))
'            oS_PS_SD920L(lTypeCount).OpenQty = Val(Trim(oDS_PS_SD920L.GetValue("U_OpenQty", i)))
'            oS_PS_SD920L(lTypeCount).Price = Val(Trim(oDS_PS_SD920L.GetValue("U_Price", i)))
'            oS_PS_SD920L(lTypeCount).LineTotal = Val(Trim(oDS_PS_SD920L.GetValue("U_LinTotal", i)))
'            oS_PS_SD920L(lTypeCount).Qty = Trim(oDS_PS_SD920L.GetValue("U_Qty", i))
'            oS_PS_SD920L(lTypeCount).BaseType = Trim(oDS_PS_SD920L.GetValue("U_BaseType", i))
'            oS_PS_SD920L(lTypeCount).BaseEntry = Trim(oDS_PS_SD920L.GetValue("U_BaseDoc", i)) 'Trim(oDS_PS_MM140H.GetValue("U_BaseDoc", i))
'            oS_PS_SD920L(lTypeCount).BaseLine = Trim(oDS_PS_SD920L.GetValue("U_BaseLine", i)) 'Trim(oDS_PS_MM140H.GetValue("U_BaseLine", i))
'            oS_PS_SD920L(lTypeCount).ULineNum = Trim(oDS_PS_SD920L.GetValue("U_ULineNum", i))
'            oS_PS_SD920L(lTypeCount).ORDRDoc = Trim(oDS_PS_SD920L.GetValue("U_ORDRDoc", i))
'            oS_PS_SD920L(lTypeCount).RDR1Line = Trim(oDS_PS_SD920L.GetValue("U_RDR1Line", i))
'            oS_PS_SD920L(lTypeCount).Check = False
'            lTypeCount = lTypeCount + 1
'        End If
'    Next i
'
'    '==========반품처리 DI_S==========
'    '1.
'
'
'    '==========반품처리 DI_E==========
'
'
'
'    '==========납품처리 DI_S==========
'
'    '==========납품처리 DI_E==========
'
'
'
'    oMat01.LoadFromDataSource
'    ProgBar01.Stop
'    If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_Commit
'
'    Set DI_oReturns = Nothing
'    Set DI_oDeliveryNotes = Nothing
'    Add_oReturns_oDeliveryNotes = True
'    Set oRecordSet01 = Nothing
'    Exit Function
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Add_oReturns_oDeliveryNotes_Error:
'    ProgBar01.Stop
'    Set ProgBar01 = Nothing
'    Set DI_oReturns = Nothing
'    Set DI_oDeliveryNotes = Nothing
'    Set oRecordSet01 = Nothing
'    If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_RollBack
'    Add_oReturns_oDeliveryNotes = False
'    If ErrNum = 1 Then
'        MDC_Com.MDC_GF_Message "Add_oReturns_oDeliveryNotes_Error:" & ErrCode & " - " & ErrMsg, "E"
'    Else
'        MDC_Com.MDC_GF_Message "Add_oReturns_oDeliveryNotes_Error:" & Err.Number & " - " & Err.Description, "E"
'    End If
'End Function

Private Function Add_oReturns_oDeliveryNotes(ChkType As Integer) As Boolean
On Error GoTo Add_oReturns_oDeliveryNotes_Error

    Dim DI_oReturns As SAPbobsCOM.Documents                 '//반품 문서객체
    Dim DI_oDeliveryNotes As SAPbobsCOM.Documents           '//납품 문서객체

    Dim i As Long
    Dim j As Long
    Dim K As Long
    Dim l As Long
    Dim m As Long
    
    Dim LineCounter As Long
    
    Dim ErrNum  As Long
    Dim ErrCode As Long
    Dim ErrMsg  As String
    Dim RetVal  As Long
    
    Dim sQry          As String
    Dim lTypeCount    As Integer   '전체 StockInfo 구조체배열의 RowCount
    Dim LineNum(1000) As Integer
    Dim Line_Count    As Long
    Dim PT_Error      As String
    Dim PT_Error_Message As String
    Dim ORDNDocDate  As Date
    Dim ODLNDocDate  As Date
    Dim DueDate      As String
    Dim ORDNDocEntry As String
    Dim ODLNDocEntry As String
    Dim TempI        As Long
    Dim BatchYN      As String
    Dim oRecordSet01  As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Company.StartTransaction

    oMat01.FlushToDataSource

    '전기일 처리(2012.01.06 송명규 추가)
    ODLNDocDate = CDate(DateAdd("M", 1, Left(oDS_PS_SD920H.GetValue("U_YYYYMM", 0), 4) & "-" & Right(oDS_PS_SD920H.GetValue("U_YYYYMM", 0), 2) & "-01")) 'Format(TempI, "0000-00-00")
    ORDNDocDate = DateAdd("D", -1, ODLNDocDate)

    For i = 0 To oMat01.RowCount - 1
        'DI API
        If Trim(oDS_PS_SD920L.GetValue("U_Check", i)) = "Y" Then
            ReDim Preserve oS_PS_SD920L(lTypeCount)
            oS_PS_SD920L(lTypeCount).LineNum = Trim(oDS_PS_SD920L.GetValue("U_LineNum", i))
            oS_PS_SD920L(lTypeCount).ODLNDocB = Trim(oDS_PS_SD920L.GetValue("U_ODLNDocB", i))
            oS_PS_SD920L(lTypeCount).DLN1LinB = Trim(oDS_PS_SD920L.GetValue("U_DLN1LinB", i))
            oS_PS_SD920L(lTypeCount).CardCode = Trim(oDS_PS_SD920L.GetValue("U_CardCode", i))
            oS_PS_SD920L(lTypeCount).CardName = Trim(oDS_PS_SD920L.GetValue("U_CardName", i))
            oS_PS_SD920L(lTypeCount).ItemCode = Trim(oDS_PS_SD920L.GetValue("U_ItemCode", i))
            oS_PS_SD920L(lTypeCount).ItemName = Trim(oDS_PS_SD920L.GetValue("U_ItemName", i))
            oS_PS_SD920L(lTypeCount).WhsCode = Trim(oDS_PS_SD920L.GetValue("U_WhsCode", i))
            oS_PS_SD920L(lTypeCount).Quantity = Val(Trim(oDS_PS_SD920L.GetValue("U_OpenQty", i))) 'Val(Trim(oDS_PS_SD920L.GetValue("U_Quantity", i)))
            oS_PS_SD920L(lTypeCount).OpenQty = Val(Trim(oDS_PS_SD920L.GetValue("U_OpenQty", i)))
            oS_PS_SD920L(lTypeCount).Price = Val(Trim(oDS_PS_SD920L.GetValue("U_Price", i)))
            oS_PS_SD920L(lTypeCount).LineTotal = Val(Trim(oDS_PS_SD920L.GetValue("U_LinTotal", i)))
            oS_PS_SD920L(lTypeCount).Qty = Trim(oDS_PS_SD920L.GetValue("U_Qty", i))
            oS_PS_SD920L(lTypeCount).BaseType = Trim(oDS_PS_SD920L.GetValue("U_BaseType", i))
            oS_PS_SD920L(lTypeCount).BaseEntry = Trim(oDS_PS_SD920L.GetValue("U_BaseDoc", i)) 'Trim(oDS_PS_MM140H.GetValue("U_BaseDoc", i))
            oS_PS_SD920L(lTypeCount).BaseLine = Trim(oDS_PS_SD920L.GetValue("U_BaseLine", i)) 'Trim(oDS_PS_MM140H.GetValue("U_BaseLine", i))
            oS_PS_SD920L(lTypeCount).ULineNum = Trim(oDS_PS_SD920L.GetValue("U_ULineNum", i))
            oS_PS_SD920L(lTypeCount).ORDRDoc = Trim(oDS_PS_SD920L.GetValue("U_ORDRDoc", i))
            oS_PS_SD920L(lTypeCount).RDR1Line = Trim(oDS_PS_SD920L.GetValue("U_RDR1Line", i))
            oS_PS_SD920L(lTypeCount).Check = False
            lTypeCount = lTypeCount + 1
        End If
    Next i

    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("반품 납품 생성!", lTypeCount, False)

    For i = 0 To (UBound(oS_PS_SD920L))
        Set DI_oReturns = Nothing
        Set DI_oDeliveryNotes = Nothing
        Set DI_oReturns = Sbo_Company.GetBusinessObject(oReturns)
        Set DI_oDeliveryNotes = Sbo_Company.GetBusinessObject(oDeliveryNotes)

        K = i
        LineNum(i + 1) = oS_PS_SD920L(i).LineNum - 1

        sQry = "Select Manbtchnum From [OITM] Where ItemCode = '" & Trim(oS_PS_SD920L(i).ItemCode) & "'"
        oRecordSet01.DoQuery sQry
        BatchYN = Trim(oRecordSet01.Fields(0).VALUE)
        If BatchYN = "Y" Then
            sQry = "        Select  BatchNum, "
            sQry = sQry & "         Quantity "
            sQry = sQry & " From    [IBT1] "
            sQry = sQry & " Where   BaseType = '15' "
            sQry = sQry & "         And BaseEntry = '" & Trim(oS_PS_SD920L(i).ODLNDocB) & "'"
            sQry = sQry & "         And BaseLinNum = '" & Trim(oS_PS_SD920L(i).DLN1LinB) & "'"

            oRecordSet01.DoQuery sQry
'            sQry = "EXECUTE [PS_SD920_02] '" & tirm(oS_PS_SD920L(i).ODLNDocB) & "', '" & tirm(oS_PS_SD920L(i).DLN1LinB) & "'"
'            oRecordSet01.DoQuery sQry
        End If

        '반품_S
        DI_oReturns.CardCode = Trim(oS_PS_SD920L(i).CardCode)
        DI_oReturns.DocDate = ORDNDocDate
        DI_oReturns.DocDueDate = ORDNDocDate
        DI_oReturns.BPL_IDAssignedToInvoice = Trim(oForm01.Items("BPLId").Specific.VALUE)
        DI_oReturns.Comments = "이월반품처리 [납품 : " & Trim(oS_PS_SD920L(i).ODLNDocB) & "]"

        '헤더 환율처리_S (2011.11.04 송명규 추가)
        Dim lclDocCur As String
        Dim lclDocRate As Double
        Dim lclQuery As String

        Dim oRecordset02 As SAPbobsCOM.Recordset
        Set oRecordset02 = Sbo_Company.GetBusinessObject(BoRecordset)
        lclQuery = "            SELECT  Currency,"
        lclQuery = lclQuery & "         Rate"
        lclQuery = lclQuery & " FROM    DLN1 "
        lclQuery = lclQuery & " WHERE   DocEntry = " & Trim(oS_PS_SD920L(i).ODLNDocB) & " "
        lclQuery = lclQuery & "         AND LineNum = " & Trim(oS_PS_SD920L(i).DLN1LinB)

        Call oRecordset02.DoQuery(lclQuery)

        lclDocCur = Trim(oRecordset02.Fields("Currency").VALUE)
        lclDocRate = Trim(oRecordset02.Fields("Rate").VALUE)

        If lclDocCur <> "KRW" Then

            DI_oReturns.DocCurrency = lclDocCur
            DI_oReturns.DocRate = lclDocRate

        End If
        '헤더 환율처리_E

        DI_oReturns.Lines.ItemCode = oS_PS_SD920L(i).ItemCode
        DI_oReturns.Lines.Quantity = oS_PS_SD920L(i).OpenQty
        DI_oReturns.Lines.WarehouseCode = oS_PS_SD920L(i).WhsCode
        DI_oReturns.Lines.UnitPrice = oS_PS_SD920L(i).Price
        DI_oReturns.Lines.LineTotal = oS_PS_SD920L(i).LineTotal
        DI_oReturns.Lines.BaseType = 15
        DI_oReturns.Lines.BaseEntry = oS_PS_SD920L(i).ODLNDocB
        DI_oReturns.Lines.BaseLine = oS_PS_SD920L(i).DLN1LinB
        DI_oReturns.Lines.UserFields("U_Qty").VALUE = oS_PS_SD920L(i).Qty
        DI_oReturns.Lines.UserFields("U_BaseType").VALUE = oS_PS_SD920L(i).BaseType '기준납품문서의 값을 그대로 저장(2016.10.20 송명규 수정)
        DI_oReturns.Lines.UserFields("U_BaseEntry").VALUE = oS_PS_SD920L(i).BaseEntry '기준납품문서의 납품처리[SD404]문서번호의 값을 그대로 저장(2016.10.20 송명규 수정)
        DI_oReturns.Lines.UserFields("U_BaseLine").VALUE = oS_PS_SD920L(i).BaseLine '기준납품문서의 납품처리[SD404]문서행번호의 값을 그대로 저장(2016.10.20 송명규 수정)

        '라인 환율처리_S
        If lclDocCur <> "KRW" Then

            DI_oReturns.Lines.Currency = lclDocCur
            DI_oReturns.Lines.rate = lclDocRate

        End If
        '라인 환율처리_E

        If BatchYN = "Y" Then
            m = 0
            Do Until oRecordSet01.EOF
                If m > 0 Then
                    DI_oReturns.Lines.BatchNumbers.Add
                End If

                DI_oReturns.Lines.BatchNumbers.BatchNumber = Trim(oRecordSet01.Fields("BatchNum").VALUE)
                DI_oReturns.Lines.BatchNumbers.Quantity = oS_PS_SD920L(i).OpenQty 'oRecordSet01.Fields("Quantity").VALUE 'oS_PS_SD920L(i).OpenQty 가용재고로 입력되도록 수정 20170307
                oRecordSet01.MoveNext
                m = m + 1
            Loop
        End If
        '반품_E

        '납품_S
        DI_oDeliveryNotes.CardCode = Trim(oS_PS_SD920L(i).CardCode)
        DI_oDeliveryNotes.DocDate = ODLNDocDate
        DI_oDeliveryNotes.DocDueDate = ODLNDocDate
        DI_oDeliveryNotes.BPL_IDAssignedToInvoice = Trim(oForm01.Items("BPLId").Specific.VALUE)
        DI_oDeliveryNotes.Comments = "이월반품처리 [판매오더 : " & Trim(oS_PS_SD920L(i).ODLNDocB) & "]"

        '헤더 환율처리_S (2011.11.04 송명규 추가)
        If lclDocCur <> "KRW" Then

            DI_oDeliveryNotes.DocCurrency = lclDocCur
            DI_oDeliveryNotes.DocRate = lclDocRate

        End If
        '헤더 환율처리_E

        DI_oDeliveryNotes.Lines.ItemCode = oS_PS_SD920L(i).ItemCode
        DI_oDeliveryNotes.Lines.Quantity = oS_PS_SD920L(i).OpenQty
        DI_oDeliveryNotes.Lines.WarehouseCode = oS_PS_SD920L(i).WhsCode
        DI_oDeliveryNotes.Lines.UnitPrice = oS_PS_SD920L(i).Price
        DI_oDeliveryNotes.Lines.LineTotal = oS_PS_SD920L(i).LineTotal
        '9.2 버전에서는 코드 오류 남(이미 마감된 판매오더 번호는 등록 불가한 것 같음, 따라서 주석 처리(2018.03.07 송명규))
'        If oS_PS_SD920L(i).ORDRDoc > 0 Then
'            DI_oDeliveryNotes.Lines.BaseType = 17
'            DI_oDeliveryNotes.Lines.BaseEntry = oS_PS_SD920L(i).ORDRDoc
'            DI_oDeliveryNotes.Lines.BaseLine = oS_PS_SD920L(i).RDR1Line
'        End If
        '9.2 버전에서는 코드 오류 남(이미 마감된 판매오더 번호는 등록 불가한 것 같음, 따라서 주석 처리(2018.03.07 송명규))
        DI_oDeliveryNotes.Lines.UserFields("U_Qty").VALUE = oS_PS_SD920L(i).Qty
        DI_oDeliveryNotes.Lines.UserFields("U_BaseType").VALUE = oS_PS_SD920L(i).BaseType '기준납품문서의 값을 그대로 저장(2016.10.20 송명규 수정)
        DI_oDeliveryNotes.Lines.UserFields("U_BaseEntry").VALUE = oS_PS_SD920L(i).BaseEntry '기준납품문서의 납품처리[SD404]문서번호의 값을 그대로 저장(2016.10.20 송명규 수정)
        DI_oDeliveryNotes.Lines.UserFields("U_BaseLine").VALUE = oS_PS_SD920L(i).BaseLine '기준납품문서의 납품처리[SD404]문서행번호의 값을 그대로 저장(2016.10.20 송명규 수정)
        
        '라인 환율처리_S
        If lclDocCur <> "KRW" Then

            DI_oDeliveryNotes.Lines.Currency = lclDocCur
            DI_oDeliveryNotes.Lines.rate = lclDocRate

        End If
        '라인 환율처리_E

        If BatchYN = "Y" Then
            m = 0
            oRecordSet01.MoveFirst
            Do Until oRecordSet01.EOF
                If m > 0 Then
                    DI_oDeliveryNotes.Lines.BatchNumbers.Add
                End If
                DI_oDeliveryNotes.Lines.BatchNumbers.BatchNumber = Trim(oRecordSet01.Fields("BatchNum").VALUE)
                DI_oDeliveryNotes.Lines.BatchNumbers.Quantity = oS_PS_SD920L(i).OpenQty 'oRecordSet01.Fields("Quantity").VALUE 'oS_PS_SD920L(i).OpenQty 가용재고로 입력되도록 수정 20170307
                'oRecordSet01.Fields("Quantity").VALUE 'oS_PS_SD920L(i).OpenQty 'oRecordSet01.Fields("Quantity").VALUE
                oRecordSet01.MoveNext
                m = m + 1
            Loop
        End If
        '납품_E

Add_Line_Again:
        If i <> lTypeCount - 1 Then
            If oS_PS_SD920L(i).ODLNDocB = oS_PS_SD920L(i + 1).ODLNDocB Then
                i = i + 1
                LineNum(i + 1) = oS_PS_SD920L(i).LineNum - 1

                sQry = "Select Manbtchnum From [OITM] Where ItemCode = '" & Trim(oS_PS_SD920L(i).ItemCode) & "'"
                oRecordSet01.DoQuery sQry
                BatchYN = Trim(oRecordSet01.Fields(0).VALUE)
                If BatchYN = "Y" Then
                    sQry = "Select BatchNum, Quantity From [IBT1] "
                    sQry = sQry & "Where BaseType = '15' And BaseEntry = '" & Trim(oS_PS_SD920L(i).ODLNDocB) & "' And BaseLinNum = '" & Trim(oS_PS_SD920L(i).DLN1LinB) & "'"
                    oRecordSet01.DoQuery sQry
        '            sQry = "EXECUTE [PS_SD920_02] '" & tirm(oS_PS_SD920L(i).ODLNDocB) & "', '" & tirm(oS_PS_SD920L(i).DLN1LinB) & "'"
        '            oRecordSet01.DoQuery sQry
                End If

                '반품
                DI_oReturns.Lines.Add
                DI_oReturns.Lines.ItemCode = oS_PS_SD920L(i).ItemCode
                DI_oReturns.Lines.Quantity = oS_PS_SD920L(i).OpenQty
                DI_oReturns.Lines.WarehouseCode = oS_PS_SD920L(i).WhsCode
                DI_oReturns.Lines.UnitPrice = oS_PS_SD920L(i).Price
                DI_oReturns.Lines.LineTotal = oS_PS_SD920L(i).LineTotal
                DI_oReturns.Lines.BaseType = 15
                DI_oReturns.Lines.BaseEntry = oS_PS_SD920L(i).ODLNDocB
                DI_oReturns.Lines.BaseLine = oS_PS_SD920L(i).DLN1LinB
                DI_oReturns.Lines.UserFields("U_Qty").VALUE = oS_PS_SD920L(i).Qty
                DI_oReturns.Lines.UserFields("U_BaseType").VALUE = oS_PS_SD920L(i).BaseType '기준납품문서의 값을 그대로 저장(2016.10.20 송명규 수정)
                DI_oReturns.Lines.UserFields("U_BaseEntry").VALUE = oS_PS_SD920L(i).BaseEntry '기준납품문서의 납품처리[SD404]문서번호의 값을 그대로 저장(2016.10.20 송명규 수정)
                DI_oReturns.Lines.UserFields("U_BaseLine").VALUE = oS_PS_SD920L(i).BaseLine '기준납품문서의 납품처리[SD404]문서행번호의 값을 그대로 저장(2016.10.20 송명규 수정)

                '환율처리_S
                If lclDocCur <> "KRW" Then

                    DI_oReturns.Lines.Currency = lclDocCur
                    DI_oReturns.Lines.rate = lclDocRate

                End If
                '환율처리_E

                If BatchYN = "Y" Then
                    m = 0
                    Do Until oRecordSet01.EOF
                        If m > 0 Then
                            DI_oReturns.Lines.BatchNumbers.Add
                        End If
                        DI_oReturns.Lines.BatchNumbers.BatchNumber = Trim(oRecordSet01.Fields("BatchNum").VALUE)
                        DI_oReturns.Lines.BatchNumbers.Quantity = oS_PS_SD920L(i).OpenQty 'oRecordSet01.Fields("Quantity").VALUE 'oS_PS_SD920L(i).OpenQty 가용재고로 입력되도록 수정 20170307
                        'oRecordSet01.Fields("Quantity").VALUE 'oS_PS_SD920L(i).OpenQty 'oRecordSet01.Fields("Quantity").VALUE
                        oRecordSet01.MoveNext
                        m = m + 1
                    Loop
                End If

                '납품
                DI_oDeliveryNotes.Lines.Add
                DI_oDeliveryNotes.Lines.ItemCode = oS_PS_SD920L(i).ItemCode
                DI_oDeliveryNotes.Lines.Quantity = oS_PS_SD920L(i).OpenQty
                DI_oDeliveryNotes.Lines.WarehouseCode = oS_PS_SD920L(i).WhsCode
                DI_oDeliveryNotes.Lines.UnitPrice = oS_PS_SD920L(i).Price
                DI_oDeliveryNotes.Lines.LineTotal = oS_PS_SD920L(i).LineTotal
                '9.2 버전에서는 코드 오류 남(이미 마감된 판매오더 번호는 등록 불가한 것 같음, 따라서 주석 처리(2018.03.07 송명규))
'                If oS_PS_SD920L(i).ORDRDoc > 0 Then
'                    DI_oDeliveryNotes.Lines.BaseType = 17
'                    DI_oDeliveryNotes.Lines.BaseEntry = oS_PS_SD920L(i).ORDRDoc
'                    DI_oDeliveryNotes.Lines.BaseLine = oS_PS_SD920L(i).RDR1Line
'                End If
                '9.2 버전에서는 코드 오류 남(이미 마감된 판매오더 번호는 등록 불가한 것 같음, 따라서 주석 처리(2018.03.07 송명규))
                DI_oDeliveryNotes.Lines.UserFields("U_Qty").VALUE = oS_PS_SD920L(i).Qty
                DI_oDeliveryNotes.Lines.UserFields("U_BaseType").VALUE = oS_PS_SD920L(i).BaseType '기준납품문서의 값을 그대로 저장(2016.10.20 송명규 수정)
                DI_oDeliveryNotes.Lines.UserFields("U_BaseEntry").VALUE = oS_PS_SD920L(i).BaseEntry '기준납품문서의 납품처리[SD040]문서번호의 값을 그대로 저장(2016.10.20 송명규 수정)
                DI_oDeliveryNotes.Lines.UserFields("U_BaseLine").VALUE = oS_PS_SD920L(i).BaseLine '기준납품문서의 납품처리[SD040]문서행번호의 값을 그대로 저장(2016.10.20 송명규 수정)
                
                '환율처리_S
                If lclDocCur <> "KRW" Then

                    DI_oDeliveryNotes.Lines.Currency = lclDocCur
                    DI_oDeliveryNotes.Lines.rate = lclDocRate

                End If
                '환율처리_E

                oRecordSet01.MoveFirst
                If BatchYN = "Y" Then
                    m = 0
                    Do Until oRecordSet01.EOF
                        If m > 0 Then
                            DI_oDeliveryNotes.Lines.BatchNumbers.Add
                        End If
                        DI_oDeliveryNotes.Lines.BatchNumbers.BatchNumber = Trim(oRecordSet01.Fields("BatchNum").VALUE)
                        DI_oDeliveryNotes.Lines.BatchNumbers.Quantity = oS_PS_SD920L(i).OpenQty 'oRecordSet01.Fields("Quantity").VALUE 'oS_PS_SD920L(i).OpenQty 가용재고로 입력되도록 수정 20170307
                        'oRecordSet01.Fields("Quantity").VALUE 'oS_PS_SD920L(i).OpenQty 'oRecordSet01.Fields("Quantity").VALUE
                        oRecordSet01.MoveNext
                        m = m + 1
                    Loop
                End If

                GoTo Add_Line_Again
            Else
                GoTo Add_DI_oReturns_oDeliveryNotes
            End If
        Else
            GoTo Add_DI_oReturns_oDeliveryNotes
        End If

Add_DI_oReturns_oDeliveryNotes:
        '반품 완료
        If DI_oReturns Is Nothing Then
        Else
            RetVal = DI_oReturns.Add
            If (0 <> RetVal) Then
                Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
                ErrNum = 1
                GoTo Add_oReturns_oDeliveryNotes_Error
            End If

            Sbo_Company.GetNewObjectCode ORDNDocEntry

            sQry = "EXECUTE [PS_Z_RETU_GI_ZSD920] '" & ORDNDocEntry & "'" ', PT_Error OUTPUT, PT_Error_Message OUTPUT"
            oRecordSet01.DoQuery sQry
        End If
'        Sbo_Company.EndTransaction wf_Commit
'        Sbo_Company.StartTransaction
        '납품 완료
        If DI_oDeliveryNotes Is Nothing Then
        Else
            RetVal = DI_oDeliveryNotes.Add
            If (0 <> RetVal) Then
                Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
                ErrNum = 1
                GoTo Add_oReturns_oDeliveryNotes_Error
            End If

            Sbo_Company.GetNewObjectCode ODLNDocEntry
        End If

        LineCounter = 0
        For j = K To i
            Call oDS_PS_SD920L.setValue("U_ORDNDoc", LineNum(l + 1), ORDNDocEntry)
            Call oDS_PS_SD920L.setValue("U_RDN1Line", LineNum(l + 1), LineCounter)
            Call oDS_PS_SD920L.setValue("U_ODLNDoc", LineNum(l + 1), ODLNDocEntry)
            Call oDS_PS_SD920L.setValue("U_DLN1Line", LineNum(l + 1), LineCounter)
            '납품문서 DLN1 (BaseEntry, BaseLine,BaseType 추가함)
            sQry = "EXECUTE [PS_Z_InsertInfo_ODLN] '" & ODLNDocEntry & "','" & LineCounter & "','" & Trim(oS_PS_SD920L(l).ORDRDoc) & "','" & Trim(oS_PS_SD920L(l).RDR1Line) & "'"
            
            oRecordSet01.DoQuery sQry
            
            LineCounter = LineCounter + 1
            l = l + 1
        Next
        
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & lTypeCount & "건의 납품 반품 문서 생성중...!"

'        Sbo_Company.EndTransaction wf_Commit
'        Sbo_Company.StartTransaction
    Next i

    oMat01.LoadFromDataSource
    ProgBar01.Stop
    If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_Commit

    Set DI_oReturns = Nothing
    Set DI_oDeliveryNotes = Nothing
    Add_oReturns_oDeliveryNotes = True
    Set oRecordSet01 = Nothing
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_oReturns_oDeliveryNotes_Error:
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    Set DI_oReturns = Nothing
    Set DI_oDeliveryNotes = Nothing
    Set oRecordSet01 = Nothing
    If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_RollBack
    Add_oReturns_oDeliveryNotes = False
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "Add_oReturns_oDeliveryNotes_Error:" & ErrCode & " - " & ErrMsg, "E"
    Else
        MDC_Com.MDC_GF_Message "Add_oReturns_oDeliveryNotes_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Function
