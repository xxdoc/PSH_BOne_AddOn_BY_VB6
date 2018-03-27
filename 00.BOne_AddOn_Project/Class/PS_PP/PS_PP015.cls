VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP015"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_PP015.cls
'//  Module         : PP
'//  Description    : 포장생산 작업지시서 발행
'//  FormType       : PS_PP015
'//  Create Date    : 2013.04.9
'//  Modified Date  :
'//  Creator        : N.G.Y
'//  Company        : Poongsan Holdings
'****************************************************************************************************************


Option Explicit

Public oFormUniqueID01 As String
Public oForm01             As SAPbouiCOM.Form
Public oMat01              As SAPbouiCOM.Matrix
Private oDS_PS_PP015H As SAPbouiCOM.DBDataSource    '등록헤더
Private oDS_PS_PP015L As SAPbouiCOM.DBDataSource    '등록라인

Private oLast_Item_UID      As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID       As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row       As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&
Private oSeq&
Private oCount&
Private oPurchase As String
Private oPQType As String


'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm(Optional sDocNum As String, Optional sBPLId As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP015.srf")
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
    
    oFormUniqueID01 = "PS_PP015_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
'    oForm01.Mode = fm_ADD_MODE
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    '화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
    oForm01.DataBrowser.BrowseBy = "DocNum"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    If sDocNum <> "" Then
        oForm01.Mode = fm_FIND_MODE
    Else
        oForm01.Mode = fm_ADD_MODE
    End If
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting
'    Call Initialization
    Call FormClear
    Call Add_MatrixRow(0, True)
        
'    FormItemEnabled
    
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1284"), True         '// 취소
    oForm01.EnableMenu ("1293"), True         '// 행삭제
    
    oForm01.Update
    oForm01.Freeze False
    
    If sDocNum <> "" Then

        Call FormItemEnabled

        oForm01.Items("DocNum").Specific.VALUE = sDocNum
        Call oForm01.Items("BPLId").Specific.Select(sBPLId, psk_ByValue)
        
        Call oForm01.Items("1").Click(ct_Regular)
        
    Else
    
        Call Initialization
        
    End If

    
    
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
    Dim SumWeight As Currency, SumQty As Long
    Dim ItemCode$, ItemName$, Size$, Qty&, Weight As Currency, Unit$, RequestDate$, DueDate$, ItemType$, RequestNo$, BPLId$
    Dim RFC_Sender$
    Dim Calculate_Weight As Double
    Dim Seq&
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE Then
                        If HeaderSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        If MatrixSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        If oForm01.Items("StdPoYN").Specific.VALUE = "Y" Then
                            Call NumberSet
                        Else
                            oForm01.Items("Number").Specific.VALUE = ""
                        End If
                        
                        oMat01.FlushToDataSource

                        oMat01.LoadFromDataSource
                        
                        Call Delete_EmptyRow
                        oLast_Mode = oForm01.Mode
                    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                        oLast_Mode = oForm01.Mode
                        
                        If HeaderSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        If MatrixSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                    ElseIf oForm01.Mode = fm_FIND_MODE Then
                        oLast_Mode = oForm01.Mode
                    End If
                End If
'et_KEY_DOWN ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "CntcCode" Then
'                        If oForm01.Items("CntcCode").Specific.VALUE = "" Then
'                            Sbo_Application.ActivateMenuItem ("7425")
'                            BubbleEvent = False
'                        End If
                    ElseIf pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "ReqNum" Then
                            If oMat01.Columns("ReqNum").Cells(pval.Row).Specific.VALUE = "" Then
                                Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                            End If
                        ElseIf pval.ColUID = "ItemCode" Then
                            If oMat01.Columns("ItemCode").Cells(pval.Row).Specific.VALUE = "" Then
                                Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                            End If
                        ElseIf pval.ColUID = "CardCode" Then
                            If oMat01.Columns("CardCode").Cells(pval.Row).Specific.VALUE = "" Then
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
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    '//종전값을 넣기위해
'                    oPurchase = Trim(oDS_PS_PP015H.GetValue("U_Purchase", 0))
'                    oPQType = Trim(oDS_PS_PP015H.GetValue("U_PQType", 0))
                    
                    If oForm01.Mode = fm_OK_MODE Then
                        If oLast_Mode = fm_UPDATE_MODE Then
                            Add_MatrixRow oMat01.RowCount, False
                            oLast_Mode = 100
                        ElseIf oLast_Mode = fm_FIND_MODE Then
                            Add_MatrixRow oMat01.RowCount, False
                            FormItemEnabled
                            oLast_Mode = 100
                        End If
                    ElseIf oForm01.Mode = fm_ADD_MODE And pval.Action_Success = True Then
                        oForm01.Mode = fm_OK_MODE
                        Call Sbo_Application.ActivateMenuItem("1282")
                    End If
                ElseIf pval.ItemUID = "Btn01" Then
'                    If HeaderSpaceLineDel = False Then
'                        BubbleEvent = False
'                        Exit Sub
'                    End If
                     
                     Call Print_Report01
                End If
            Case et_KEY_DOWN: '//2
'et_COMBO_SELECT ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_COMBO_SELECT: '//5
                If pval.ItemUID = "StdPoYN" Then
                    If oForm01.Items("StdPoYN").Specific.VALUE = "N" Then
                        oForm01.Items("Number").Specific.VALUE = ""
                    Else
                        Call NumberSet
                    End If
                End If
                If pval.ItemUID = "BPLId" Then
                    Call NumberSet
                End If
                
            Case et_CLICK: '//6
            
                If pval.ItemUID = "Mat01" Then
        
                    If pval.Row > 0 Then
                    
                        oLast_Item_UID = pval.ItemUID
                        oLast_Col_UID = pval.ColUID
                        oLast_Col_Row = pval.Row
                        
                        Call oMat01.SelectRow(pval.Row, True, False)
                        
                    End If
                    
                End If
            
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
'et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "DocDate" Then
                        Call NumberSet
                    ElseIf pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "ReqNum" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        ElseIf pval.ColUID = "ItemCode" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        ElseIf pval.ColUID = "CardCode" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        End If
                    End If
                End If
            Case et_MATRIX_LOAD: '//11
                
'                Add_MatrixRow oMat01.RowCount, False
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
                Set oDS_PS_PP015H = Nothing
                Set oDS_PS_PP015L = Nothing
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
    Set ProgressBar01 = Nothing
    oForm01.Freeze False
    If ErrNum = 101 Then
        ErrNum = 0
        MDC_Com.MDC_GF_Message "Raise_ItemEvent_Error:" & Err.Number & " - " & Err.Description, "E"
        BubbleEvent = False
    Else
        MDC_Com.MDC_GF_Message "Raise_ItemEvent_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_MenuEvent_Error
    Dim i&
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
'[1284:취소] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1284": '취소
                FormItemEnabled
                oForm01.Items("DocNum").Click ct_Regular
            Case "1286": '닫기
'[1293:행삭제] //////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1293": '행삭제
                If oMat01.RowCount <> oMat01.VisualRowCount Then
                    For i = 0 To oMat01.VisualRowCount - 1
                        oMat01.Columns("LineNum").Cells(i + 1).Specific.VALUE = i + 1
                    Next i
                    
                    oMat01.FlushToDataSource
                    oDS_PS_PP015L.RemoveRecord oDS_PS_PP015L.Size - 1       '// Mat01에 마지막라인(빈라인) 삭제
                    oMat01.Clear
                    oMat01.LoadFromDataSource
                End If
'[1281:찾기] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1281": '찾기
                oForm01.Freeze True
                Call FormItemEnabled
                oForm01.Items("DocNum").Click ct_Regular
                
                '//아이디별 사업장 세팅
                oForm01.Items("BPLId").Specific.Select MDC_PS_Common.User_BPLId, psk_ByValue
                oForm01.Freeze False
'[1282:추가] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1282": '추가
                oDS_PS_PP015H.setValue "U_DocDate", 0, Format(Now, "YYYYMMDD")
                Call Initialization
                Call FormItemEnabled
                Call FormClear
                
                Call Add_MatrixRow(0, True)
                
'                oForm01.Items("BPLId").Click ct_Collapsed
'[1288~1291:네비게이션] /////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                Call FormItemEnabled
                If oMat01.VisualRowCount > 0 Then
                    If oMat01.Columns("ReqNum").Cells(oMat01.VisualRowCount).Specific.VALUE <> "" Then
                        Add_MatrixRow oMat01.RowCount, False
                    End If
                End If
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
    Set oDS_PS_PP015H = oForm01.DataSources.DBDataSources("@PS_PP015H")
    Set oDS_PS_PP015L = oForm01.DataSources.DBDataSources("@PS_PP015L")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    
    oDS_PS_PP015H.setValue "U_DocDate", 0, Format(Now, "yyyymmdd")
    
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
    sQry = "SELECT BPLId, BPLName From [OBPL] order by 1"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    Set oCombo = oForm01.Items("StdPoYN").Specific
    oCombo.ValidValues.Add "Y", "PO정식발행"
    oCombo.ValidValues.Add "N", "기PO발행분저장"
    oCombo.Select "Y", psk_ByValue
    
    
    '//품목대분류
    sQry = "SELECT Code, Name From [@PSH_ITMBSORT] Order by Code"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oMat01.Columns("ItmBsort").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    '//품목중분류
    sQry = "SELECT U_Code, U_CodeName From [@PSH_ITMMSORT] Order by U_Code"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oMat01.Columns("ItmMsort").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
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
    
    '//아이디별 사업장 세팅
    Set oCombo = oForm01.Items("BPLId").Specific
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
    oForm01.Items("Send").Specific.VALUE = "영업"
    oForm01.Items("Receive").Specific.VALUE = "생산"
    
    Call NumberSet
    
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
    If (oForm01.Mode = fm_ADD_MODE) Then
        oForm01.Items("DocNum").Enabled = False
        
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("DocDate").Enabled = True
                
        oForm01.Items("Mat01").Enabled = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        oForm01.Items("DocNum").Enabled = True
        
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("DocDate").Enabled = True
        
        oForm01.Items("Mat01").Enabled = False
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        oForm01.Items("DocNum").Enabled = False
        
        oForm01.Items("BPLId").Enabled = False
        oForm01.Items("DocDate").Enabled = True
        oForm01.Items("Mat01").Enabled = True
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub FormClear()
On Error GoTo FormClear_Error
    Dim DocNum As String
    DocNum = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP015'", "")
    If DocNum = 0 Then
        oForm01.Items("DocNum").Specific.VALUE = 1
    Else
        oForm01.Items("DocNum").Specific.VALUE = DocNum
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormClear_Error:
    MDC_Com.MDC_GF_Message "FormClear_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_PP015L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_PP015L.Offset = oRow
    oDS_PS_PP015L.setValue "U_LineNum", oRow, oRow + 1
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
    Dim SumWeight As Currency, SumQty As Long
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
        Case "CntcCode"
'            sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" & Trim(oDS_PS_PP015H.GetValue("U_CntcCode", 0)) & "'"
'            oRecordSet01.DoQuery sQry
'
'            oDS_PS_PP015H.setValue "U_CntcName", 0, Trim(oRecordSet01.Fields(0).VALUE)
        Case "Mat01"
            If oCol = "ReqNum" Then
                oForm01.Freeze True
                If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 0) And Trim(oMat01.Columns("ReqNum").Cells(oRow).Specific.VALUE) <> "" Then
                    oMat01.FlushToDataSource
                    Add_MatrixRow oMat01.RowCount, False
                    oMat01.Columns("ReqNum").Cells(oRow).Click ct_Regular
                End If
                sQry = " select OrdrNum = Convert(Nvarchar(10),t.DocNum) + '-' + Convert(Nvarchar(10),t1.LineNum), "
                sQry = sQry + " t1.ItemCode, ItemName = t1.Dscription, t2.U_ItmBsort, t2.U_ItmMsort, "
                sQry = sQry + " t.DocNum, t1.LineNum, t1.Quantity, t.DocDate, DueDate = t1.ShipDate, "
                sQry = sQry + " t.CardCode , t.CardName "
                sQry = sQry + " from ORDR t Inner Join RDR1 t1 On t.DocEntry = t1.DocEntry and t1.LineStatus = 'O' "
                sQry = sQry + " Inner Join OITM t2 On t1.ItemCode = t2.ItemCode "
                sQry = sQry + " where Convert(Nvarchar(10),t.DocNum) + '-' + Convert(Nvarchar(10),t1.LineNum) = '" & Trim(oMat01.Columns("ReqNum").Cells(oRow).Specific.VALUE) & "'"
                
'                sQry = " Select A.U_RegNum, A.U_ItemCode,A.U_ItemName, B.U_ItmBsort, B.U_ItmMsort, c.DocNum, c.LineNum, c.Quantity, c.DocDate, c.DueDate, A.U_CardCode, A.U_CardName "
'                sQry = sQry + " From [@PS_SD010H] AS A Inner Join OITM AS B On A.U_ItemCode = B.ItemCode "
'                sQry = sQry + " Left Join ( Select DocLinNum = Convert(Nvarchar(10),t.DocNum) + '-' + Convert(Nvarchar(10),t1.LineNum), t.DocNum, t1.LineNum, t.DocDate, DueDate = t1.ShipDate, t1.Quantity"
'                sQry = sQry + " From ORDR t Inner Join RDR1 t1 On t.DocEntry = t1.DocEntry ) c On A.U_SjDocLin = c.DocLinNum"
'                sQry = sQry + " Where A.U_Status = 'O' And A.U_RegNum = '" & Trim(oMat01.Columns("ReqNum").Cells(oRow).Specific.VALUE) & "'"

                oRecordSet01.DoQuery sQry
                
                If oRecordSet01.RecordCount = 0 Then
                    MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
                    Set oRecordSet01 = Nothing
                    oForm01.Freeze False
                    Exit Sub
                End If
    
                oMat01.Columns("ItemCode").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("ItemCode").VALUE)
                oMat01.Columns("ItemName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("ItemName").VALUE)
                oMat01.Columns("ItmBsort").Cells(oRow).Specific.Select Trim(oRecordSet01.Fields("U_ItmBsort").VALUE)
                oMat01.Columns("ItmMsort").Cells(oRow).Specific.Select Trim(oRecordSet01.Fields("U_ItmMsort").VALUE)
                
                oMat01.Columns("SjDocNum").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("DocNum").VALUE)
                oMat01.Columns("SjLinNum").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("LineNum").VALUE)
                oMat01.Columns("SjQty").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("Quantity").VALUE)
                oMat01.Columns("DueDate").Cells(oRow).Specific.VALUE = Trim(Format(oRecordSet01.Fields("DueDate").VALUE, "YYYYMMDD"))
                oMat01.Columns("SjDcDate").Cells(oRow).Specific.VALUE = Trim(Format(oRecordSet01.Fields("DocDate").VALUE, "YYYYMMDD"))
                oMat01.Columns("SjDuDate").Cells(oRow).Specific.VALUE = Trim(Format(oRecordSet01.Fields("DueDate").VALUE, "YYYYMMDD"))
                oMat01.Columns("CardCode").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("CardCode").VALUE)
                oMat01.Columns("CardName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("CardName").VALUE)
                
                oMat01.Columns("ReqNum").Cells(oRow).Click ct_Regular
                
                oForm01.Freeze False
            ElseIf oCol = "ItemCode" Then
                oForm01.Freeze True
                If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 0) And Trim(oMat01.Columns("ItemCode").Cells(oRow).Specific.VALUE) <> "" Then
                    oMat01.FlushToDataSource
                    Add_MatrixRow oMat01.RowCount, False
                    oMat01.Columns("ItemCode").Cells(oRow).Click ct_Regular
                End If
                sQry = " select t.ItemCode, t.ItemName, t.U_ItmBsort, t.U_ItmMsort "
                sQry = sQry + " from OITM t "
                sQry = sQry + " where t.ItemCode = '" & Trim(oMat01.Columns("ItemCode").Cells(oRow).Specific.VALUE) & "'"
                
                oRecordSet01.DoQuery sQry
                
                If oRecordSet01.RecordCount = 0 Then
                    MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
                    Set oRecordSet01 = Nothing
                    oForm01.Freeze False
                    Exit Sub
                End If
    
                oMat01.Columns("ItemName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("ItemName").VALUE)
                oMat01.Columns("ItmBsort").Cells(oRow).Specific.Select Trim(oRecordSet01.Fields("U_ItmBsort").VALUE)
                oMat01.Columns("ItmMsort").Cells(oRow).Specific.Select Trim(oRecordSet01.Fields("U_ItmMsort").VALUE)
                
                oMat01.Columns("ItemCode").Cells(oRow).Click ct_Regular
                
                oForm01.Freeze False
            ElseIf oCol = "CardCode" Then
                oForm01.Freeze True
                sQry = "Select CardName from OCRD Where CardCode = '" & Trim(oMat01.Columns("CardCode").Cells(oRow).Specific.VALUE) & "'"
                oRecordSet01.DoQuery sQry
                
                If oRecordSet01.RecordCount = 0 Then
                    MDC_Com.MDC_GF_Message "거래처명 조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
                    oMat01.Columns("CardName").Cells(oRow).Specific.VALUE = ""
                    Set oRecordSet01 = Nothing
                    oForm01.Freeze False
                    Exit Sub
                End If
    
                oMat01.Columns("CardName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("CardName").VALUE)
                
                oMat01.Columns("CardCode").Cells(oRow).Click ct_Regular
                
                oForm01.Freeze False
            End If
    End Select

'    Set oRecordset = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
FlushToItemValue_Error:
    oForm01.Freeze False
    MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum          As Integer
    
    ErrNum = 0

    '// Check
    Select Case True
        Case oDS_PS_PP015H.GetValue("U_BPLId", 0) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_PP015H.GetValue("U_DocDate", 0) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "사업장은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "지시일은 필수사항입니다. 확인하세요.", "E"
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

    ErrNum = 0
    
    oMat01.FlushToDataSource

    '// 라인
    If oMat01.VisualRowCount = 0 Then
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    ElseIf oMat01.VisualRowCount = 1 And oDS_PS_PP015L.GetValue("U_ReqNum", 0) = "" Then
        ErrNum = 2
        GoTo MatrixSpaceLineDel_Error
    End If
    
    
    For i = 0 To oMat01.VisualRowCount - 2
        Select Case True
            Case oDS_PS_PP015L.GetValue("U_ItemCode", i) = ""
                ErrNum = 3
                GoTo MatrixSpaceLineDel_Error
            Case oDS_PS_PP015L.GetValue("U_DueDate", i) = 0
                ErrNum = 4
                GoTo MatrixSpaceLineDel_Error
            Case oDS_PS_PP015L.GetValue("U_CardCode", i) = ""
                ErrNum = 5
                GoTo MatrixSpaceLineDel_Error
        End Select
    Next
    oMat01.LoadFromDataSource

    Set oRecordSet01 = Nothing
    MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
MatrixSpaceLineDel_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Or ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "라인 데이터가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "품목코드는 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "납기일은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "거래처는 필수사항입니다. 확인하세요.", "E"
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
        If Trim(oDS_PS_PP015L.GetValue("U_ItemCode", i)) = "" Then
            oDS_PS_PP015L.RemoveRecord i   '// Mat01에 마지막라인(빈라인) 삭제
        End If
    Next i
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Delete_EmptyRow_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub


Private Sub Print_Report01()
On Error GoTo Print_Report01_Error
    Dim i               As Integer
    Dim ErrNum          As Integer
    Dim WinTitle        As String
    Dim ReportName      As String
    Dim oText(1)        As String
    
    Dim DocNum$, PQType$, Purchase$
    
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call ConnectODBC
    
    DocNum = Trim(oDS_PS_PP015H.GetValue("DocNum", 0))
    
    
    '// Crystal
    
    WinTitle = "[PS_PP015]" & "생산지시서"
    ReportName = "PS_PP015_01.RPT"
    
    '// Formula 수식필드
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    
'    If Purchase = "10" Then
'        oText(1) = "작번 / 품명 / Main 품명"
'    ElseIf Purchase = "20" Then
'        oText(1) = "사용처"
'    End If
'
'    For i = 1 To 1
'        If Len("" & i & "") = 1 Then
'            gRpt_Formula(i) = "F0" & i & ""
'        Else
'            gRpt_Formula(i) = "F" & i & ""
'        End If
'        gRpt_Formula_Value(i) = oText(i)
'    Next i
'
    '// SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)

    '// 조회조건문
    sQry = "EXEC PS_PP015_01 '" & DocNum & "'"
    oRecordSet01.DoQuery sQry
    If oRecordSet01.RecordCount = 0 Then
        ErrNum = 1
        GoTo Print_Report01_Error
    End If

    '// Action
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
    End If
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Report01_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다.확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "Print_Report01_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Private Sub NumberSet()
On Error GoTo NumberSet_Error
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLId As String
    Dim YM As String
    Dim Cnt As String
    
    BPLId = oForm01.Items("BPLId").Specific.VALUE
    YM = Left(oForm01.Items("DocDate").Specific.VALUE, 6)
    
    
    sQry = " Select right('0' + Convert(Nvarchar(2), Convert(integer,Right(Max(Isnull(U_Number,'00')),2)) + 1),2) "
    sQry = sQry + " From [@PS_PP015H] "
    sQry = sQry + " Where U_BPLId = '" & BPLId & "' And Convert(char(6),U_DocDate,112) = '" & YM & "' and Isnull(U_Number,'') <> '' and Canceled = 'N' "
    oRecordSet01.DoQuery sQry
    
    Cnt = oRecordSet01.Fields(0).VALUE
    If Trim(Cnt) = "" Then
        Cnt = "01"
    End If
    oForm01.Items("Number").Specific.VALUE = YM + "-" + Cnt
    
    
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
NumberSet_Error:
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "NumberSet_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub
