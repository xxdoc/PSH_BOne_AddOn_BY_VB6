VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_SD010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_SD010.cls
'//  Module         : SD
'//  Description    : 생산요청
'//  FormType       : PS_SD010
'//  Create Date    : 2010.09.28
'//  Modified Date  :
'//  Creator        : Ryu Yung Jo
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01  As String
Public oForm01          As SAPbouiCOM.Form
Public oMat01           As SAPbouiCOM.Matrix
Private oDS_PS_SD010H   As SAPbouiCOM.DBDataSource    '등록헤더

Private oLast_Item_UID  As String   '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID   As String   '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row   As Long     '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&
Private oLast_RightClick_RegNum&

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i           As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01   As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_SD010.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    
    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_SD010_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    
    oForm01.Freeze True
    Call PS_SD010_CreateItems
    Call ComboBox_Setting
    Call Initialization
    Call RadioButton_Setting
    Call Add_MatrixRow(0, True)
    Call LoadCaption
    Call FormItemEnabled
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1285"), False        '// 복원
    oForm01.EnableMenu ("1284"), True         '// 취소
    oForm01.EnableMenu ("1293"), True         '// 행삭제
    oForm01.EnableMenu ("1299"), True '행닫기
        
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
    
    Dim ChildForm01 As Variant
    Set ChildForm01 = New PS_SM010
    
    Dim ItemCode$, ItemName$, Size$, Qty&, Weight As Currency, Unit$, RequestDate$, DueDate$, ItemType$, RequestNo$
    Dim RFC_Sender$
    Dim Seq&
    
    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "Btn01" Then
                    If oForm01.Mode = fm_ADD_MODE Then
                        If HeaderSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        If MatrixSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
'                        If Add_ProductionDemand(pval) = False Then
'                            BubbleEvent = False
'                            Exit Sub
'                        End If

                        If PS_SD010_AddData = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        oMat01.Clear
                        oMat01.FlushToDataSource
                        oMat01.LoadFromDataSource
                        Add_MatrixRow 0, True
                        
'                        Call Delete_EmptyRow
                        oLast_Mode = oForm01.Mode
                    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                        If MatrixSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
'                        If Update_ProductionDemand(pval) = False Then
'                            BubbleEvent = False
'                            Exit Sub
'                        End If
                        
                        If PS_SD010_AddData = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        oForm01.Mode = fm_OK_MODE
                        Call LoadCaption
                    ElseIf oForm01.Mode = fm_FIND_MODE Then
                        oLast_Mode = oForm01.Mode
                    End If
                ElseIf pval.ItemUID = "Btn02" Then
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    oForm01.Mode = fm_OK_MODE
                    Call LoadCaption
                    Call LoadData
                ElseIf pval.ItemUID = "Btn03" Then
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    Call PS_SD010_Print_Report01
                    
                End If
'et_KEY_DOWN ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "CntcCode" Then
                        If oForm01.Items("CntcCode").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "CardCode" Then
                        If oForm01.Items("CardCode").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "ItemCode" Then
                        If oForm01.Items("ItemCode").Specific.VALUE = "" Then
                            Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row)
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "ItemCode" Then
                            If oMat01.Columns("ItemCode").Cells(pval.Row).Specific.VALUE = "" Then
                                If Trim(oForm01.Items("BPLId").Specific.VALUE) = "" Or Trim(oForm01.Items("ItmBSort").Specific.VALUE) = "" Then
                                    MDC_Com.MDC_GF_Message "사업장 또는 품목대분류를 먼저 선택하세요.", "E"
                                    BubbleEvent = False
                                    Exit Sub
                                Else
                                    Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row, "", Trim(oForm01.Items("ItmBSort").Specific.VALUE))
                                    BubbleEvent = False
                                End If
                            End If
                        ElseIf pval.ColUID = "SjDocLin" Then
                            If oMat01.Columns("SjDocLin").Cells(pval.Row).Specific.VALUE = "" Then
                                If Trim(oForm01.Items("BPLId").Specific.VALUE) = "" Or Trim(oForm01.Items("ItmBSort").Specific.VALUE) = "" Or Trim(oForm01.Items("CardCode").Specific.VALUE) = "" Then
                                    MDC_Com.MDC_GF_Message "사업장 또는 품목대분류 또는 고객코드를 먼저 선택하세요.", "E"
                                    BubbleEvent = False
                                    Exit Sub
                                Else
                                    Sbo_Application.ActivateMenuItem ("7425")
                                    BubbleEvent = False
                                End If
                            End If
                        ElseIf pval.ColUID = "CntcCode" Then
                            If oMat01.Columns("CntcCode").Cells(pval.Row).Specific.VALUE = "" Then
                                Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                            End If
                        ElseIf pval.ColUID = "WhsCode" Then
                            If oMat01.Columns("WhsCode").Cells(pval.Row).Specific.VALUE = "" Then
                                Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                            End If
                        End If
                    End If
                End If
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
            Case et_ITEM_PRESSED: '//1
            Case et_KEY_DOWN: '//2
'et_COMBO_SELECT ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_COMBO_SELECT: '//5
                If pval.ItemUID = "BPLId" Or pval.ItemUID = "ItmBSort" Then
                    oForm01.Freeze True
                    oMat01.Clear
                    oDS_PS_SD010H.Clear
                    oMat01.FlushToDataSource
                    Call Add_MatrixRow(0, True)
                    oForm01.Freeze False
                ElseIf pval.ItemUID = "Mat01" Then
                    If oForm01.Mode = fm_ADD_MODE Then
                    Else
                        oForm01.Mode = fm_UPDATE_MODE
                        Call LoadCaption
                    End If
                End If
'et_CLICK //////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_CLICK: '//6
                If pval.ItemUID = "PRType01" And oForm01.Mode = fm_ADD_MODE Then
                    oForm01.Freeze True
                    oMat01.Clear
                    oMat01.FlushToDataSource
                    oMat01.LoadFromDataSource
                    
                    Call Add_MatrixRow(0, True)
                    oMat01.Columns("SjDocLin").Editable = False
                    oMat01.Columns("ItemCode").Editable = True
                    oForm01.Freeze False
'                    End If
                ElseIf pval.ItemUID = "PRType02" And oForm01.Mode = fm_ADD_MODE Then
                    oForm01.Freeze True
                    oMat01.Clear
                    oMat01.FlushToDataSource
                    oMat01.LoadFromDataSource
                    
                    Call Add_MatrixRow(0, True)
                    oMat01.Columns("SjDocLin").Editable = True
                    oMat01.Columns("ItemCode").Editable = False
                    oForm01.Freeze False
'                    End If
                End If
                
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
'et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "CntcCode" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "CardCode" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "ItemCode" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "SjDocLin" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        ElseIf pval.ColUID = "ItemCode" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        ElseIf pval.ColUID = "ReQty" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        ElseIf pval.ColUID = "CntcCode" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        ElseIf pval.ColUID = "WhsCode" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        End If
                        
                        If oForm01.Mode = fm_ADD_MODE Then
                        Else
                            oForm01.Mode = fm_UPDATE_MODE
                            Call LoadCaption
                        End If
                    End If
                End If
            Case et_MATRIX_LOAD: '//11
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
                Set oDS_PS_SD010H = Nothing
        End Select
    End If
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
    Set ProgressBar01 = Nothing
    Set oRecordSet01 = Nothing
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
    Dim ReturnValue&
    Dim sQry         As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim ErrNum       As String
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
'[1293:행삭제] //////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1293": '행삭제
                If oForm01.Mode <> fm_ADD_MODE Then
                    ReturnValue = Sbo_Application.MessageBox("해당 라인의 생산요청을 삭제합니다. 삭제 후 복원할 수 없습니다. 삭제하시겠습니까?", 1, "&확인", "&취소")
                    Select Case ReturnValue
                        Case 1:
                            '생산의뢰접수여부 체크
                            sQry = "        SELECT  COUNT(*) AS [Cnt]"
                            sQry = sQry & " FROM    [@PS_PP010H]"
                            sQry = sQry & " WHERE   U_RegNum = '" & oLast_RightClick_RegNum & "'"
                            sQry = sQry & "         AND [Status] = 'O'"
                            Call oRecordSet01.DoQuery(sQry)
                        
                            If oRecordSet01.Fields("Cnt").VALUE > 0 Then '생산의뢰접수가 존재하면
                                ErrNum = "1"
                                BubbleEvent = False
                                GoTo Raise_MenuEvent_Error
                            End If
                        
                            sQry = "Delete [@PS_SD010H] Where DocEntry = '" & oLast_RightClick_RegNum & "'"
                            oRecordSet01.DoQuery sQry
                            
                            oLast_RightClick_RegNum = 0
                        Case 2:
                            Sbo_Application.MessageBox ("실행이 취소되었습니다.")
                            BubbleEvent = False
                            Exit Sub
                    End Select
                End If

            Case "1299": '행닫기(취소)
            
                If oForm01.Mode <> fm_ADD_MODE Then
                    
                    '생산의뢰접수여부 체크
                    sQry = "        SELECT  COUNT(*) AS [Cnt]"
                    sQry = sQry & " FROM    [@PS_PP010H]"
                    sQry = sQry & " WHERE   U_RegNum = '" & oLast_RightClick_RegNum & "'"
                    sQry = sQry & "         AND [Status] = 'O'"
                    Call oRecordSet01.DoQuery(sQry)
                    
                    If oRecordSet01.Fields(0).VALUE = 0 Then
                        ReturnValue = Sbo_Application.MessageBox("해당 라인의 생산요청을 닫기(취소)합니다. 진행하시겠습니까?", 1, "&확인", "&취소")
                        Select Case ReturnValue
                            Case 1:
                                sQry = "        UPDATE  [@PS_SD010H] "
                                sQry = sQry & " SET     Status = 'C',"
                                sQry = sQry & "         Canceled = 'Y',"
                                sQry = sQry & "         UpdateDate = GETDATE(),"
                                sQry = sQry & "         UserSign = '" & Sbo_Company.UserSignature & "'"
                                sQry = sQry & " Where   U_RegNum = '" & oLast_RightClick_RegNum & "'"
                                oRecordSet01.DoQuery sQry
    
                                oLast_RightClick_RegNum = 0
                                
                                Call MDC_Com.MDC_GF_Message("해당 생산요청건이 닫기(취소)되었습니다.", "W")
                            Case 2:
                                Sbo_Application.MessageBox ("실행이 취소되었습니다.")
                                BubbleEvent = False
                                Exit Sub
                        End Select
                    Else
                        MDC_Com.MDC_GF_Message "생산의뢰접수된 생산요청입니다. 닫기(취소)할 수 없습니다.", "E"
                        BubbleEvent = False
                        Exit Sub
                    End If
                End If
            
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
                        oMat01.Columns("DocNum").Cells(i + 1).Specific.VALUE = i + 1
                    Next i

                    oMat01.FlushToDataSource
                    oDS_PS_SD010H.RemoveRecord oDS_PS_SD010H.Size - 1       '// Mat01에 마지막라인(빈라인) 삭제
                    oMat01.Clear
                    oMat01.LoadFromDataSource
                    oForm01.Freeze False
                End If
'[1281:찾기] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1281": '찾기
                FormItemEnabled
                oForm01.Mode = fm_OK_MODE
                Call LoadCaption
'                oForm01.Items("OrdType").Click ct_Regular
'[1282:추가] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1282": '추가
                oForm01.Freeze True
                If oForm01.DataSources.UserDataSources("RadioBtn01").VALUE = "10" Then
                    oMat01.Columns("SjDocLin").Editable = False
                    oMat01.Columns("ItemCode").Editable = True
                ElseIf oForm01.DataSources.UserDataSources("RadioBtn01").VALUE = "20" Then
                    oMat01.Columns("SjDocLin").Editable = True
                    oMat01.Columns("ItemCode").Editable = False
                End If
                
                oForm01.Mode = fm_ADD_MODE
                oMat01.Clear
                oMat01.FlushToDataSource
                oMat01.LoadFromDataSource
                Add_MatrixRow 0, True
                FormItemEnabled
                Call LoadCaption
                oForm01.Freeze False
'[1288~1291:네비게이션] /////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                FormItemEnabled
                If oMat01.VisualRowCount > 0 Then
                    If oMat01.Columns("CGNo").Cells(oMat01.VisualRowCount).Specific.VALUE <> "" Then
                        If oDS_PS_SD010H.GetValue("Status", 0) = "O" Then
                            Add_MatrixRow oMat01.RowCount, False
                        End If
                    End If
                End If
  
        End Select
    End If
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_MenuEvent_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = "1" Then
        Call MDC_Com.MDC_GF_Message("생산의뢰접수가 등록된 작번입니다. 삭제할 수 없습니다.", "E")
    Else
        Call MDC_Com.MDC_GF_Message("Raise_MenuEvent_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
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
        If eventInfo.Row > 0 And oForm01.Mode <> fm_ADD_MODE Then
            oLast_RightClick_RegNum = Trim(oMat01.Columns("RegNum").Cells(eventInfo.Row).Specific.VALUE)
        End If
    ElseIf (eventInfo.BeforeAction = False) Then
        '//작업
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD010_CreateItems()
On Error GoTo PS_SD010_CreateItems_Error

    '//디비데이터 소스 개체 할당
    Set oDS_PS_SD010H = oForm01.DataSources.DBDataSources("@PS_SD010H")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    
    '사업장
    Call oForm01.DataSources.UserDataSources.Add("BPLId", dt_SHORT_TEXT, 5)
    Call oForm01.Items("BPLId").Specific.DataBind.SetBound(True, "", "BPLId")
    
    '품목대분류
    Call oForm01.DataSources.UserDataSources.Add("ItmBSort", dt_SHORT_TEXT, 5)
    Call oForm01.Items("ItmBSort").Specific.DataBind.SetBound(True, "", "ItmBSort")
    
    '고객
    Call oForm01.DataSources.UserDataSources.Add("CardCode", dt_SHORT_TEXT, 20)
    Call oForm01.Items("CardCode").Specific.DataBind.SetBound(True, "", "CardCode")
    
    '고객명
    Call oForm01.DataSources.UserDataSources.Add("CardName", dt_SHORT_TEXT, 50)
    Call oForm01.Items("CardName").Specific.DataBind.SetBound(True, "", "CardName")
    
    '담당자
    Call oForm01.DataSources.UserDataSources.Add("CntcCode", dt_SHORT_TEXT, 20)
    Call oForm01.Items("CntcCode").Specific.DataBind.SetBound(True, "", "CntcCode")
    
    '담당자명
    Call oForm01.DataSources.UserDataSources.Add("CntcName", dt_SHORT_TEXT, 50)
    Call oForm01.Items("CntcName").Specific.DataBind.SetBound(True, "", "CntcName")
    
    '요청일자(시작)
    Call oForm01.DataSources.UserDataSources.Add("RegDateFr", dt_DATE)
    Call oForm01.Items("RegDateFr").Specific.DataBind.SetBound(True, "", "RegDateFr")
    
    '요청일자(종료)
    Call oForm01.DataSources.UserDataSources.Add("RegDateTo", dt_DATE)
    Call oForm01.Items("RegDateTo").Specific.DataBind.SetBound(True, "", "RegDateTo")
    
    '요청번호(시작)
    Call oForm01.DataSources.UserDataSources.Add("RegNumFr", dt_SHORT_TEXT, 20)
    Call oForm01.Items("RegNumFr").Specific.DataBind.SetBound(True, "", "RegNumFr")
    
    '요청번호(종료)
    Call oForm01.DataSources.UserDataSources.Add("RegNumTo", dt_SHORT_TEXT, 20)
    Call oForm01.Items("RegNumTo").Specific.DataBind.SetBound(True, "", "RegNumTo")
    
    '품목코드
    Call oForm01.DataSources.UserDataSources.Add("ItemCode", dt_SHORT_TEXT, 20)
    Call oForm01.Items("ItemCode").Specific.DataBind.SetBound(True, "", "ItemCode")
    
    '품목명
    Call oForm01.DataSources.UserDataSources.Add("ItemName", dt_SHORT_TEXT, 50)
    Call oForm01.Items("ItemName").Specific.DataBind.SetBound(True, "", "ItemName")
    
    '상태
    Call oForm01.DataSources.UserDataSources.Add("Status", dt_SHORT_TEXT, 10)
    Call oForm01.Items("Status").Specific.DataBind.SetBound(True, "", "Status")
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_SD010_CreateItems_Error:
    MDC_Com.MDC_GF_Message "PS_SD010_CreateItems_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub ComboBox_Setting()
On Error GoTo ComboBox_Setting_Error
    '//콤보에 기본값설정
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    '// 사업장
    Set oCombo = oForm01.Items("BPLId").Specific
    sQry = "SELECT BPLId, BPLName From [OBPL] order by 1"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oMat01.Columns("BPLId").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    '//품목대분류
    Set oCombo = oForm01.Items("ItmBSort").Specific
    sQry = "SELECT Code, Name From [@PSH_ITMBSORT] Where U_PudYN = 'Y' Order by Code"
    oRecordSet01.DoQuery sQry
'    oCombo.ValidValues.Add "ALL", "ALL"
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oMat01.Columns("ItmBSort").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
        
    '//상태
    Set oCombo = oForm01.Items("Status").Specific
    oCombo.ValidValues.Add "O", "열기"
    oCombo.ValidValues.Add "C", "닫기"
    
    '//요청구분
    oMat01.Columns("ReGbn").ValidValues.Add "10", "계획생산요청"
    oMat01.Columns("ReGbn").ValidValues.Add "20", "수주생산요청"
    
    '//품목타입
'    Set oCombo = oForm01.Items("ItemType").Specific
'    sQry = "SELECT Code, Name From [@PSH_SHAPE] Order by Code"
'    oRecordset01.DoQuery sQry
'    oCombo.ValidValues.Add "ALL", "ALL"
'    Do Until oRecordset01.EOF
'        oCombo.ValidValues.Add Trim(oRecordset01.Fields(0).Value), Trim(oRecordset01.Fields(1).Value)
'        oRecordset01.MoveNext
'    Loop
    
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
    oForm01.Items("CntcCode").Specific.VALUE = MDC_PS_Common.User_MSTCOD
    
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

Sub RadioButton_Setting()
On Error GoTo RadioButton_Setting_Error
    '//라디오에 기본값설정
    Dim optBtn          As SAPbouiCOM.OptionBtn

    '//라디오버튼
    oForm01.DataSources.UserDataSources.Add "RadioBtn01", dt_SHORT_TEXT, 2
    Set optBtn = oForm01.Items("PRType01").Specific
    optBtn.ValOn = "10"
    optBtn.ValOff = "0"
    optBtn.DataBind.SetBound True, , "RadioBtn01"
    optBtn.Selected = True
    
    Set optBtn = oForm01.Items("PRType02").Specific
    optBtn.ValOn = "20"
    optBtn.ValOff = "0"
    optBtn.DataBind.SetBound True, , "RadioBtn01"
    optBtn.GroupWith ("PRType01")
    
    oMat01.Columns("SjDocLin").Editable = False
    oMat01.Columns("ItemCode").Editable = True
        
    Set optBtn = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
RadioButton_Setting_Error:
    Set optBtn = Nothing
    MDC_Com.MDC_GF_Message "RadioButton_Setting_Error:" & Err.Number & " - " & Err.Description, "E"
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
        oForm01.Items("CardCode").Click ct_Regular
        oForm01.Items("PRType01").Enabled = True
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("ItmBSort").Enabled = True
        oForm01.Items("CardCode").Enabled = True
        oForm01.Items("CntcCode").Enabled = True
        
        oForm01.Items("RegDateFr").Enabled = False
        oForm01.Items("RegDateTo").Enabled = False
        oForm01.Items("RegNumFr").Enabled = False
        oForm01.Items("RegNumTo").Enabled = False
        oForm01.Items("ItemCode").Enabled = False
        oForm01.Items("Status").Enabled = False
        oForm01.Items("Btn02").Enabled = False
        
        oMat01.Columns("ItemName").Editable = False
        oMat01.Columns("RegNum").Editable = False
'        oMat01.Columns("SjQty").Editable = False
        oMat01.Columns("SjWeight").Editable = False
        oMat01.Columns("UnWeight").Editable = False
        oMat01.Columns("Price").Editable = False
        oMat01.Columns("LinTotal").Editable = False
'        oMat01.Columns("Quality").Editable = False
        oMat01.Columns("BPLId").Editable = False
        oMat01.Columns("CntcCode").Editable = False
        oMat01.Columns("CntcName").Editable = False
        
        oMat01.Columns("ReGbn").Editable = False
        oMat01.Columns("CardCode").Editable = False
        oMat01.Columns("CardName").Editable = False
        
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        oForm01.Items("CardCode").Click ct_Regular
        oForm01.Items("PRType01").Enabled = True
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("ItmBSort").Enabled = True
        oForm01.Items("CardCode").Enabled = True
        oForm01.Items("CntcCode").Enabled = True
        
        oForm01.Items("RegDateFr").Enabled = True
        oForm01.Items("RegDateTo").Enabled = True
        oForm01.Items("RegNumFr").Enabled = True
        oForm01.Items("RegNumTo").Enabled = True
        oForm01.Items("ItemCode").Enabled = True
        oForm01.Items("Status").Enabled = True
        oForm01.Items("Btn02").Enabled = True
        
        oMat01.Columns("ItemName").Editable = False
        oMat01.Columns("RegNum").Editable = False
'        oMat01.Columns("SjQty").Editable = False
        oMat01.Columns("SjWeight").Editable = False
        oMat01.Columns("UnWeight").Editable = False
        oMat01.Columns("Price").Editable = False
        oMat01.Columns("LinTotal").Editable = False
'        oMat01.Columns("Quality").Editable = False
        oMat01.Columns("BPLId").Editable = False
        oMat01.Columns("CntcCode").Editable = True
        oMat01.Columns("CntcName").Editable = False
        
        oMat01.Columns("ReGbn").Editable = False
        oMat01.Columns("CardCode").Editable = False
        oMat01.Columns("CardName").Editable = False
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_SD010H.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_SD010H.Offset = oRow
    oDS_PS_SD010H.setValue "DocNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error
    Dim i&
    Dim ErrNum              As Integer
    Dim sQry                As String
    Dim Calculate_Weight    As Double
    Dim oRecordSet01        As SAPbobsCOM.Recordset
    Dim ItemCode$, Qty&
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
        Case "CntcCode"
            sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" & Trim(oForm01.Items("CntcCode").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            oForm01.Items("CntcName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
        Case "CardCode"
            sQry = "Select CardName From OCRD Where CardCode = '" & Trim(oForm01.Items("CardCode").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            oForm01.Items("CardName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
        Case "ItemCode"
            sQry = "Select ItemName From OITM Where ItemCode = '" & Trim(oForm01.Items("ItemCode").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            oForm01.Items("ItemName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
        Case "Mat01"
            If oCol = "SjDocLin" Then
                Dim SjDocLin$
                oForm01.Freeze True
                If oForm01.Mode = fm_ADD_MODE Then
                    If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 0) And Trim(oMat01.Columns("SjDocLin").Cells(oRow).Specific.VALUE) <> "" Then
                        oMat01.FlushToDataSource
                        Add_MatrixRow oMat01.RowCount, False
                    End If
                End If
                
                SjDocLin = Trim(oMat01.Columns("SjDocLin").Cells(oRow).Specific.VALUE)
                sQry = "Select b.ItemCode, b.Dscription, c.U_ItmBSort, b.U_Qty, b.Quantity, b.U_UnWeight, b.Price, b.LineTotal, c.U_Size, "
                sQry = sQry & "c.U_Spec3, c.U_Quality, a.BPLId, b.WhsCode, d.WhsName, a.CardCode, a.CardName, b.ShipDate, c.InvntryUom "
                sQry = sQry & "From [ORDR] a Inner Join [RDR1] b On a.DocEntry = b.DocEntry "
                sQry = sQry & "Inner Join [OITM] c On c.ItemCode = b.ItemCode Inner Join [OWHS] d On d.WhsCode = b.WhsCode "
                sQry = sQry & "Where a.DocNum = Left('" & SjDocLin & "', CharIndex('-', '" & SjDocLin & "') - 1) And b.U_LineNum = Right('" & SjDocLin & "', Len('" & SjDocLin & "') - CharIndex('-', '" & SjDocLin & "'))"
                oRecordSet01.DoQuery sQry
                                
                oMat01.Columns("ItemCode").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("ItemCode").VALUE)
                oMat01.Columns("ItemName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("Dscription").VALUE)
                oMat01.Columns("RegDate").Cells(oRow).Specific.VALUE = Format(Date, "YYYYMMDD")
                oMat01.Columns("DueDate").Cells(oRow).Specific.VALUE = Format(Trim(oRecordSet01.Fields("ShipDate").VALUE), "YYYYMMDD")
'                oMat01.Columns("ItmBSort").Cells(oRow).Specific.Select Trim(oRecordset01.Fields("U_ItmBSort").Value), psk_ByValue
'                oMat01.Columns("SjQty").Cells(oRow).Specific.Value = Trim(oRecordSet01.Fields("U_Qty").Value)
                oMat01.Columns("SjWeight").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("Quantity").VALUE)
                oMat01.Columns("SjUnit").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("InvntryUom").VALUE)
                oMat01.Columns("UnWeight").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_UnWeight").VALUE)
                oMat01.Columns("Price").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("Price").VALUE)
                oMat01.Columns("LinTotal").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("LineTotal").VALUE)
'                oMat01.Columns("Size").Cells(oRow).Specific.Value = Trim(oRecordSet01.Fields("U_Size").Value)
'                oMat01.Columns("Spec3").Cells(oRow).Specific.Value = Trim(oRecordSet01.Fields("U_Spec3").Value)
'                oMat01.Columns("Quality").Cells(oRow).Specific.Value = Trim(oRecordSet01.Fields("U_Quality").Value)
                oMat01.Columns("BPLId").Cells(oRow).Specific.Select Trim(oRecordSet01.Fields("BPLId").VALUE), psk_ByValue
'                oMat01.Columns("ReQty").Cells(oRow).Specific.Value = Trim(oRecordSet01.Fields("U_Qty").Value)
                oMat01.Columns("ReWeight").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("Quantity").VALUE)
                oMat01.Columns("Unit").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("InvntryUom").VALUE)
'                oMat01.Columns("ReQty").Cells(oRow).Specific.Select Trim(oRecordSet01.Fields("U_Qty").Value), psk_ByValue
'                oMat01.Columns("WhsCode").Cells(oRow).Specific.Value = Trim(oRecordSet01.Fields("WhsCode").Value)
'                oMat01.Columns("WhsName").Cells(oRow).Specific.Value = Trim(oRecordSet01.Fields("WhsName").Value)
                oMat01.Columns("CardCode").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("CardCode").VALUE)
                oMat01.Columns("CardName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("CardName").VALUE)
                
                oMat01.Columns("DueDate").Cells(oRow).Click ct_Regular
                oForm01.Freeze False
            ElseIf oCol = "ItemCode" Then
                oForm01.Freeze True
                If oForm01.Mode = fm_ADD_MODE Then
                    If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 0) And Trim(oMat01.Columns("ItemCode").Cells(oRow).Specific.VALUE) <> "" Then
                        oMat01.FlushToDataSource
                        Add_MatrixRow oMat01.RowCount, False
                        
                        oMat01.Columns("RegDate").Cells(oRow).Specific.VALUE = Format(Date, "YYYYMMDD")
                        oMat01.Columns("DueDate").Cells(oRow).Specific.VALUE = Format(Date, "YYYYMMDD")
'                        oMat01.Columns("Auto").Cells(oRow).Specific.Select "N"
'                        oMat01.Columns("QCYN").Cells(oRow).Specific.Select "N"
'                        oMat01.Columns("ReType").Cells(oRow).Specific.Select "10"
'                        oMat01.Columns("OKYN").Cells(oRow).Specific.Select "N"
                    End If
                End If
                
                sQry = "Select ItemName, U_UnWeight, InvntryUom From OITM Where ItemCode = '" & Trim(oMat01.Columns("ItemCode").Cells(oRow).Specific.VALUE) & "'"
                oRecordSet01.DoQuery sQry
                oMat01.Columns("ItemName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                oMat01.Columns("UnWeight").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_UnWeight").VALUE)
                oMat01.Columns("SjUnit").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("InvntryUom").VALUE)
                oMat01.Columns("Unit").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("InvntryUom").VALUE)
'                oMat01.Columns("ItmBSort").Cells(oRow).Specific.Select Trim(oRecordset01.Fields(1).Value), psk_ByValue
                                
                oMat01.Columns("ItemCode").Cells(oRow).Click ct_Regular
                oForm01.Freeze False
            ElseIf oCol = "ReQty" Then
                oForm01.Freeze True
                oMat01.FlushToDataSource
'                ItemCode = Trim(oDS_PS_SD010H.GetValue("U_ItemCode", oRow - 1))
'                Qty = oDS_PS_SD010H.GetValue("U_ReQty", oRow - 1)
'
'                Calculate_Weight = MDC_PS_Common.Calculate_Weight(ItemCode, Qty, Trim(oForm01.Items("BPLId").Specific.Value))
'
''                            oMat01.Columns("Weight").Cells(pval.Row).Specific.Value = Calculate_Weight
'                oDS_PS_SD010H.setValue "U_ReWeight", oRow - 1, Calculate_Weight
'                oMat01.LoadFromDataSource
                If oDS_PS_SD010H.GetValue("U_ReQty", oRow - 1) <> "" Then
                    Qty = oDS_PS_SD010H.GetValue("U_ReQty", oRow - 1)
                    oDS_PS_SD010H.setValue "U_ReWeight", oRow - 1, Qty
                Else
                    oDS_PS_SD010H.setValue "U_ReWeight", oRow - 1, 0
                End If
                oMat01.LoadFromDataSource
                
                oMat01.Columns("ReQty").Cells(oRow).Click
                oForm01.Freeze False
            ElseIf oCol = "CntcCode" Then
                sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" & Trim(oMat01.Columns("CntcCode").Cells(oRow).Specific.VALUE) & "'"
                oRecordSet01.DoQuery sQry
            
                oMat01.Columns("CntcName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            ElseIf oCol = "WhsCode" Then
                sQry = "Select WhsName From [OWHS] Where WhsCode = '" & Trim(oMat01.Columns("WhsCode").Cells(oRow).Specific.VALUE) & "'"
                oRecordSet01.DoQuery sQry
            
                oMat01.Columns("WhsName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
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
        Case Trim(oForm01.Items("BPLId").Specific.VALUE) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ItmBSort").Specific.VALUE) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
    End Select
    
    If oForm01.Mode = fm_ADD_MODE Then
        If Trim(oForm01.Items("CntcCode").Specific.VALUE) = "" Then
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
        ElseIf oForm01.DataSources.UserDataSources("RadioBtn01").VALUE = "20" Then
            If Trim(oForm01.Items("CardCode").Specific.VALUE) = "" Then
                ErrNum = 4
                GoTo HeaderSpaceLineDel_Error
            End If
        End If
    End If

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "사업장은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "품목대분류는 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "담당자는 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "수주생산요청시 고객은 필수사항입니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function

Private Function MatrixSpaceLineDel() As Boolean
On Error GoTo MatrixSpaceLineDel_Error
    Dim i               As Long, j&
    Dim ErrNum          As Integer
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    Dim sQry            As String

    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    ErrNum = 0

    oMat01.FlushToDataSource

    '// 라인
    If oMat01.VisualRowCount = 0 Then
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    End If
    
    If oMat01.VisualRowCount = 1 And Trim(oDS_PS_SD010H.GetValue("U_ItemCode", 0)) = "" Then
        j = 0
    Else
        j = oMat01.VisualRowCount
    End If
    
    For i = 0 To j - 2
        Select Case True
            Case oDS_PS_SD010H.GetValue("U_ItemCode", i) = ""
                ErrNum = 2
                GoTo MatrixSpaceLineDel_Error
            Case oDS_PS_SD010H.GetValue("U_ReWeight", i) = 0
                ErrNum = 3
                GoTo MatrixSpaceLineDel_Error
            Case oDS_PS_SD010H.GetValue("U_RegDate", i) = ""
                ErrNum = 4
                GoTo MatrixSpaceLineDel_Error
            Case oDS_PS_SD010H.GetValue("U_DueDate", i) = ""
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
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "라인 데이터가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 품목코드가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 생산요청수량이 0입니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 생산요청일자가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 납기일자가 없습니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function

Sub Delete_EmptyRow()
On Error GoTo Delete_EmptyRow_Error
    Dim i&
    
'    oMat01.FlushToDataSource
'
'    For i = 0 To oMat01.VisualRowCount - 1
'        If Trim(oDS_PS_MM010L.GetValue("U_CGNo", i)) = "" Then
'            oDS_PS_MM010L.RemoveRecord i   '// Mat01에 마지막라인(빈라인) 삭제
'        End If
'    Next i
'
'    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Delete_EmptyRow_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub LoadCaption()
On Error GoTo LoadCaption_Error
    If oForm01.Mode = fm_ADD_MODE Then
        oForm01.Items("Btn01").Specific.Caption = "추가"
    ElseIf oForm01.Mode = fm_OK_MODE Then
        oForm01.Items("Btn01").Specific.Caption = "확인"
    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
        oForm01.Items("Btn01").Specific.Caption = "갱신"
    End If

    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadCaption_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Function PS_SD010_AddData()
'******************************************************************************
'Function ID : PS_SD010_AddData()
'해당모듈    : PS_SD010
'기능        : 데이터 INSERT, UPDATE(기존 데이터가 존재하면 UPDATE, 아니면 INSERT)
'인수        : 없음
'반환값      : 없음
'특이사항    : 선행프로세스 일자 체크
'******************************************************************************
On Error GoTo PS_SD010_AddData_Error

    Dim loopCount As Integer
    Dim ErrNum    As String
    Dim ErrOrdNum As String '선행프로세스보다 일자가 빨라서 저장되지 않는 작번을 저장
    
    Dim Query01   As String
    Dim Query02   As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim RecordSet02 As SAPbobsCOM.Recordset
    Set RecordSet02 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim SjDocLin   As String '수주문서-행번호
    Dim ItemCode   As String '품목코드
    Dim ItemName   As String '품목명
    Dim RegDate    As String '요청일자
    Dim DueDate    As String '납기
    Dim RegNum     As String '요청번호
    Dim ItmBsort   As String '품목대분류
    Dim SjWeight   As Currency '수주중량
    Dim SjUnit     As String '수주단위
    Dim UnWeight   As Currency '단위중량
    Dim Price      As Currency '단가
    Dim LinTotal   As Currency '금액
    Dim BPLID      As String '사업장코드
    Dim CntcCode   As String '담당자사번
    Dim CntcName   As String '담당자성명
    Dim ReWeight   As Currency '생산요청수량
    Dim Unit       As String '단위
    Dim Comments   As String '비고
    Dim ReGbn      As String '요청구분(계획생산요청, 수주생산요청)
    Dim CardCode   As String '고객
    Dim CardName   As String '고객명
    Dim Status     As String '상태
    
    Dim BaseEntry  As String
    Dim BaseLine   As String
    Dim DocType    As String
    Dim CurDocDate As String
    
    Dim Entry()    As String
    Dim MinusNum   As Integer '화면모드에 따라 VisualRowCount에서 빼줄 수
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("저장 중...", oMat01.VisualRowCount - 1, False)

    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
    ReGbn = Trim(oForm01.DataSources.UserDataSources("RadioBtn01").VALUE)
    ItmBsort = Trim(oForm01.Items("ItmBSort").Specific.VALUE)
    CntcCode = Trim(oForm01.Items("CntcCode").Specific.VALUE)
    CntcName = Trim(oForm01.Items("CntcName").Specific.VALUE)
    CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
    CardName = Trim(oForm01.Items("CardName").Specific.VALUE)

    If oForm01.Mode = fm_ADD_MODE Then
        MinusNum = 2
    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
        MinusNum = 1
    End If

    Call oMat01.FlushToDataSource
    For loopCount = 0 To oMat01.VisualRowCount - MinusNum
    
        SjDocLin = Trim(oDS_PS_SD010H.GetValue("U_SjDocLin", loopCount))
        ItemCode = Trim(oDS_PS_SD010H.GetValue("U_ItemCode", loopCount))
        ItemName = Trim(oDS_PS_SD010H.GetValue("U_ItemName", loopCount))
        RegDate = Trim(oDS_PS_SD010H.GetValue("U_RegDate", loopCount))
        DueDate = Trim(oDS_PS_SD010H.GetValue("U_DueDate", loopCount))
        RegNum = Trim(oDS_PS_SD010H.GetValue("U_RegNum", loopCount))
        SjWeight = Trim(oDS_PS_SD010H.GetValue("U_SjWeight", loopCount))
        SjUnit = Trim(oDS_PS_SD010H.GetValue("U_SjUnit", loopCount))
        UnWeight = Trim(oDS_PS_SD010H.GetValue("U_UnWeight", loopCount))
        Price = Trim(oDS_PS_SD010H.GetValue("U_Price", loopCount))
        LinTotal = Trim(oDS_PS_SD010H.GetValue("U_LinTotal", loopCount))
        ReWeight = Trim(oDS_PS_SD010H.GetValue("U_ReWeight", loopCount))
        Unit = Trim(oDS_PS_SD010H.GetValue("U_Unit", loopCount))
        Comments = Trim(oDS_PS_SD010H.GetValue("U_Comments", loopCount))
        Status = "O"
        
        Query01 = "         EXEC PS_SD010_03 '"
        Query01 = Query01 & SjDocLin & "','" '수주문서-행번호
        Query01 = Query01 & ItemCode & "','" '품목코드
        Query01 = Query01 & ItemName & "','" '품목명
        Query01 = Query01 & RegDate & "','" '요청일자
        Query01 = Query01 & DueDate & "','" '납기
        Query01 = Query01 & RegNum & "','" '요청번호
        Query01 = Query01 & ItmBsort & "','" '품목대분류
        Query01 = Query01 & SjWeight & "','" '수주중량
        Query01 = Query01 & SjUnit & "','" '수주단위
        Query01 = Query01 & UnWeight & "','" '단위중량
        Query01 = Query01 & Price & "','" '단가
        Query01 = Query01 & LinTotal & "','" '금액
        Query01 = Query01 & BPLID & "','" '사업장코드
        Query01 = Query01 & CntcCode & "','" '담당자사번
        Query01 = Query01 & CntcName & "','" '담당자성명
        Query01 = Query01 & ReWeight & "','" '생산요청수량
        Query01 = Query01 & Unit & "','" '단위
        Query01 = Query01 & Comments & "','" '비고
        Query01 = Query01 & ReGbn & "','" '요청구분(계획생산요청, 수주생산요청)
        Query01 = Query01 & CardCode & "','" '고객
        Query01 = Query01 & CardName & "','" '고객명
        Query01 = Query01 & Status & "'" '상태
        
        If ReGbn = "20" Then '수주생산요청일 때만 선행프로세스 대비 일자 체크
        
            '선행프로세스 대비 일자체크_S
            Entry = Split(SjDocLin, "-")
            BaseEntry = Entry(0)
            BaseLine = Entry(1)
            DocType = "PS_SD010"
            CurDocDate = RegDate
         
            Query02 = "         EXEC PS_Z_CHECK_DATE '"
            Query02 = Query02 & BaseEntry & "','"
            Query02 = Query02 & BaseLine & "','"
            Query02 = Query02 & DocType & "','"
            Query02 = Query02 & CurDocDate & "'"
        
            Call RecordSet02.DoQuery(Query02)
            '선행프로세스 대비 일자체크_E
        
            If RecordSet02.Fields("ReturnValue").VALUE = "True" Then
             
                Call RecordSet01.DoQuery(Query01) '등록
             
            Else
             
                ErrOrdNum = ErrOrdNum & " [" & ItemCode & "]"
             
            End If
            
        Else '계획생산요청일 때는 수주가 없기때문에 선행프로세스 대비 일자 체크 없이 등록
        
            Call RecordSet01.DoQuery(Query01) '등록
            
        End If
        
    Next

    If ErrOrdNum <> "" Then '하나라도 선행프로세스 일자가 빠른 작번이 있으면
        ErrNum = "1"
        GoTo PS_SD010_AddData_Error
    End If

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Call MDC_Com.MDC_GF_Message("전체 저장 완료!", "S")
    
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    
    PS_SD010_AddData = True

    Exit Function

PS_SD010_AddData_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    PS_SD010_AddData = False
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    
    If ErrNum = "1" Then
    
        Call Sbo_Application.MessageBox("생산요청일은 수주일과 같거나 늦어야합니다. 확인하십시오." & Chr(13) & ErrOrdNum, 1)
        
        '등록되지 않은 작번이 있어도 화면 Clear_S
        Call oMat01.Clear
        Call oMat01.FlushToDataSource
        Call oMat01.LoadFromDataSource
        Call Add_MatrixRow(0, True)
        '등록되지 않은 작번이 있어도 화면 Clear_E
        
    Else
    
        Call MDC_Com.MDC_GF_Message("PS_SD010_AddData_Error:" & Err.Number & " - " & Err.Description, "E")
    
    End If
    
End Function

Public Function Add_ProductionDemand(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
'******************************************************************************
'Function ID : Add_ProductionDemand()
'해당모듈    : PS_SD010
'기능        : 데이터 INSERT
'인수        : 없음
'반환값      : 없음
'특이사항    : 2017.02.17 패치 이후 사용 안함. PS_SD010_AddDate 로 대체 (2017.02.27 송명규)
'******************************************************************************
On Error GoTo Add_ProductionDemand_Error

    Dim i           As Integer
    Dim sQry        As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry$, ItemCode$, ItemName$, SjQty&, SjWeight As Currency, ItmBsort$, Price As Currency, LinTotal As Currency, BPLID$
    Dim RegDate$, CntcCode$, CntcName$, UnWeight As Currency, Size$, Spec3$, Quality$, ReQty&, ReWeight As Currency, DueDate$
    Dim WhsCode$, WhsName$, Comments$, SjDocLin$, ReGbn$, CardCode$, CardName$, SjUnit$, Unit$
    
    oMat01.FlushToDataSource
    
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
    ReGbn = Trim(oForm01.DataSources.UserDataSources("RadioBtn01").VALUE)
    ItmBsort = Trim(oForm01.Items("ItmBSort").Specific.VALUE)
    CntcCode = Trim(oForm01.Items("CntcCode").Specific.VALUE)
    CntcName = Trim(oForm01.Items("CntcName").Specific.VALUE)
    CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
    CardName = Trim(oForm01.Items("CardName").Specific.VALUE)
    
    For i = 0 To oMat01.RowCount - 2
        RegDate = Trim(oDS_PS_SD010H.GetValue("U_RegDate", i))
        DueDate = Trim(oDS_PS_SD010H.GetValue("U_DueDate", i))
        
        sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_SD010H] where Left(DocEntry, 6) = Left('" & RegDate & "', 6)"
        'sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_SD010H] where Left(Convert(char, U_RegDate, 112), 6) = Left('" & RegDate & "', 6)"
        RecordSet01.DoQuery sQry
        If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
            DocEntry = Left(RegDate, 6) + "0001"
        Else
            DocEntry = Trim(RecordSet01.Fields(0).VALUE) + 1
        End If
        ItemCode = Trim(oDS_PS_SD010H.GetValue("U_ItemCode", i))
        ItemName = MDC_PS_Common.Make_ItemName(Trim(oDS_PS_SD010H.GetValue("U_ItemName", i)))
'        If Trim(oDS_PS_SD010H.GetValue("U_SjQty", i)) = "" Then
'            SjQty = 0
'        Else
'            SjQty = Trim(oDS_PS_SD010H.GetValue("U_SjQty", i))
'        End If
        SjWeight = Trim(oDS_PS_SD010H.GetValue("U_SjWeight", i))
        SjUnit = Trim(oDS_PS_SD010H.GetValue("U_SjUnit", i))
        UnWeight = Trim(oDS_PS_SD010H.GetValue("U_UnWeight", i))
        Price = Trim(oDS_PS_SD010H.GetValue("U_Price", i))
        LinTotal = Trim(oDS_PS_SD010H.GetValue("U_LinTotal", i))
'        Size = Trim(oDS_PS_SD010H.GetValue("U_Size", i))
'        Spec3 = Trim(oDS_PS_SD010H.GetValue("U_Spec3", i))
'        Quality = Trim(oDS_PS_SD010H.GetValue("U_Quality", i))
'        ReQty = Trim(oDS_PS_SD010H.GetValue("U_ReQty", i))
        ReWeight = Trim(oDS_PS_SD010H.GetValue("U_ReWeight", i))
        Unit = Trim(oDS_PS_SD010H.GetValue("U_Unit", i))
'        WhsCode = Trim(oDS_PS_SD010H.GetValue("U_WhsCode", i))
'        WhsName = Trim(oDS_PS_SD010H.GetValue("U_WhsName", i))
        Comments = Trim(oDS_PS_SD010H.GetValue("U_Comments", i))
        SjDocLin = Trim(oDS_PS_SD010H.GetValue("U_SjDocLin", i))
        
        sQry = "INSERT INTO [@PS_SD010H]"
        sQry = sQry & " ("
        sQry = sQry & " DocEntry,"
        sQry = sQry & " DocNum,"
        sQry = sQry & " U_SjDocLin,"
        sQry = sQry & " U_ItemCode,"
        sQry = sQry & " U_ItemName,"
        sQry = sQry & " U_RegDate,"
        sQry = sQry & " U_DueDate,"
        sQry = sQry & " U_RegNum,"
        sQry = sQry & " U_ItmBSort,"
'        sQry = sQry & " U_SjQty,"
        sQry = sQry & " U_SjWeight,"
        sQry = sQry & " U_SjUnit,"
        sQry = sQry & " U_UnWeight,"
        sQry = sQry & " U_Price,"
        sQry = sQry & " U_LinTotal,"
'        sQry = sQry & " U_Size,"
'        sQry = sQry & " U_Spec3,"
'        sQry = sQry & " U_Quality,"
        sQry = sQry & " U_BPLId,"
        sQry = sQry & " U_CntcCode,"
        sQry = sQry & " U_CntcName,"
'        sQry = sQry & " U_ReQty,"
        sQry = sQry & " U_ReWeight,"
        sQry = sQry & " U_Unit,"
'        sQry = sQry & " U_WhsCode,"
'        sQry = sQry & " U_WhsName,"
        sQry = sQry & " U_Comments,"
        sQry = sQry & " U_ReGbn,"
        sQry = sQry & " U_CardCode,"
        sQry = sQry & " U_CardName,"
        sQry = sQry & " U_Status,"
        sQry = sQry & " CreateDate"
        sQry = sQry & " ) "
        sQry = sQry & "VALUES("
        sQry = sQry & DocEntry & ","
        sQry = sQry & DocEntry & ","
        sQry = sQry & "'" & SjDocLin & "',"
        sQry = sQry & "'" & ItemCode & "',"
        sQry = sQry & "'" & ItemName & "',"
        sQry = sQry & "'" & RegDate & "',"
        sQry = sQry & "'" & DueDate & "',"
        sQry = sQry & "'" & DocEntry & "',"
        sQry = sQry & "'" & ItmBsort & "',"
'        sQry = sQry & "'" & SjQty & "',"
        sQry = sQry & "'" & SjWeight & "',"
        sQry = sQry & "'" & SjUnit & "',"
        sQry = sQry & "'" & UnWeight & "',"
        sQry = sQry & "'" & Price & "',"
        sQry = sQry & "'" & LinTotal & "',"
'        sQry = sQry & "'" & Size & "',"
'        sQry = sQry & "'" & Spec3 & "',"
'        sQry = sQry & "'" & Quality & "',"
        sQry = sQry & "'" & BPLID & "',"
        sQry = sQry & "'" & CntcCode & "',"
        sQry = sQry & "'" & CntcName & "',"
'        sQry = sQry & "'" & ReQty & "',"
        sQry = sQry & "'" & ReWeight & "',"
        sQry = sQry & "'" & Unit & "',"
'        sQry = sQry & "'" & WhsCode & "',"
'        sQry = sQry & "'" & WhsName & "',"
        sQry = sQry & "'" & Comments & "',"
        sQry = sQry & "'" & ReGbn & "',"
        sQry = sQry & "'" & CardCode & "',"
        sQry = sQry & "'" & CardName & "',"
        sQry = sQry & "'O',"
        sQry = sQry & " GETDATE()"
        sQry = sQry & ")"
        RecordSet01.DoQuery sQry
    Next
    
    MDC_Com.MDC_GF_Message "생산요청완료!", "S"
    
    Set RecordSet01 = Nothing
    Add_ProductionDemand = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_ProductionDemand_Error:
    Add_ProductionDemand = False
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "Add_ProductionDemand_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Function Update_ProductionDemand(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
'******************************************************************************
'Function ID : Update_ProductionDemand()
'해당모듈    : PS_SD010
'기능        : 데이터 UPDATE
'인수        : 없음
'반환값      : 없음
'특이사항    : 2017.02.17 패치 이후 사용 안함. PS_SD010_AddDate 로 대체 (2017.02.17 송명규)
'******************************************************************************
On Error GoTo Update_ProductionDemand_Error
    Dim i As Integer
    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry$, ItemCode$, ItemName$, RegDate$, DueDate$, ItmBsort$, CntcCode$, CntcName$
    Dim ReQty&, ReWeight As Currency, WhsCode$, WhsName$, Comments$, Status$, RegNum$
    
    oMat01.FlushToDataSource
    
    For i = 0 To oMat01.RowCount - 1
        ItemCode = Trim(oDS_PS_SD010H.GetValue("U_ItemCode", i))
        ItemName = Make_ItemName(Trim(oDS_PS_SD010H.GetValue("U_ItemName", i)))
        RegDate = Trim(oDS_PS_SD010H.GetValue("U_RegDate", i))
        DueDate = Trim(oDS_PS_SD010H.GetValue("U_DueDate", i))
        ItmBsort = Trim(oDS_PS_SD010H.GetValue("U_ItmBSort", i))
        CntcCode = Trim(oDS_PS_SD010H.GetValue("U_CntcCode", i))
        CntcName = Trim(oDS_PS_SD010H.GetValue("U_CntcName", i))
'        ReQty = Trim(oDS_PS_SD010H.GetValue("U_ReQty", i))
        ReWeight = Trim(oDS_PS_SD010H.GetValue("U_ReWeight", i))
'        WhsCode = Trim(oDS_PS_SD010H.GetValue("U_WhsCode", i))
'        WhsName = Trim(oDS_PS_SD010H.GetValue("U_WhsName", i))
        Comments = Trim(oDS_PS_SD010H.GetValue("U_Comments", i))
        Status = Trim(oDS_PS_SD010H.GetValue("U_Status", i))
        RegNum = Trim(oDS_PS_SD010H.GetValue("U_RegNum", i))
        
        sQry = "        UPDATE  [@PS_SD010H] "
        sQry = sQry & " SET     U_ItemCode = '" & ItemCode & "', "
        sQry = sQry & "         U_ItemName = '" & ItemName & "', "
        sQry = sQry & "         U_RegDate = '" & RegDate & "', "
        sQry = sQry & "         U_DueDate = '" & DueDate & "', "
        sQry = sQry & "         U_ItmBSort = '" & ItmBsort & "', "
        sQry = sQry & "         U_CntcCode = '" & CntcCode & "', "
        sQry = sQry & "         U_CntcName = '" & CntcName & "', "
'        sQry = sQry & "U_ReQty = '" & ReQty & "', "
        sQry = sQry & "         U_ReWeight = '" & ReWeight & "', "
'        sQry = sQry & "U_WhsCode = '" & WhsCode & "', "
'        sQry = sQry & "U_WhsName = '" & WhsName & "', "
        sQry = sQry & "         U_Comments = '" & Comments & "', "
        sQry = sQry & "         U_Status = '" & Status & "' "
        sQry = sQry & " Where   DocEntry = '" & RegNum & "' "
    
        RecordSet01.DoQuery sQry
    Next
'    sQry = sQry & " U_DueDate,"

    MDC_Com.MDC_GF_Message "구매요청 갱신 완료!", "S"
    
    Set RecordSet01 = Nothing
    Update_ProductionDemand = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Update_ProductionDemand_Error:
    Update_ProductionDemand = False
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "Update_ProductionDemand_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Sub LoadData()
On Error GoTo LoadData_Error
    Dim i As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ReGbn$, BPLID$, ItmBsort$, CardCode$, CntcCode$, RegDateFr$, RegDateTo$, RegNumFr$, RegNumTo$, ItemCode$, Status$
        
    ReGbn = Trim(oForm01.DataSources.UserDataSources("RadioBtn01").VALUE)
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
    ItmBsort = Trim(oForm01.Items("ItmBSort").Specific.VALUE)
    CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
    CntcCode = Trim(oForm01.Items("CntcCode").Specific.VALUE)
    RegDateFr = Trim(oForm01.Items("RegDateFr").Specific.VALUE)
    RegDateTo = Trim(oForm01.Items("RegDateTo").Specific.VALUE)
    RegNumFr = Trim(oForm01.Items("RegNumFr").Specific.VALUE)
    RegNumTo = Trim(oForm01.Items("RegNumTo").Specific.VALUE)
    ItemCode = Trim(oForm01.Items("ItemCode").Specific.VALUE)
    Status = Trim(oForm01.Items("Status").Specific.VALUE)
    
    If BPLID = "" Then BPLID = "%"
    If ItmBsort = "" Or ItmBsort = "ALL" Then ItmBsort = "%"
    If CardCode = "" Then CardCode = "%"
    If CntcCode = "" Then CntcCode = "%"
    If RegDateFr = "" Then RegDateFr = "19000101"
    If RegDateTo = "" Then RegDateTo = "20991231"
    If RegNumFr = "" Then RegNumFr = "0000000000"
    If RegNumTo = "" Then RegNumTo = "9999999999"
    If ItemCode = "" Then ItemCode = "%"
    If Status = "" Then Status = "%"
    
    sQry = "      EXEC [PS_SD010_01] '"
    sQry = sQry & ReGbn & "','"
    sQry = sQry & BPLID & "','"
    sQry = sQry & ItmBsort & "','"
    sQry = sQry & CardCode & "','"
    sQry = sQry & CntcCode & "','"
    sQry = sQry & RegDateFr & "','"
    sQry = sQry & RegDateTo & "','"
    sQry = sQry & RegNumFr & "','"
    sQry = sQry & RegNumTo & "','"
    sQry = sQry & ItemCode & "','"
    sQry = sQry & Status & "'"
    
    Call oRecordSet01.DoQuery(sQry)
    
    oMat01.Clear
    oDS_PS_SD010H.Clear
    
    If (oRecordSet01.RecordCount = 0) Then
        MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
        Set oRecordSet01 = Nothing
        Exit Sub
    End If
    
    oForm01.Freeze True
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_SD010H.Size Then
            oDS_PS_SD010H.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_SD010H.Offset = i
        oDS_PS_SD010H.setValue "DocNum", i, i + 1
        oDS_PS_SD010H.setValue "U_SjDocLin", i, Trim(oRecordSet01.Fields("U_SjDocLin").VALUE)
        oDS_PS_SD010H.setValue "U_ItemCode", i, Trim(oRecordSet01.Fields("U_ItemCode").VALUE)
        oDS_PS_SD010H.setValue "U_ItemName", i, Trim(oRecordSet01.Fields("U_ItemName").VALUE)
        oDS_PS_SD010H.setValue "U_RegDate", i, Format(Trim(oRecordSet01.Fields("U_RegDate").VALUE), "YYYYMMDD")
        oDS_PS_SD010H.setValue "U_DueDate", i, Format(Trim(oRecordSet01.Fields("U_DueDate").VALUE), "YYYYMMDD")
        oDS_PS_SD010H.setValue "U_RegNum", i, Trim(oRecordSet01.Fields("U_RegNum").VALUE)
        oDS_PS_SD010H.setValue "U_ItmBSort", i, Trim(oRecordSet01.Fields("U_ItmBSort").VALUE)
'        oDS_PS_SD010H.setValue "U_SjQty", i, Trim(oRecordSet01.Fields("U_SjQty").Value)
        oDS_PS_SD010H.setValue "U_SjWeight", i, Trim(oRecordSet01.Fields("U_SjWeight").VALUE)
        oDS_PS_SD010H.setValue "U_UnWeight", i, Trim(oRecordSet01.Fields("U_UnWeight").VALUE)
        oDS_PS_SD010H.setValue "U_Price", i, Trim(oRecordSet01.Fields("U_Price").VALUE)
        oDS_PS_SD010H.setValue "U_LinTotal", i, Trim(oRecordSet01.Fields("U_LinTotal").VALUE)
'        oDS_PS_SD010H.setValue "U_Size", i, Trim(oRecordSet01.Fields("U_Size").Value)
'        oDS_PS_SD010H.setValue "U_Spec3", i, Trim(oRecordSet01.Fields("U_Spec3").Value)
'        oDS_PS_SD010H.setValue "U_Quality", i, Trim(oRecordSet01.Fields("U_Quality").Value)
        oDS_PS_SD010H.setValue "U_BPLId", i, Trim(oRecordSet01.Fields("U_BPLId").VALUE)
        oDS_PS_SD010H.setValue "U_CntcCode", i, Trim(oRecordSet01.Fields("U_CntcCode").VALUE)
        oDS_PS_SD010H.setValue "U_CntcName", i, Trim(oRecordSet01.Fields("U_CntcName").VALUE)
'        oDS_PS_SD010H.setValue "U_ReQty", i, Trim(oRecordSet01.Fields("U_ReQty").Value)
        oDS_PS_SD010H.setValue "U_ReWeight", i, Trim(oRecordSet01.Fields("U_ReWeight").VALUE)
'        oDS_PS_SD010H.setValue "U_WhsCode", i, Trim(oRecordSet01.Fields("U_WhsCode").Value)
'        oDS_PS_SD010H.setValue "U_WhsName", i, Trim(oRecordSet01.Fields("U_WhsName").Value)
        oDS_PS_SD010H.setValue "U_Comments", i, Trim(oRecordSet01.Fields("U_Comments").VALUE)
        oDS_PS_SD010H.setValue "U_ReGbn", i, Trim(oRecordSet01.Fields("U_ReGbn").VALUE)
        oDS_PS_SD010H.setValue "U_CardCode", i, Trim(oRecordSet01.Fields("U_CardCode").VALUE)
        oDS_PS_SD010H.setValue "U_CardName", i, Trim(oRecordSet01.Fields("U_CardName").VALUE)
        oDS_PS_SD010H.setValue "U_Status", i, Trim(oRecordSet01.Fields("U_Status").VALUE)
        
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
    oForm01.Freeze False
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "LoadData_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Function PS_SD010_CheckDate() As Boolean
'******************************************************************************
'Function ID : PS_SD010_CheckDate()
'해당모듈    : PS_SD010
'기능        : 선행프로세스와 일자 비교
'인수        : 없음
'반환값      : True-선행프로세스보다 일자가 같거나 느릴 경우, False-선행프로세스보다 일자가 빠를 경우
'특이사항    : 구현만 하고 사용은 안함
'******************************************************************************
On Error GoTo PS_SD010_CheckDate_Error

    Dim Query01      As String
    Dim loopCount    As Integer
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BaseEntry    As String
    Dim BaseLine     As String
    Dim DocType      As String
    Dim CurDocDate   As String
    
    Dim Entry()      As String
    
    For loopCount = 0 To oMat01.VisualRowCount - 2
    
        Entry = Split(oMat01.Columns("SjDocLin").Cells(loopCount).Specific.VALUE, "-")
        BaseEntry = Entry(0)
        BaseLine = Entry(1)
        DocType = "PS_SD010"
        CurDocDate = oMat01.Columns("RegDate").Cells(loopCount).Specific.VALUE
    
        Query01 = "         EXEC PS_Z_CHECK_DATE '"
        Query01 = Query01 & BaseEntry & "','"
        Query01 = Query01 & BaseLine & "','"
        Query01 = Query01 & DocType & "','"
        Query01 = Query01 & CurDocDate & "'"
    
        Call oRecordSet01.DoQuery(Query01)
    
'        If oRecordSet01.Fields("ReturnValue") = False Then
'
'        End
    
    Next
    
    Set oRecordSet01 = Nothing

    Exit Function
PS_SD010_CheckDate_Error:
    PS_SD010_CheckDate = False
    Call Sbo_Application.SetStatusBarMessage("PS_SD010_CheckDate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)
End Function

Private Sub PS_SD010_Print_Report01()
On Error GoTo PS_SD010_Print_Report01_Error

    Dim WinTitle     As String
    Dim ReportName   As String
    Dim sQry         As String
    Dim sQry01       As String
    Dim i            As Long
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ReGbn$, BPLID$, ItmBsort$, CardCode$, CntcCode$, RegDateFr$, RegDateTo$, RegNumFr$, RegNumTo$, ItemCode$, Status$
        
    ReGbn = Trim(oForm01.DataSources.UserDataSources("RadioBtn01").VALUE)
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
    ItmBsort = Trim(oForm01.Items("ItmBSort").Specific.VALUE)
    CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
    CntcCode = Trim(oForm01.Items("CntcCode").Specific.VALUE)
    RegDateFr = Trim(oForm01.Items("RegDateFr").Specific.VALUE)
    RegDateTo = Trim(oForm01.Items("RegDateTo").Specific.VALUE)
    RegNumFr = Trim(oForm01.Items("RegNumFr").Specific.VALUE)
    RegNumTo = Trim(oForm01.Items("RegNumTo").Specific.VALUE)
    ItemCode = Trim(oForm01.Items("ItemCode").Specific.VALUE)
    Status = Trim(oForm01.Items("Status").Specific.VALUE)
    
    If BPLID = "" Then BPLID = "%"
    If ItmBsort = "" Or ItmBsort = "ALL" Then ItmBsort = "%"
    If CardCode = "" Then CardCode = "%"
    If CntcCode = "" Then CntcCode = "%"
    If RegDateFr = "" Then RegDateFr = "19000101"
    If RegDateTo = "" Then RegDateTo = "20991231"
    If RegNumFr = "" Then RegNumFr = "0000000000"
    If RegNumTo = "" Then RegNumTo = "9999999999"
    If ItemCode = "" Then ItemCode = "%"
    If Status = "" Then Status = "%"
    
    Call ConnectODBC
    WinTitle = "[PS_SD010_02] 생산PO현황"
    ReportName = "PS_SD010_02.rpt"
    
    sQry = "      EXEC [PS_SD010_02] '"
    sQry = sQry & ReGbn & "','"
    sQry = sQry & BPLID & "','"
    sQry = sQry & ItmBsort & "','"
    sQry = sQry & CardCode & "','"
    sQry = sQry & CntcCode & "','"
    sQry = sQry & RegDateFr & "','"
    sQry = sQry & RegDateTo & "','"
    sQry = sQry & RegNumFr & "','"
    sQry = sQry & RegNumTo & "','"
    sQry = sQry & ItemCode & "','"
    sQry = sQry & Status & "'"
    
    ReDim gRpt_Formula(3)
    ReDim gRpt_Formula_Value(3)
    
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    gRpt_Formula(1) = "BPLId"
    If Trim(oForm01.Items("BPLId").Specific.VALUE) = "" Then
        gRpt_Formula_Value(1) = "전체"
    Else
        
        sQry01 = "SELECT BPLName FROM [OBPL] WHERE BPLId = '" & Trim(oForm01.Items("BPLId").Specific.VALUE) & "'"
        Call oRecordSet01.DoQuery(sQry01)
        gRpt_Formula_Value(1) = oRecordSet01.Fields(0).VALUE
        
    End If
    
    gRpt_Formula(2) = "RegDateFr"
    gRpt_Formula_Value(2) = Format(oForm01.Items("RegDateFr").Specific.VALUE, "####-##-##")
    
    gRpt_Formula(3) = "RegDateTo"
    gRpt_Formula_Value(3) = Format(oForm01.Items("RegDateTo").Specific.VALUE, "####-##-##")
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    Set oRecordSet01 = Nothing
    
    Exit Sub
    
PS_SD010_Print_Report01_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD010_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
