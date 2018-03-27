VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_MM095"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_MM095.cls
'//  Module         : MM
'//  Description    : 휘팅 원자재 불출 등록
'//  FormType       : PS_MM095
'//  Create Date    : 2010.11.25
'//  Modified Date  :
'//  Creator        : Ryu Yung Jo
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01              As SAPbouiCOM.Form
Public oMat01               As SAPbouiCOM.Matrix
Public oMat02               As SAPbouiCOM.Matrix
Private oDS_PS_MM095H       As SAPbouiCOM.DBDataSource    '등록헤더
Private oDS_PS_MM095L       As SAPbouiCOM.DBDataSource    '등록라인
Private oDS_PS_USERDS01     As SAPbouiCOM.DBDataSource    '등록라인

Private oLast_Item_UID      As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID       As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row       As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oClick_ColRow       As Long
Private oLast_Mode&

Private oQty As Long
Private oWeight As Currency

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_MM095.srf")
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
    
    oFormUniqueID01 = "PS_MM095_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    '화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
    oForm01.DataBrowser.BrowseBy = "DocNum"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting
    Call Initialization
    Call FormClear
        
    Call FormItemEnabled
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1284"), True         '// 취소
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
    
    Dim ItemCode$, ItemName$, Size$, Qty&, Weight As Currency, Unit$, RequestDate$, DueDate$, ItemType$, RequestNo$, UnWeight As Currency
    Dim RFC_Sender$
    Dim Calculate_Weight As Double
    Dim Seq&
    
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
                        
                        '//애드온 자료만 만들때 Add_oInventoryGenExit(1) 로하면 됨
                        '//정상 불출은 Add_oInventoryGenExit(2)
                        If oForm01.Mode = fm_ADD_MODE Then
'                        If oForm01.Mode = fm_UPDATE_MODE Then '//애드온만 만들어졌을때 DI만 생성할때(출고번호는 Update해줘야함)
                            If Add_oInventoryGenExit(2) = False Then
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                        
                        If oForm01.Mode = fm_UPDATE_MODE Then
                            Call Update_PP040L("UPDATE")
                        End If
                    End If
                End If
            Case et_KEY_DOWN: '//2
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
                        oMat02.Clear
                        oDS_PS_USERDS01.Clear
                        oForm01.Mode = fm_OK_MODE
                        Call Sbo_Application.ActivateMenuItem("1282")
                    End If
                End If
            Case et_KEY_DOWN: '//2
'et_COMBO_SELECT ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_COMBO_SELECT: '//5
                If pval.ItemChanged = True Then
                    If oForm01.Mode <> fm_FIND_MODE Then
                        If pval.ItemUID = "BPLId" Or pval.ItemUID = "OrdGbn" Or pval.ItemUID = "Div" Then
                            Call LoadData
                        End If
                    End If
                End If
'et_CLICK ///////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_CLICK: '//6
                If pval.ItemUID = "Mat01" And pval.ColUID = "LineNum" And oForm01.Mode = fm_ADD_MODE Then
                    oClick_ColRow = pval.Row
                    Call LoadData_Mat02(oClick_ColRow)
                End If
'et_DOUBLE_CLICK ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_DOUBLE_CLICK: '//7
                If pval.ItemUID = "Mat02" Then
                    If oMat02.Columns("ItemCode").Cells(pval.Row).Specific.VALUE <> "" And oMat02.Columns("ItemName").Cells(pval.Row).Specific.VALUE <> "" And oMat02.Columns("WhsCode").Cells(pval.Row).Specific.VALUE <> "" Then
'                        If oMat02.Columns("Qty").Cells(pval.Row).Specific.Value = 0 Or oMat02.Columns("Weight").Cells(pval.Row).Specific.Value = 0 Then
'                            MDC_Com.MDC_GF_Message "원자재 재고 데이터의 수량 혹은 중량이 0입니다. 관리자에게 문의 바랍니다.", "E"
'                        Else
                            oForm01.Freeze True
                            
                            oMat01.FlushToDataSource
                            
                            oDS_PS_MM095L.setValue "U_CItemCod", oClick_ColRow - 1, oMat02.Columns("ItemCode").Cells(pval.Row).Specific.VALUE
                            oDS_PS_MM095L.setValue "U_CItemNam", oClick_ColRow - 1, oMat02.Columns("ItemName").Cells(pval.Row).Specific.VALUE
                            oDS_PS_MM095L.setValue "U_WhsCode", oClick_ColRow - 1, oMat02.Columns("WhsCode").Cells(pval.Row).Specific.VALUE
                            oDS_PS_MM095L.setValue "U_CQty", oClick_ColRow - 1, oMat02.Columns("Qty").Cells(pval.Row).Specific.VALUE
                            oDS_PS_MM095L.setValue "U_CWeight", oClick_ColRow - 1, oMat02.Columns("Weight").Cells(pval.Row).Specific.VALUE
                            oDS_PS_MM095L.setValue "U_UnWeight", oClick_ColRow - 1, oMat02.Columns("UnWeight").Cells(pval.Row).Specific.VALUE '//원재료단중
                            
                            'oMat01.Columns("CItemCod").Cells(oClick_ColRow).Specific.VALUE = oMat02.Columns("ItemCode").Cells(pval.Row).Specific.VALUE
'                            oMat01.Columns("CItemNam").Cells(oClick_ColRow).Specific.VALUE = oMat02.Columns("ItemName").Cells(pval.Row).Specific.VALUE
'                            oMat01.Columns("WhsCode").Cells(oClick_ColRow).Specific.VALUE = oMat02.Columns("WhsCode").Cells(pval.Row).Specific.VALUE
'                            oMat01.Columns("CQty").Cells(oClick_ColRow).Specific.VALUE = oMat02.Columns("Qty").Cells(pval.Row).Specific.VALUE
'                            oMat01.Columns("CWeight").Cells(oClick_ColRow).Specific.VALUE = oMat02.Columns("Weight").Cells(pval.Row).Specific.VALUE
                            '
                            If oMat02.Columns("UnWeight").Cells(pval.Row).Specific.VALUE = 1 Then
                                '//원재료 수량으로 불출
                                oDS_PS_MM095L.setValue "U_IssueQty", oClick_ColRow - 1, oMat01.Columns("PQty").Cells(oClick_ColRow).Specific.VALUE
                                oDS_PS_MM095L.setValue "U_IssueWt", oClick_ColRow - 1, oMat01.Columns("PQty").Cells(oClick_ColRow).Specific.VALUE
                            Else
                                '//원재료 단중으로 계산
                                oDS_PS_MM095L.setValue "U_IssueQty", oClick_ColRow - 1, 0
                                oDS_PS_MM095L.setValue "U_IssueWt", oClick_ColRow - 1, Round(oMat01.Columns("PQty").Cells(oClick_ColRow).Specific.VALUE * oMat02.Columns("UnWeight").Cells(pval.Row).Specific.VALUE, 2)
                            End If
'                            If oMat01.Columns("PItemCod").Cells(oClick_ColRow).Specific.VALUE = "102010002" Or oMat01.Columns("PItemCod").Cells(oClick_ColRow).Specific.VALUE = "102010005" Or oMat01.Columns("PItemCod").Cells(oClick_ColRow).Specific.VALUE = "102010006" Or oMat01.Columns("PItemCod").Cells(oClick_ColRow).Specific.VALUE = "102010003" Or oMat01.Columns("PItemCod").Cells(oClick_ColRow).Specific.VALUE = "102010004" Or oMat01.Columns("PItemCod").Cells(oClick_ColRow).Specific.VALUE = "102010007" Then
'                                oDS_PS_MM095L.setValue "U_IssueQty", oClick_ColRow - 1, oMat01.Columns("PQty").Cells(oClick_ColRow).Specific.VALUE
'                                oDS_PS_MM095L.setValue "U_IssueWt", oClick_ColRow - 1, oMat01.Columns("PQty").Cells(oClick_ColRow).Specific.VALUE
                                
                                
'                                oMat01.Columns("IssueWt").Cells(oClick_ColRow).Specific.VALUE = oMat01.Columns("PQty").Cells(oClick_ColRow).Specific.VALUE
'                                oMat01.Columns("IssueQty").Cells(oClick_ColRow).Specific.VALUE = oMat01.Columns("PQty").Cells(oClick_ColRow).Specific.VALUE
'                            End If
                            oMat01.LoadFromDataSource
                            oForm01.Freeze False
                            '해당라인으로 이동
                            If oClick_ColRow = oMat01.RowCount Then
                                Call oMat01.SelectRow(oClick_ColRow, True, False)
                            Else
                                Call oMat01.SelectRow(oClick_ColRow + 1, True, False)
                            End If
                        
'                            oQty = oMat02.Columns("Qty").Cells(pval.Row).Specific.Value
'                            oWeight = oMat02.Columns("Weight").Cells(pval.Row).Specific.Value
'                        End If
                    Else
                        MDC_Com.MDC_GF_Message "원자재 재고 데이터가 불확실 합니다. 관리자에게 문의 바랍니다.", "E"
                    End If
                End If
            Case et_MATRIX_LINK_PRESSED '//8
'et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "IssueQty" Then
                            oForm01.Freeze True
                            oMat01.FlushToDataSource
                            ItemCode = Trim(oDS_PS_MM095L.GetValue("U_CItemCod", pval.Row - 1))
                            If ItemCode = "" Then
                                oDS_PS_MM095L.setValue "U_IssueQty", pval.Row - 1, 0
                                oDS_PS_MM095L.setValue "U_IssueWt", pval.Row - 1, 0
                                oMat01.LoadFromDataSource
                                oMat01.Columns("IssueQty").Cells(pval.Row).Click ct_Regular
                                MDC_Com.MDC_GF_Message "불출할 원자재를 아래 매트릭스에서 먼저 선택해 주세요.", "E"
                            Else
                                Qty = oDS_PS_MM095L.GetValue("U_CQty", pval.Row - 1)
                                Weight = oDS_PS_MM095L.GetValue("U_CWeight", pval.Row - 1)
                                If Qty = 0 Then
                                    UnWeight = 0
                                Else
                                    UnWeight = Weight / Qty
                                End If
                                
                                Calculate_Weight = oDS_PS_MM095L.GetValue("U_IssueQty", pval.Row - 1) * UnWeight
                                oDS_PS_MM095L.setValue "U_IssueWt", pval.Row - 1, Calculate_Weight
                                oMat01.LoadFromDataSource
                                oMat01.Columns("IssueQty").Cells(pval.Row).Click ct_Regular
'                                Qty = oDS_PS_MM095L.GetValue("U_IssueQty", pval.Row - 1)
'
'                                Calculate_Weight = MDC_PS_Common.Calculate_Weight(ItemCode, Qty,Trim(oForm01.Items("BPLId").Specific.Value))
'
'                                oDS_PS_MM095L.setValue "U_IssueWt", pval.Row - 1, Calculate_Weight
'                                oMat01.LoadFromDataSource
'                                oMat01.Columns("IssueQty").Cells(pval.Row).Click ct_Regular
                            End If
                            oForm01.Freeze False
                        ElseIf pval.ColUID = "IssueWt" Then
                            oForm01.Freeze True
                            oMat01.FlushToDataSource
                            ItemCode = Trim(oDS_PS_MM095L.GetValue("U_CItemCod", pval.Row - 1))
                            If ItemCode = "" Then
                                oDS_PS_MM095L.setValue "U_IssueQty", pval.Row - 1, 0
                                oDS_PS_MM095L.setValue "U_IssueWt", pval.Row - 1, 0
                                oMat01.LoadFromDataSource
                                oMat01.Columns("IssueQty").Cells(pval.Row).Click ct_Regular
                                MDC_Com.MDC_GF_Message "불출할 원자재를 아래 매트릭스에서 먼저 선택해 주세요.", "E"
                            End If
                            oForm01.Freeze False
                        
                        End If
                    ElseIf pval.ItemUID = "DocDate" Then
                        Call LoadData
                    End If
                End If
            Case et_MATRIX_LOAD: '//11
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
'et_FORM_RESIZE /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_FORM_RESIZE '//20
                oForm01.Items("Mat02").Top = oForm01.Items("Comments").Top - 105
                oForm01.Items("Mat02").Left = 6
                oForm01.Items("Mat02").Width = oForm01.Width - 18 '788
                oForm01.Items("Mat02").Height = 90
                oMat02.Columns("ItemCode").Width = 100
                oMat02.Columns("ItemName").Width = 200
                oMat02.Columns("WhsCode").Width = 100
                oMat02.Columns("WhsName").Width = 100
                oMat02.Columns("Qty").Width = 100
                oMat02.Columns("Weight").Width = 100
                
                oForm01.Items("Mat01").Top = 61
                oForm01.Items("Mat01").Left = 6
                oForm01.Items("Mat01").Width = oForm01.Width - 18
                oForm01.Items("Mat01").Height = oForm01.Height - 280
                
                
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
            Case et_LOST_FOCUS: '//4
'et_FORM_UNLOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_FORM_UNLOAD: '//17
                RemoveForms oFormUniqueID01
                Set oForm01 = Nothing
                Set oMat01 = Nothing
                Set oDS_PS_MM095H = Nothing
                Set oDS_PS_MM095L = Nothing
                Set oDS_PS_USERDS01 = Nothing
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
    Set ProgressBar01 = Nothing
    If oForm01 Is Nothing Then
    Else
        oForm01.Freeze False
    End If
    MDC_Com.MDC_GF_Message "Raise_ItemEvent_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_MenuEvent_Error
    Dim i&
    Dim sReturnValue As Integer
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
'[1284:취소] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1284": '취소
                If Trim(oDS_PS_MM095H.GetValue("Canceled", 0)) = "N" Then
                    sReturnValue = Sbo_Application.MessageBox("해당 문서가 취소됩니다 계속하시겠습니까??", 1, "&확인", "&취소")
                    Select Case sReturnValue
                        Case 1:
                            If Add_oInventoryGenEntry(1) = False Then
                                BubbleEvent = False
                                Exit Sub
                            Else
                                If Add_oInventoryGenEntry(2) = False Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            End If
                        Case 2:
                            Sbo_Application.MessageBox ("실행이 취소되었습니다.")
                            BubbleEvent = False
                    End Select
                Else
                    MDC_Com.MDC_GF_Message "취소된 생산실적입니다. 확인하세요.", "E"
                    BubbleEvent = False
                    Exit Sub
                End If
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
                    oDS_PS_MM095L.RemoveRecord oDS_PS_MM095L.Size - 1       '// Mat01에 마지막라인(빈라인) 삭제
                    oMat01.Clear
                    oMat01.LoadFromDataSource
                    
'                    If oMat01.Columns("GADocLin").Cells(oMat01.RowCount).Specific.Value <> "" Then
'                        Call Add_MatrixRow(oMat01.RowCount, False)
'                    End If
                End If
'[1281:찾기] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1281": '찾기
                Call FormItemEnabled
                oForm01.Items("DocNum").Click ct_Regular
'[1282:추가] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1282": '추가
                oForm01.Freeze True
                Call Initialization
                Call FormItemEnabled
                Call FormClear
                oForm01.Items("BPLId").Specific.Select "0", psk_Index
                oForm01.Items("OrdGbn").Specific.Select "0", psk_Index
                oDS_PS_MM095H.setValue "U_DocDate", 0, Format(Now, "YYYYMMDD")
                oForm01.Freeze False
'[1288~1291:네비게이션] /////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                Call FormItemEnabled
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_MenuEvent_Error:
    MDC_Com.MDC_GF_Message "Raise_MenuEvent_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub Raise_FormDataEvent(ByRef FormUID As String, ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_FormDataEvent_Error
    If (BusinessObjectInfo.BeforeAction = True) Then '//BeforeAction = True
        Select Case BusinessObjectInfo.EventType
            Case et_FORM_DATA_LOAD: '//33
            Case et_FORM_DATA_ADD: '//34
            
'                If Add_oInventoryGenExit(2) = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
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
    Set oDS_PS_MM095H = oForm01.DataSources.DBDataSources("@PS_MM095H")
    Set oDS_PS_MM095L = oForm01.DataSources.DBDataSources("@PS_MM095L")
    Set oDS_PS_USERDS01 = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    Set oMat02 = oForm01.Items("Mat02").Specific
    
        
    oDS_PS_MM095H.setValue "U_DocDate", 0, Format(Now, "yyyymmdd")
    
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
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    
    
    '//작업구분
    Set oCombo = oForm01.Items("OrdGbn").Specific
    oCombo.ValidValues.Add "-", "선택"
    sQry = "SELECT Code, Name From [@PSH_ITMBSORT] Where Code in ('102', '111', '601','602') And U_PudYN = 'Y' Order by Code"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    oCombo.Select "1", psk_Index
    
    '//조회기준
    Set oCombo = oForm01.Items("Div").Specific
    oCombo.ValidValues.Add "1", "전체"
    oCombo.ValidValues.Add "2", "기준일자"
    oCombo.Select "1", psk_Index
'
    '//작업구분
    oMat01.Columns("WorkGbn").ValidValues.Add "10", "자가"
    oMat01.Columns("WorkGbn").ValidValues.Add "20", "정밀"
    oMat01.Columns("WorkGbn").ValidValues.Add "30", "외주"
    
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
    If (oForm01.Mode = fm_ADD_MODE) Then
        oForm01.Items("DocNum").Enabled = False

        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("OrdGbn").Enabled = True
        oForm01.Items("DocDate").Enabled = True
        
        oMat01.Columns("IssueYN").Editable = True
        oMat01.Columns("IssueQty").Editable = True
        oMat01.Columns("IssueWt").Editable = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        oForm01.Items("DocNum").Enabled = True

        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("OrdGbn").Enabled = True
        oForm01.Items("DocDate").Enabled = True
        
        oMat01.Columns("IssueYN").Editable = False
        oMat01.Columns("IssueQty").Editable = False
        oMat01.Columns("IssueWt").Editable = False
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        oForm01.Items("DocNum").Enabled = False
        
        oForm01.Items("BPLId").Enabled = False
        oForm01.Items("OrdGbn").Enabled = False
        oForm01.Items("DocDate").Enabled = False

        oMat01.Columns("IssueYN").Editable = False
        oMat01.Columns("IssueQty").Editable = False
        oMat01.Columns("IssueWt").Editable = False
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub FormClear()
On Error GoTo FormClear_Error
    Dim DocNum As String
    DocNum = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM095'", "")
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
        oDS_PS_MM095L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_MM095L.Offset = oRow
    oDS_PS_MM095L.setValue "U_LineNum", oRow, oRow + 1
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
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
        Case "Mat01"
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
        Case oDS_PS_MM095H.GetValue("U_BPLId", 0) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_MM095H.GetValue("U_OrdGbn", 0) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_MM095H.GetValue("U_DocDate", 0) = ""
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "사업장은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "작업구분은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "전기일은 필수사항입니다. 확인하세요.", "E"
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

    '// 라인
    If oMat01.VisualRowCount = 0 Then
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    End If
    
    oForm01.Freeze True
First_Step:
    oMat01.FlushToDataSource
    For i = 0 To oMat01.VisualRowCount - 1
        If oDS_PS_MM095L.GetValue("U_IssueWt", i) = 0 Then
            oDS_PS_MM095L.RemoveRecord i
            oMat01.LoadFromDataSource
            GoTo First_Step
        End If
    Next i
    oMat01.LoadFromDataSource
    
    oMat01.FlushToDataSource
    For i = 0 To oMat01.VisualRowCount - 1
        oDS_PS_MM095L.setValue "U_LineNum", i, i + 1
    Next i
    oMat01.LoadFromDataSource
    
    oForm01.Freeze False
    
    Set oRecordSet01 = Nothing
    MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
MatrixSpaceLineDel_Error:
    oForm01.Freeze False
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Or ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "라인 데이터가 없습니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function

Public Sub LoadData()
On Error GoTo LoadData_Error
    Dim i As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLID$, OrdGbn$
    Dim DocDate$, Div$
    
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
    OrdGbn = Trim(oForm01.Items("OrdGbn").Specific.VALUE)
    DocDate = Trim(oForm01.Items("DocDate").Specific.VALUE)
    Div = Trim(oForm01.Items("Div").Specific.VALUE)
           
    sQry = "EXEC [PS_MM095_01] '" & BPLID & "','" & OrdGbn & "','" & DocDate & "','" & Div & "'"
    oRecordSet01.DoQuery sQry
    
    oMat01.Clear
    oDS_PS_MM095L.Clear
    
    If (oRecordSet01.RecordCount = 0) Then
        MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
        Set oRecordSet01 = Nothing
        Exit Sub
    End If
    
    oForm01.Freeze True
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_MM095L.Size Then
            oDS_PS_MM095L.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_MM095L.Offset = i
        oDS_PS_MM095L.setValue "U_LineNum", i, i + 1
        oDS_PS_MM095L.setValue "u_DocDate", i, Format(Trim(oRecordSet01.Fields("u_DocDate").VALUE), "YYYYMMDD")
        oDS_PS_MM095L.setValue "U_WorkGbn", i, Trim(oRecordSet01.Fields("U_WorkGbn").VALUE)
        oDS_PS_MM095L.setValue "U_OrdNum", i, Trim(oRecordSet01.Fields("U_OrdNum").VALUE)
        oDS_PS_MM095L.setValue "U_CpCode", i, Trim(oRecordSet01.Fields("U_CpCode").VALUE)
        oDS_PS_MM095L.setValue "U_CpName", i, Trim(oRecordSet01.Fields("U_CpName").VALUE)
        oDS_PS_MM095L.setValue "U_PItemCod", i, Trim(oRecordSet01.Fields("U_PItemCod").VALUE)
        oDS_PS_MM095L.setValue "U_PItemNam", i, Trim(oRecordSet01.Fields("U_PItemNam").VALUE)
        oDS_PS_MM095L.setValue "U_PQty", i, Trim(oRecordSet01.Fields("U_PQty").VALUE)
        oDS_PS_MM095L.setValue "U_PWeight", i, Trim(oRecordSet01.Fields("U_PWeight").VALUE)
        oDS_PS_MM095L.setValue "U_CItemCod", i, Trim(oRecordSet01.Fields("U_CItemCod").VALUE)
        oDS_PS_MM095L.setValue "U_CItemNam", i, Trim(oRecordSet01.Fields("U_CItemNam").VALUE)
        oDS_PS_MM095L.setValue "U_CutCount", i, Trim(oRecordSet01.Fields("U_CutCount").VALUE)
        oDS_PS_MM095L.setValue "U_IssueYN", i, Trim(oRecordSet01.Fields("U_IssueYN").VALUE)
        oDS_PS_MM095L.setValue "U_IssueQty", i, 0
        oDS_PS_MM095L.setValue "U_PreQty", i, Trim(oRecordSet01.Fields("U_PreQty").VALUE)
        oDS_PS_MM095L.setValue "U_PreWt", i, Trim(oRecordSet01.Fields("U_PreWt").VALUE)
''        oDS_PS_MM095L.setValue "U_WhsCode", i, Trim(oRecordSet01.Fields("U_WhsCode").Value)
        oDS_PS_MM095L.setValue "U_PP040Doc", i, Trim(oRecordSet01.Fields("U_PP040Doc").VALUE)
        oDS_PS_MM095L.setValue "U_PP040Lin", i, Trim(oRecordSet01.Fields("U_PP040Lin").VALUE)

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

Public Sub LoadData_Mat02(sRow As Long)
On Error GoTo LoadData_Mat02_Error
    Dim i As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ItemCode$, CItemCod$
    
    ItemCode = Trim(oMat01.Columns("PItemCod").Cells(sRow).Specific.VALUE)
    CItemCod = Trim(oMat01.Columns("CItemCod").Cells(sRow).Specific.VALUE)
    
    If Trim(CItemCod) = "" Then CItemCod = "%"
    
    sQry = "Select a.U_ItemCod2, a.U_ItemNam2, b.WhsCode, c.WhsName, b.U_Qty, b.OnHand, "
    sQry = sQry & " UnWeight = Case When Isnull(a.U_UnWeight,0) = 0 Then 1 When a.U_UnWeight = 0 Then 1 Else a.U_UnWeight End "
    sQry = sQry & "From [@PS_PP005H] a Inner Join [OITW] b On a.U_ItemCod2 = b.ItemCode Inner Join [OWHS] c On b.WhsCode = c.WhsCode "
    sQry = sQry & "Where a.U_ItemCod1 = '" & ItemCode & "' And b.OnHand > 0"
    sQry = sQry & " And a.U_ItemCod2 Like '" & CItemCod & "'"
    oRecordSet01.DoQuery sQry
    
    oMat02.Clear
    oDS_PS_USERDS01.Clear
    
    If (oRecordSet01.RecordCount = 0) Then
        MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
        Set oRecordSet01 = Nothing
        Exit Sub
    End If
    
    oForm01.Freeze True
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_USERDS01.Size Then
            oDS_PS_USERDS01.InsertRecord (i)
        End If
        
        oMat02.AddRow
        oDS_PS_USERDS01.Offset = i
        oDS_PS_USERDS01.setValue "U_LineNum", i, i + 1
        oDS_PS_USERDS01.setValue "U_ColReg01", i, Trim(oRecordSet01.Fields("U_ItemCod2").VALUE)
        oDS_PS_USERDS01.setValue "U_ColRgL01", i, Trim(oRecordSet01.Fields("U_ItemNam2").VALUE)
        oDS_PS_USERDS01.setValue "U_ColReg02", i, Trim(oRecordSet01.Fields("WhsCode").VALUE)
        oDS_PS_USERDS01.setValue "U_ColReg03", i, Trim(oRecordSet01.Fields("WhsName").VALUE)
        oDS_PS_USERDS01.setValue "U_ColNum01", i, Trim(oRecordSet01.Fields("U_Qty").VALUE)
        oDS_PS_USERDS01.setValue "U_ColQty01", i, Trim(oRecordSet01.Fields("OnHand").VALUE)
        oDS_PS_USERDS01.setValue "U_ColRat01", i, Trim(oRecordSet01.Fields("UnWeight").VALUE)
        
        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
    Next
    oMat02.LoadFromDataSource
    oMat02.AutoResizeColumns
    ProgBar01.Stop
    oForm01.Freeze False
  
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadData_Mat02_Error:
    ProgBar01.Stop
    oForm01.Freeze False
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "LoadData_Mat02_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub Update_PP040L(sStatus As String)
On Error GoTo Update_PP040L_Error
    Dim i As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocDate$, DocNum$, CItemCod$, IssueQty As Long, IssueWt As Currency, WhsCode$, IssueYN$
    Dim PP040Doc$, PP040Lin$
        
    oForm01.Freeze True
    If sStatus = "ADD" Then
        IssueYN = "Y"
    ElseIf sStatus = "CANCEL" Then
        IssueYN = "N"
    End If
    
    For i = 0 To oMat01.VisualRowCount - 1
        CItemCod = oMat01.Columns("CItemCod").Cells(i + 1).Specific.VALUE
        IssueQty = oMat01.Columns("IssueQty").Cells(i + 1).Specific.VALUE
        IssueWt = oMat01.Columns("IssueWt").Cells(i + 1).Specific.VALUE
        WhsCode = oMat01.Columns("WhsCode").Cells(i + 1).Specific.VALUE
        If CItemCod <> "" And IssueQty >= 0 And IssueWt <> 0 And WhsCode <> "" Then
            If Trim(oMat01.Columns("IssueYN").Cells(i + 1).Specific.VALUE) = "Y" Then
                PP040Doc = Trim(oMat01.Columns("PP040Doc").Cells(i + 1).Specific.VALUE)
                PP040Lin = Trim(oMat01.Columns("PP040Lin").Cells(i + 1).Specific.VALUE)
                sQry = "Update [@PS_PP040L] Set U_IssueYN = '" & IssueYN & "' Where DocEntry = '" & PP040Doc & "' And LineId = '" & PP040Lin & "'"
                oRecordSet01.DoQuery sQry
            End If
        End If
    Next i
    oForm01.Freeze False
  
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Update_PP040L_Error:
    oForm01.Freeze False
    If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_RollBack
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "Update_PP040L_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Function Add_oInventoryGenExit(ChkType As Integer) As Boolean
On Error GoTo Add_oInventoryGenExit_Error
    Dim DI_oInventoryGenExit As SAPbobsCOM.Documents         '//재고출고 문서 객체
    
    Dim i&, j&
    Dim ErrNum          As Integer
    Dim ErrCode         As Long
    Dim ErrMsg          As String
    Dim RetVal&
    
    Dim DocDate$, DocNum$, CItemCod$, IssueQty As Long, IssueWt As Currency, WhsCode$
    Dim sDocEntry$
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
     
    
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    oMat01.FlushToDataSource
    
    DocDate = Format(oDS_PS_MM095H.GetValue("U_DocDate", 0), "0000-00-00")
    If ChkType = 2 Then Call FormClear
    DocNum = Trim(oDS_PS_MM095H.GetValue("DocNum", 0))
    
        
    If Trim(oDS_PS_MM095H.GetValue("U_OIGEDoc", 0)) = "" Then
        Sbo_Company.StartTransaction
        Set DI_oInventoryGenExit = Nothing
        Set DI_oInventoryGenExit = Sbo_Company.GetBusinessObject(oInventoryGenExit)
        
        With DI_oInventoryGenExit
            .DocDate = DocDate
            .TaxDate = DocDate
            .Comments = "원재료 불출 등록(" & DocNum & ") 출고"
            
            j = 0
            For i = 0 To oMat01.VisualRowCount - 1
                CItemCod = oMat01.Columns("CItemCod").Cells(i + 1).Specific.VALUE
                IssueQty = oMat01.Columns("IssueQty").Cells(i + 1).Specific.VALUE
                IssueWt = oMat01.Columns("IssueWt").Cells(i + 1).Specific.VALUE
                WhsCode = oMat01.Columns("WhsCode").Cells(i + 1).Specific.VALUE
            
                If CItemCod <> "" And IssueQty >= 0 And IssueWt <> 0 And WhsCode <> "" Then
                    If j > 0 Then .Lines.Add
                    .Lines.SetCurrentLine j
                    .Lines.ItemCode = CItemCod
                    .Lines.WarehouseCode = WhsCode
        '            .Lines.AccountCode = Trim(sAccount)
                    .Lines.Quantity = IssueWt
                    .Lines.UserFields("U_Qty").VALUE = IssueQty
                    
                    j = j + 1
                End If
            Next i
        End With
            
        '// 완료
        RetVal = DI_oInventoryGenExit.Add
        If (0 <> RetVal) Then
            Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
            ErrNum = 1
            GoTo Add_oInventoryGenExit_Error
        End If
        
        If ChkType = 1 Then
            Sbo_Company.EndTransaction wf_RollBack
        ElseIf ChkType = 2 Then
            Sbo_Company.GetNewObjectCode sDocEntry
            oDS_PS_MM095H.setValue "U_OIGEDoc", 0, sDocEntry
            Call Update_PP040L("ADD")
            Sbo_Company.EndTransaction wf_Commit
        End If
    End If
    
    Set oRecordSet = Nothing
    Set DI_oInventoryGenExit = Nothing
    Add_oInventoryGenExit = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_oInventoryGenExit_Error:
    Set oRecordSet = Nothing
    Set DI_oInventoryGenExit = Nothing
    If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_RollBack
    Add_oInventoryGenExit = False
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "Add_oInventoryGenExit_Error:" & ErrCode & " - " & ErrMsg, "E"
    Else
        MDC_Com.MDC_GF_Message "Add_oInventoryGenExit_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Function

Private Function Add_oInventoryGenEntry(ChkType As Integer) As Boolean
On Error GoTo Add_oInventoryGenEntry_Error
    Dim DI_oInventoryGenEntry As SAPbobsCOM.Documents         '//재고출고 문서 객체
    
    Dim i&, j&
    Dim ErrNum          As Integer
    Dim ErrCode         As Long
    Dim ErrMsg          As String
    Dim RetVal&
    
    Dim DocDate$, DocNum$, CItemCod$, IssueQty As Long, IssueWt As Currency, WhsCode$
    Dim sDocEntry$
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim OIGEDoc As String
    
    
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    oMat01.FlushToDataSource
    
    DocDate = Format(oDS_PS_MM095H.GetValue("U_DocDate", 0), "0000-00-00")
    DocNum = Trim(oDS_PS_MM095H.GetValue("DocNum", 0))
    OIGEDoc = Trim(oDS_PS_MM095H.GetValue("U_OIGEDoc", 0))
    
    If Trim(oDS_PS_MM095H.GetValue("U_OIGNDoc", 0)) = "" Then
        Sbo_Company.StartTransaction
        Set DI_oInventoryGenEntry = Nothing
        Set DI_oInventoryGenEntry = Sbo_Company.GetBusinessObject(oInventoryGenEntry)
        
        With DI_oInventoryGenEntry
            .DocDate = DocDate
            .TaxDate = DocDate
            .Comments = "원재료 불출 등록 출고 취소 (" & DocNum & ") 입고"
            
            .UserFields("U_CancDoc").VALUE = OIGEDoc '//입고취소 문서번호
            
            j = 0
            For i = 0 To oMat01.VisualRowCount - 1
                CItemCod = oMat01.Columns("CItemCod").Cells(i + 1).Specific.VALUE
                IssueQty = oMat01.Columns("IssueQty").Cells(i + 1).Specific.VALUE
                IssueWt = oMat01.Columns("IssueWt").Cells(i + 1).Specific.VALUE
                WhsCode = oMat01.Columns("WhsCode").Cells(i + 1).Specific.VALUE
            
                If CItemCod <> "" And IssueQty >= 0 And IssueWt <> 0 And WhsCode <> "" Then
                    If j > 0 Then .Lines.Add
                    .Lines.SetCurrentLine j
                    .Lines.ItemCode = CItemCod
                    .Lines.WarehouseCode = WhsCode
        '            .Lines.AccountCode = Trim(sAccount)
                    .Lines.Quantity = IssueWt
                    .Lines.UserFields("U_Qty").VALUE = IssueQty
                    
                    j = j + 1
                End If
            Next i
        End With
            
        '// 완료
        RetVal = DI_oInventoryGenEntry.Add
        If (0 <> RetVal) Then
            Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
            ErrNum = 1
            GoTo Add_oInventoryGenEntry_Error
        End If
        
        If ChkType = 1 Then
            Sbo_Company.EndTransaction wf_RollBack
        ElseIf ChkType = 2 Then
            Sbo_Company.GetNewObjectCode sDocEntry
            oDS_PS_MM095H.setValue "U_OIGNDoc", 0, sDocEntry
            Call Update_PP040L("CANCEL")
            Sbo_Company.EndTransaction wf_Commit
        End If
    End If
    
    Set oRecordSet = Nothing
    Set DI_oInventoryGenEntry = Nothing
    Add_oInventoryGenEntry = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_oInventoryGenEntry_Error:
    Set oRecordSet = Nothing
    Set DI_oInventoryGenEntry = Nothing
    If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_RollBack
    Add_oInventoryGenEntry = False
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "Add_oInventoryGenEntry_Error:" & ErrCode & " - " & ErrMsg, "E"
    Else
        MDC_Com.MDC_GF_Message "Add_oInventoryGenEntry_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Function
