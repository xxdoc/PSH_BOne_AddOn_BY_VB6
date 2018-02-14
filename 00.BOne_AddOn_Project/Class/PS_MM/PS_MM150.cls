VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_MM150"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_MM150.cls
'//  Module         : MM
'//  Description    : 외주업체 출고등록
'//  FormType       : PS_MM150
'//  Create Date    : 2010.10.13
'//  Modified Date  :
'//  Creator        : Ryu Yung Jo
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01             As SAPbouiCOM.Form
Public oMat01              As SAPbouiCOM.Matrix
Private oDS_PS_MM150H As SAPbouiCOM.DBDataSource    '등록헤더

Private oLast_Item_UID      As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID       As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row       As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private Type PS_PP040s
    PP030HNo As String
    PP030MNo As String
    OrdNum As String
    OrdGbn As String
    ItemCode As String
    ItemName As String
    Sequence As String
    CpCode As String
    CpName As String
    Chk As String
End Type

Private PS_PP040() As PS_PP040s
    
Private Type PS_PP040DocEntrys
    PP040DocEntry As String
End Type

Private PS_PP040DocEntry() As PS_PP040DocEntrys

Private oMenuUID$
Private oLast_RightClick_DocEntry$, oLast_RightClick_InFlag$

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm(Optional MenuUID$)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_MM150.srf")
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
    
    oFormUniqueID01 = "PS_MM150_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    
    oForm01.Freeze True
    oMenuUID = MenuUID
    Call CreateItems
    Call ComboBox_Setting(oMenuUID)
    Call Initialization
    Call FormItemEnabled(oMenuUID)
    
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
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    
                    If MatrixSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    
                    If Trim(oForm01.Items("Status").Specific.Value) = "N" Then '미출고
                        If Add_PS_MM150H(pval) = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                    ElseIf Trim(oForm01.Items("Status").Specific.Value) = "Y" Then '출고
                        If oMenuUID = "PS_MM150_1" Then
                            Sbo_Company.StartTransaction
                            If Add_PS_PP040(pval) = False Then
                                BubbleEvent = False
                                Exit Sub
                            End If
                            If Update_PS_MM150H(pval) = False Then
                                BubbleEvent = False
                                Exit Sub
                            End If
                        Else
                            Sbo_Company.StartTransaction
                            If Update_PS_MM150H(pval) = False Then
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                    End If
                    
                    oMat01.Clear
                    oDS_PS_MM150H.Clear
                    oMat01.FlushToDataSource
                    oMat01.LoadFromDataSource
                ElseIf pval.ItemUID = "Btn02" Then
                    oForm01.Freeze True
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    Call LoadData
                    Call FormItemEnabled(oMenuUID)
                    oForm01.Freeze False
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
                    oDS_PS_MM150H.Clear
                    oMat01.FlushToDataSource
                    oMat01.LoadFromDataSource
                    oForm01.Freeze False
'                ElseIf pval.ItemUID = "Mat01" And pval.ColUID = "ReStatus" Then
'                    oMat01.FlushToDataSource
'                    If Trim(oDS_PS_MM150H.GetValue("U_ReStatus", pval.Row)) = "Y" Then
'
'                    End If
                End If
'et_CLICK ///////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_CLICK: '//6
'                If pval.ItemUID = "Mat01" And oLast_Col_Row > 0 And pval.ColUID = "Check" Then
'                    oForm01.Freeze True
'                    oMat01.FlushToDataSource
'                    Check = Trim(oDS_PS_MM150H.GetValue("U_Check", pval.Row - 1))
'                    OutDoc = Trim(oDS_PS_MM150H.GetValue("U_OutDoc", pval.Row - 1))
'                    For i = 0 To oMat01.VisualRowCount - 1
'                        If Trim(oDS_PS_MM150H.GetValue("U_OutDoc", i)) = OutDoc Then
'                            oDS_PS_MM150H.setValue "U_Check", i, Check
'                        End If
'                    Next i
'                    oMat01.LoadFromDataSource
'                    oForm01.Freeze False
'                End If
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
                Set oDS_PS_MM150H = Nothing
                oMenuUID = ""
        End Select
    End If
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
    If oForm01 Is Nothing Then
    Else
        oForm01.Freeze False
    End If
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
'[1293:행삭제] //////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1293": '행삭제
                If oLast_RightClick_InFlag <> "Y" Then
                    ReturnValue = Sbo_Application.MessageBox("해당 라인의 데이터를 삭제합니다. 삭제 후 복원할 수 없습니다. 삭제하시겠습니까?", 1, "&확인", "&취소")
                    Select Case ReturnValue
                        Case 1:
                            sQry = "Delete [@PS_MM150H] Where DocEntry = '" & oLast_RightClick_DocEntry & "'"
                            oRecordSet01.DoQuery sQry
                            
                            oLast_RightClick_DocEntry = 0
                        Case 2:
                            Sbo_Application.MessageBox ("실행이 취소되었습니다.")
                            BubbleEvent = False
                            Exit Sub
                    End Select
                Else
                    MDC_Com.MDC_GF_Message "해당 문서는 출고확인이 되었기 때문에 삭제할 수 없습니다.", "E"
                    BubbleEvent = False
                    Exit Sub
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
                        oMat01.Columns("DocNum").Cells(i + 1).Specific.Value = i + 1
                    Next i

                    oMat01.FlushToDataSource
                    oDS_PS_MM150H.RemoveRecord oDS_PS_MM150H.Size - 1       '// Mat01에 마지막라인(빈라인) 삭제
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
            oLast_RightClick_DocEntry = Trim(oMat01.Columns("DocEntry").Cells(eventInfo.Row).Specific.Value)
            oLast_RightClick_InFlag = Trim(oMat01.Columns("InStatus").Cells(eventInfo.Row).Specific.Value)
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
    Set oDS_PS_MM150H = oForm01.DataSources.DBDataSources("@PS_MM150H")
    
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

Sub ComboBox_Setting(Optional MenuUID$)
On Error GoTo ComboBox_Setting_Error
    '//콤보에 기본값설정
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    '// 사업장
    Set oCombo = oForm01.Items("BPLId").Specific
    sQry = "SELECT BPLId, BPLName From [OBPL]"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).Value), Trim(oRecordSet01.Fields(1).Value)
        oRecordSet01.MoveNext
    Loop
    
    '//작업구분
    Set oCombo = oForm01.Items("OutGbn").Specific
    oCombo.ValidValues.Add "%", "전체"
    oCombo.ValidValues.Add "10", "원재료"
    oCombo.ValidValues.Add "20", "재공"
    oCombo.Select "10", psk_ByValue
    
    '//상태
    Set oCombo = oForm01.Items("Status").Specific
    If MenuUID = "PS_MM150_1" Then '자체 확인
        oCombo.ValidValues.Add "Y", "출고"
    Else
        oCombo.ValidValues.Add "N", "미출고"
        oCombo.ValidValues.Add "Y", "출고"
    End If
    oCombo.Select 0, psk_Index
    
    '//입고확인
    Set oCombo = oForm01.Items("InStatus").Specific
    oCombo.ValidValues.Add "N", "미확인"
    oCombo.ValidValues.Add "Y", "확인"
    oCombo.Select 0, psk_Index
    oMat01.Columns("InStatus").ValidValues.Add "N", "미확인"
    oMat01.Columns("InStatus").ValidValues.Add "Y", "확인"
    
    '//완료여부
    oMat01.Columns("ReStatus").ValidValues.Add "N", "미완료"
    oMat01.Columns("ReStatus").ValidValues.Add "Y", "완료"
    
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

Sub FormItemEnabled(Optional MenuUID$)
On Error GoTo FormItemEnabled_Error
    If MenuUID = "PS_MM150_1" Then '자체 확인
        If Trim(oForm01.Items("Status").Specific.Selected.Value) = "Y" Then
            oForm01.Items("Mat01").Enabled = True
            oMat01.Columns("Check").Editable = False
            oMat01.Columns("OutQty").Editable = False
            oMat01.Columns("OutWt").Editable = False
            oMat01.Columns("DocDate").Editable = False
            oMat01.Columns("ReStatus").Editable = False
            
            oMat01.Columns("InStatus").Editable = True
        ElseIf Trim(oForm01.Items("Status").Specific.Selected.Value) = "N" Then
            oForm01.Items("Mat01").Enabled = True
            oMat01.Columns("Check").Editable = False
            oMat01.Columns("OutQty").Editable = False
            oMat01.Columns("OutWt").Editable = False
            oMat01.Columns("DocDate").Editable = False
            oMat01.Columns("ReStatus").Editable = False
            
            oMat01.Columns("InStatus").Editable = True
        End If
    Else '외주 입력
        If Trim(oForm01.Items("Status").Specific.Selected.Value) = "Y" Then
            oForm01.Items("Mat01").Enabled = True
            oMat01.Columns("Check").Editable = False
            oMat01.Columns("OutQty").Editable = False
            oMat01.Columns("OutWt").Editable = False
            oMat01.Columns("DocDate").Editable = False
            oMat01.Columns("ReStatus").Editable = True
            
            oMat01.Columns("InStatus").Editable = False
        ElseIf Trim(oForm01.Items("Status").Specific.Selected.Value) = "N" Then
            oForm01.Items("Mat01").Enabled = True
            oMat01.Columns("Check").Editable = True
            oMat01.Columns("OutQty").Editable = True
            oMat01.Columns("OutWt").Editable = True
            oMat01.Columns("DocDate").Editable = True
            oMat01.Columns("ReStatus").Editable = True
            
            oMat01.Columns("InStatus").Editable = False
        End If
    End If
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    If oForm01 Is Nothing Then
    Else
        oForm01.Freeze False
    End If
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_MM150H.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_MM150H.Offset = oRow
    oDS_PS_MM150H.setValue "DocNum", oRow, oRow + 1
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
    If oForm01 Is Nothing Then
    Else
        oForm01.Freeze False
    End If
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
'            Case oDS_PS_MM150H.GetValue("U_ItemCode", i) = ""
'                ErrNum = 2
'                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_MM150H.GetValue("U_Qty", i) = ""
'                ErrNum = 3
'                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_MM150H.GetValue("U_DocDate", i) = ""
'                ErrNum = 4
'                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_MM150H.GetValue("U_DueDate", i) = ""
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
'        OutDoc = Trim(oDS_PS_MM150H.GetValue("U_OutDoc", i))
'        OutLine = Trim(oDS_PS_MM150H.GetValue("U_OutLine", i))
'        ItemCode = Trim(oDS_PS_MM150H.GetValue("U_ItemCode", i))
'        ItemName = Trim(oDS_PS_MM150H.GetValue("U_ItemName", i))
'        Size = Trim(oDS_PS_MM150H.GetValue("U_Size", i))
'        Mark = Trim(oDS_PS_MM150H.GetValue("U_Mark", i))
'        OutItmCd = Trim(oDS_PS_MM150H.GetValue("U_OutItmCd", i))
'        OutItmNm = Trim(oDS_PS_MM150H.GetValue("U_OutItmNm", i))
'        OutGbn = Trim(oDS_PS_MM150H.GetValue("U_OutGbn", i))
'        OutQty = Trim(oDS_PS_MM150H.GetValue("U_OutQty", i))
'        OutWt = Trim(oDS_PS_MM150H.GetValue("U_OutWt", i))
'        InQty = Trim(oDS_PS_MM150H.GetValue("U_InQty", i))
'        InWt = Trim(oDS_PS_MM150H.GetValue("U_InWt", i))
'        JakNum = Trim(oDS_PS_MM150H.GetValue("U_JakNum", i))
'        DocDate = Trim(oDS_PS_MM150H.GetValue("U_DocDate", i))
'
''        If Trim(oDS_PS_MM150H.GetValue("U_SjQty", i)) = "" Then
''            SjQty = 0
''        Else
''            SjQty = Trim(oDS_PS_MM150H.GetValue("U_SjQty", i))
''        End If
'
'        '// DocEntry
'        sQry = "Select IsNULL(Max(DocEntry), 0) From [@PS_MM150H] "
'        RecordSet01.DoQuery sQry
'        If Trim(RecordSet01.Fields(0).Value) = 0 Then
'            DocEntry = 1
'        Else
'            DocEntry = Trim(RecordSet01.Fields(0).Value) + 1
'        End If
'
'        sQry = "INSERT INTO [@PS_MM150H]"
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
    
    Dim BPLId$, OutDoc$, OutGbn$, CardCode$, DocDateFr$, DocDateTo$, Status$, InStatus$
        
    BPLId = Trim(oForm01.Items("BPLId").Specific.Value)
    OutDoc = Trim(oForm01.Items("OutDoc").Specific.Value)
    OutGbn = Trim(oForm01.Items("OutGbn").Specific.Value)
    CardCode = Trim(oForm01.Items("CardCode").Specific.Value)
    DocDateFr = Trim(oForm01.Items("DocDateFr").Specific.Value)
    DocDateTo = Trim(oForm01.Items("DocDateTo").Specific.Value)
    Status = Trim(oForm01.Items("Status").Specific.Value)
    InStatus = Trim(oForm01.Items("InStatus").Specific.Value)
    
    If BPLId = "" Then BPLId = "%"
    If OutDoc = "" Then OutDoc = "%"
    If OutGbn = "" Then OutGbn = "%"
    If CardCode = "" Then CardCode = "%"
    If DocDateFr = "" Then DocDateFr = "19000101"
    If DocDateTo = "" Then DocDateTo = "20991231"
    
    sQry = "EXEC [PS_MM150_01] '" & BPLId & "','" & OutDoc & "','" & OutGbn & "','" & CardCode & "','" & DocDateFr & "','" & DocDateTo & "','" & Status & "','" & InStatus & "'"
    oRecordSet01.DoQuery sQry
    
    oMat01.Clear
    oDS_PS_MM150H.Clear
    
    If oRecordSet01.RecordCount = 0 Then
        MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
        Set oRecordSet01 = Nothing
        Exit Sub
    End If
    
'    oForm01.Freeze True
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_MM150H.Size Then
            oDS_PS_MM150H.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_MM150H.Offset = i
        oDS_PS_MM150H.setValue "DocNum", i, i + 1
        oDS_PS_MM150H.setValue "U_OutDoc", i, Trim(oRecordSet01.Fields("U_OutDoc").Value)
        oDS_PS_MM150H.setValue "U_OutLine", i, Trim(oRecordSet01.Fields("U_OutLine").Value)
        oDS_PS_MM150H.setValue "U_ItemCode", i, Trim(oRecordSet01.Fields("U_ItemCode").Value)
        oDS_PS_MM150H.setValue "U_ItemName", i, Trim(oRecordSet01.Fields("U_ItemName").Value)
        oDS_PS_MM150H.setValue "U_Size", i, Trim(oRecordSet01.Fields("U_Size").Value)
        oDS_PS_MM150H.setValue "U_Mark", i, Trim(oRecordSet01.Fields("U_Mark").Value)
        oDS_PS_MM150H.setValue "U_OutItmCd", i, Trim(oRecordSet01.Fields("U_OutItmCd").Value)
        oDS_PS_MM150H.setValue "U_OutItmNm", i, Trim(oRecordSet01.Fields("U_OutItmNm").Value)
        oDS_PS_MM150H.setValue "U_OutGbn", i, Trim(oRecordSet01.Fields("U_OutGbn").Value)
'        oDS_PS_MM150H.setValue "U_OutWhCd", i, Trim(oRecordset01.Fields("U_OutWhCd").Value)
'        oDS_PS_MM150H.setValue "U_OutWhNm", i, Trim(oRecordset01.Fields("U_OutWhNm").Value)
'        oDS_PS_MM150H.setValue "U_InWhCd", i, Trim(oRecordset01.Fields("U_InWhCd").Value)
'        oDS_PS_MM150H.setValue "U_InWhNm", i, Trim(oRecordset01.Fields("U_InWhNm").Value)
        oDS_PS_MM150H.setValue "U_InQty", i, Trim(oRecordSet01.Fields("U_InQty").Value)
        oDS_PS_MM150H.setValue "U_InWt", i, Trim(oRecordSet01.Fields("U_InWt").Value)
'        oDS_PS_MM150H.setValue "U_OutDate", i, Format(Trim(oRecordset01.Fields("U_OutDate").Value), "YYYYMMDD")
        oDS_PS_MM150H.setValue "U_JakNum", i, Trim(oRecordSet01.Fields("U_JakNum").Value)
        oDS_PS_MM150H.setValue "U_CardCode", i, Trim(oRecordSet01.Fields("U_CardCode").Value)
        oDS_PS_MM150H.setValue "U_CardName", i, Trim(oRecordSet01.Fields("U_CardName").Value)
        
        If Trim(oForm01.Items("Status").Specific.Selected.Value) = "Y" Then
            oDS_PS_MM150H.setValue "U_ScrapWt", i, Trim(oRecordSet01.Fields("U_ScrapWt").Value)
            oDS_PS_MM150H.setValue "U_DocDate", i, Format(Trim(oRecordSet01.Fields("U_DocDate").Value), "YYYYMMDD")
            oDS_PS_MM150H.setValue "U_OutQty", i, Trim(oRecordSet01.Fields("U_OutQty").Value)
            oDS_PS_MM150H.setValue "U_OutWt", i, Trim(oRecordSet01.Fields("U_OutWt").Value)
            oDS_PS_MM150H.setValue "U_MM140Doc", i, Trim(oRecordSet01.Fields("U_MM140Doc").Value)
            oDS_PS_MM150H.setValue "U_InStatus", i, Trim(oRecordSet01.Fields("U_InStatus").Value)
            oDS_PS_MM150H.setValue "U_ReStatus", i, Trim(oRecordSet01.Fields("U_ReStatus").Value)
            oDS_PS_MM150H.setValue "U_InFlag", i, Trim(oRecordSet01.Fields("U_InFlag").Value)
            oDS_PS_MM150H.setValue "U_ReFlag", i, Trim(oRecordSet01.Fields("U_ReFlag").Value)
            oDS_PS_MM150H.setValue "DocEntry", i, Trim(oRecordSet01.Fields("DocEntry").Value)
            
            oDS_PS_MM150H.setValue "U_CPWt", i, Trim(oRecordSet01.Fields("U_CPWt").Value)
            oDS_PS_MM150H.setValue "U_CPWtName", i, Trim(oRecordSet01.Fields("U_CPWtName").Value)
            oDS_PS_MM150H.setValue "U_MUseQty", i, Trim(oRecordSet01.Fields("U_MUseQty").Value)
            oDS_PS_MM150H.setValue "U_MUseWt", i, Trim(oRecordSet01.Fields("U_MUseWt").Value)
            oDS_PS_MM150H.setValue "U_MTUseQty", i, Trim(oRecordSet01.Fields("U_MTUseQty").Value)
            oDS_PS_MM150H.setValue "U_MTUseWt", i, Trim(oRecordSet01.Fields("U_MTUseWt").Value)
            oDS_PS_MM150H.setValue "U_MDUseQty", i, Trim(oRecordSet01.Fields("U_MDUseQty").Value)
            oDS_PS_MM150H.setValue "U_MDUseWt", i, Trim(oRecordSet01.Fields("U_MDUseWt").Value)
        ElseIf Trim(oForm01.Items("Status").Specific.Selected.Value) = "N" Then
            oDS_PS_MM150H.setValue "U_ScrapWt", i, 0
            oDS_PS_MM150H.setValue "U_DocDate", i, Format(Date, "YYYYMMDD")
            oDS_PS_MM150H.setValue "U_MM140Doc", i, Trim(oRecordSet01.Fields("DocNum").Value)
            oDS_PS_MM150H.setValue "U_InStatus", i, "N"
            oDS_PS_MM150H.setValue "U_ReStatus", i, "N"
            oDS_PS_MM150H.setValue "U_InFlag", i, "N"
            oDS_PS_MM150H.setValue "U_ReFlag", i, "N"
            
            oDS_PS_MM150H.setValue "U_CPWt", i, 0
            oDS_PS_MM150H.setValue "U_CPWtName", i, ""
            oDS_PS_MM150H.setValue "U_MUseQty", i, 0
            oDS_PS_MM150H.setValue "U_MUseWt", i, 0
            oDS_PS_MM150H.setValue "U_MTUseQty", i, Trim(oRecordSet01.Fields("MTUseQty").Value)
            oDS_PS_MM150H.setValue "U_MTUseWt", i, Trim(oRecordSet01.Fields("MTUseWt").Value)
            oDS_PS_MM150H.setValue "U_MDUseQty", i, Trim(oRecordSet01.Fields("MDUseQty").Value)
            oDS_PS_MM150H.setValue "U_MDUseWt", i, Trim(oRecordSet01.Fields("MDUseQty").Value)
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
    If oForm01 Is Nothing Then
    Else
        oForm01.Freeze False
    End If
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "LoadData_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Function Add_PS_PP040(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
On Error GoTo Add_PS_PP040_Error
    Dim i&, j&
    Dim ErrNum&, RowNum&
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim lTypeCount As Integer
    
    Dim BPLId$, JakNum$, OutQty As Long, OutWt As Currency
    Dim PP040H_DocEntry$, PP030M_LineId$, PP030M_U_Sequence$, PP030M_U_CpCode$, PP030M_U_CpName$, DocType$, AutoKey&
    Dim iTemp01&, sTemp01$
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    oMat01.FlushToDataSource
    
    BPLId = Trim(oForm01.Items("BPLId").Specific.Value)
    For i = 0 To oMat01.RowCount - 1
        If Trim(oDS_PS_MM150H.GetValue("U_InStatus", i)) = "Y" And Trim(oDS_PS_MM150H.GetValue("U_InFlag", i)) <> "Y" Then
            JakNum = Trim(oDS_PS_MM150H.GetValue("U_JakNum", i))
            OutQty = Trim(oDS_PS_MM150H.GetValue("U_OutQty", i))
            OutWt = Trim(oDS_PS_MM150H.GetValue("U_OutWt", i))
            
            sQry = "EXEC [PS_MM150_02] '" & JakNum & "','10'"
            oRecordSet01.DoQuery sQry
            
            If oRecordSet01.RecordCount = 0 Then
                ErrNum = 1
                GoTo Add_PS_PP040_Error
            End If
                        
            lTypeCount = 0
            Do Until oRecordSet01.EOF
                ReDim Preserve PS_PP040(lTypeCount)
                
                PS_PP040(lTypeCount).PP030HNo = Trim(oRecordSet01.Fields("DocEntry").Value)
                PS_PP040(lTypeCount).PP030MNo = Trim(oRecordSet01.Fields("LineId").Value)
                PS_PP040(lTypeCount).OrdNum = JakNum
                PS_PP040(lTypeCount).Sequence = Trim(oRecordSet01.Fields("U_Sequence").Value)
                PS_PP040(lTypeCount).OrdGbn = Trim(oRecordSet01.Fields("U_OrdGbn").Value)
                PS_PP040(lTypeCount).ItemCode = Trim(oRecordSet01.Fields("U_ItemCode").Value)
                PS_PP040(lTypeCount).ItemName = Trim(oRecordSet01.Fields("U_ItemName").Value)
                PS_PP040(lTypeCount).CpCode = Trim(oRecordSet01.Fields("U_CpCode").Value)
                PS_PP040(lTypeCount).CpName = Trim(oRecordSet01.Fields("U_CpName").Value)
                PS_PP040(lTypeCount).Chk = "N"
                
                lTypeCount = lTypeCount + 1
                oRecordSet01.MoveNext
            Loop
                        
            '// DocEntry
            sQry = "Select AutoKey From [ONNM] Where ObjectCode = 'PS_PP040'"
            oRecordSet01.DoQuery sQry
            PP040H_DocEntry = Trim(oRecordSet01.Fields("AutoKey").Value)
            AutoKey = PP040H_DocEntry + 1
            
            ReDim Preserve PS_PP040DocEntry(i)
            PS_PP040DocEntry(i).PP040DocEntry = PP040H_DocEntry
            
            '//문서타입 하드코딩
            If Trim(PS_PP040(0).OrdGbn) = "104" Or Trim(PS_PP040(0).OrdGbn) = "107" Then
                DocType = "20"
            Else
                DocType = "10"
            End If
            
            '// Insert PS_PP040H
            sQry = "INSERT INTO [@PS_PP040H]"
            sQry = sQry & " ("
            sQry = sQry & " DocEntry,"
            sQry = sQry & " DocNum,"
            sQry = sQry & " Period,"
            sQry = sQry & " Series,"
            sQry = sQry & " Object,"
            sQry = sQry & " UserSign,"
            sQry = sQry & " CreateDate,"
            sQry = sQry & " CreateTime,"
            sQry = sQry & " DataSource,"
            sQry = sQry & " U_OrdType,"
            sQry = sQry & " U_OrdGbn,"
            sQry = sQry & " U_BPLId,"
            sQry = sQry & " U_ItemCode,"
            sQry = sQry & " U_ItemName,"
            sQry = sQry & " U_OrdNum,"
            sQry = sQry & " U_OrdSub1,"
            sQry = sQry & " U_OrdSub2,"
            sQry = sQry & " U_PP030HNo,"
            sQry = sQry & " U_DocType,"
            sQry = sQry & " Canceled,"
            sQry = sQry & " U_DocDate"
            sQry = sQry & " ) "
            
            sQry = sQry & "VALUES("
            sQry = sQry & "'" & PP040H_DocEntry & "',"
            sQry = sQry & "'" & PP040H_DocEntry & "',"
            sQry = sQry & "'22',"
            sQry = sQry & "'-1',"
            sQry = sQry & "'PS_PP040',"
            sQry = sQry & "'1',"
            sQry = sQry & "'" & Format(Date, "YYYYMMDD") & "',"
            sQry = sQry & "'1000',"
            sQry = sQry & "'I',"
            sQry = sQry & "'30',"
            sQry = sQry & "'" & PS_PP040(0).OrdGbn & "',"
            sQry = sQry & "'" & BPLId & "',"
            sQry = sQry & "'" & PS_PP040(0).ItemCode & "',"
            sQry = sQry & "'" & PS_PP040(0).ItemName & "',"
            sQry = sQry & "'" & JakNum & "',"
            sQry = sQry & "'000',"
            sQry = sQry & "'00',"
            sQry = sQry & "'" & PS_PP040(0).PP030HNo & "',"
            sQry = sQry & "'" & DocType & "',"
            sQry = sQry & "'N',"
            sQry = sQry & "'" & Format(Date, "YYYYMMDD") & "'"
            sQry = sQry & ")"
            oRecordSet01.DoQuery sQry
            
            '// Insert PS_PP040M
            sQry = "INSERT INTO [@PS_PP040M]"
            sQry = sQry & " ("
            sQry = sQry & " DocEntry,"
            sQry = sQry & " LineId,"
            sQry = sQry & " VisOrder,"
            sQry = sQry & " Object,"
            sQry = sQry & " U_LineNum,"
            sQry = sQry & " U_LineId,"
            sQry = sQry & " U_WorkCode,"
            sQry = sQry & " U_WorkName"
            sQry = sQry & " ) "
            
            sQry = sQry & "VALUES("
            sQry = sQry & "'" & PP040H_DocEntry & "',"
            sQry = sQry & "'1',"
            sQry = sQry & "'0',"
            sQry = sQry & "'PS_PP040',"
            sQry = sQry & "'1',"
            sQry = sQry & "'1',"
            sQry = sQry & "'9999999',"
            sQry = sQry & "'조정'"
            sQry = sQry & ")"
            oRecordSet01.DoQuery sQry
                
            For j = 0 To (UBound(PS_PP040))
                If PS_PP040(j).Chk = "N" Then
                    iTemp01 = j + 1
                    sTemp01 = PS_PP040(j).PP030HNo & "-" & PS_PP040(j).PP030MNo
                                    
                    '// Insert PS_PP040L
                    sQry = "INSERT INTO [@PS_PP040L]"
                    sQry = sQry & " ("
                    sQry = sQry & " DocEntry,"
                    sQry = sQry & " LineId,"
                    sQry = sQry & " VisOrder,"
                    sQry = sQry & " Object,"
                    sQry = sQry & " U_LineNum,"
                    sQry = sQry & " U_LineId,"
                    sQry = sQry & " U_OrdMgNum,"
                    sQry = sQry & " U_Sequence,"
                    sQry = sQry & " U_CpCode,"
                    sQry = sQry & " U_CpName,"
                    sQry = sQry & " U_OrdGbn,"
                    sQry = sQry & " U_BPLId,"
                    sQry = sQry & " U_ItemCode,"
                    sQry = sQry & " U_ItemName,"
                    sQry = sQry & " U_OrdNum,"
                    sQry = sQry & " U_OrdSub1,"
                    sQry = sQry & " U_OrdSub2,"
                    sQry = sQry & " U_PP030HNo,"
                    sQry = sQry & " U_PP030MNo,"
                    sQry = sQry & " U_PQty,"
                    sQry = sQry & " U_PWeight,"
                    sQry = sQry & " U_YQty,"
                    sQry = sQry & " U_YWeight,"
                    sQry = sQry & " U_NQty,"
                    sQry = sQry & " U_NWeight,"
                    sQry = sQry & " U_PSum"
                    sQry = sQry & " ) "
                    
                    sQry = sQry & "VALUES("
                    sQry = sQry & "'" & PP040H_DocEntry & "',"
                    sQry = sQry & "'" & iTemp01 & "',"
                    sQry = sQry & "'" & j & "',"
                    sQry = sQry & "'PS_PP040',"
                    sQry = sQry & "'" & iTemp01 & "',"
                    sQry = sQry & "'" & iTemp01 & "',"
                    sQry = sQry & "'" & sTemp01 & "',"
                    sQry = sQry & "'" & PS_PP040(j).Sequence & "',"
                    sQry = sQry & "'" & PS_PP040(j).CpCode & "',"
                    sQry = sQry & "'" & PS_PP040(j).CpName & "',"
                    sQry = sQry & "'" & PS_PP040(j).OrdGbn & "',"
                    sQry = sQry & "'" & BPLId & "',"
                    sQry = sQry & "'" & PS_PP040(j).ItemCode & "',"
                    sQry = sQry & "'" & PS_PP040(j).ItemName & "',"
                    sQry = sQry & "'" & JakNum & "',"
                    sQry = sQry & "'000',"
                    sQry = sQry & "'00',"
                    sQry = sQry & "'" & PS_PP040(j).PP030HNo & "',"
                    sQry = sQry & "'" & PS_PP040(j).PP030MNo & "',"
                    sQry = sQry & "'" & OutQty & "',"
                    sQry = sQry & "'" & OutWt & "',"
                    sQry = sQry & "'" & OutQty & "',"
                    sQry = sQry & "'" & OutWt & "',"
                    sQry = sQry & "'" & 0 & "',"
                    sQry = sQry & "'" & 0 & "',"
                    sQry = sQry & "'" & 0 & "'"
                    sQry = sQry & ")"
                    oRecordSet01.DoQuery sQry
                    
                    '// Insert PS_PP040N
                    sQry = "INSERT INTO [@PS_PP040N]"
                    sQry = sQry & " ("
                    sQry = sQry & " DocEntry,"
                    sQry = sQry & " LineId,"
                    sQry = sQry & " VisOrder,"
                    sQry = sQry & " Object,"
                    sQry = sQry & " U_LineNum,"
                    sQry = sQry & " U_LineId,"
                    sQry = sQry & " U_OrdMgNum,"
                    sQry = sQry & " U_CpCode,"
                    sQry = sQry & " U_CpName"
                    sQry = sQry & " ) "
                    
                    sQry = sQry & "VALUES("
                    sQry = sQry & "'" & PP040H_DocEntry & "',"
                    sQry = sQry & "'" & iTemp01 & "',"
                    sQry = sQry & "'" & j & "',"
                    sQry = sQry & "'PS_PP040',"
                    sQry = sQry & "'" & iTemp01 & "',"
                    sQry = sQry & "'" & iTemp01 & "',"
                    sQry = sQry & "'" & sTemp01 & "',"
                    sQry = sQry & "'" & PS_PP040(j).CpCode & "',"
                    sQry = sQry & "'" & PS_PP040(j).CpName & "'"
                    sQry = sQry & ")"
                    oRecordSet01.DoQuery sQry
                    
                    PS_PP040(j).Chk = "Y"
                End If
            Next
            sQry = "Update [ONNM] Set AutoKey = '" & AutoKey & "' Where ObjectCode = 'PS_PP040'"
            oRecordSet01.DoQuery sQry
        End If
    Next
    
    MDC_Com.MDC_GF_Message "외주업체 출고등록 완료!", "S"
    
    Set oRecordSet01 = Nothing
    Add_PS_PP040 = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_PS_PP040_Error:
    If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_RollBack
    Add_PS_PP040 = False
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "작업지시가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
    Else
        MDC_Com.MDC_GF_Message "Add_PS_PP040_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Function

Public Function Add_PS_MM150H(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
On Error GoTo Add_PS_MM150H_Error
    Dim i As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    
    Dim DocEntry$, BPLId$, OutDoc$, OutLine$, ItemCode$, ItemName$, Size$, Mark$, OutItmCd$, OutItmNm$, OutGbn$, CardCode, CardName$, OutDate$
    Dim OutQty As Long, OutWt As Currency, InQty As Long, InWt As Currency, JakNum$, DocDate$, InStatus$, ReStatus$, MM140Doc$, PP040Doc$, ScrapWt As Currency
    Dim CPWt As Currency, CPWtName As String, MUseQty As Currency, MUseWt As Currency, MTUseQty As Currency, MTUseWt As Currency, MDUseQty As Currency, MDUseWt As Currency
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    oMat01.FlushToDataSource
    
    BPLId = Trim(oForm01.Items("BPLId").Specific.Value)
    For i = 0 To oMat01.RowCount - 1
        If Trim(oDS_PS_MM150H.GetValue("U_Check", i)) = "Y" Then
            InStatus = Trim(oDS_PS_MM150H.GetValue("U_InStatus", i))
            OutDoc = Trim(oDS_PS_MM150H.GetValue("U_OutDoc", i))
            OutLine = Trim(oDS_PS_MM150H.GetValue("U_OutLine", i))
            ItemCode = Trim(oDS_PS_MM150H.GetValue("U_ItemCode", i))
            ItemName = Trim(oDS_PS_MM150H.GetValue("U_ItemName", i))
            Size = Trim(oDS_PS_MM150H.GetValue("U_Size", i))
            Mark = Trim(oDS_PS_MM150H.GetValue("U_Mark", i))
            OutItmCd = Trim(oDS_PS_MM150H.GetValue("U_OutItmCd", i))
            OutItmNm = Trim(oDS_PS_MM150H.GetValue("U_OutItmNm", i))
            OutGbn = Trim(oDS_PS_MM150H.GetValue("U_OutGbn", i))
    '        OutWhCd = Trim(oDS_PS_MM150H.GetValue("U_OutWhCd", i))
    '        OutWhNm = Trim(oDS_PS_MM150H.GetValue("U_OutWhNm", i))
    '        InWhCd = Trim(oDS_PS_MM150H.GetValue("U_InWhCd", i))
    '        InWhNm = Trim(oDS_PS_MM150H.GetValue("U_InWhNm", i))
            InQty = Trim(oDS_PS_MM150H.GetValue("U_InQty", i))
            InWt = Trim(oDS_PS_MM150H.GetValue("U_InWt", i))
'            OutQty = Trim(oDS_PS_MM150H.GetValue("U_OutQty", i))
            If Trim(oDS_PS_MM150H.GetValue("OutQty", i)) = "" Then
                OutQty = 0
            Else
                OutQty = Trim(oDS_PS_MM150H.GetValue("U_SjQty", i))
            End If
            OutWt = Trim(oDS_PS_MM150H.GetValue("U_OutWt", i))
            ReStatus = Trim(oDS_PS_MM150H.GetValue("U_ReStatus", i))
            JakNum = Trim(oDS_PS_MM150H.GetValue("U_JakNum", i))
'            OutDate = Trim(oDS_PS_MM150H.GetValue("U_OutDate", i))
            DocDate = Trim(oDS_PS_MM150H.GetValue("U_DocDate", i))
            CardCode = Trim(oDS_PS_MM150H.GetValue("U_CardCode", i))
            CardName = Trim(oDS_PS_MM150H.GetValue("U_CardName", i))
            MM140Doc = Trim(oDS_PS_MM150H.GetValue("U_MM140Doc", i))
            PP040Doc = ""
'            ScrapWt = Trim(oDS_PS_MM150H.GetValue("U_ScrapWt", i))
    
    '        If Trim(oDS_PS_MM150H.GetValue("U_SjQty", i)) = "" Then
    '            SjQty = 0
    '        Else
    '            SjQty = Trim(oDS_PS_MM150H.GetValue("U_SjQty", i))
    '        End If
            
            CPWt = Trim(oDS_PS_MM150H.GetValue("U_CPWt", i))
            CPWtName = Trim(oDS_PS_MM150H.GetValue("U_CPWtName", i))
            MUseQty = Trim(oDS_PS_MM150H.GetValue("U_MUseQty", i))
            MUseWt = Trim(oDS_PS_MM150H.GetValue("U_MUseWt", i))
            MTUseQty = Trim(oDS_PS_MM150H.GetValue("U_MTUseQty", i))
            MTUseWt = Trim(oDS_PS_MM150H.GetValue("U_MTUseWt", i))
            MDUseQty = Trim(oDS_PS_MM150H.GetValue("U_MDUseQty", i))
            MDUseWt = Trim(oDS_PS_MM150H.GetValue("U_MDUseWt", i))
            
            '// DocEntry
            sQry = "Select IsNULL(Max(DocEntry), 0) From [@PS_MM150H] "
            oRecordSet01.DoQuery sQry
            If Trim(oRecordSet01.Fields(0).Value) = 0 Then
                DocEntry = 1
            Else
                DocEntry = Trim(oRecordSet01.Fields(0).Value) + 1
            End If
            
            sQry = "INSERT INTO [@PS_MM150H]"
            sQry = sQry & " ("
            sQry = sQry & " DocEntry,"
            sQry = sQry & " DocNum,"
            sQry = sQry & " U_BPLId,"
            sQry = sQry & " U_InStatus,"
            sQry = sQry & " U_OutDoc,"
            sQry = sQry & " U_OutLine,"
            sQry = sQry & " U_ItemCode,"
            sQry = sQry & " U_ItemName,"
            sQry = sQry & " U_Size,"
            sQry = sQry & " U_Mark,"
            sQry = sQry & " U_OutItmCd,"
            sQry = sQry & " U_OutItmNm,"
            sQry = sQry & " U_OutGbn,"
'            sQry = sQry & " U_OutWhCd,"
'            sQry = sQry & " U_OutWhNm,"
'            sQry = sQry & " U_InWhCd,"
'            sQry = sQry & " U_InWhNm,"
            sQry = sQry & " U_InQty,"
            sQry = sQry & " U_InWt,"
            sQry = sQry & " U_OutQty,"
            sQry = sQry & " U_OutWt,"
            sQry = sQry & " U_CPWt,"
            sQry = sQry & " U_CPWtName,"
            sQry = sQry & " U_MUseQty,"
            sQry = sQry & " U_MUseWt,"
            sQry = sQry & " U_MTUseQty,"
            sQry = sQry & " U_MTUseWt,"
            sQry = sQry & " U_MDUseQty,"
            sQry = sQry & " U_MDUseWt,"
'            sQry = sQry & " U_ScrapWt,"
            sQry = sQry & " U_ReStatus,"
            sQry = sQry & " U_JakNum,"
'            sQry = sQry & " U_OutDate,"
            sQry = sQry & " U_DocDate,"
            sQry = sQry & " U_CardCode,"
            sQry = sQry & " U_CardName,"
            sQry = sQry & " U_MM140Doc,"
            sQry = sQry & " U_PP040Doc"
            sQry = sQry & " ) "
            sQry = sQry & "VALUES("
            sQry = sQry & DocEntry & ","
            sQry = sQry & DocEntry & ","
            sQry = sQry & "'" & BPLId & "',"
            sQry = sQry & "'" & InStatus & "',"
            sQry = sQry & "'" & OutDoc & "',"
            sQry = sQry & "'" & OutLine & "',"
            sQry = sQry & "'" & ItemCode & "',"
            sQry = sQry & "'" & ItemName & "',"
            sQry = sQry & "'" & Size & "',"
            sQry = sQry & "'" & Mark & "',"
            sQry = sQry & "'" & OutItmCd & "',"
            sQry = sQry & "'" & OutItmNm & "',"
            sQry = sQry & "'" & OutGbn & "',"
'            sQry = sQry & "'" & OutWhCd & "',"
'            sQry = sQry & "'" & OutWhNm & "',"
'            sQry = sQry & "'" & InWhCd & "',"
'            sQry = sQry & "'" & InWhNm & "',"
            sQry = sQry & "'" & InQty & "',"
            sQry = sQry & "'" & InWt & "',"
            sQry = sQry & "'" & OutQty & "',"
            sQry = sQry & "'" & OutWt & "',"
            sQry = sQry & "'" & CPWt & "',"
            sQry = sQry & "'" & CPWtName & "',"
            sQry = sQry & "'" & MUseQty & "',"
            sQry = sQry & "'" & MUseWt & "',"
            sQry = sQry & "'" & MTUseQty & "',"
            sQry = sQry & "'" & MTUseWt & "',"
            sQry = sQry & "'" & MDUseQty & "',"
            sQry = sQry & "'" & MDUseWt & "',"
'            sQry = sQry & "'" & ScrapWt & "',"
            sQry = sQry & "'" & ReStatus & "',"
            sQry = sQry & "'" & JakNum & "',"
'            sQry = sQry & "'" & OutDate & "',"
            sQry = sQry & "'" & DocDate & "',"
            sQry = sQry & "'" & CardCode & "',"
            sQry = sQry & "'" & CardName & "',"
            sQry = sQry & "'" & MM140Doc & "',"
            sQry = sQry & "'" & PP040Doc & "'"
            sQry = sQry & ")"
            oRecordSet01.DoQuery sQry
            
            If ReStatus = "Y" Then
                sQry = "Select Sum(IsNull(a.U_OutWt, 0) - IsNUll(b.U_ReWt, 0)) From [@PS_MM150H] a Left Join [@PS_MM155H] b On a.DocEntry =  b.U_MM150Doc "
                sQry = sQry & "Where a.U_OutDoc = '" & OutDoc & "' And a.U_OutLine = '" & OutLine & "'"
                oRecordSet01.DoQuery sQry
                ScrapWt = InWt - oRecordSet01.Fields(0).Value
                
                sQry = "Update [@PS_MM150H] Set U_ReFlag = 'Y', U_ScrapWt = '" & ScrapWt & "' Where DocEntry = '" & DocEntry & "'"
                oRecordSet01.DoQuery sQry
            End If
        End If
    Next
    
    MDC_Com.MDC_GF_Message "외주업체 출고등록 완료!", "S"
    
    
    
    Set oRecordSet01 = Nothing
    Add_PS_MM150H = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_PS_MM150H_Error:
    Add_PS_MM150H = False
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "Add_PS_MM150H_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Function Update_PS_MM150H(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
On Error GoTo Update_PS_MM150H_Error
    Dim i As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    
    Dim DocEntry$, BPLId$, OutDoc$, OutLine$, ItemCode$, ItemName$, Size$, Mark$, OutItmCd$, OutItmNm$, OutGbn$, CardCode, CardName$, OutDate$
    Dim OutQty As Long, OutWt As Currency, InQty As Long, InWt As Currency, JakNum$, DocDate$, InStatus$, ReStatus$, MM140Doc$, PP040Doc$, MM150Doc$, ScrapWt As Currency
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    oMat01.FlushToDataSource
    
    BPLId = Trim(oForm01.Items("BPLId").Specific.Value)
    
    If oMenuUID = "PS_MM150_1" Then
        For i = 0 To oMat01.RowCount - 1
            If Trim(oDS_PS_MM150H.GetValue("U_InStatus", i)) = "Y" And Trim(oDS_PS_MM150H.GetValue("U_InFlag", i)) <> "Y" Then
                InStatus = Trim(oDS_PS_MM150H.GetValue("U_InStatus", i))
                OutDoc = Trim(oDS_PS_MM150H.GetValue("U_OutDoc", i))
                OutLine = Trim(oDS_PS_MM150H.GetValue("U_OutLine", i))
                JakNum = Trim(oDS_PS_MM150H.GetValue("U_JakNum", i))
                PP040Doc = PS_PP040DocEntry(i).PP040DocEntry
                DocEntry = Trim(oDS_PS_MM150H.GetValue("DocEntry", i))
                
                sQry = "UPDATE [@PS_MM150H] "
                sQry = sQry & "SET "
                sQry = sQry & "U_InStatus = '" & InStatus & "', "
                sQry = sQry & "U_PP040Doc = '" & PP040Doc & "' "
                sQry = sQry & "Where DocEntry = '" & DocEntry & "'"
            
                oRecordSet01.DoQuery sQry
                
                If InStatus = "Y" Then
                    sQry = "Update [@PS_MM150H] Set U_InFlag = 'Y' Where DocEntry = '" & DocEntry & "'"
                    oRecordSet01.DoQuery sQry
                End If
            End If
        Next
    Else
        For i = 0 To oMat01.RowCount - 1
            If Trim(oDS_PS_MM150H.GetValue("U_ReStatus", i)) = "Y" And Trim(oDS_PS_MM150H.GetValue("U_ReFlag", i)) <> "Y" Then
                ReStatus = Trim(oDS_PS_MM150H.GetValue("U_ReStatus", i))
                OutDoc = Trim(oDS_PS_MM150H.GetValue("U_OutDoc", i))
                OutLine = Trim(oDS_PS_MM150H.GetValue("U_OutLine", i))
                JakNum = Trim(oDS_PS_MM150H.GetValue("U_JakNum", i))
                InWt = Trim(oDS_PS_MM150H.GetValue("U_InWt", i))
                OutWt = Trim(oDS_PS_MM150H.GetValue("U_OutWt", i))
                DocEntry = Trim(oDS_PS_MM150H.GetValue("DocEntry", i))
                    
                sQry = "UPDATE [@PS_MM150H] "
                sQry = sQry & "SET "
                sQry = sQry & "U_ReStatus = '" & ReStatus & "' "
                sQry = sQry & "Where DocEntry = '" & DocEntry & "'"
            
                oRecordSet01.DoQuery sQry
            
                If ReStatus = "Y" Then
                    sQry = "Select Sum(IsNull(a.U_OutWt, 0) - IsNUll(b.U_ReWt, 0)) From [@PS_MM150H] a Left Join [@PS_MM155H] b On a.DocEntry =  b.U_MM150Doc "
                    sQry = sQry & "Where a.U_OutDoc = '" & OutDoc & "' And a.U_OutLine = '" & OutLine & "'"
                    oRecordSet01.DoQuery sQry
                    ScrapWt = InWt - oRecordSet01.Fields(0).Value
                
                    sQry = "Update [@PS_MM150H] Set U_ReFlag = 'Y', U_ScrapWt = '" & ScrapWt & "' Where DocEntry = '" & DocEntry & "'"
                    oRecordSet01.DoQuery sQry
                End If
            End If
        Next
    End If
'        If Trim(oDS_PS_MM150H.GetValue("U_Check", i)) = "Y" Then
'            InStatus = Trim(oDS_PS_MM150H.GetValue("U_InStatus", i))
'            OutDoc = Trim(oDS_PS_MM150H.GetValue("U_OutDoc", i))
'            OutLine = Trim(oDS_PS_MM150H.GetValue("U_OutLine", i))
'            ItemCode = Trim(oDS_PS_MM150H.GetValue("U_ItemCode", i))
'            ItemName = Trim(oDS_PS_MM150H.GetValue("U_ItemName", i))
'            Size = Trim(oDS_PS_MM150H.GetValue("U_Size", i))
'            Mark = Trim(oDS_PS_MM150H.GetValue("U_Mark", i))
'            OutItmCd = Trim(oDS_PS_MM150H.GetValue("U_OutItmCd", i))
'            OutItmNm = Trim(oDS_PS_MM150H.GetValue("U_OutItmNm", i))
'            OutGbn = Trim(oDS_PS_MM150H.GetValue("U_OutGbn", i))
'    '        OutWhCd = Trim(oDS_PS_MM150H.GetValue("U_OutWhCd", i))
'    '        OutWhNm = Trim(oDS_PS_MM150H.GetValue("U_OutWhNm", i))
'    '        InWhCd = Trim(oDS_PS_MM150H.GetValue("U_InWhCd", i))
'    '        InWhNm = Trim(oDS_PS_MM150H.GetValue("U_InWhNm", i))
'            InQty = Trim(oDS_PS_MM150H.GetValue("U_InQty", i))
'            InWt = Trim(oDS_PS_MM150H.GetValue("U_InWt", i))
'            OutQty = Trim(oDS_PS_MM150H.GetValue("U_OutQty", i))
'            OutWt = Trim(oDS_PS_MM150H.GetValue("U_OutWt", i))
'            ReStatus = Trim(oDS_PS_MM150H.GetValue("U_ReStatus", i))
'            JakNum = Trim(oDS_PS_MM150H.GetValue("U_JakNum", i))
''            OutDate = Trim(oDS_PS_MM150H.GetValue("U_OutDate", i))
'            DocDate = Trim(oDS_PS_MM150H.GetValue("U_DocDate", i))
'            CardCode = Trim(oDS_PS_MM150H.GetValue("U_CardCode", i))
'            CardName = Trim(oDS_PS_MM150H.GetValue("U_CardName", i))
'            MM140Doc = Trim(oDS_PS_MM150H.GetValue("U_MM140Doc", i))
'            PP040Doc = PS_PP040DocEntry(i).PP040DocEntry
'
'    '        If Trim(oDS_PS_MM150H.GetValue("U_SjQty", i)) = "" Then
'    '            SjQty = 0
'    '        Else
'    '            SjQty = Trim(oDS_PS_MM150H.GetValue("U_SjQty", i))
'    '        End If
'
'            '// DocEntry
'            sQry = "Select IsNULL(Max(DocEntry), 0) From [@PS_MM150H] "
'            oRecordset01.DoQuery sQry
'            If Trim(oRecordset01.Fields(0).Value) = 0 Then
'                DocEntry = 1
'            Else
'                DocEntry = Trim(oRecordset01.Fields(0).Value) + 1
'            End If
'
'            sQry = "INSERT INTO [@PS_MM150H]"
'            sQry = sQry & " ("
'            sQry = sQry & " DocEntry,"
'            sQry = sQry & " DocNum,"
'            sQry = sQry & " U_BPLId,"
'            sQry = sQry & " U_InStatus,"
'            sQry = sQry & " U_OutDoc,"
'            sQry = sQry & " U_OutLine,"
'            sQry = sQry & " U_ItemCode,"
'            sQry = sQry & " U_ItemName,"
'            sQry = sQry & " U_Size,"
'            sQry = sQry & " U_Mark,"
'            sQry = sQry & " U_OutItmCd,"
'            sQry = sQry & " U_OutItmNm,"
'            sQry = sQry & " U_OutGbn,"
''            sQry = sQry & " U_OutWhCd,"
''            sQry = sQry & " U_OutWhNm,"
''            sQry = sQry & " U_InWhCd,"
''            sQry = sQry & " U_InWhNm,"
'            sQry = sQry & " U_InQty,"
'            sQry = sQry & " U_InWt,"
'            sQry = sQry & " U_OutQty,"
'            sQry = sQry & " U_OutWt,"
'            sQry = sQry & " U_ReStatus,"
'            sQry = sQry & " U_JakNum,"
''            sQry = sQry & " U_OutDate,"
'            sQry = sQry & " U_DocDate,"
'            sQry = sQry & " U_CardCode,"
'            sQry = sQry & " U_CardName,"
'            sQry = sQry & " U_MM140Doc,"
'            sQry = sQry & " U_PP040Doc"
'            sQry = sQry & " ) "
'            sQry = sQry & "VALUES("
'            sQry = sQry & DocEntry & ","
'            sQry = sQry & DocEntry & ","
'            sQry = sQry & "'" & BPLId & "',"
'            sQry = sQry & "'" & InStatus & "',"
'            sQry = sQry & "'" & OutDoc & "',"
'            sQry = sQry & "'" & OutLine & "',"
'            sQry = sQry & "'" & ItemCode & "',"
'            sQry = sQry & "'" & ItemName & "',"
'            sQry = sQry & "'" & Size & "',"
'            sQry = sQry & "'" & Mark & "',"
'            sQry = sQry & "'" & OutItmCd & "',"
'            sQry = sQry & "'" & OutItmNm & "',"
'            sQry = sQry & "'" & OutGbn & "',"
''            sQry = sQry & "'" & OutWhCd & "',"
''            sQry = sQry & "'" & OutWhNm & "',"
''            sQry = sQry & "'" & InWhCd & "',"
''            sQry = sQry & "'" & InWhNm & "',"
'            sQry = sQry & "'" & InQty & "',"
'            sQry = sQry & "'" & InWt & "',"
'            sQry = sQry & "'" & OutQty & "',"
'            sQry = sQry & "'" & OutWt & "',"
'            sQry = sQry & "'" & ReStatus & "',"
'            sQry = sQry & "'" & JakNum & "',"
''            sQry = sQry & "'" & OutDate & "',"
'            sQry = sQry & "'" & DocDate & "',"
'            sQry = sQry & "'" & CardCode & "',"
'            sQry = sQry & "'" & CardName & "',"
'            sQry = sQry & "'" & MM140Doc & "',"
'            sQry = sQry & "'" & PP040Doc & "'"
'            sQry = sQry & ")"
'            oRecordset01.DoQuery sQry
'        End If
    
    If oMenuUID = "PS_MM150_1" Then
        MDC_Com.MDC_GF_Message "작업일보 등록 완료!", "S"
    Else
        MDC_Com.MDC_GF_Message "외주업체 출고등록 갱신 완료!", "S"
    End If
        
    If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_Commit
    Set oRecordSet01 = Nothing
    Update_PS_MM150H = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Update_PS_MM150H_Error:
    If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_RollBack
    Update_PS_MM150H = False
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "Update_PS_MM150H_Error:" & Err.Number & " - " & Err.Description, "E"
End Function
