VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_QM008"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_QM008.cls
'//  Module : 품질관리>(분말)제품검사성적서 등록
'//  FormType : PS_QM008
'//  Create Date : 2014.10.30
'//  Modified Date :
'//  Creator : NGY
'//  Company : Poongsan Holdings
'****************************************************************************************************************
Option Explicit
Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_QM008H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_QM008L As SAPbouiCOM.DBDataSource '등록라인

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

'//사용자구조체
Private Type ItemInformations
    ItemCode As String
    BatchNum As String
    Quantity As Long
    OPORNo As Long
    POR1No As Long
    Check As Boolean
    OPDNNo As Long
    PDN1No As Long
End Type
Private ItemInformation() As ItemInformations
Private ItemInformationCount As Long

Private oDocType01 As String
Private oDocEntry01 As String
Private oFormMode01 As SAPbouiCOM.BoFormMode

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01 As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_QM008.srf")
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
    
    oFormUniqueID01 = "PS_QM008_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    oForm01.DataBrowser.BrowseBy = "DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_QM008_CreateItems
    Call PS_QM008_ComboBox_Setting
    Call PS_QM008_Initial_Setting
    Call PS_QM008_CF_ChooseFromList
    Call PS_QM008_EnableMenus
    Call PS_QM008_SetDocument(oFromDocEntry01)
    Call PS_QM008_FormResize
    
    oForm01.EnableMenu ("1283"), True         '// 삭제
    oForm01.EnableMenu ("1287"), True         '// 복제
    oForm01.EnableMenu ("1286"), True        '// 닫기
    oForm01.EnableMenu ("1284"), True        '// 취소
    oForm01.EnableMenu ("1293"), True         '// 행삭제
    
    oForm01.Update
    oForm01.Freeze False
    
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

    Dim i As Integer
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim InspNo As String
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    

    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1283": '삭제
                If oForm01.Mode = fm_OK_MODE Then
                    InspNo = oForm01.Items("InspNo").Specific.VALUE
                    '검사삭제시 검사완료 Update
                    sQry = "Update [Z_PACKING_PD] Set QM006YN = 'N' Where InspNo = '" & InspNo & "'"
                    oRecordSet.DoQuery sQry
                End If
                
            Case "1284": '취소
                If oForm01.Mode = fm_OK_MODE Then
                    InspNo = oForm01.Items("InspNo").Specific.VALUE
                    '검사삭제시 검사완료 Update
                    sQry = "Update [Z_PACKING_PD] Set QM006YN = 'N' Where InspNo = '" & InspNo & "'"
                    oRecordSet.DoQuery sQry
                End If
            Case "1286": '닫기
            Case "1293": '행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
                Call PS_QM008_FormItemEnabled '//UDO방식
            Case "1282": '추가
                Call PS_QM008_FormItemEnabled '//UDO방식
                Call PS_QM008_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                Call PS_QM008_FormItemEnabled
            Case "1287"
            
                oForm01.Freeze True
                Call oDS_PS_QM008H.setValue("DocEntry", 0, "")
                
                For i = 0 To oMat01.VisualRowCount - 1
                    Call oMat01.FlushToDataSource
                    Call oDS_PS_QM008H.setValue("DocEntry", i, "")
                    Call oMat01.LoadFromDataSource
                Next i
                oForm01.Freeze False
                
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
'        If pval.ItemUID = "Mat01" And pval.Row > 0 And pval.Row <= oMat01.RowCount Then
'            Dim MenuCreationParams01 As SAPbouiCOM.MenuCreationParams
'            Set MenuCreationParams01 = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
'            MenuCreationParams01.Type = SAPbouiCOM.BoMenuType.mt_STRING
'            MenuCreationParams01.uniqueID = "MenuUID"
'            MenuCreationParams01.String = "메뉴명"
'            MenuCreationParams01.Enabled = True
'            Call Sbo_Application.Menus.Item("1280").SubMenus.AddEx(MenuCreationParams01)
'        End If
    ElseIf pval.BeforeAction = False Then
'        If pval.ItemUID = "Mat01" And pval.Row > 0 And pval.Row <= oMat01.RowCount Then
'                Call Sbo_Application.Menus.RemoveEx("MenuUID")
'        End If
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
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim InspNo As String
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    
    If pval.BeforeAction = True Then
        If pval.ItemUID = "PS_QM008" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If PS_QM008_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                InspNo = oForm01.Items("InspNo").Specific.VALUE
                '//검사입력되면 검사완료 Sign Update
                
                sQry = "Update [Z_PACKING_PD] Set QM006YN = 'Y' Where InspNo = '" & InspNo & "'"
                oRecordSet.DoQuery sQry
                
                oDocEntry01 = oForm01.Items("DocEntry").Specific.VALUE
                oFormMode01 = oForm01.Mode
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                If PS_QM008_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                oDocEntry01 = oForm01.Items("DocEntry").Specific.VALUE
                oFormMode01 = oForm01.Mode
                
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "Button01" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                Call PS_QM008_Print_Report01
            ElseIf oForm01.Mode = fm_OK_MODE Then
                Call PS_QM008_Print_Report01
            End If
        End If
        If pval.ItemUID = "Button02" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                 Call PS_QM008_Print_Report02
            ElseIf oForm01.Mode = fm_OK_MODE Then
                Call PS_QM008_Print_Report02
            End If
        End If
        If pval.ItemUID = "Button03" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                Call PS_QM008_Print_Report02
            ElseIf oForm01.Mode = fm_OK_MODE Then
                Call PS_QM008_Print_Report03
            End If
        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_QM008" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_QM008_FormItemEnabled
                    Call PS_QM008_AddMatrixRow(0, True) '//UDO방식일때
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_QM008_FormItemEnabled
                End If
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
        If pval.CharPressed = 9 Then
            If pval.ItemUID = "CntcCode" Then
                If oForm01.Items("CntcCode").Specific.VALUE = "" Then
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                End If
            ElseIf pval.ItemUID = "ItemCode" Then
                If oForm01.Items("ItemCode").Specific.VALUE = "" Then
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                End If
            ElseIf pval.ItemUID = "CardCode" Then
                If oForm01.Items("CardCode").Specific.VALUE = "" Then
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                End If
            ElseIf pval.ItemUID = "InspNo" Then
                If oForm01.Items("InspNo").Specific.VALUE = "" Then
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                End If
            ElseIf pval.ItemUID = "CardSeq" Then
                If oForm01.Items("CardSeq").Specific.VALUE = "" Then
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                End If
            End If
        End If
        
    ElseIf pval.BeforeAction = False Then
    
    
    End If
    
    Exit Sub
Raise_EVENT_KEY_DOWN_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_KEY_DOWN_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_COMBO_SELECT(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_COMBO_SELECT_Error
    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
    
        If pval.ItemChanged = True Then
            
        End If
    End If
    Call oForm01.Freeze(False)
    Exit Sub
Raise_EVENT_COMBO_SELECT_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_COMBO_SELECT_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CLICK(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CLICK_Error
    If pval.BeforeAction = True Then
'        If pval.ItemUID = "Mat01" Then
'            If pval.Row > 0 Then
'                Call oMat01.SelectRow(pval.Row, True, False)
'            End If
'        End If
        If pval.ItemUID = "Mat01" Then
            If pval.Row > 0 Then
                oLastItemUID01 = pval.ItemUID
                oLastColUID01 = pval.ColUID
                oLastColRow01 = pval.Row
                
                Call oMat01.SelectRow(pval.Row, True, False)
            End If
        Else
            oLastItemUID01 = pval.ItemUID
            oLastColUID01 = ""
            oLastColRow01 = 0
        End If
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_CLICK_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
    Dim oTempClass As Variant
    If pval.BeforeAction = True Then
        If pval.ItemUID = "Mat01" Then
            
        End If
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_MATRIX_LINK_PRESSED_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LINK_PRESSED_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_VALIDATE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_VALIDATE_Error

    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim i As Long
    Dim YQty As Long
    Dim AQLYN As String
    
    Dim PS_PH030() As String
    
    Dim ErrNum          As Integer
    Dim SPEC_MIN        As Currency
    Dim SPEC_MAX        As Currency
    Dim PVALUE          As Currency
    
    Dim CardCode        As String
    Dim ItemCode        As String
    
    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then
            '검사치수 CHECK
            If (pval.ItemUID = "Mat01") Then
               If pval.ColUID = "Value" Then
                  
                  If oMat01.Columns("InspBal").Cells(pval.Row).Specific.VALUE <> "Y" Then
                     
                     PVALUE = Val(oMat01.Columns("Value").Cells(pval.Row).Specific.VALUE)
                     SPEC_MIN = Val(oMat01.Columns("InspMin").Cells(pval.Row).Specific.VALUE)
                     SPEC_MAX = Val(oMat01.Columns("InspMax").Cells(pval.Row).Specific.VALUE)
        
                     If SPEC_MIN = 0 And SPEC_MAX = 0 Then
                        If PVALUE <> 0 Then
                           ErrNum = 100
                           GoTo Raise_EVENT_VALIDATE_Error
                        End If
                     ElseIf PVALUE < SPEC_MIN Or PVALUE > SPEC_MAX Then
                            ErrNum = 100
                            GoTo Raise_EVENT_VALIDATE_Error
                     End If
                  End If
    
               End If
            
            ElseIf (pval.ItemUID = "InspNo") Then  '검사의뢰번호
                oForm01.Freeze True
                sQry = "Select Itemcode, Quantity = Sum(Quantity), CardName, CardSeq From [Z_PACKING_PD] Where QM006YN = 'N' AND InspNo = '" & Trim(oForm01.Items("InspNo").Specific.VALUE) & "' Group by Itemcode, CardName, CardSeq "
                oRecordSet01.DoQuery sQry
                oDS_PS_QM008H.setValue "U_Weight", 0, Trim(oRecordSet01.Fields(1).VALUE)
                oForm01.Items("ItemCode").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                
                If Trim(oRecordSet01.Fields(2).VALUE) <> "" Then
                    If Trim(oRecordSet01.Fields(3).VALUE) = "00" Then
                        oForm01.Items("Remark").Specific.VALUE = Trim(oRecordSet01.Fields(2).VALUE)
                    Else
                        oForm01.Items("Remark").Specific.VALUE = Trim(oRecordSet01.Fields(2).VALUE) + " #" + Trim(oRecordSet01.Fields(3).VALUE)
                    End If
                Else
                    oForm01.Items("Remark").Specific.VALUE = ""
                End If
                
                oForm01.Items("ItemCode").Enabled = False
                oForm01.Freeze False
                
            ElseIf (pval.ItemUID = "CntcCode") Then  '사번
                oForm01.Freeze True
                Call oDS_PS_QM008H.setValue("U_CntcName", 0, MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"))
                oForm01.Freeze False
                
            ElseIf (pval.ItemUID = "ItemCode") Then  '품목코드
                oForm01.Freeze True
                
                sQry = "Select ItemName, FrgnName, U_Size From OITM Where ItemCode = '" & Trim(oForm01.Items("ItemCode").Specific.VALUE) & "'"
                oRecordSet01.DoQuery sQry
                oDS_PS_QM008H.setValue "U_ItemName", 0, Trim(oRecordSet01.Fields(0).VALUE)
                oDS_PS_QM008H.setValue "U_FrgnName", 0, Trim(oRecordSet01.Fields(1).VALUE)
                oDS_PS_QM008H.setValue "U_Size", 0, Trim(oRecordSet01.Fields(2).VALUE)
                
                CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
                ItemCode = Trim(oForm01.Items("ItemCode").Specific.VALUE)
                sQry = "Select Count(*), Max(U_CardSeq) From [@PS_QM007H] Where U_CardCode = '" & CardCode & "' And U_ItemCode = '" & ItemCode & "'"
                oRecordSet01.DoQuery sQry
                
                If Trim(oRecordSet01.Fields(0).VALUE) = 1 Then
                   oDS_PS_QM008H.setValue "U_CardSeq", 0, Trim(oRecordSet01.Fields(1).VALUE)
                   oForm01.Items("CardSeq").Enabled = False
                   Call LoadData
                Else
                   oForm01.Items("CardSeq").Enabled = True
                   oDS_PS_QM008H.setValue "U_CardSeq", 0, ""
                   
                End If
                
'                If Trim(oForm01.Items("ItemCode").Specific.VALUE) <> "" And Trim(oForm01.Items("CardCode").Specific.VALUE) <> "" Then
'                   Call LoadData
'                End If
                
                oForm01.Freeze False
            ElseIf (pval.ItemUID = "CardCode") Then  '거래처코드
                oForm01.Freeze True
                
                sQry = "select cardname from ocrd where cardtype='C' and cardcode = '" & Trim(oDS_PS_QM008H.GetValue("U_CardCode", 0)) & "'"
                oRecordSet01.DoQuery sQry
                oDS_PS_QM008H.setValue "U_CardName", 0, Trim(oRecordSet01.Fields(0).VALUE)
                
                
                CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
                ItemCode = Trim(oForm01.Items("ItemCode").Specific.VALUE)
                sQry = "Select Count(*), Max(U_CardSeq) From [@PS_QM007H] Where U_CardCode = '" & CardCode & "' And U_ItemCode = '" & ItemCode & "'"
                oRecordSet01.DoQuery sQry
                
                If Trim(oRecordSet01.Fields(0).VALUE) = 1 Then
                   oDS_PS_QM008H.setValue "U_CardSeq", 0, Trim(oRecordSet01.Fields(1).VALUE)
                   oForm01.Items("CardSeq").Enabled = False
                   Call LoadData
                Else
                   oForm01.Items("CardSeq").Enabled = True
                   oDS_PS_QM008H.setValue "U_CardSeq", 0, ""
                End If
'                If Trim(oForm01.Items("ItemCode").Specific.VALUE) <> "" And Trim(oForm01.Items("CardCode").Specific.VALUE) <> "" Then
'                   Call LoadData
'                End If
                
                oForm01.Freeze False
            ElseIf (pval.ItemUID = "CardSeq") Then '거래처 순번
                If Trim(oForm01.Items("ItemCode").Specific.VALUE) <> "" And Trim(oForm01.Items("CardCode").Specific.VALUE) <> "" And Trim(oForm01.Items("CardSeq").Specific.VALUE) <> "" Then
                   Call LoadData
                End If
            End If
        End If
        
        
    ElseIf pval.BeforeAction = False Then
    
    End If
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
Raise_EVENT_VALIDATE_Error:
    If ErrNum = 100 Then
        Sbo_Application.MessageBox ("검사치수와 검사규격을 확인하여 주십시오.")
    Else
        Sbo_Application.SetStatusBarMessage "Raise_EVENT_VALIDATE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
End Sub

Private Sub Raise_EVENT_MATRIX_LOAD(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_MATRIX_LOAD_Error
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        Call PS_QM008_FormItemEnabled
        Call PS_QM008_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
        Call oMat01.AutoResizeColumns
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_QM008_FormResize
    End If
    Exit Sub
Raise_EVENT_RESIZE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_RESIZE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CHOOSE_FROM_LIST(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CHOOSE_FROM_LIST_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
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
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_QM008L.RemoveRecord(oDS_PS_QM008L.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PS_QM008_AddMatrixRow(0)
            Else
                If Trim(oDS_PS_QM008L.GetValue("U_InspItem", oMat01.RowCount - 1)) <> "" Then
                    Call PS_QM008_AddMatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_QM008_CreateItems() As Boolean
On Error GoTo PS_QM008_CreateItems_Error
   
    Set oDS_PS_QM008H = oForm01.DataSources.DBDataSources("@PS_QM008H")
    Set oDS_PS_QM008L = oForm01.DataSources.DBDataSources("@PS_QM008L")
    Set oMat01 = oForm01.Items("Mat01").Specific
    
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    Exit Function
PS_QM008_CreateItems_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM008_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_QM008_ComboBox_Setting()
On Error GoTo PS_QM008_ComboBox_Setting_Error
        
    '//콤보에 기본값설정
    Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    '사업장
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", False, False)
    
    '합부판정
    Set oCombo = oForm01.Items("PASSYN").Specific
    oCombo.ValidValues.Add "Y", "합격"
    oCombo.ValidValues.Add "N", "불합격"
    oCombo.ValidValues.Add "S", "특채"
    oCombo.Select 0, psk_Index
    
    Set oCombo = oForm01.Items("VIEWYN").Specific
    oCombo.ValidValues.Add "Y", "출력"
    oCombo.ValidValues.Add "N", "미출력"
    oCombo.Select 0, psk_Index
        
        
    Set oCombo = oForm01.Items("SOVIEWYN").Specific
    oCombo.ValidValues.Add "Y", "출력"
    oCombo.ValidValues.Add "N", "미출력"
    oCombo.Select 1, psk_Index
    
    
    
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing

    Exit Sub
PS_QM008_ComboBox_Setting_Error:
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM008_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM008_CF_ChooseFromList()
On Error GoTo PS_QM008_CF_ChooseFromList_Error
    Exit Sub
PS_QM008_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM008_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM008_FormItemEnabled()
On Error GoTo PS_QM008_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
'
        'oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("Mat01").Enabled = True
        Call PS_QM008_FormClear '//UDO방식
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", False) '//추가
        oForm01.Items("CntcCode").Specific.VALUE = MDC_PS_Common.User_MSTCOD()
        oForm01.Items("CardCode").Specific.VALUE = "12549"
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
        oForm01.Items("DocEntry").Specific.VALUE = ""
        oForm01.Items("DocEntry").Enabled = True
        oForm01.Items("Mat01").Enabled = False
        Call oForm01.EnableMenu("1281", False) '//찾기
        Call oForm01.EnableMenu("1282", True) '//추가
        
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
        
        oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("Mat01").Enabled = True
        
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_QM008_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM008_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM008_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_QM008_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_QM008L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_QM008L.Offset = oRow
    oDS_PS_QM008L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_QM008_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM008_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM008_FormClear()
On Error GoTo PS_QM008_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_QM008'", "")
    If DocEntry = "" Or DocEntry = "0" Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_QM008_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM008_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM008_Initial_Setting()
On Error GoTo PS_QM008_Initial_Setting_Error

    '사업장
    Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
    
    '일자
    oForm01.Items("DocDate").Specific.VALUE = Format(Now, "YYYYMMDD")
    oForm01.Items("PPYN").Specific.VALUE = "N"
    oForm01.Items("QMYN").Specific.VALUE = "N"
    '검사의뢰번호 포커서
    Call oForm01.Items("InspNo").Click
    
    
    
    Exit Sub
PS_QM008_Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM008_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM008_EnableMenus()
On Error GoTo PS_QM008_EnableMenus_Error
    '//메뉴활성화
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    Call MDC_GP_EnableMenus(oForm01, False, False, True, True, False, True, True, True, True, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_QM008_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM008_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM008_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_QM008_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_QM008_FormItemEnabled
        Call PS_QM008_AddMatrixRow(0, True) '//UDO방식일때
    Else
        oForm01.Mode = fm_FIND_MODE
        Call PS_QM008_FormItemEnabled
        oForm01.Items("DocEntry").Specific.VALUE = oFromDocEntry01
        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_QM008_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM008_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Function PS_QM008_DataValidCheck() As Boolean
On Error GoTo PS_QM008_DataValidCheck_Error
    PS_QM008_DataValidCheck = False
    Dim i As Long
    Dim ErrNum          As Integer
    Dim SPEC_MIN        As Currency
    Dim SPEC_MAX        As Currency
    Dim PVALUE          As Currency
    
    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_QM008_FormClear
    End If
    
    '일자 미입력시
    If oForm01.Items("DocDate").Specific.VALUE = "" Then
        Call Sbo_Application.SetStatusBarMessage("검사일자가 입력되지 않았습니다.", bmt_Short, True)
        PS_QM008_DataValidCheck = False
        Exit Function
    End If
    
    '고객순번 미입력시
    If oForm01.Items("CardSeq").Specific.VALUE = "" Then
        Call Sbo_Application.SetStatusBarMessage("고객번호가 입력되지 않았습니다.", bmt_Short, True)
        PS_QM008_DataValidCheck = False
        Exit Function
    End If
    
    '라인정보 미입력 시
    If oMat01.VisualRowCount = 1 Then
        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
        PS_QM008_DataValidCheck = False
        Exit Function
    End If
    
    If oForm01.Items("PASSYN").Specific.VALUE = "Y" Then
        '합격일때만 검사치수CHECK
        For i = 1 To oMat01.VisualRowCount - 1
        
            If oMat01.Columns("InspBal").Cells(i).Specific.VALUE <> "Y" Then
               PVALUE = Val(oMat01.Columns("Value").Cells(i).Specific.VALUE)
               SPEC_MIN = Val(oMat01.Columns("InspMin").Cells(i).Specific.VALUE)
               SPEC_MAX = Val(oMat01.Columns("InspMax").Cells(i).Specific.VALUE)
            
               If SPEC_MIN = 0 And SPEC_MAX = 0 Then
                  If PVALUE <> 0 Then
                     Call Sbo_Application.SetStatusBarMessage("검사치수를 확인 하십시요..", bmt_Short, True)
                     Call oMat01.Columns("Value").Cells(i).Click
                     PS_QM008_DataValidCheck = False
                     Exit Function
                  End If
               ElseIf PVALUE < SPEC_MIN Or PVALUE > SPEC_MAX Then
                     Call Sbo_Application.SetStatusBarMessage("검사치수를 확인 하십시요..", bmt_Short, True)
                     Call oMat01.Columns("Value").Cells(i).Click
                     PS_QM008_DataValidCheck = False
                     Exit Function
               End If
            End If
            
        Next
        
        '
        If oForm01.Items("IpdoChk").Specific.VALUE = "Y" Then
            '입도분포CHECK = 합이 100  ----------------------------------
            PVALUE = 0
            For i = 1 To oMat01.VisualRowCount - 1
            
                If oMat01.Columns("InspItem").Cells(i).Specific.VALUE = "입도분포" Then
                   PVALUE = PVALUE + Val(oMat01.Columns("Value").Cells(i).Specific.VALUE)
                End If
                
            Next
            
            If PVALUE <> 100 Then
               Call Sbo_Application.SetStatusBarMessage("입도분포의 합이(100)이 아닙니다. 확인 하십시요..", bmt_Short, True)
               PS_QM008_DataValidCheck = False
               Exit Function
            End If
            '-----------------------------------------------------------------------------------------------------------
        End If
    End If
    
    
    Call oMat01.FlushToDataSource
    Call oDS_PS_QM008L.RemoveRecord(oDS_PS_QM008L.Size - 1)
    Call oMat01.LoadFromDataSource

    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_QM008_FormClear
    End If
    
    PS_QM008_DataValidCheck = True
    Exit Function
PS_QM008_DataValidCheck_Error:
    PS_QM008_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_QM008_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_QM008_FormResize()
On Error GoTo PS_QM008_FormResize_Error
    
    Exit Sub
PS_QM008_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM008_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_QM008_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PS_QM008_Validate_Error
    PS_QM008_Validate = True
    Dim i, j As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If ValidateType = "수정" Then
        '//삭제된 행을 찾아서 삭제가능성 검사 , 만약 입력된행이 수정이 불가능하도록 변경이 필요하다면 삭제된행 찾는구문 제거
    ElseIf ValidateType = "행삭제" Then
        '//행삭제전 행삭제가능여부검사
    ElseIf ValidateType = "취소" Then
    End If
    Set RecordSet01 = Nothing
    Exit Function
PS_QM008_Validate_Exit:
    Set RecordSet01 = Nothing
    Exit Function
PS_QM008_Validate_Error:
    PS_QM008_Validate = False
    Sbo_Application.SetStatusBarMessage "PS_QM008_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Public Sub LoadData()
On Error GoTo LoadData_Error
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim i As Integer
    Dim Sintern As String
    Dim IpdoChk As String
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sQry = "Select b.U_InspItem, b.U_InspItNm, b.U_InspUnit, U_InspSpec, b.U_InspMeth, b.U_InspMin, b.U_InspMax, b.U_InspBal, a.U_Sintern, a.U_IpdoChk "
    sQry = sQry & " From [@PS_QM007H] a INNER JOIN [@PS_QM007L] b ON a.DocEntry = b.DocEntry AND a.Canceled = 'N' "
    sQry = sQry & "Where a.U_ItemCode = '" & Trim(oForm01.Items("ItemCode").Specific.VALUE) & "' "
    sQry = sQry & "  AND a.U_CardCode = '" & Trim(oForm01.Items("CardCode").Specific.VALUE) & "' "
    sQry = sQry & "  And a.U_CardSeq  = '" & Trim(oForm01.Items("CardSeq").Specific.VALUE) & "' "
    sQry = sQry & "  AND b.U_UseYN = 'Y' Order By b.U_Seqno "
    oRecordSet01.DoQuery sQry
    
    oDS_PS_QM008L.Clear
    oMat01.Clear
    oMat01.FlushToDataSource
    
    i = 0
    Do Until oRecordSet01.EOF
        oDS_PS_QM008L.InsertRecord i
        oDS_PS_QM008L.setValue "U_LineNum", i, i + 1
        oDS_PS_QM008L.setValue "U_InspItem", i, Trim(oRecordSet01.Fields(0).VALUE)
        oDS_PS_QM008L.setValue "U_InspItNm", i, Trim(oRecordSet01.Fields(1).VALUE)
        oDS_PS_QM008L.setValue "U_InspUnit", i, Trim(oRecordSet01.Fields(2).VALUE)
        oDS_PS_QM008L.setValue "U_InspSpec", i, Trim(oRecordSet01.Fields(3).VALUE)
        oDS_PS_QM008L.setValue "U_InspMeth", i, Trim(oRecordSet01.Fields(4).VALUE)
        oDS_PS_QM008L.setValue "U_InspMin", i, Trim(oRecordSet01.Fields(5).VALUE)
        oDS_PS_QM008L.setValue "U_InspMax", i, Trim(oRecordSet01.Fields(6).VALUE)
        oDS_PS_QM008L.setValue "U_InspBal", i, Trim(oRecordSet01.Fields(7).VALUE)
        Sintern = Trim(oRecordSet01.Fields(8).VALUE)
        IpdoChk = Trim(oRecordSet01.Fields(9).VALUE)
        i = i + 1
        oRecordSet01.MoveNext
    Loop
    
    oDS_PS_QM008H.setValue "U_Sintern", 0, Sintern
    oDS_PS_QM008H.setValue "U_IpdoChk", 0, IpdoChk
    
    oMat01.LoadFromDataSource
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadData_Error:
    oForm01.Freeze False
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "LoadData_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub PS_QM008_Print_Report01()
On Error GoTo PS_QM008_Print_Report01_Error
    
    Dim WinTitle                    As String
    Dim ReportName                  As String
    Dim sQry                        As String
    Dim Sub_sQry                    As String
    Dim i                           As Long
    Dim DocEntry                    As String
    Dim VIEWYN                      As String
    Dim SoviewYN                    As String
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    If oForm01.Items("ProgStat").Specific.Selected.VALUE <> "3" Then
'        Call MDC_Com.MDC_GF_Message("문서상태가 납품이 아닙니다.", "W")
'        Exit Sub
'    End If

    Call ConnectODBC
    
    DocEntry = oForm01.Items("DocEntry").Specific.VALUE
    VIEWYN = oForm01.Items("VIEWYN").Specific.VALUE
    SoviewYN = oForm01.Items("SOVIEWYN").Specific.VALUE
    
    WinTitle = "[PS_QM008_01] 검사성적서 출력(한글)"
    
    ReportName = "PS_QM008_01.rpt"
       
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    
    '// SubReport
    ReDim gRpt_SRptSqry(13)
    ReDim gRpt_SRptName(13)
    
    ReDim gRpt_SFormula(13, 1)
    ReDim gRpt_SFormula_Value(13, 1)
    
    gRpt_SFormula(1, 1) = ""
    gRpt_SFormula_Value(1, 1) = ""
    
    gRpt_SFormula(2, 1) = ""
    gRpt_SFormula_Value(2, 1) = ""
    
    gRpt_SFormula(3, 1) = ""
    gRpt_SFormula_Value(3, 1) = ""
    
    gRpt_SFormula(4, 1) = ""
    gRpt_SFormula_Value(4, 1) = ""
    
    gRpt_SFormula(5, 1) = ""
    gRpt_SFormula_Value(5, 1) = ""
    
    gRpt_SFormula(6, 1) = ""
    gRpt_SFormula_Value(6, 1) = ""
    
    gRpt_SFormula(7, 1) = ""
    gRpt_SFormula_Value(7, 1) = ""
    
    gRpt_SFormula(8, 1) = ""
    gRpt_SFormula_Value(8, 1) = ""
    
    gRpt_SFormula(9, 1) = ""
    gRpt_SFormula_Value(9, 1) = ""
    
    gRpt_SFormula(10, 1) = ""
    gRpt_SFormula_Value(10, 1) = ""
    
    gRpt_SFormula(11, 1) = ""
    gRpt_SFormula_Value(11, 1) = ""
    
    gRpt_SFormula(12, 1) = ""
    gRpt_SFormula_Value(12, 1) = ""
    
    gRpt_SFormula(13, 1) = ""
    gRpt_SFormula_Value(13, 1) = ""
    
    
    Sub_sQry = "EXEC [PS_QM008_02] '" & DocEntry & "','" & VIEWYN & "'"
    gRpt_SRptSqry(1) = Sub_sQry
    gRpt_SRptName(1) = "PS_QM008_SUB_01"
    
    Sub_sQry = "EXEC [PS_QM008_03] '" & DocEntry & "','" & VIEWYN & "'"
    gRpt_SRptSqry(2) = Sub_sQry
    gRpt_SRptName(2) = "PS_QM008_SUB_02"
    
    Sub_sQry = "EXEC [PS_QM008_04]  '" & DocEntry & "','" & VIEWYN & "'"
    gRpt_SRptSqry(3) = Sub_sQry
    gRpt_SRptName(3) = "PS_QM008_SUB_03"
    
    Sub_sQry = "EXEC [PS_QM008_05]  '" & DocEntry & "','" & VIEWYN & "'"
    gRpt_SRptSqry(6) = Sub_sQry
    gRpt_SRptName(6) = "PS_QM008_SUB_06"

    Sub_sQry = "EXEC [PS_QM008_03] '" & DocEntry & "','" & VIEWYN & "'"
    gRpt_SRptSqry(4) = Sub_sQry
    gRpt_SRptName(4) = "PS_QM008_SUB_04"
    
    Sub_sQry = "EXEC [PS_QM008_02] '" & DocEntry & "','" & VIEWYN & "'"
    gRpt_SRptSqry(5) = Sub_sQry
    gRpt_SRptName(5) = "PS_QM008_SUB_05"
    
    Sub_sQry = "EXEC [PS_QM008_02] '" & DocEntry & "','" & VIEWYN & "'"
    gRpt_SRptSqry(7) = Sub_sQry
    gRpt_SRptName(7) = "PS_QM008_SUB_07"
    
    Sub_sQry = "EXEC [PS_QM008_03] '" & DocEntry & "','" & VIEWYN & "'"
    gRpt_SRptSqry(8) = Sub_sQry
    gRpt_SRptName(8) = "PS_QM008_SUB_08"
    
    Sub_sQry = "EXEC [PS_QM008_04]  '" & DocEntry & "','" & VIEWYN & "'"
    gRpt_SRptSqry(9) = Sub_sQry
    gRpt_SRptName(9) = "PS_QM008_SUB_09"
    
    Sub_sQry = "EXEC [PS_QM008_04] '" & DocEntry & "','" & VIEWYN & "'"
    gRpt_SRptSqry(10) = Sub_sQry
    gRpt_SRptName(10) = "PS_QM008_SUB_10"
    
    Sub_sQry = "EXEC [PS_QM008_03] '" & DocEntry & "','" & VIEWYN & "'"
    gRpt_SRptSqry(11) = Sub_sQry
    gRpt_SRptName(11) = "PS_QM008_SUB_11"
    
    Sub_sQry = "EXEC [PS_QM008_02] '" & DocEntry & "','" & VIEWYN & "'"
    gRpt_SRptSqry(12) = Sub_sQry
    gRpt_SRptName(12) = "PS_QM008_SUB_12"
    
    Sub_sQry = "EXEC [PS_QM008_06]  '" & DocEntry & "','" & VIEWYN & "'"
    gRpt_SRptSqry(13) = Sub_sQry
    gRpt_SRptName(13) = "PS_QM008_SUB_13"
    
    
     sQry = "EXEC PS_QM008_01 '" & DocEntry & "','" & SoviewYN & "'"
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    Exit Sub
PS_QM008_Print_Report01_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM008_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM008_Print_Report02()
On Error GoTo PS_QM008_Print_Report01_Error
    
    Dim WinTitle                    As String
    Dim ReportName                  As String
    Dim sQry                        As String
    Dim Sub_sQry                    As String
    Dim i                           As Long
    Dim DocEntry                    As String
    Dim VIEWYN                      As String
    Dim SoviewYN                    As String
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    If oForm01.Items("ProgStat").Specific.Selected.VALUE <> "3" Then
'        Call MDC_Com.MDC_GF_Message("문서상태가 납품이 아닙니다.", "W")
'        Exit Sub
'    End If

    Call ConnectODBC
    
    DocEntry = oForm01.Items("DocEntry").Specific.VALUE
    VIEWYN = oForm01.Items("VIEWYN").Specific.VALUE
    SoviewYN = oForm01.Items("SOVIEWYN").Specific.VALUE
    
    WinTitle = "[PS_QM008_02] 검사성적서 출력(한문)"
    
    ReportName = "PS_QM008_02.rpt"
       
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    
    '// SubReport
    ReDim gRpt_SRptSqry(3)
    ReDim gRpt_SRptName(3)
    
    ReDim gRpt_SFormula(3, 1)
    ReDim gRpt_SFormula_Value(3, 1)
    
    gRpt_SFormula(1, 1) = ""
    gRpt_SFormula_Value(1, 1) = ""
    
    gRpt_SFormula(2, 1) = ""
    gRpt_SFormula_Value(2, 1) = ""
    
    gRpt_SFormula(3, 1) = ""
    gRpt_SFormula_Value(3, 1) = ""
    
    
    
    Sub_sQry = "EXEC [PS_QM008_04] '" & DocEntry & "','" & VIEWYN & "'"
    gRpt_SRptSqry(1) = Sub_sQry
    gRpt_SRptName(1) = "PS_QM008_SUB_01"
    
    Sub_sQry = "EXEC [PS_QM008_03] '" & DocEntry & "','" & VIEWYN & "'"
    gRpt_SRptSqry(2) = Sub_sQry
    gRpt_SRptName(2) = "PS_QM008_SUB_02"
    
    Sub_sQry = "EXEC [PS_QM008_02]  '" & DocEntry & "','" & VIEWYN & "'"
    gRpt_SRptSqry(3) = Sub_sQry
    gRpt_SRptName(3) = "PS_QM008_SUB_03"
    
    
    sQry = "EXEC PS_QM008_01 '" & DocEntry & "','" & SoviewYN & "'"
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    Exit Sub
PS_QM008_Print_Report01_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM008_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM008_Print_Report03()
On Error GoTo PS_QM008_Print_Report03_Error
    
    Dim WinTitle                    As String
    Dim ReportName                  As String
    Dim sQry                        As String
    Dim Sub_sQry                    As String
    Dim i                           As Long
    Dim DocEntry                    As String
    Dim VIEWYN                      As String
    Dim SoviewYN                    As String
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    If oForm01.Items("ProgStat").Specific.Selected.VALUE <> "3" Then
'        Call MDC_Com.MDC_GF_Message("문서상태가 납품이 아닙니다.", "W")
'        Exit Sub
'    End If

    Call ConnectODBC
    
    DocEntry = oForm01.Items("DocEntry").Specific.VALUE
    VIEWYN = oForm01.Items("VIEWYN").Specific.VALUE
    SoviewYN = oForm01.Items("SOVIEWYN").Specific.VALUE
    
    WinTitle = "[PS_QM008_03] 검사성적서 출력(영문)"
    
    ReportName = "PS_QM008_03.rpt"
       
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    
    '// SubReport
    ReDim gRpt_SRptSqry(9)
    ReDim gRpt_SRptName(9)
    
    ReDim gRpt_SFormula(9, 1)
    ReDim gRpt_SFormula_Value(9, 1)
    
    gRpt_SFormula(1, 1) = ""
    gRpt_SFormula_Value(1, 1) = ""
    
    gRpt_SFormula(2, 1) = ""
    gRpt_SFormula_Value(2, 1) = ""
    
    gRpt_SFormula(3, 1) = ""
    gRpt_SFormula_Value(3, 1) = ""
    
    gRpt_SFormula(4, 1) = ""
    gRpt_SFormula_Value(4, 1) = ""
    
    gRpt_SFormula(5, 1) = ""
    gRpt_SFormula_Value(5, 1) = ""
    
    gRpt_SFormula(6, 1) = ""
    gRpt_SFormula_Value(6, 1) = ""
    
    gRpt_SFormula(7, 1) = ""
    gRpt_SFormula_Value(7, 1) = ""
    
    gRpt_SFormula(8, 1) = ""
    gRpt_SFormula_Value(8, 1) = ""
    
    gRpt_SFormula(9, 1) = ""
    gRpt_SFormula_Value(9, 1) = ""
    
    
    Sub_sQry = "EXEC [PS_QM008_02] '" & DocEntry & "','" & VIEWYN & "'"
    gRpt_SRptSqry(1) = Sub_sQry
    gRpt_SRptName(1) = "PS_QM008_SUB_01"
    
    Sub_sQry = "EXEC [PS_QM008_03] '" & DocEntry & "','" & VIEWYN & "'"
    gRpt_SRptSqry(2) = Sub_sQry
    gRpt_SRptName(2) = "PS_QM008_SUB_02"
    
    Sub_sQry = "EXEC [PS_QM008_04]  '" & DocEntry & "','" & VIEWYN & "'"
    gRpt_SRptSqry(3) = Sub_sQry
    gRpt_SRptName(3) = "PS_QM008_SUB_03"
    
    Sub_sQry = "EXEC [PS_QM008_05]  '" & DocEntry & "','" & VIEWYN & "'"
    gRpt_SRptSqry(6) = Sub_sQry
    gRpt_SRptName(6) = "PS_QM008_SUB_06"

    Sub_sQry = "EXEC [PS_QM008_03] '" & DocEntry & "','" & VIEWYN & "'"
    gRpt_SRptSqry(4) = Sub_sQry
    gRpt_SRptName(4) = "PS_QM008_SUB_04"
    
    Sub_sQry = "EXEC [PS_QM008_02] '" & DocEntry & "','" & VIEWYN & "'"
    gRpt_SRptSqry(5) = Sub_sQry
    gRpt_SRptName(5) = "PS_QM008_SUB_05"
    
    Sub_sQry = "EXEC [PS_QM008_02] '" & DocEntry & "','" & VIEWYN & "'"
    gRpt_SRptSqry(7) = Sub_sQry
    gRpt_SRptName(7) = "PS_QM008_SUB_07"
    
    Sub_sQry = "EXEC [PS_QM008_03] '" & DocEntry & "','" & VIEWYN & "'"
    gRpt_SRptSqry(8) = Sub_sQry
    gRpt_SRptName(8) = "PS_QM008_SUB_08"
    
    Sub_sQry = "EXEC [PS_QM008_04]  '" & DocEntry & "','" & VIEWYN & "'"
    gRpt_SRptSqry(9) = Sub_sQry
    gRpt_SRptName(9) = "PS_QM008_SUB_09"
    
    
     sQry = "EXEC PS_QM008_01 '" & DocEntry & "','" & SoviewYN & "'"
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    Exit Sub
PS_QM008_Print_Report03_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM008_Print_Report03_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
