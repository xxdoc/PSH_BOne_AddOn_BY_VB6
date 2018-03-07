VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_GA054"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_GA054.cls
'//  Module : �ѹ�����>����ȯ�����>�ǰ���������>�ǰ������̷µ��
'//  Desc :
'//  FormType : PS_GA054
'//  Create Date : 2012.09.03
'//  Modified Date :
'//  Creator : Song Myoung gyu
'//  Company : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_GA054H As SAPbouiCOM.DBDataSource '������
Private oDS_PS_GA054L As SAPbouiCOM.DBDataSource '��϶���

Private oLastItemUID01 As String 'Ŭ�������� ������ ������ ������ Uid��
Private oLastColUID01 As String '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Col�� Uid��
Private oLastColRow01 As Long '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Row��

'//����ڱ���ü
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
' .srf ���Ϸκ��� ���� �ε��Ѵ�.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01 As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_GA054.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    
    '��Ʈ������ Ÿ��Ʋ���̿� �����̸� ����
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_GA054_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//���߰�
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '�� �Ҵ�
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    oForm01.DataBrowser.BrowseBy = "DocEntry" '//UDO����϶�
        
    oForm01.Freeze True
    Call PS_GA054_CreateItems
    Call PS_GA054_ComboBox_Setting
    Call PS_GA054_Initial_Setting
    Call PS_GA054_CF_ChooseFromList
    Call PS_GA054_EnableMenus
    Call PS_GA054_SetDocument(oFromDocEntry01)
    Call PS_GA054_FormResize
    
    oForm01.EnableMenu ("1283"), False         '// ����
    oForm01.EnableMenu ("1287"), True         '// ����
    oForm01.EnableMenu ("1286"), False        '// �ݱ�
    oForm01.EnableMenu ("1284"), True        '// ���
    oForm01.EnableMenu ("1293"), True         '// �����
    
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

    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '���
            Case "1286": '�ݱ�
            Case "1293": '�����
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": 'ã��
            Case "1282": '�߰�
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '���
            Case "1286": '�ݱ�
            Case "1293": '�����
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": 'ã��
                Call PS_GA054_FormItemEnabled '//UDO���
            Case "1282": '�߰�
                Call PS_GA054_FormItemEnabled '//UDO���
                Call PS_GA054_AddMatrixRow(0, True) '//UDO���
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
            
                Call oForm01.Freeze(True)
                Call PS_GA054_FormItemEnabled
'                Call oMat01.AutoResizeColumns
                Call PS_GA054_SetTitle(oForm01.Items("Cls").Specific.VALUE)
                Call oForm01.Freeze(False)
            
            Case "1287"
            
                oForm01.Freeze True
                Call oDS_PS_GA054H.setValue("DocEntry", 0, "")
                
                For i = 0 To oMat01.VisualRowCount - 1
                    Call oMat01.FlushToDataSource
                    Call oDS_PS_GA054H.setValue("DocEntry", i, "")
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
'            MenuCreationParams01.String = "�޴���"
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
    If pval.BeforeAction = True Then
        If pval.ItemUID = "PS_GA054" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If PS_GA054_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                oDocEntry01 = oForm01.Items("DocEntry").Specific.VALUE
                oFormMode01 = oForm01.Mode
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                If PS_GA054_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                oDocEntry01 = oForm01.Items("DocEntry").Specific.VALUE
                oFormMode01 = oForm01.Mode
                
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_GA054" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_GA054_FormItemEnabled
                    Call PS_GA054_AddMatrixRow(0, True) '//UDO����϶�
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_GA054_FormItemEnabled
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
    
        If pval.ItemUID = "Mat01" Then
        
            If pval.ColUID = "CntcCode" Then
            
                Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat01", "CntcCode") '��� ���˼�ġ����
            
            ElseIf pval.ColUID = "CodeLv1" Then
            
                If pval.CharPressed = 9 Then '���� ������ ��츸
            
                    Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat01", "CodeLv1") '��ȯ�ߺз� ���˼�ġ����
            
                End If
                
            ElseIf pval.ColUID = "CodeLv2" Then
            
                If pval.CharPressed = 9 Then '���� ������ ��츸
                
                    Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat01", "CodeLv2") '��ȯ�Һз� ���˼�ġ����
                
                End If
                
            End If
        
'        ElseIf pval.ItemUID = "RegCd" Then
'
'            Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "RegCd", "") '����� ���˼�ġ ����
'
'        ElseIf pval.ItemUID = "ReqCd" Then
'
'            Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ReqCd", "") '��û�� ���˼�ġ ����
        
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
            
            If pval.ItemUID = "Cls" Then
            
                Call PS_GA054_SetTitle(oForm01.Items("Cls").Specific.VALUE)
            
            End If
            
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

    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset

    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then

            If (pval.ItemUID = "Mat01") Then
                If pval.ColUID = "CntcCode" Then
                    oMat01.FlushToDataSource
                    
                    Call oDS_PS_GA054L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                    Call oDS_PS_GA054L.setValue("U_CntcName", pval.Row - 1, MDC_GetData.Get_ReData("U_FULLNAME", "Code", "[@PH_PY001A]", "'" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'")) '����
                    Call oDS_PS_GA054L.setValue("U_TeamCode", pval.Row - 1, MDC_GetData.Get_ReData("U_TeamCode", "Code", "[@PH_PY001A]", "'" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'")) '�Ҽ���
                    Call oDS_PS_GA054L.setValue("U_RspCode", pval.Row - 1, MDC_GetData.Get_ReData("U_RspCode", "Code", "[@PH_PY001A]", "'" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'")) '�ҼӴ��

                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_GA054L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                        PS_GA054_AddMatrixRow (pval.Row)
                    End If
                    oMat01.LoadFromDataSource
                    
                ElseIf pval.ColUID = "CodeLv1" Then
                
                    oMat01.FlushToDataSource
                    Call oDS_PS_GA054L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                    Call oDS_PS_GA054L.setValue("U_NameLv1", pval.Row - 1, MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_GA050L]", "'" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'", " AND Code = '4'")) '�ߺз���
                    oMat01.LoadFromDataSource
                    
                ElseIf pval.ColUID = "CodeLv2" Then
                
                    oMat01.FlushToDataSource
                    Call oDS_PS_GA054L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                    Call oDS_PS_GA054L.setValue("U_NameLv2", pval.Row - 1, MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_GA050L]", "'" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'", " AND Code = '5'")) '�Һз���
                    oMat01.LoadFromDataSource
                
                End If
                
'                Call oMat01.Columns(pval.ColUID).Cells(pval.Row).Click(ct_Regular)
                oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
            Else
                
            End If
        End If
        
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
        Call PS_GA054_FormItemEnabled
        Call PS_GA054_AddMatrixRow(oMat01.VisualRowCount) '//UDO���
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
        Call PS_GA054_FormResize
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
            '//������� ��������ɿ��ΰ˻�
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_GA054L.RemoveRecord(oDS_PS_GA054L.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PS_GA054_AddMatrixRow(0)
            Else
                If Trim(oDS_PS_GA054L.GetValue("U_CntcCode", oMat01.RowCount - 1)) <> "" Then
                    Call PS_GA054_AddMatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_GA054_CreateItems() As Boolean
On Error GoTo PS_GA054_CreateItems_Error
   
    Set oDS_PS_GA054H = oForm01.DataSources.DBDataSources("@PS_GA054H")
    Set oDS_PS_GA054L = oForm01.DataSources.DBDataSources("@PS_GA054L")
    Set oMat01 = oForm01.Items("Mat01").Specific
    
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    Exit Function
PS_GA054_CreateItems_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA054_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_GA054_ComboBox_Setting()
On Error GoTo PS_GA054_ComboBox_Setting_Error
        
    '//�޺��� �⺻������
    Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_GA054", "Mat01", "UseYN", "Y", "Y")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_GA054", "Mat01", "UseYN", "N", "N")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("UseYN"), "PS_GA054", "Mat01", "UseYN")

    '�����
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", False, False)
    
    '����
    sQry = "            SELECT  U_Code,"
    sQry = sQry & "             U_CodeNm"
    sQry = sQry & "  FROM   [@PS_GA050L]"
    sQry = sQry & "  WHERE  Code = '8'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    
    Call MDC_SetMod.Set_ComboList(oForm01.Items("Cls").Specific, sQry, "01", False, False)
    
    '��Ʈ����-�Ҽ���_S
    sQry = "           SELECT      b.U_Code, "
    sQry = sQry & "                b.U_CodeNm"
    sQry = sQry & " FROM       [@PS_HR200H] a"
    sQry = sQry & "                INNER JOIN"
    sQry = sQry & "                [@PS_HR200L] b"
    sQry = sQry & "                    On a.Code = b.Code"
    sQry = sQry & "                    And a.Name = '�μ�'"

    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("TeamCode"), sQry)
    '��Ʈ����-�Ҽ���_E

    '��Ʈ����-�ҼӴ���޺�_S
    sQry = "           SELECT      b.U_Code, "
    sQry = sQry & "                b.U_CodeNm"
    sQry = sQry & " FROM       [@PS_HR200H] a"
    sQry = sQry & "                INNER JOIN"
    sQry = sQry & "                [@PS_HR200L] b"
    sQry = sQry & "                    On a.Code = b.Code"
    sQry = sQry & "                    And a.Name = '���'"

    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("RspCode"), sQry)
    '��Ʈ����-�ҼӴ���޺�_E
    
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing

    Exit Sub
PS_GA054_ComboBox_Setting_Error:
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA054_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_GA054_CF_ChooseFromList()
On Error GoTo PS_GA054_CF_ChooseFromList_Error
    Exit Sub
PS_GA054_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA054_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_GA054_FormItemEnabled()
On Error GoTo PS_GA054_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//����忡���� �����ۼ���
'
        'oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("Cls").Enabled = True
        oForm01.Items("Mat01").Enabled = True
        Call PS_GA054_FormClear '//UDO���
        Call oForm01.EnableMenu("1281", True) '//ã��
        Call oForm01.EnableMenu("1282", False) '//�߰�

    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//����忡���� �����ۼ���
        oForm01.Items("DocEntry").Specific.VALUE = ""
        oForm01.Items("DocEntry").Enabled = True
        oForm01.Items("Cls").Enabled = True
        oForm01.Items("Mat01").Enabled = False
        Call oForm01.EnableMenu("1281", False) '//ã��
        Call oForm01.EnableMenu("1282", True) '//�߰�
        
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//����忡���� �����ۼ���
        
        oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("Cls").Enabled = False
        oForm01.Items("Mat01").Enabled = True
        
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_GA054_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA054_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_GA054_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_GA054_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//���߰�����
        oDS_PS_GA054L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_GA054L.Offset = oRow
    oDS_PS_GA054L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_GA054_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA054_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_GA054_FormClear()
On Error GoTo PS_GA054_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_GA054'", "")
    If DocEntry = "" Or DocEntry = "0" Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_GA054_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA054_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_GA054_Initial_Setting()
On Error GoTo PS_GA054_Initial_Setting_Error

    ' �����
    Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)

    '�������� ����
    oForm01.Items("StdDt").Specific.VALUE = Format(Now, "YYYYMMDD")

    Exit Sub
PS_GA054_Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA054_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA054_EnableMenus()
On Error GoTo PS_GA054_EnableMenus_Error
    '//�޴�Ȱ��ȭ
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//�޴�����
    Call MDC_GP_EnableMenus(oForm01, False, False, True, True, False, True, True, True, True, False, False, False, False, False, False) '//�޴�����
    Exit Sub
PS_GA054_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA054_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA054_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_GA054_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_GA054_FormItemEnabled
        Call PS_GA054_AddMatrixRow(0, True) '//UDO����϶�
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_GA054_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.VALUE = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_GA054_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA054_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Function PS_GA054_DataValidCheck() As Boolean
On Error GoTo PS_GA054_DataValidCheck_Error
    PS_GA054_DataValidCheck = False
    Dim i As Long
    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_GA054_FormClear
    End If
    
    '�������� ���Է� ��
    If oForm01.Items("StdDt").Specific.VALUE = "" Then
        Call Sbo_Application.SetStatusBarMessage("�������ڰ� �Էµ��� �ʾҽ��ϴ�.", bmt_Short, True)
        PS_GA054_DataValidCheck = False
        Exit Function
    End If
    
    '�����(��ġ��) ���Է� ��
    Dim ErrorString As String
     
    If oForm01.Items("ConName").Specific.VALUE = "" Then
    
        If oForm01.Items("Cls").Specific.Selected.VALUE = "01" Then
            ErrorString = "����ڰ� �Էµ��� �ʾҽ��ϴ�."
        Else
            ErrorString = "��ġ�ڰ� �Էµ��� �ʾҽ��ϴ�."
        End If
    
        Call Sbo_Application.SetStatusBarMessage(ErrorString, bmt_Short, True)
        PS_GA054_DataValidCheck = False
        Exit Function
    End If
    
    '�������� ���Է� ��
    If oMat01.VisualRowCount = 1 Then
        Sbo_Application.SetStatusBarMessage "������ �������� �ʽ��ϴ�.", bmt_Short, True
        PS_GA054_DataValidCheck = False
        Exit Function
    End If
    
    For i = 1 To oMat01.VisualRowCount - 1
        
        '���
        If (oMat01.Columns("CntcCode").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "����� �ʼ��Դϴ�.", bmt_Short, True
            oMat01.Columns("CntcCode").Cells(i).Click ct_Regular
            PS_GA054_DataValidCheck = False
            Exit Function
        End If

        '����
        If (oMat01.Columns("Age").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "������ �ʼ��Դϴ�.", bmt_Short, True
            oMat01.Columns("Age").Cells(i).Click ct_Regular
            PS_GA054_DataValidCheck = False
            Exit Function
        End If
        
    Next
    
    Call oMat01.FlushToDataSource
    Call oDS_PS_GA054L.RemoveRecord(oDS_PS_GA054L.Size - 1)
    Call oMat01.LoadFromDataSource

    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_GA054_FormClear
    End If
    
    PS_GA054_DataValidCheck = True
    Exit Function
PS_GA054_DataValidCheck_Error:
    PS_GA054_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_GA054_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Private Sub PS_GA054_MTX01()
On Error GoTo PS_GA054_MTX01_Error
    '//��Ʈ������ ������ �ε�
    Call oForm01.Freeze(True)
    Dim i As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    Param01 = oForm01.Items("Param01").Specific.VALUE
    Param02 = oForm01.Items("Param01").Specific.VALUE
    Param03 = oForm01.Items("Param01").Specific.VALUE
    Param04 = oForm01.Items("Param01").Specific.VALUE
    
    Query01 = "SELECT 10"
    Call RecordSet01.DoQuery(Query01)
    
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("����� �������� �ʽ��ϴ�.", "E")
        GoTo PS_GA054_MTX01_Exit
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("��ȸ����!", RecordSet01.RecordCount, False)
    
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDS_PS_GA054L.InsertRecord (i)
        End If
        oDS_PS_GA054L.Offset = i
        Call oDS_PS_GA054L.setValue("U_COL01", i, RecordSet01.Fields(0).VALUE)
        Call oDS_PS_GA054L.setValue("U_COL02", i, RecordSet01.Fields(1).VALUE)
        RecordSet01.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & RecordSet01.RecordCount & "�� ��ȸ��...!"
    Next
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    oForm01.Update
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_GA054_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Exit Sub
PS_GA054_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA054_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PS_GA054_FormResize()
On Error GoTo PS_GA054_FormResize_Error
    
    Exit Sub
PS_GA054_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA054_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_GA054_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PS_GA054_Validate_Error
    PS_GA054_Validate = True
    Dim i, j As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If ValidateType = "����" Then
        '//������ ���� ã�Ƽ� �������ɼ� �˻� , ���� �Էµ����� ������ �Ұ����ϵ��� ������ �ʿ��ϴٸ� �������� ã�±��� ����
    ElseIf ValidateType = "�����" Then
        '//������� ��������ɿ��ΰ˻�
    ElseIf ValidateType = "���" Then
    End If
    Set RecordSet01 = Nothing
    Exit Function
PS_GA054_Validate_Exit:
    Set RecordSet01 = Nothing
    Exit Function
PS_GA054_Validate_Error:
    PS_GA054_Validate = False
    Sbo_Application.SetStatusBarMessage "PS_GA054_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_GA054_SetTitle(ByVal Cls As String)
On Error GoTo PS_GA054_SetTitle_Error

    If Cls = "01" Then '�ǰ���� ���� ��
    
        'Matrix �ʵ� Visible ����
        oMat01.Columns("CodeLv1").Visible = False '��ȯ�ߺз�
        oMat01.Columns("NameLv1").Visible = False '��ȯ�ߺз���
        oMat01.Columns("CodeLv2").Visible = False '��ȯ�Һз�
        oMat01.Columns("NameLv2").Visible = False '��ȯ�Һз���
        oMat01.Columns("Cts").Visible = True '��㳻��
        oMat01.Columns("AfterCts").Visible = True '��� �� ��ġ
        
        oForm01.Items("Static04").Specific.Caption = "�����"
        Call oDS_PS_GA054H.setValue("U_ConName", 0, "")

        Call oMat01.AutoResizeColumns
    
    Else '��ǰ���� ���� ��
    
        'Matrix �ʵ� Visible ����
        oMat01.Columns("CodeLv1").Visible = True '��ȯ�ߺз�
        oMat01.Columns("NameLv1").Visible = True '��ȯ�ߺз���
        oMat01.Columns("CodeLv2").Visible = True '��ȯ�Һз�
        oMat01.Columns("NameLv2").Visible = True '��ȯ�Һз���
        oMat01.Columns("Cts").Visible = False '����
        oMat01.Columns("AfterCts").Visible = False '��� �� ��ġ
        
        oForm01.Items("Static04").Specific.Caption = "��ġ��"
        Call oDS_PS_GA054H.setValue("U_ConName", 0, MDC_GetData.Get_ReData("U_FULLNAME", "U_MSTCOD", "[OHEM]", "'" & MDC_PS_Common.User_MSTCOD & "'"))
                
        Call oMat01.AutoResizeColumns
    
    End If

    Exit Sub
PS_GA054_SetTitle_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA054_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub





