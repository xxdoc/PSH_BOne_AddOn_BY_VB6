VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_SD001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_SD001.cls
'//  Module         : SD
'//  Desc           : ������,���� ǰ���ڵ� ���
'//  FormType       : PS_SD001
'//  Create Date    : 2010.10.05
'//  Modified Date  :
'//  Creator        : Dong sub Kim
'//  Copyright  (c) Poongsan Holdings
'//********************************************************************************
Option Explicit         '��� ���ؿ��� �ش��ϴ� ����� ��� ������ ���������� �����ϵ��� �� �� ����մϴ�.

Public oFormUniqueID01      As String
Public oForm01              As SAPbouiCOM.Form
Public oMat01               As SAPbouiCOM.Matrix
Private oDS_PS_SD001H       As SAPbouiCOM.DBDataSource    '������ 'Represents a SAP Business One database table to be attached to a form.
Private oDS_PS_SD001L       As SAPbouiCOM.DBDataSource    '��϶���
Private oDS_PS_SD001M       As SAPbouiCOM.DBDataSource

Private oLast_Item_UID   As String                        'Ŭ�������� ������ ������ ������ Uid��
Private oLast_Col_UID    As String                        '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Col�� Uid��
Private oLast_Col_Row    As Long                          '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Row��
Private oSeq&
Private TmpCode As String
Private oRadioBtn As String
Private oCardCode As String
Private oCpNaming As String
Private oWhsCode As String
Private oItmBsort As String
Private oItmMsort As String


'Private oFormMode As String

'*******************************************************************
' .srf ���Ϸκ��� ���� �ε��Ѵ�.
'*******************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_SD001.srf")
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
    
    oFormUniqueID01 = "PS_SD001_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//���߰�
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '�� �Ҵ�
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    ' ������ �Ӽ� � ��ü ������ �����մϴ�

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
 
    oForm01.Freeze True
    Call CreateItems
    Call Initial_Setting
    
    oForm01.EnableMenu ("1281"), False     '// ã��
    oForm01.EnableMenu ("1282"), True      '// �߰�
    oForm01.EnableMenu ("1283"), True      '// ����
    oForm01.EnableMenu ("1287"), True      '// ����
    oForm01.EnableMenu ("1284"), False     '// ���
    oForm01.EnableMenu ("1288"), False
    oForm01.EnableMenu ("1289"), False
    oForm01.EnableMenu ("1290"), False
    oForm01.EnableMenu ("1291"), False
    
    Call ComboBox_Setting
'    Call CF_ChooseFromList
'    Call FormItemEnabled
    
        
    oForm01.Update      ' Reloads the form from the application.
                        ' Use this method to update an object from the database.
    oForm01.Freeze False
    oForm01.Visible = True  ' Indicates whether the form is visible.
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
    Dim sQry As String
    Dim ErrNum As Integer
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim loopCount As Integer
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "Btn02" Then
                    
                    If oForm01.Mode = fm_ADD_MODE Or oForm01.Mode = fm_UPDATE_MODE Then
                        If HeaderSpaceLineDel02 = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        If (oForm01.Mode = fm_ADD_MODE) Then
                            If Create_Itemcode(1) = True Then
                                 MDC_Com.MDC_GF_Message "�ű� ���� ��� �۾��� �����Ͽ����ϴ�.", "S"
                                 FormItem_Clear
                            Else
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                    End If
                    
                ElseIf pval.ItemUID = "Btn01" Then      ' ǰ���ڵ� ���� ��ư
                     If oForm01.Mode = fm_ADD_MODE Then
                        If HeaderSpaceLineDel01 = False Then
                            BubbleEvent = False
                            Exit Sub
                        Else
                            Call Create_Form_Code
                        End If
                     End If
                 End If
                
            Case et_KEY_DOWN: '//2
            
                ' ����
                If oForm01.Items("CardCode").Specific.VALUE = "" Then
                    If pval.ItemUID = "CardCode" And pval.CharPressed = 9 Then
                        oForm01.Items("CardCode").Click ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
                                
                ' �⺻â��
                If oForm01.Items("WhsCode").Specific.VALUE = "" Then
                    If pval.ItemUID = "WhsCode" And pval.CharPressed = 9 Then
                        oForm01.Items("WhsCode").Click ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
                
                ' ǰ���з�
                If oForm01.Items("ItmBsort").Specific.VALUE = "" Then
                    If pval.ItemUID = "ItmBsort" And pval.CharPressed = 9 Then
                        oForm01.Items("ItmBsort").Click ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
                
                ' ǰ���ߺз�
                If oForm01.Items("ItmMsort").Specific.VALUE = "" Then
                    If pval.ItemUID = "ItmMsort" And pval.CharPressed = 9 Then
                        oForm01.Items("ItmMsort").Click ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
                
                ' �԰�
                If oForm01.Items("Unit1").Specific.VALUE = "" Then
                    If pval.ItemUID = "Unit1" And pval.CharPressed = 9 Then
                        oForm01.Items("Unit1").Click ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
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
                 oLast_Item_UID = pval.ItemUID
            Case et_LOST_FOCUS: '//4
            Case et_FORM_UNLOAD: '//17
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
            '//��Ʈ������ ������ �ε�
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
              If pval.ItemUID = "Btn02" Then
                If oForm01.Mode = fm_ADD_MODE And pval.Action_Success = True Then
                      oForm01.Mode = fm_OK_MODE
                      Call Sbo_Application.ActivateMenuItem("1282")
                End If
              End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
                If pval.ItemUID = "CpNaming" Then
                    If oForm01.Items("CpNaming").Specific.VALUE = "G" Then
                        oForm01.Items("ItmBsort").Specific.VALUE = "105"
                        oForm01.Items("ItmMsort").Specific.VALUE = "10504"
                    ElseIf oForm01.Items("CpNaming").Specific.VALUE = "M" Then
                        oForm01.Items("ItmBsort").Specific.VALUE = "105"
                        oForm01.Items("ItmMsort").Specific.VALUE = "10501"
                    ElseIf oForm01.Items("CpNaming").Specific.VALUE = "P" Then
                        oForm01.Items("ItmBsort").Specific.VALUE = "105"
                        oForm01.Items("ItmMsort").Specific.VALUE = "10502"
                    ElseIf oForm01.Items("CpNaming").Specific.VALUE = "T" Then
                        oForm01.Items("ItmBsort").Specific.VALUE = "105"
                        oForm01.Items("ItmMsort").Specific.VALUE = "10503"
                    ElseIf oForm01.Items("CpNaming").Specific.VALUE = "J" Then
                        oForm01.Items("ItmBsort").Specific.VALUE = "106"
                        oForm01.Items("ItmMsort").Specific.VALUE = "10601"
                    End If
                End If
                
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
                ' ����
                If pval.ItemUID = "CardCode" And pval.ItemChanged = True Then
                    sQry = "Select CardName From [OCRD] Where CardCode = '" & Trim(oForm01.Items("CardCode").Specific.VALUE) & "'"
                    oRecordSet01.DoQuery sQry
                    oForm01.Items("CardName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                    
                        If oForm01.Items("TeamCode").Specific.ValidValues.Count > 0 Then
                            For loopCount = oForm01.Items("TeamCode").Specific.ValidValues.Count - 1 To 0 Step -1
                                Call oForm01.Items("TeamCode").Specific.ValidValues.Remove(loopCount, psk_Index)
                            Next
                        End If
                        
                        Call oForm01.Items("TeamCode").Specific.ValidValues.Add("%", "����")
                        sQry = "            SELECT      U_Minor AS [Code],"
                        sQry = sQry & "                 U_CdName As [Name]"
                        sQry = sQry & "  FROM       [@PS_SY001L]"
                        sQry = sQry & "  WHERE      Code = 'I001'"
                        sQry = sQry & "                 AND U_UseYN = 'Y'"
                        sQry = sQry & "                 AND U_RelCd = '" & Trim(oForm01.Items("CardCode").Specific.VALUE) & "'"
                        sQry = sQry & "  ORDER BY  U_Seq"
                        Call MDC_SetMod.Set_ComboList(oForm01.Items("TeamCode").Specific, sQry, "", False, False)
                        Call oForm01.Items("TeamCode").Specific.Select(0, psk_Index)
                    
                End If
                
                ' �⺻â��
                If pval.ItemUID = "WhsCode" And pval.ItemChanged = True Then
                    sQry = "Select WhsName From [OWHS] Where WhsCode = '" & Trim(oForm01.Items("WhsCode").Specific.VALUE) & "'"
                    oRecordSet01.DoQuery sQry
                    oForm01.Items("WhsName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                End If
                
                ' ǰ���з�
                If pval.ItemUID = "ItmBsort" And pval.ItemChanged = True Then
                    sQry = "Select Name From [@PSH_ITMBSORT] Where Code = '" & Trim(oForm01.Items("ItmBsort").Specific.VALUE) & "'"
                    oRecordSet01.DoQuery sQry
                    oForm01.Items("ItmBname").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                End If
                
                ' ǰ���ߺз�
                If pval.ItemUID = "ItmMsort" And pval.ItemChanged = True Then
                    sQry = "Select U_CodeName From [@PSH_ITMMSORT] Where U_Code = '" & Trim(oForm01.Items("ItmMsort").Specific.VALUE) & "'"
                    oRecordSet01.DoQuery sQry
                    oForm01.Items("ItmMname").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                End If
                
            Case et_MATRIX_LOAD: '//11
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
            Case et_FORM_RESIZE '//20
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
                oLast_Item_UID = pval.ItemUID
            Case et_LOST_FOCUS: '//4
            Case et_FORM_UNLOAD: '//17
                RemoveForms oFormUniqueID01
                Set oForm01 = Nothing
                Set oMat01 = Nothing
        End Select
    End If
    Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Raise_ItemEvent_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "ǰ���ڵ���� ��ư�� ������ ���� ������ ���� �����ϼ���.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "ǰ���ڵ���� ��ư�� ������ ���� ǰ�񱸺��� ���� �����ϼ���.", "E"
    Else
        Sbo_Application.SetStatusBarMessage "Raise_ItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
End Sub


Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_MenuEvent_Error
    Dim i&
    If (pval.BeforeAction = True) Then      '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '���
            Case "1286": '�ݱ�
            Case "1281": 'ã��
            Case "1282": '�߰�
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
            Case "1293"     '�����
        End Select
        
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '���
            Case "1286": '�ݱ�
            Case "1281": 'ã��
            Case "1282": '�߰�
                FormItem_Clear
            Case "1287": '����
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
            Case "1293"     '�����
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

Public Sub Raise_RightClickEvent(ByRef FormUID As String, ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_RightClickEvent_Error
    If (eventInfo.BeforeAction = True) Then
        '//�۾�
    ElseIf (eventInfo.BeforeAction = False) Then
        '//�۾�
    End If
    Exit Sub
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function CreateItems() As Boolean
On Error GoTo CreateItems_Error
    Dim optBtn          As SAPbouiCOM.OptionBtn
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    
    Call oForm01.DataSources.UserDataSources.Add("MatDate", dt_DATE, 10)
    oForm01.Items("MatDate").Specific.DataBind.SetBound True, "", "MatDate"
    oForm01.DataSources.UserDataSources.Item("MatDate").VALUE = Format(Now, "YYYYMMDD")

    ' Radio Button ó��'''''''''''''''''''''''''''''''''''''''''''''''''''
    Call oForm01.DataSources.UserDataSources.Add("RadioBtn", dt_SHORT_TEXT, 1)
    
    Set optBtn = oForm01.Items("Ra_Sale").Specific ' ����
    optBtn.ValOn = "A"
    optBtn.ValOff = "0"
    optBtn.DataBind.SetBound True, , "RadioBtn"
    optBtn.Selected = True
    
    Set optBtn = oForm01.Items("Ra_Samp").Specific ' �ߺ�
    optBtn.ValOn = "B"
    optBtn.ValOff = "0"
    optBtn.DataBind.SetBound True, , "RadioBtn"
    optBtn.GroupWith ("Ra_Sale")
    
    Set optBtn = oForm01.Items("Ra_AS").Specific    ' AS
    optBtn.ValOn = "C"
    optBtn.ValOff = "0"
    optBtn.DataBind.SetBound True, , "RadioBtn"
    optBtn.GroupWith ("Ra_Sale")
    
    Set optBtn = oForm01.Items("Ra_Multi").Specific    ' ��Ƽ
    optBtn.ValOn = "D"
    optBtn.ValOff = "0"
    optBtn.DataBind.SetBound True, , "RadioBtn"
    optBtn.GroupWith ("Ra_Sale")
    
    Set optBtn = oForm01.Items("Ra_Sin").Specific    ' �ŵ�
    optBtn.ValOn = "E"
    optBtn.ValOff = "0"
    optBtn.DataBind.SetBound True, , "RadioBtn"
    optBtn.GroupWith ("Ra_Sale")
    
    Set optBtn = oForm01.Items("Ra_RND").Specific 'R&D(����)
    optBtn.ValOn = "R"
    optBtn.ValOff = "0"
    optBtn.DataBind.SetBound True, , "RadioBtn"
    optBtn.GroupWith ("Ra_Sale")
    
    Set optBtn = Nothing
    
    ' ǰ�񱸺�
    oForm01.Items("CpNaming").Specific.ValidValues.Add "D", "�Ӱ���"
    oForm01.Items("CpNaming").Specific.ValidValues.Add "G", "������"
    oForm01.Items("CpNaming").Specific.ValidValues.Add "J", "����"
    oForm01.Items("CpNaming").Specific.ValidValues.Add "M", "���"
    oForm01.Items("CpNaming").Specific.ValidValues.Add "P", "��ǰ"
    oForm01.Items("CpNaming").Specific.ValidValues.Add "T", "����"
    
    oForm01.Items("WhsCode").Specific.VALUE = "102"
    oForm01.Items("WhsName").Specific.VALUE = "����"
    
    
    Exit Function
CreateItems_Error:
    Set oRecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub Initial_Setting()
On Error GoTo Initial_Setting_Error
 ' �����
'    Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
 ' �μ���
'    oForm01.Items("RepName").Specific.Value = MDC_PS_Common.User_MSTCOD()
    
    Exit Sub
Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub ComboBox_Setting()
On Error GoTo ComboBox_Setting_Error
    
    Dim loopCount As Integer
    Dim sQry As String
    
    If oForm01.Items("TeamCode").Specific.ValidValues.Count > 0 Then
        For loopCount = oForm01.Items("TeamCode").Specific.ValidValues.Count - 1 To 0 Step -1
            Call oForm01.Items("TeamCode").Specific.ValidValues.Remove(loopCount, psk_Index)
        Next
    End If
    
    Call oForm01.Items("TeamCode").Specific.ValidValues.Add("%", "����")
    sQry = "            SELECT      U_Minor AS [Code],"
    sQry = sQry & "                 U_CdName As [Name]"
    sQry = sQry & "  FROM       [@PS_SY001L]"
    sQry = sQry & "  WHERE      Code = 'I001'"
    sQry = sQry & "                 AND U_UseYN = 'Y'"
    'sQry = sQry & "                 AND U_Char2 = '" & CLTCOD & "'"
    sQry = sQry & "  ORDER BY  U_Seq"
    Call MDC_SetMod.Set_ComboList(oForm01.Items("TeamCode").Specific, sQry, "", False, False)
    Call oForm01.Items("TeamCode").Specific.Select(0, psk_Index)
    
    Call oForm01.Items("YearPdYN").Specific.ValidValues.Add("%", "����")
    Call oForm01.Items("YearPdYN").Specific.ValidValues.Add("Y", "Y")
    Call oForm01.Items("YearPdYN").Specific.ValidValues.Add("N", "N")
    Call oForm01.Items("YearPdYN").Specific.Select(0, psk_Index)
    '//�޺��� �⺻������
    Exit Sub
ComboBox_Setting_Error:
    Sbo_Application.SetStatusBarMessage "ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub CF_ChooseFromList()
On Error GoTo CF_ChooseFromList_Error
    '//ChooseFromList ����
    Exit Sub
CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub FormItemEnabled()
On Error GoTo FormItemEnabled_Error
    If (oForm01.Mode = fm_ADD_MODE) Then
        
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        
    ElseIf (oForm01.Mode = fm_OK_MODE) Then

    End If
    Exit Sub
FormItemEnabled_Error:
    Sbo_Application.SetStatusBarMessage "FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub FormItem_Clear()
On Error GoTo FormItem_Clear_Error
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//����忡���� �����ۼ���
'        oForm01.Items("CardCode").Specific.Value = ""
'        oForm01.Items("CardName").Specific.Value = ""
'        oForm01.Items("WhsCode").Specific.Value = ""
'        oForm01.Items("WhsName").Specific.Value = ""
'        oForm01.Items("ItemName").Specific.Value = ""
'        oForm01.Items("ItmBsort").Specific.Value = ""
'        oForm01.Items("ItmBname").Specific.Value = ""
'        oForm01.Items("ItmMsort").Specific.Value = ""
        oForm01.Items("ItemName").Specific.VALUE = ""
        oForm01.Items("Spec1").Specific.VALUE = ""
        oForm01.Items("Unit1").Specific.VALUE = ""
        oForm01.Items("ItemCode").Specific.VALUE = ""
'        oForm01.Items("MatDate").Specific.VALUE = ""
        
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        
    ElseIf (oForm01.Mode = fm_OK_MODE) Then

    End If
    Exit Sub
FormItem_Clear_Error:
    Sbo_Application.SetStatusBarMessage "FormItem_Clear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub AddMatrixRow(ByVal oSeq As Integer, ByVal oRow As Long, Optional RowIserted As Boolean)
'On Error GoTo AddMatrixRow_Error
   Select Case oSeq
       Case 0:
            'oMat01.AddRow
            'oMat01.Columns("LineNum").Cells(oRow + 1).Specific.Value = oRow + 1
            oDS_PS_SD001L.setValue "U_LineNum", oRow, oRow + 1
            oMat01.LoadFromDataSource
       Case 1:
            oDS_PS_SD001L.InsertRecord oRow
            oDS_PS_SD001L.setValue "U_LineNum", oRow, oRow + 1
            oMat01.LoadFromDataSource
   End Select
'AddMatrixRow_Error:
'    Sbo_Application.SetStatusBarMessage "AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function DataValidCheck() As Boolean
On Error GoTo DataValidCheck_Error
    '//��ȿ���˻�
    Exit Function
DataValidCheck_Error:
    Sbo_Application.SetStatusBarMessage "DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Private Function HeaderSpaceLineDel02() As Boolean
On Error GoTo HeaderSpaceLineDel02_Error
    Dim ErrNum      As Integer
    Dim DocNum      As String
    
    ErrNum = 0
    
    '// Check
    Select Case True
'        Case oForm01.Items("BPLId").Specific.Value = ""
'             ErrNum = 1
'             GoTo HeaderSpaceLineDel02_Error
        Case oForm01.DataSources.UserDataSources("RadioBtn").VALUE = "A" And oForm01.Items("CardCode").Specific.VALUE = ""
             ErrNum = 2
             GoTo HeaderSpaceLineDel02_Error
        Case oForm01.Items("CpNaming").Specific.VALUE = ""
             ErrNum = 3
             GoTo HeaderSpaceLineDel02_Error
        Case oForm01.Items("WhsCode").Specific.VALUE = ""
             ErrNum = 4
             GoTo HeaderSpaceLineDel02_Error
        Case oForm01.Items("ItemName").Specific.VALUE = ""
             ErrNum = 5
             GoTo HeaderSpaceLineDel02_Error
        Case oForm01.Items("ItmBsort").Specific.VALUE = ""
             ErrNum = 6
             GoTo HeaderSpaceLineDel02_Error
        Case oForm01.Items("ItmMsort").Specific.VALUE = ""
             ErrNum = 7
             GoTo HeaderSpaceLineDel02_Error
        Case oForm01.Items("Spec1").Specific.VALUE = ""
             ErrNum = 8
             GoTo HeaderSpaceLineDel02_Error
        Case oForm01.Items("Unit1").Specific.VALUE = ""
             ErrNum = 9
             GoTo HeaderSpaceLineDel02_Error
        Case oForm01.Items("MatDate").Specific.VALUE = ""
             ErrNum = 10
             GoTo HeaderSpaceLineDel02_Error
        Case oForm01.Items("ItemCode").Specific.VALUE = ""
             ErrNum = 11
             GoTo HeaderSpaceLineDel02_Error
    End Select
    
    HeaderSpaceLineDel02 = True
Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel02_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "������� �ʼ��Է� �����Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "������ ��� �����ڵ�� �ʼ��Է� �����Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "ǰ�񱸺��� �ʼ��Է� �����Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "�⺻â���� �ʼ��Է� �����Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "ǰ���̸��� �ʼ��Է� �����Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "ǰ���з��� �ʼ��Է� �����Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 7 Then
        MDC_Com.MDC_GF_Message "ǰ���ߺз��� �ʼ��Է� �����Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 8 Then
        MDC_Com.MDC_GF_Message "�԰��� �ʼ��Է� �����Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 9 Then
        MDC_Com.MDC_GF_Message "������ �ʼ��Է� �����Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 10 Then
        MDC_Com.MDC_GF_Message "�۸����ڴ� �ʼ��Է� �����Դϴ�. Ȯ���ϼ���", "E"
    ElseIf ErrNum = 11 Then
        MDC_Com.MDC_GF_Message "ǰ���ڵ���� ��ư�� ���� ��������.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel02_Error:" & Err.Description, "E"
    End If
    HeaderSpaceLineDel02 = False
End Function

Private Function HeaderSpaceLineDel01() As Boolean
On Error GoTo HeaderSpaceLineDel01_Error
    Dim ErrNum      As Integer
    Dim DocNum      As String
    
    ErrNum = 0
    
    '// Check
    Select Case True
'        Case oForm01.Items("BPLId").Specific.Value = ""
'             ErrNum = 1
'             GoTo HeaderSpaceLineDel01_Error
        Case oForm01.DataSources.UserDataSources("RadioBtn").VALUE = "A" And oForm01.Items("CardCode").Specific.VALUE = ""
             ErrNum = 2
             GoTo HeaderSpaceLineDel01_Error
        Case oForm01.Items("CpNaming").Specific.VALUE = ""
             ErrNum = 3
             GoTo HeaderSpaceLineDel01_Error
        Case oForm01.Items("WhsCode").Specific.VALUE = ""
             ErrNum = 4
             GoTo HeaderSpaceLineDel01_Error
        Case oForm01.Items("ItemName").Specific.VALUE = ""
             ErrNum = 5
             GoTo HeaderSpaceLineDel01_Error
        Case oForm01.Items("ItmBsort").Specific.VALUE = ""
             ErrNum = 6
             GoTo HeaderSpaceLineDel01_Error
        Case oForm01.Items("ItmMsort").Specific.VALUE = ""
             ErrNum = 7
             GoTo HeaderSpaceLineDel01_Error
        Case oForm01.Items("Spec1").Specific.VALUE = ""
             ErrNum = 8
             GoTo HeaderSpaceLineDel01_Error
        Case oForm01.Items("Unit1").Specific.VALUE = ""
             ErrNum = 9
             GoTo HeaderSpaceLineDel01_Error
        Case oForm01.Items("MatDate").Specific.VALUE = ""
             ErrNum = 10
             GoTo HeaderSpaceLineDel01_Error
    End Select
    
    HeaderSpaceLineDel01 = True
Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel01_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "������� �ʼ��Է� �����Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "������ ��� �����ڵ�� �ʼ��Է� �����Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "ǰ�񱸺��� �ʼ��Է� �����Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "�⺻â���� �ʼ��Է� �����Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "ǰ���̸��� �ʼ��Է� �����Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "ǰ���з��� �ʼ��Է� �����Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 7 Then
        MDC_Com.MDC_GF_Message "ǰ���ߺз��� �ʼ��Է� �����Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 8 Then
        MDC_Com.MDC_GF_Message "�԰��� �ʼ��Է� �����Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 9 Then
        MDC_Com.MDC_GF_Message "������ �ʼ��Է� �����Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 10 Then
        MDC_Com.MDC_GF_Message "�۸����ڴ� �ʼ��Է� �����Դϴ�. Ȯ���ϼ���", "E"
    ElseIf ErrNum = 11 Then
        MDC_Com.MDC_GF_Message "ǰ���ڵ���� ��ư�� ���� ��������.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel02_Error:" & Err.Description, "E"
    End If
    HeaderSpaceLineDel01 = False
End Function

Private Sub FlushToItemValue(ByVal oUID As String, ByVal oCID As String, Optional oRow As Long, Optional oCol As String)
    Dim i$
    Dim oRecordSet01        As SAPbobsCOM.Recordset
    Dim sQry                As String
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
        Case "Mat01"
            Select Case oCID
                Case "CpCode"
                    If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 2) And _
                    Trim(oMat01.Columns("CpCode").Cells(oRow).Specific.VALUE) <> "" Then
                        oMat01.FlushToDataSource
                        AddMatrixRow 1, oMat01.RowCount, True
                        oMat01.Columns("CpCode").Cells(oRow).Click ct_Regular
                    End If
           End Select
    End Select
    Set oRecordSet01 = Nothing
End Sub


Private Function Create_Itemcode(ChkType As Integer) As Boolean
On Error GoTo Create_Itemcode_Error

    Dim oItem01 As SAPbobsCOM.Items
    Dim ErrCode       As Long
    Dim ErrMsg        As String
    Dim ItemName      As String
    Dim RetVal&

    Set oItem01 = Sbo_Company.GetBusinessObject(oItems)
    
    If (Sbo_Company.InTransaction = True) Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    
    Sbo_Company.StartTransaction
    ItemName = oForm01.Items("ItemName").Specific.VALUE & " " & oForm01.Items("Spec1").Specific.VALUE  '�������ڵ�
    oItem01.ItemCode = oForm01.Items("ItemCode").Specific.VALUE
    oItem01.ItemName = ItemName                                                           '�������̸�
    oItem01.ForeignName = oForm01.Items("ItemName").Specific.VALUE
    oItem01.DefaultWarehouse = oForm01.Items("WhsCode").Specific.VALUE                    '�⺻â��
    oItem01.ItemsGroupCode = "102"                                                        'ǰ��׷�
    oItem01.UserFields("U_ItmBsort").VALUE = oForm01.Items("ItmBsort").Specific.VALUE     'ǰ���з�
    oItem01.UserFields("U_ItmMsort").VALUE = oForm01.Items("ItmMsort").Specific.VALUE     'ǰ���ߺз�
    oItem01.UserFields("U_Spec1").VALUE = oForm01.Items("Spec1").Specific.VALUE           '�԰�
    oItem01.UserFields("U_Size").VALUE = oForm01.Items("Spec1").Specific.VALUE            '������
    oItem01.PurchaseUnit = oForm01.Items("Unit1").Specific.VALUE                          '����óǥ������
    oItem01.SalesUnit = oForm01.Items("Unit1").Specific.VALUE                             '�Ǹ�óǥ������
    oItem01.UserFields("U_ObasUnit").VALUE = "101"                                        '���Ա��ش���
    oItem01.UserFields("U_SbasUnit").VALUE = "101"                                        '�Ǹű��ش���
    
    oItem01.UserFields("U_TeamCode").VALUE = oForm01.Items("TeamCode").Specific.VALUE '�����Ҽ�
    oItem01.UserFields("U_YearPdYN").VALUE = oForm01.Items("YearPdYN").Specific.VALUE '����ǰ����
    
    RetVal = oItem01.Add
    
    If (0 <> RetVal) Then
        Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
        GoTo Create_Itemcode_Error
    End If

    If (Sbo_Company.InTransaction = True) Then
        Sbo_Company.EndTransaction wf_Commit
    End If
'
'        oRadioBtn = oForm01.Items("RadioBtn").Specific.Value
        oCardCode = oForm01.Items("CardCode").Specific.VALUE
        oCpNaming = oForm01.Items("CpNaming").Specific.VALUE
        oWhsCode = oForm01.Items("WhsCode").Specific.VALUE
        oItmBsort = oForm01.Items("ItmBsort").Specific.VALUE
        oItmMsort = oForm01.Items("ItmMsort").Specific.VALUE
    
    Set oItem01 = Nothing
    Create_Itemcode = True
    
Exit Function
Create_Itemcode_Error:
    Set oItem01 = Nothing
    If Sbo_Company.InTransaction Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Create_Itemcode = False
    MDC_Com.MDC_GF_Message "Create_oItem:" & Err.Description & ErrMsg, "E"
End Function

Sub Create_Form_Code()
On Error GoTo Create_Form_Code_Error
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim sQry As String
    Dim sTxt As String
    Dim sCho As String
    Dim sSeq As String
    Dim sDate As String
    Dim ErrNum As Integer
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If oForm01.DataSources.UserDataSources("RadioBtn").VALUE = "A" Then
        sQry = "SELECT U_CdNaming FROM [OCRD] WHERE CardCode = '" & Trim(oForm01.Items("CardCode").Specific.VALUE) & "'"
        oRecordSet01.DoQuery sQry
        
        If oRecordSet01.Fields(0).VALUE = "" Then
            ErrNum = 1
            GoTo Create_Form_Code_Error
        Else
            sTxt = oRecordSet01.Fields(0).VALUE
        End If
        
    ElseIf oForm01.DataSources.UserDataSources("RadioBtn").VALUE = "B" Then
        sTxt = "S"
        
    ElseIf oForm01.DataSources.UserDataSources("RadioBtn").VALUE = "C" Then
        sTxt = "Z"
        
    ElseIf oForm01.DataSources.UserDataSources("RadioBtn").VALUE = "D" Then
        sTxt = "F"
        
    ElseIf oForm01.DataSources.UserDataSources("RadioBtn").VALUE = "E" Then
        sTxt = "E"
        
    ElseIf oForm01.DataSources.UserDataSources("RadioBtn").VALUE = "R" Then
        sTxt = "R"
        
    End If
    
    sDate = oForm01.Items("MatDate").Specific.VALUE
    sDate = Mid(sDate, 1, 6)
    sCho = sTxt & oForm01.Items("CpNaming").Specific.VALUE
    
    sQry = "SELECT MAX(SubString(ItemCode,3,9)) FROM OITM WHERE SubString(ItemCode,1,2) = '" & sCho & "' AND "
    sQry = sQry & "SubString(ItemCode,3,6) = '" & sDate & "'"
    oRecordSet01.DoQuery sQry
    If oRecordSet01.Fields(0).VALUE = "" Then
        oForm01.Items("ItemCode").Specific.VALUE = sTxt & oForm01.Items("CpNaming").Specific.VALUE & sDate & "001"
    Else
        sSeq = Trim(oRecordSet01.Fields(0).VALUE) + 1
        oForm01.Items("ItemCode").Specific.VALUE = sTxt & oForm01.Items("CpNaming").Specific.VALUE & sSeq
    End If
    Exit Sub
    
Create_Form_Code_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "�ŷ�ó ������ �����Ϳ� �����ڵ尡 �����Ǿ����ϴ�.", "E"
    End If
End Sub
