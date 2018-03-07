VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_SD081"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_SD081.cls
'//  Module         : SD
'//  Description    : �����ѵ� �ʰ���û > �ŷ�ó�� �ѵ� ��ȸ
'//  FormType       : PS_SD081
'//  Create Date    : 2010.10.21
'//  Modified Date  :
'//  Creator        : Ryu Yung Jo
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_SD081H As SAPbouiCOM.DBDataSource '������
Private oDS_PS_SD081L As SAPbouiCOM.DBDataSource '��϶���

Public oBaseForm01 As SAPbouiCOM.Form '//�θ���
Public oBaseItemUID01 As String
Public oBaseColUID01 As String
Public oBaseColRow01 As Long

Private oLast_Item_UID      As String                     'Ŭ�������� ������ ������ ������ Uid��
Private oLast_Col_UID       As String                     '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Col�� Uid��
Private oLast_Col_Row       As Long                       '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Row��

'*******************************************************************
' .srf ���Ϸκ��� ���� �ε��Ѵ�.
'*******************************************************************
Public Sub LoadForm(ByRef oForm02 As SAPbouiCOM.Form, ByVal oItemUID02 As String, ByVal oColUID02 As String, ByVal oColRow02 As Long)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_SD081.srf")
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
    
    oFormUniqueID01 = "PS_SD081_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//���߰�
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '�� �Ҵ�
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO����϶�
        
    oForm01.Freeze True
    Set oBaseForm01 = oForm02
    oBaseItemUID01 = oItemUID02
    oBaseColUID01 = oColUID02
    oBaseColRow01 = oColRow02
    
    Call CreateItems
    Call ComboBox_Setting
    Call Initialization
'    Call FormItemEnabled
    
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
    Set oForm01 = Nothing
    Sbo_Application.SetStatusBarMessage "Form_Load Error:" & Err.Description, bmt_Short, True
End Sub

'****************************************************************************************************************
'// ItemEventHander
'****************************************************************************************************************
Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_ItemEvent_Error
    Dim i&
    Dim ErrNum&
    Dim TempForm01 As Variant
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "Btn01" Then
                    Call SetBaseForm
                ElseIf pval.ItemUID = "Btn02" Then
                    Call LoadData
                End If
'et_KEY_DOWN ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "CardCode" Then
                        If oForm01.Items(pval.ItemUID).Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
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
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
'et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "CardCode" Then
                        FlushToItemValue pval.ItemUID
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
                Set oDS_PS_SD081H = Nothing
                Set oDS_PS_SD081L = Nothing
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
    MDC_Com.MDC_GF_Message "Raise_ItemEvent_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_MenuEvent_Error
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '���
            Case "1286": '�ݱ�
            Case "1293": '�����
            Case "1281": 'ã��
            Case "1282": '�߰�
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '���
            Case "1286": '�ݱ�
            Case "1293": '�����
            Case "1281": 'ã��
            Case "1282": '�߰�
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
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

Public Sub Raise_RightClickEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_RightClickEvent_Error
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub CreateItems()
On Error GoTo CreateItems_Error
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_SD081L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    Call oForm01.DataSources.UserDataSources.Add("BPLId", dt_SHORT_TEXT, 10)
    oForm01.Items("BPLId").Specific.DataBind.SetBound True, "", "BPLId"
'    oForm01.DataSources.UserDataSources.Item("BPLId").Value = oBaseForm01.Items("BPLId").Specific.Value

    Call oForm01.DataSources.UserDataSources.Add("CardCode", dt_SHORT_TEXT, 10)
    oForm01.Items("CardCode").Specific.DataBind.SetBound True, "", "CardCode"
'    oForm01.DataSources.UserDataSources.Item("CardCode").Value = oBaseForm01.Items("CardCode").Specific.Value

    Call oForm01.DataSources.UserDataSources.Add("CardName", dt_SHORT_TEXT, 100)
    oForm01.Items("CardName").Specific.DataBind.SetBound True, "", "CardName"
'    oForm01.DataSources.UserDataSources.Item("CardName").Value = oBaseForm01.Items("CardCode").Specific.Value

    Call oForm01.DataSources.UserDataSources.Add("DocDate", dt_DATE, 8)
    oForm01.Items("DocDate").Specific.DataBind.SetBound True, "", "DocDate"
    oForm01.DataSources.UserDataSources.Item("DocDate").VALUE = Date  'Format(Now, "YYYY-MM") & "-01"
    
        
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
CreateItems_Error:
    MDC_Com.MDC_GF_Message "CreateItems_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub ComboBox_Setting()
On Error GoTo ComboBox_Setting_Error
    '//�޺��� �⺻������
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    '// �����
    Set oCombo = oForm01.Items("BPLId").Specific
    sQry = "SELECT BPLId, BPLName From [OBPL]"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    oCombo.Select "4", psk_ByValue
    
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
    
    '//���̵� ����� ����
    Set oCombo = oForm01.Items("BPLId").Specific
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
    '//���̵� ��� ����
'    oForm01.Items("CntcCode").Specific.Value = MDC_PS_Common.User_MSTCOD
    
    '//���̵� �μ� ����
'    Set oCombo = oForm01.Items("DeptCode").Specific
'    oCombo.Select MDC_PS_Common.User_DeptCode, psk_ByValue
    Set oCombo = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Initialization_Error:
    Set oCombo = Nothing
    MDC_Com.MDC_GF_Message "Initialization_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error
    Dim i&
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
        Case "CardCode"
            sQry = "Select CardName From [OCRD] Where CardCode = '" & Trim(oForm01.DataSources.UserDataSources("CardCode").VALUE) & "'"
            oRecordSet01.DoQuery sQry
            
            oForm01.DataSources.UserDataSources("CardName").VALUE = Trim(oRecordSet01.Fields(0).VALUE)
    End Select

    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FlushToItemValue_Error:
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub LoadData()
On Error GoTo LoadData_Error
    Dim i As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLId$, CardCode$, DocDate
        
    BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
    CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
    DocDate = Trim(oForm01.Items("DocDate").Specific.VALUE)
    
    
    If BPLId = "" Then BPLId = "%"
    If CardCode = "" Then CardCode = "%"
    
    If DocDate = "" Then
        MDC_Com.MDC_GF_Message "�������ڰ� �����ϴ�. Ȯ���ϼ���.:" & Err.Number & " - " & Err.Description, "W"
        Exit Sub
    End If
    
    
    sQry = "EXEC [S139_hando] '" & CardCode & "','" & DocDate & "'"
    oRecordSet01.DoQuery sQry
    
    oMat01.Clear
    oDS_PS_SD081L.Clear
    
    If oRecordSet01.RecordCount = 0 Then
        MDC_Com.MDC_GF_Message "��ȸ ����� �����ϴ�. Ȯ���ϼ���.:" & Err.Number & " - " & Err.Description, "W"
        Set oRecordSet01 = Nothing
        Exit Sub
    End If
    
    oForm01.Freeze True
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("��ȸ����!", oRecordSet01.RecordCount, False)
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_SD081L.Size Then
            oDS_PS_SD081L.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_SD081L.Offset = i
        oDS_PS_SD081L.setValue "U_LineNum", i, i + 1
        oDS_PS_SD081L.setValue "U_ColReg02", i, Trim(oRecordSet01.Fields("CardCode").VALUE)
        oDS_PS_SD081L.setValue "U_ColReg03", i, Trim(oRecordSet01.Fields("CardName").VALUE)
        oDS_PS_SD081L.setValue "U_ColSum01", i, Trim(oRecordSet01.Fields("U_CreditP").VALUE) '���翩�� �ݾ�
        oDS_PS_SD081L.setValue "U_ColSum02", i, Trim(oRecordSet01.Fields("U_MiSuP").VALUE) '�̼���
        oDS_PS_SD081L.setValue "U_ColSum06", i, Trim(oRecordSet01.Fields("U_ArfAmt").VALUE)  '����
        oDS_PS_SD081L.setValue "U_ColSum07", i, Trim(oRecordSet01.Fields("U_Budo").VALUE)  '�ε�
        oDS_PS_SD081L.setValue "U_ColSum08", i, Trim(oRecordSet01.Fields("U_MisuTot").VALUE)  'ä�ǰ�
        
        'Trim (oRecordSet01.Fields("U_MiSuP").VALUE) '�̼��ݾ�
        'Trim(oRecordSet01.Fields("U_ArfAmt").VALUE) '�����ܾ�
        'Trim(oRecordSet01.Fields("U_Budo").VALUE) '�ε�����
                
        oDS_PS_SD081L.setValue "U_ColSum03", i, Trim(oRecordSet01.Fields("U_Balance").VALUE)
        oDS_PS_SD081L.setValue "U_ColSum04", i, Trim(oRecordSet01.Fields("U_OutPreP").VALUE)
        oDS_PS_SD081L.setValue "U_ColSum05", i, Trim(oRecordSet01.Fields("OverAmt").VALUE)

        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "�� ��ȸ��...!"
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

Private Sub SetBaseForm()
On Error GoTo SetBaseForm_Error
    Dim i As Long
    Dim j As Long
    Dim ErrNum As Long
    Dim sRow As Long
    
    Dim oBaseMat01 As SAPbouiCOM.Matrix
    Dim oBaseDS_PS_SD080L As SAPbouiCOM.DBDataSource
    Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific
    Set oBaseDS_PS_SD080L = oBaseForm01.DataSources.DBDataSources("@PS_SD080L")
    
    oBaseForm01.Freeze True
    oBaseMat01.Clear
    oBaseMat01.FlushToDataSource
    oBaseMat01.LoadFromDataSource
    
    oMat01.FlushToDataSource
    sRow = 0
    For i = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PS_SD081L.GetValue("U_ColReg01", i)) = "Y" Then
            If sRow + 1 > oBaseDS_PS_SD080L.Size Then
                oBaseDS_PS_SD080L.InsertRecord (sRow)
            End If
            
            oBaseMat01.AddRow
            oBaseDS_PS_SD080L.Offset = sRow
            oBaseDS_PS_SD080L.setValue "U_LineNum", sRow, sRow + 1
            oBaseDS_PS_SD080L.setValue "U_CardCode", sRow, Trim(oDS_PS_SD081L.GetValue("U_ColReg02", i))
            oBaseDS_PS_SD080L.setValue "U_CardName", sRow, Trim(oDS_PS_SD081L.GetValue("U_ColReg03", i))
            oBaseDS_PS_SD080L.setValue "U_CreditP", sRow, Trim(oDS_PS_SD081L.GetValue("U_ColSum01", i))
            oBaseDS_PS_SD080L.setValue "U_MiSuP", sRow, Trim(oDS_PS_SD081L.GetValue("U_ColSum02", i))
            oBaseDS_PS_SD080L.setValue "U_Balance", sRow, Trim(oDS_PS_SD081L.GetValue("U_ColSum03", i))
            oBaseDS_PS_SD080L.setValue "U_Bill", sRow, Trim(oDS_PS_SD081L.GetValue("U_ColSum06", i))
            oBaseDS_PS_SD080L.setValue "U_Budo", sRow, Trim(oDS_PS_SD081L.GetValue("U_ColSum07", i))
            oBaseDS_PS_SD080L.setValue "U_TotAmt", sRow, Trim(oDS_PS_SD081L.GetValue("U_ColSum08", i))
            
            oBaseDS_PS_SD080L.setValue "U_OutPreP", sRow, Trim(oDS_PS_SD081L.GetValue("U_ColSum04", i))
            oBaseDS_PS_SD080L.setValue "U_RequestP", sRow, Trim(oDS_PS_SD081L.GetValue("U_ColSum05", i))
            
            
            oBaseColRow01 = oBaseColRow01 + 1
            sRow = sRow + 1
        End If
    Next i

    oBaseMat01.LoadFromDataSource
    oBaseForm01.Freeze False
    Call oForm01.Close
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
SetBaseForm_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message " ", "E"
    Else
        MDC_Com.MDC_GF_Message "SetBaseForm_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub