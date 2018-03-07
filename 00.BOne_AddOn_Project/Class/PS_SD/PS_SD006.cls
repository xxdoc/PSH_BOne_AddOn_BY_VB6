VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_SD006"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_SD006.cls
'//  Module         : SD
'//  Description    : �ǸŰ�ȹ���(�и�)
'//  FormType       : PS_SD006
'//  Create Date    : 2014.09.01
'//  Modified Date  :
'//  Creator        : N.G.Y
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01             As SAPbouiCOM.Form
Public oMat01             As SAPbouiCOM.Matrix
Private oDS_PS_SD006H As SAPbouiCOM.DBDataSource    '������
Private oDS_PS_SD006L As SAPbouiCOM.DBDataSource    '��϶���

Private oLast_Item_UID      As String                     'Ŭ�������� ������ ������ ������ Uid��
Private oLast_Col_UID       As String                     '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Col�� Uid��
Private oLast_Col_Row       As Long                       '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Row��

Private oLast_Mode&

Private oYYYYMM$

'****************************************************************************************************************
' .srf ���Ϸκ��� ���� �ε��Ѵ�.
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_SD006.srf")
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
    
    oFormUniqueID01 = "PS_SD006_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//���߰�
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '�� �Ҵ�
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
  
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    'ȭ��Ű��(ȭ�鿡�� ����Ű���� ��� �ִ� �������� Uid��)
    oForm01.DataBrowser.BrowseBy = "Code"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting
    Call Initialization
    Call FormItemEnabled
    
    
    
    oForm01.EnableMenu ("1283"), True         '// ����
    oForm01.EnableMenu ("1287"), True         '// ����
    oForm01.EnableMenu ("1286"), False        '// �ݱ�
    oForm01.EnableMenu ("1284"), False        '// ���
    oForm01.EnableMenu ("1293"), True         '// �����
        
    oForm01.Update
        
    oForm01.Freeze False
    oForm01.Visible = True
    
     '���� Ŀ����ġ
    oForm01.ActiveItem = "CardCode"
    
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
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    Dim ItemCode        As String
    Dim rate As Currency
    Dim StdWgt As Currency
    Dim T_Wgt1 As Double, T_Wgt2 As Double, T_Wgt3 As Double, T_Wgt4 As Double, T_Wgt5 As Double, T_Wgt6 As Double
    Dim T_Wgt7 As Double, T_Wgt8 As Double, T_Wgt9 As Double, T_Wgt10 As Double, T_Wgt11 As Double, T_Wgt12 As Double
    Dim T_Amt1 As Double, T_Amt2 As Double, T_Amt3 As Double, T_Amt4 As Double, T_Amt5 As Double, T_Amt6 As Double
    Dim T_Amt7 As Double, T_Amt8 As Double, T_Amt9 As Double, T_Amt10 As Double, T_Amt11 As Double, T_Amt12 As Double
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Dim ChildForm01 As Variant
    Set ChildForm01 = New PS_SM010
    
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
                        
                        If oForm01.Mode = fm_ADD_MODE Then
                        
                            Dim Code$, CardCode$, Year$
                            CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
                            Year = Trim(oForm01.Items("Year").Specific.VALUE)
                            
                            Code = Right(Year, 2) + CardCode
                                                    
                            Call oDS_PS_SD006H.setValue("Code", 0, Code)
                        End If
                        
                    End If
                End If
'et_KEY_DOWN ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "CardCode" Then
                        If oForm01.Items("CardCode").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "Mat01" Then
                        
                        If pval.ColUID = "ItemCode" Then
                            If oMat01.Columns("ItemCode").Cells(pval.Row).Specific.VALUE = "" Then
                                Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row)
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
                    If oForm01.Mode = fm_ADD_MODE And pval.Action_Success = True Then
                        oForm01.Mode = fm_OK_MODE
                        Call Sbo_Application.ActivateMenuItem("1282")
                    End If
                End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_MATRIX_LINK_PRESSED '//8
'et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "CardCode" Then   '//�ŷ�ó����ȸ
                        oForm01.Freeze True
                        
                        sQry = "SELECT CardName FROM [OCRD] WHERE CardCode = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
                        oRecordSet01.DoQuery sQry
                        oForm01.Items("CardName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
 '                       Call LoadData
                        oForm01.Freeze False
                    End If
                
                    If pval.ItemUID = "Mat01" Then
                        oMat01.FlushToDataSource
                        If pval.ColUID = "ItemCode" Then
                            If (pval.Row = oMat01.RowCount Or oMat01.VisualRowCount = 0) And Trim(oMat01.Columns("ItemCode").Cells(pval.Row).Specific.VALUE) <> "" Then
                                oMat01.FlushToDataSource
                                Call Add_MatrixRow(oMat01.RowCount, False)
                                oMat01.Columns("ItemCode").Cells(pval.Row).Click ct_Regular
                            End If

                            ItemCode = Trim(oDS_PS_SD006L.GetValue("U_ItemCode", pval.Row - 1))
                            sQry = "Select ItemName "
                            sQry = sQry & " From OITM "
                            sQry = sQry & " Where ItemCode = '" & ItemCode & "'"
                            oRecordSet01.DoQuery sQry

                            oForm01.Freeze True
                            oMat01.FlushToDataSource

                            If oRecordSet01.RecordCount = 0 Then

                                '��Ʈ������ �����͸� ���� ���ε��ϸ� �̺�Ʈ�� ����Ǳ� ������ DataSource�� ���ε��ϴ� ������� ����(2011.11.22 �۸���)
                                Call oDS_PS_SD006L.setValue("U_ItemName", pval.Row - 1, "")
                                
                                MDC_Com.MDC_GF_Message "��ȸ ����� �����ϴ�. Ȯ���ϼ���.:" & Err.Number & " - " & Err.Description, "W"

                                Set oRecordSet01 = Nothing
                                oMat01.LoadFromDataSource
                                oForm01.Freeze False
                                Exit Sub

                            End If

                            '��Ʈ������ �����͸� ���� ���ε��ϸ� �̺�Ʈ�� ����Ǳ� ������ DataSource�� ���ε��ϴ� ������� ����(2011.11.22 �۸���)
                            Call oDS_PS_SD006L.setValue("U_ItemName", pval.Row - 1, Trim(oRecordSet01.Fields("ItemName").VALUE)) 'ǰ���̸�
                            
                            oMat01.LoadFromDataSource
                            oForm01.Freeze False
                        Else
                            If pval.ColUID = "Wgt9" Or pval.ColUID = "Prc9" Then
                                oMat01.Columns("Amt9").Cells(i + 1).Specific.VALUE = oMat01.Columns("Wgt9").Cells(i + 1).Specific.VALUE * oMat01.Columns("Prc9").Cells(i + 1).Specific.VALUE
                            End If
                            For i = 0 To oMat01.VisualRowCount - 1
                                T_Wgt1 = T_Wgt1 + oMat01.Columns("Wgt1").Cells(i + 1).Specific.VALUE
                                T_Wgt2 = T_Wgt2 + oMat01.Columns("Wgt2").Cells(i + 1).Specific.VALUE
                                T_Wgt3 = T_Wgt3 + oMat01.Columns("Wgt3").Cells(i + 1).Specific.VALUE
                                T_Wgt4 = T_Wgt4 + oMat01.Columns("Wgt4").Cells(i + 1).Specific.VALUE
                                T_Wgt5 = T_Wgt5 + oMat01.Columns("Wgt5").Cells(i + 1).Specific.VALUE
                                T_Wgt6 = T_Wgt6 + oMat01.Columns("Wgt6").Cells(i + 1).Specific.VALUE
                                T_Wgt7 = T_Wgt7 + oMat01.Columns("Wgt7").Cells(i + 1).Specific.VALUE
                                T_Wgt8 = T_Wgt8 + oMat01.Columns("Wgt8").Cells(i + 1).Specific.VALUE
                                T_Wgt9 = T_Wgt9 + oMat01.Columns("Wgt9").Cells(i + 1).Specific.VALUE
                                T_Wgt10 = T_Wgt10 + oMat01.Columns("Wgt10").Cells(i + 1).Specific.VALUE
                                T_Wgt11 = T_Wgt11 + oMat01.Columns("Wgt11").Cells(i + 1).Specific.VALUE
                                T_Wgt12 = T_Wgt12 + oMat01.Columns("Wgt12").Cells(i + 1).Specific.VALUE
                                
                                T_Amt1 = T_Amt1 + oMat01.Columns("Amt1").Cells(i + 1).Specific.VALUE
                                T_Amt2 = T_Amt2 + oMat01.Columns("Amt2").Cells(i + 1).Specific.VALUE
                                T_Amt3 = T_Amt3 + oMat01.Columns("Amt3").Cells(i + 1).Specific.VALUE
                                T_Amt4 = T_Amt4 + oMat01.Columns("Amt4").Cells(i + 1).Specific.VALUE
                                T_Amt5 = T_Amt5 + oMat01.Columns("Amt5").Cells(i + 1).Specific.VALUE
                                T_Amt6 = T_Amt6 + oMat01.Columns("Amt6").Cells(i + 1).Specific.VALUE
                                T_Amt7 = T_Amt7 + oMat01.Columns("Amt7").Cells(i + 1).Specific.VALUE
                                T_Amt8 = T_Amt8 + oMat01.Columns("Amt8").Cells(i + 1).Specific.VALUE
                                T_Amt9 = T_Amt9 + oMat01.Columns("Amt9").Cells(i + 1).Specific.VALUE
                                T_Amt10 = T_Amt10 + oMat01.Columns("Amt10").Cells(i + 1).Specific.VALUE
                                T_Amt11 = T_Amt11 + oMat01.Columns("Amt11").Cells(i + 1).Specific.VALUE
                                T_Amt12 = T_Amt12 + oMat01.Columns("Amt12").Cells(i + 1).Specific.VALUE
                                
                            Next i
                            
                            oForm01.Items("T_Wgt1").Specific.VALUE = T_Wgt1
                            oForm01.Items("T_Wgt2").Specific.VALUE = T_Wgt2
                            oForm01.Items("T_Wgt3").Specific.VALUE = T_Wgt3
                            oForm01.Items("T_Wgt4").Specific.VALUE = T_Wgt4
                            oForm01.Items("T_Wgt5").Specific.VALUE = T_Wgt5
                            oForm01.Items("T_Wgt6").Specific.VALUE = T_Wgt6
                            oForm01.Items("T_Wgt7").Specific.VALUE = T_Wgt7
                            oForm01.Items("T_Wgt8").Specific.VALUE = T_Wgt8
                            oForm01.Items("T_Wgt9").Specific.VALUE = T_Wgt9
                            oForm01.Items("T_Wgt10").Specific.VALUE = T_Wgt10
                            oForm01.Items("T_Wgt11").Specific.VALUE = T_Wgt11
                            oForm01.Items("T_Wgt12").Specific.VALUE = T_Wgt12
                            
                            oForm01.Items("T_Amt1").Specific.VALUE = T_Amt1
                            oForm01.Items("T_Amt2").Specific.VALUE = T_Amt2
                            oForm01.Items("T_Amt3").Specific.VALUE = T_Amt3
                            oForm01.Items("T_Amt4").Specific.VALUE = T_Amt4
                            oForm01.Items("T_Amt5").Specific.VALUE = T_Amt5
                            oForm01.Items("T_Amt6").Specific.VALUE = T_Amt6
                            oForm01.Items("T_Amt7").Specific.VALUE = T_Amt7
                            oForm01.Items("T_Amt8").Specific.VALUE = T_Amt8
                            oForm01.Items("T_Amt9").Specific.VALUE = T_Amt9
                            oForm01.Items("T_Amt10").Specific.VALUE = T_Amt10
                            oForm01.Items("T_Amt11").Specific.VALUE = T_Amt11
                            oForm01.Items("T_Amt12").Specific.VALUE = T_Amt12
                        
                        End If
                    End If
                 End If
                 
            Case et_MATRIX_LOAD: '//11
                
                 For i = 0 To oMat01.VisualRowCount - 1
                    T_Wgt1 = T_Wgt1 + oMat01.Columns("Wgt1").Cells(i + 1).Specific.VALUE
                    T_Wgt2 = T_Wgt2 + oMat01.Columns("Wgt2").Cells(i + 1).Specific.VALUE
                    T_Wgt3 = T_Wgt3 + oMat01.Columns("Wgt3").Cells(i + 1).Specific.VALUE
                    T_Wgt4 = T_Wgt4 + oMat01.Columns("Wgt4").Cells(i + 1).Specific.VALUE
                    T_Wgt5 = T_Wgt5 + oMat01.Columns("Wgt5").Cells(i + 1).Specific.VALUE
                    T_Wgt6 = T_Wgt6 + oMat01.Columns("Wgt6").Cells(i + 1).Specific.VALUE
                    T_Wgt7 = T_Wgt7 + oMat01.Columns("Wgt7").Cells(i + 1).Specific.VALUE
                    T_Wgt8 = T_Wgt8 + oMat01.Columns("Wgt8").Cells(i + 1).Specific.VALUE
                    T_Wgt9 = T_Wgt9 + oMat01.Columns("Wgt9").Cells(i + 1).Specific.VALUE
                    T_Wgt10 = T_Wgt10 + oMat01.Columns("Wgt10").Cells(i + 1).Specific.VALUE
                    T_Wgt11 = T_Wgt11 + oMat01.Columns("Wgt11").Cells(i + 1).Specific.VALUE
                    T_Wgt12 = T_Wgt12 + oMat01.Columns("Wgt12").Cells(i + 1).Specific.VALUE
                    
                    T_Amt1 = T_Amt1 + oMat01.Columns("Amt1").Cells(i + 1).Specific.VALUE
                    T_Amt2 = T_Amt2 + oMat01.Columns("Amt2").Cells(i + 1).Specific.VALUE
                    T_Amt3 = T_Amt3 + oMat01.Columns("Amt3").Cells(i + 1).Specific.VALUE
                    T_Amt4 = T_Amt4 + oMat01.Columns("Amt4").Cells(i + 1).Specific.VALUE
                    T_Amt5 = T_Amt5 + oMat01.Columns("Amt5").Cells(i + 1).Specific.VALUE
                    T_Amt6 = T_Amt6 + oMat01.Columns("Amt6").Cells(i + 1).Specific.VALUE
                    T_Amt7 = T_Amt7 + oMat01.Columns("Amt7").Cells(i + 1).Specific.VALUE
                    T_Amt8 = T_Amt8 + oMat01.Columns("Amt8").Cells(i + 1).Specific.VALUE
                    T_Amt9 = T_Amt9 + oMat01.Columns("Amt9").Cells(i + 1).Specific.VALUE
                    T_Amt10 = T_Amt10 + oMat01.Columns("Amt10").Cells(i + 1).Specific.VALUE
                    T_Amt11 = T_Amt11 + oMat01.Columns("Amt11").Cells(i + 1).Specific.VALUE
                    T_Amt12 = T_Amt12 + oMat01.Columns("Amt12").Cells(i + 1).Specific.VALUE
                    
                Next i
                
                oForm01.Items("T_Wgt1").Specific.VALUE = T_Wgt1
                oForm01.Items("T_Wgt2").Specific.VALUE = T_Wgt2
                oForm01.Items("T_Wgt3").Specific.VALUE = T_Wgt3
                oForm01.Items("T_Wgt4").Specific.VALUE = T_Wgt4
                oForm01.Items("T_Wgt5").Specific.VALUE = T_Wgt5
                oForm01.Items("T_Wgt6").Specific.VALUE = T_Wgt6
                oForm01.Items("T_Wgt7").Specific.VALUE = T_Wgt7
                oForm01.Items("T_Wgt8").Specific.VALUE = T_Wgt8
                oForm01.Items("T_Wgt9").Specific.VALUE = T_Wgt9
                oForm01.Items("T_Wgt10").Specific.VALUE = T_Wgt10
                oForm01.Items("T_Wgt11").Specific.VALUE = T_Wgt11
                oForm01.Items("T_Wgt12").Specific.VALUE = T_Wgt12
                
                oForm01.Items("T_Amt1").Specific.VALUE = T_Amt1
                oForm01.Items("T_Amt2").Specific.VALUE = T_Amt2
                oForm01.Items("T_Amt3").Specific.VALUE = T_Amt3
                oForm01.Items("T_Amt4").Specific.VALUE = T_Amt4
                oForm01.Items("T_Amt5").Specific.VALUE = T_Amt5
                oForm01.Items("T_Amt6").Specific.VALUE = T_Amt6
                oForm01.Items("T_Amt7").Specific.VALUE = T_Amt7
                oForm01.Items("T_Amt8").Specific.VALUE = T_Amt8
                oForm01.Items("T_Amt9").Specific.VALUE = T_Amt9
                oForm01.Items("T_Amt10").Specific.VALUE = T_Amt10
                oForm01.Items("T_Amt11").Specific.VALUE = T_Amt11
                oForm01.Items("T_Amt12").Specific.VALUE = T_Amt12
                
            
                
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
                Set oDS_PS_SD006H = Nothing
                Set oDS_PS_SD006L = Nothing
        End Select
    End If
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
    oForm01.Freeze False
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "Raise_ItemEvent_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_MenuEvent_Error
    Dim i&
    
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
'[1293:�����] //////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1293": '�����
                If oMat01.RowCount <> oMat01.VisualRowCount Then
                    For i = 0 To oMat01.VisualRowCount - 1
                        oMat01.Columns("LineNum").Cells(i + 1).Specific.VALUE = i + 1
                    Next i
                    
                    oMat01.FlushToDataSource
                    oDS_PS_SD006L.RemoveRecord oDS_PS_SD006L.Size - 1       '// Mat01�� ����������(�����) ����
                    oMat01.Clear
                    oMat01.LoadFromDataSource
                End If
            Case "1281": 'ã��
                oForm01.Freeze True
                Call FormItemEnabled
                Call Initialization
'                oForm01.Items("CycleCod").Click ct_Regular
                oForm01.Freeze False
            Case "1282": '�߰�
                oForm01.Freeze True
                Call FormItemEnabled
                Call Initialization
                oForm01.Freeze False
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
                oForm01.Freeze True
                Call FormItemEnabled
'                If oMat01.VisualRowCount > 0 Then
'                    If oMat01.Columns("CycleCod").Cells(oMat01.VisualRowCount).Specific.Value <> "" Then
'                        Add_MatrixRow oMat01.RowCount, False
'                    End If
'                End If
                oForm01.Freeze False
            Case "1287": '// ����
                oForm01.Freeze True
                oDS_PS_SD006H.setValue "Code", 0, ""
                
                For i = 0 To oMat01.VisualRowCount - 1
                    oMat01.FlushToDataSource
                    oDS_PS_SD006L.setValue "Code", i, ""
                    oMat01.LoadFromDataSource
                Next i
                oForm01.Freeze False
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
        '//�۾�
    ElseIf (eventInfo.BeforeAction = False) Then
        '//�۾�
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub CreateItems()
On Error GoTo CreateItems_Error
    '//������� �ҽ� ��ü �Ҵ�
    Set oDS_PS_SD006H = oForm01.DataSources.DBDataSources("@PS_SD006H")
    Set oDS_PS_SD006L = oForm01.DataSources.DBDataSources("@PS_SD006L")
    
    '// ��Ʈ���� ��ü �Ҵ�
    Set oMat01 = oForm01.Items("Mat01").Specific
    
    Call oForm01.DataSources.UserDataSources.Add("T_Wgt1", dt_QUANTITY)
    oForm01.Items("T_Wgt1").Specific.DataBind.SetBound True, "", "T_Wgt1"
    Call oForm01.DataSources.UserDataSources.Add("T_Wgt2", dt_QUANTITY)
    oForm01.Items("T_Wgt2").Specific.DataBind.SetBound True, "", "T_Wgt2"
    Call oForm01.DataSources.UserDataSources.Add("T_Wgt3", dt_QUANTITY)
    oForm01.Items("T_Wgt3").Specific.DataBind.SetBound True, "", "T_Wgt3"
    Call oForm01.DataSources.UserDataSources.Add("T_Wgt4", dt_QUANTITY)
    oForm01.Items("T_Wgt4").Specific.DataBind.SetBound True, "", "T_Wgt4"
    Call oForm01.DataSources.UserDataSources.Add("T_Wgt5", dt_QUANTITY)
    oForm01.Items("T_Wgt5").Specific.DataBind.SetBound True, "", "T_Wgt5"
    Call oForm01.DataSources.UserDataSources.Add("T_Wgt6", dt_QUANTITY)
    oForm01.Items("T_Wgt6").Specific.DataBind.SetBound True, "", "T_Wgt6"
    Call oForm01.DataSources.UserDataSources.Add("T_Wgt7", dt_QUANTITY)
    oForm01.Items("T_Wgt7").Specific.DataBind.SetBound True, "", "T_Wgt7"
    Call oForm01.DataSources.UserDataSources.Add("T_Wgt8", dt_QUANTITY)
    oForm01.Items("T_Wgt8").Specific.DataBind.SetBound True, "", "T_Wgt8"
    Call oForm01.DataSources.UserDataSources.Add("T_Wgt9", dt_QUANTITY)
    oForm01.Items("T_Wgt9").Specific.DataBind.SetBound True, "", "T_Wgt9"
    Call oForm01.DataSources.UserDataSources.Add("T_Wgt10", dt_QUANTITY)
    oForm01.Items("T_Wgt10").Specific.DataBind.SetBound True, "", "T_Wgt10"
    Call oForm01.DataSources.UserDataSources.Add("T_Wgt11", dt_QUANTITY)
    oForm01.Items("T_Wgt11").Specific.DataBind.SetBound True, "", "T_Wgt11"
    Call oForm01.DataSources.UserDataSources.Add("T_Wgt12", dt_QUANTITY)
    oForm01.Items("T_Wgt12").Specific.DataBind.SetBound True, "", "T_Wgt12"
    
    Call oForm01.DataSources.UserDataSources.Add("T_Amt1", dt_SUM)
    oForm01.Items("T_Amt1").Specific.DataBind.SetBound True, "", "T_Amt1"
    Call oForm01.DataSources.UserDataSources.Add("T_Amt2", dt_SUM)
    oForm01.Items("T_Amt2").Specific.DataBind.SetBound True, "", "T_Amt2"
    Call oForm01.DataSources.UserDataSources.Add("T_Amt3", dt_SUM)
    oForm01.Items("T_Amt3").Specific.DataBind.SetBound True, "", "T_Amt3"
    Call oForm01.DataSources.UserDataSources.Add("T_Amt4", dt_SUM)
    oForm01.Items("T_Amt4").Specific.DataBind.SetBound True, "", "T_Amt4"
    Call oForm01.DataSources.UserDataSources.Add("T_Amt5", dt_SUM)
    oForm01.Items("T_Amt5").Specific.DataBind.SetBound True, "", "T_Amt5"
    Call oForm01.DataSources.UserDataSources.Add("T_Amt6", dt_SUM)
    oForm01.Items("T_Amt6").Specific.DataBind.SetBound True, "", "T_Amt6"
    Call oForm01.DataSources.UserDataSources.Add("T_Amt7", dt_SUM)
    oForm01.Items("T_Amt7").Specific.DataBind.SetBound True, "", "T_Amt7"
    Call oForm01.DataSources.UserDataSources.Add("T_Amt8", dt_SUM)
    oForm01.Items("T_Amt8").Specific.DataBind.SetBound True, "", "T_Amt8"
    Call oForm01.DataSources.UserDataSources.Add("T_Amt9", dt_SUM)
    oForm01.Items("T_Amt9").Specific.DataBind.SetBound True, "", "T_Amt9"
    Call oForm01.DataSources.UserDataSources.Add("T_Amt10", dt_SUM)
    oForm01.Items("T_Amt10").Specific.DataBind.SetBound True, "", "T_Amt10"
    Call oForm01.DataSources.UserDataSources.Add("T_Amt11", dt_SUM)
    oForm01.Items("T_Amt11").Specific.DataBind.SetBound True, "", "T_Amt11"
    Call oForm01.DataSources.UserDataSources.Add("T_Amt12", dt_SUM)
    oForm01.Items("T_Amt12").Specific.DataBind.SetBound True, "", "T_Amt12"

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
    
    oForm01.Items("Year").Specific.VALUE = Format(Now, "YYYY")
    Call Add_MatrixRow(0, True)
    Set oCombo = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Initialization_Error:
    Set oCombo = Nothing
    MDC_Com.MDC_GF_Message "Initialization_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub CF_ChooseFromList()
On Error GoTo CF_ChooseFromList_Error
    '//ChooseFromList ����
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
CF_ChooseFromList_Error:
    MDC_Com.MDC_GF_Message "CF_ChooseFromList_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub FormItemEnabled()
On Error GoTo FormItemEnabled_Error
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub


Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum          As Integer
    Dim DocNum          As String

    ErrNum = 0

    '// Check
    Select Case True
        Case oDS_PS_SD006H.GetValue("U_CardCode", 0) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_SD006H.GetValue("U_BPLId", 0) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_SD006H.GetValue("U_Year", 0) = ""
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
        
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "�ŷ�ó�� �ʼ��Է»����Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "������� �ʼ��Է»����Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "�⵵�� �ʼ��Է»����Դϴ�. Ȯ���ϼ���.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function


Sub Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo Add_MatrixRow_Error
    If RowIserted = False Then '//���߰�����
        oDS_PS_SD006L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_SD006L.Offset = oRow
    oDS_PS_SD006L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub



