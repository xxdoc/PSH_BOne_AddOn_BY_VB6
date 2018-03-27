VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_MM070"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_MM070.cls
'//  Module         : MM
'//  Description    : 검수입고
'//  FormType       : PS_MM070
'//  Create Date    : 2010.09.01
'//  Modified Date  :
'//  Creator        : Ryu Yung Jo
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01         As SAPbouiCOM.Form
Public oMat01          As SAPbouiCOM.Matrix
Private oDS_PS_MM070H  As SAPbouiCOM.DBDataSource    '등록헤더
Private oDS_PS_MM070L  As SAPbouiCOM.DBDataSource    '등록라인

Private oLast_Item_UID As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID  As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row  As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&
Private oErrNum&
Private oQEYesNo       As String
Private oPurchase      As String

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error

    Dim i           As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01   As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_MM070.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)

    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_MM070_" & GetTotalFormsCount
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
    Call Add_MatrixRow(0, True)
                  
    FormItemEnabled
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1284"), True         '// 취소
    oForm01.EnableMenu ("1293"), True         '// 행삭제
'    oForm01.EnableMenu ("1299"), True         '// 행닫기
        
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
    Dim ChildForm01 As Variant
    Set ChildForm01 = New PS_SM010
    Dim i&
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    Dim sQry$
    Dim vReturnValue$
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
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
                        
                        
                        If oForm01.Items("DocDate").Specific.VALUE < oForm01.Items("DueDate").Specific.VALUE Then
                            vReturnValue = Sbo_Application.MessageBox("납품일보다 검수일이 빠릅니다. 계속하겠습니까?", 1, "&확인", "&취소")
                            If vReturnValue = 1 Then
                                '검수입고 문서만 등록 시 주석처리_S
                                If oForm01.Mode = fm_ADD_MODE Then
                                    If Add_oPurchaseDeliveryNotes(1) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                                '검수입고 문서만 등록 시 주석처리_E
                                
                                oQEYesNo = oForm01.Items("QEYesNo").Specific.VALUE
                                oPurchase = oForm01.Items("Purchase").Specific.VALUE
                            Else
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                    End If
                    oLast_Mode = oForm01.Mode
                End If
                
'et_KEY_DOWN ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "CardCode" Then
'                        oMat01.Columns("Auto").Editable = False
                        If oForm01.Items("CardCode").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "CntcCode" Then
                        If oForm01.Items("CntcCode").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "GADocLin" Then
                            If oMat01.Columns("GADocLin").Cells(pval.Row).Specific.VALUE = "" Then
                                If Trim(oForm01.Items("BPLId").Specific.VALUE) = "" Or Trim(oForm01.Items("QEYesNo").Specific.VALUE) = "" Or Trim(oForm01.Items("Purchase").Specific.VALUE) = "" Then
                                    MDC_Com.MDC_GF_Message "사업장, 품질검수여부 또는 구매구분을 먼저 선택하세요.", "E"
                                    BubbleEvent = False
                                    Exit Sub
                                Else
                                    Sbo_Application.ActivateMenuItem ("7425")
                                    BubbleEvent = False
                                End If
                            End If
                        ElseIf pval.ColUID = "ItemCode" Then
                            If oMat01.Columns("ItemCode").Cells(pval.Row).Specific.VALUE = "" Then
                                Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row)
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
                
'et_COMBO_SELECT ////////////'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
            
                If pval.ItemUID = "Mat01" Then
                
                    If pval.ColUID = "BadCode1" Then
                    
                        Dim sCount&, sSeq&
                        oForm01.Freeze True
                        sCount = oMat01.Columns("BadCode2").Cells(pval.Row).Specific.ValidValues.Count
                        sSeq = sCount
                        For i = 1 To sCount
                            oMat01.Columns("BadCode2").Cells(pval.Row).Specific.ValidValues.Remove sSeq - 1, psk_Index
                            sSeq = sSeq - 1
                        Next i

                        sQry = "SELECT Distinct Convert(int, b.U_SmalCode) As U_SmalCode, b.U_SmalName From [@PS_PP003H] a Inner Join [@PS_PP003L] b On a.DocEntry = b.DocEntry "
                        sQry = sQry & "Where a.U_BigCode = '1' And a.U_MidCode = '" & Trim(oMat01.Columns("BadCode1").Cells(pval.Row).Specific.VALUE) & "' Order by Convert(int, b.U_SmalCode) "
                        oRecordSet01.DoQuery sQry
                        'oForm01.Items("ItmMSort").Specific.ValidValues.Add "ALL", "ALL"
                        Do Until oRecordSet01.EOF
                            oMat01.Columns("BadCode2").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
                            oRecordSet01.MoveNext
                        Loop
                        oForm01.Freeze False
                    
'                        Sbo_Application.MessageBox ("테스트")

                    End If
                
                End If
            
            Case et_MATRIX_LOAD: '//11
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
'et_FORM_RESIZE//////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_FORM_RESIZE '//20
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
            Case et_LOST_FOCUS: '//4
'et_FORM_UNLOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
                        
                        '검수여부 및 품의형태 세팅 로직 수정(2013.05.23 송명규 수정)
                        If oQEYesNo <> "" Then
                            Call oForm01.Items("QEYesNo").Specific.Select(oQEYesNo, psk_ByValue)
                        Else
                            Call oForm01.Items("QEYesNo").Specific.Select("N", psk_ByValue)
                        End If
                        
                        If oPurchase <> "" Then
                            Call oForm01.Items("Purchase").Specific.Select(oPurchase, psk_ByValue)
                        Else
                            Call oForm01.Items("Purchase").Specific.Select("%", psk_ByValue)
                        End If
                        '검수여부 및 품의형태 세팅 로직 수정(2013.05.23 송명규 수정)
                        
                    End If
                End If
            Case et_KEY_DOWN: '//2
'et_COMBO_SELECT ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_COMBO_SELECT: '//5
                If pval.ItemUID = "Purchase" Or pval.ItemUID = "BPLId" Then
                    oForm01.Freeze True
                    oMat01.Clear
                    oDS_PS_MM070L.Clear
                    If oForm01.Mode = fm_ADD_MODE Then
                        Call Add_MatrixRow(0, False)
                    End If
                    
                    'If Trim(oForm01.Items("Purchase").Specific.VALUE) = "30" Or Trim(oForm01.Items("Purchase").Specific.VALUE) = "40" Then
                    If Trim(oForm01.Items("Purchase").Specific.VALUE) = "30" Or Trim(oForm01.Items("Purchase").Specific.VALUE) = "40" Or Trim(oForm01.Items("Purchase").Specific.VALUE) = "60" Then
                    
                        If Trim(oForm01.Items("Purchase").Specific.VALUE) = "40" Then '외주제작품의를 검수입고하고자 할 때

                            Dim TeamCode As String

                            TeamCode = MDC_PS_Common.User_TeamCode

                            If TeamCode = "1400" Or TeamCode = "2600" Then '1400 : 창원사업장 품보팀, 2600 : 부산사업장 품보팀

                            Else

                                Call MDC_Com.MDC_GF_Message("외주제작품의 검수입고는 품질보증팀 담당자만 가능합니다.", "E")
                                Call oDS_PS_MM070H.setValue("U_Purchase", 0, "%")
                                BubbleEvent = False

                            End If

                        End If
                    
'                        oMat01.Columns("ItemName").Editable = True
'                        oMat01.Columns("OutSize").Editable = True
'                        oMat01.Columns("OutUnit").Editable = True
'                        oMat01.Columns("WhsCode").Editable = False
                    Else
'                        oMat01.Columns("ItemName").Editable = False
'                        oMat01.Columns("OutSize").Editable = False
'                        oMat01.Columns("OutUnit").Editable = False
'                        oMat01.Columns("WhsCode").Editable = True
                    End If
                    oForm01.Freeze False
                ElseIf pval.ItemUID = "Mat01" And pval.ColUID = "BadCode1" Then
'                    Dim sCount&, sSeq&
'                    oForm01.Freeze True
'                    sCount = oMat01.Columns("BadCode2").Cells(pval.Row).Specific.ValidValues.Count
'                    sSeq = sCount
'                    For i = 1 To sCount
'                        oMat01.Columns("BadCode2").Cells(pval.Row).Specific.ValidValues.Remove sSeq - 1, psk_Index
'                        sSeq = sSeq - 1
'                    Next i
'
'                    sQry = "SELECT Distinct Convert(int, b.U_SmalCode) As U_SmalCode, b.U_SmalName From [@PS_PP003H] a Inner Join [@PS_PP003L] b On a.DocEntry = b.DocEntry "
'                    sQry = sQry & "Where a.U_BigCode = '1' And a.U_MidCode = '" & Trim(oMat01.Columns("BadCode1").Cells(pval.Row).Specific.VALUE) & "' Order by Convert(int, b.U_SmalCode) "
'                    oRecordSet01.DoQuery sQry
'                    'oForm01.Items("ItmMSort").Specific.ValidValues.Add "ALL", "ALL"
'                    Do Until oRecordSet01.EOF
'                        oMat01.Columns("BadCode2").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
'                        oRecordSet01.MoveNext
'                    Loop
'                    oForm01.Freeze False
                End If
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
'et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "CardCode" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "CntcCode" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "GADocLin" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                            If oErrNum = 1 Then
                                oErrNum = 0
                                BubbleEvent = False
                            End If
                        ElseIf pval.ColUID = "ItemCode" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        ElseIf pval.ColUID = "WhsCode" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        ElseIf pval.ColUID = "InQty" Or pval.ColUID = "Qty" Or pval.ColUID = "RealWt" Or pval.ColUID = "Weight" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        End If
                    End If
                End If
            Case et_MATRIX_LOAD: '//11
                Add_MatrixRow oMat01.RowCount, False
                FormItemEnabled
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
            Case "1284": '취소
                If Add_oPurchaseReturns(1) = False Then
                    BubbleEvent = False
                    Exit Sub
                Else
                    oLast_Mode = 101
                End If
            Case "1286": '닫기
                If Close_oPurchaseDeliveryNotes(1) = False Then
                    BubbleEvent = False
                    Exit Sub
                Else
                    oLast_Mode = 102
                End If
            Case "1293": '행삭제
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
                FormItemEnabled
                oForm01.Items("DocNum").Click ct_Regular
            Case "1286": '닫기
            Case "1293": '행삭제
                If oMat01.RowCount <> oMat01.VisualRowCount Then
                    For i = 0 To oMat01.VisualRowCount - 1
                        oMat01.Columns("LineNum").Cells(i + 1).Specific.VALUE = i + 1
                    Next i
                    
                    oMat01.FlushToDataSource
                    oDS_PS_MM070L.RemoveRecord oDS_PS_MM070L.Size - 1       '// Mat01에 마지막라인(빈라인) 삭제
                    oMat01.Clear
                    oMat01.LoadFromDataSource
                    
                    If oMat01.Columns("GADocLin").Cells(oMat01.RowCount).Specific.VALUE <> "" Then
                        Call Add_MatrixRow(oMat01.RowCount, False)
                    End If
                End If
            
            Case "1281": '찾기
                FormItemEnabled
                oForm01.Items("DocNum").Click ct_Regular
                
            Case "1282": '추가
                oForm01.Freeze True
                Call FormItemEnabled
                Call Initialization
                Call FormClear
                oDS_PS_MM070H.setValue "U_DocDate", 0, Format(Now, "YYYYMMDD")
                Call Add_MatrixRow(0, True)
'                oForm01.Items("BPLId").Click ct_Collapsed
                oForm01.Freeze False
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                FormItemEnabled
                If oMat01.VisualRowCount > 0 Then
                    If oMat01.Columns("GADocLin").Cells(oMat01.VisualRowCount).Specific.VALUE <> "" Then
                        If oDS_PS_MM070H.GetValue("Status", 0) = "O" Then
                            Add_MatrixRow oMat01.RowCount, False
                        End If
                    End If
                End If
  
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
                '검수입고 문서만 등록 시 주석처리_S
                If Add_oPurchaseDeliveryNotes(2) = False Then
                    BubbleEvent = False
                    Exit Sub
                Else
                    Call Delete_EmptyRow '검수입고 문서만 등록 시 이 행은 주석 제외
                End If
                '검수입고 문서만 등록 시 주석처리_E

            Case et_FORM_DATA_UPDATE: '//35
                If oLast_Mode = 101 Then
                    If Add_oPurchaseReturns(2) = False Then
                        oLast_Mode = 0
                        BubbleEvent = False
                        Exit Sub
                    Else
                        oLast_Mode = 0
                    End If
                ElseIf oLast_Mode = 102 Then '//닫기
                    If Close_oPurchaseDeliveryNotes(2) = False Then
                        oLast_Mode = 0
                        BubbleEvent = False
                        Exit Sub
                    Else
                        oLast_Mode = 0
                    End If
                End If
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
    Set oDS_PS_MM070H = oForm01.DataSources.DBDataSources("@PS_MM070H")
    Set oDS_PS_MM070L = oForm01.DataSources.DBDataSources("@PS_MM070L")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    
    oDS_PS_MM070H.setValue "U_DocDate", 0, Format(Now, "yyyymmdd")
    
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
    
    '//품의형태
'    Set oCombo = oForm01.Items("POType").Specific
'    sQry = "SELECT Code, Name From [@PSH_RETYPE]"
'    oRecordSet01.DoQuery sQry
'    Do Until oRecordSet01.EOF
'        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).Value), Trim(oRecordSet01.Fields(1).Value)
'        oRecordSet01.MoveNext
'    Loop
    
    '// 품목구분
    Set oCombo = oForm01.Items("Purchase").Specific
    sQry = "           SELECT       '%' AS [Code],"
    sQry = sQry & "                 '선택' AS [Name]"
    sQry = sQry & " UNION ALL"
    sQry = sQry & " SELECT       Code, "
    sQry = sQry & "                 Name "
    sQry = sQry & " FROM        [@PSH_ORDTYP] "
    sQry = sQry & " ORDER BY   Code"
    
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oMat01.Columns("ItemGpCd").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    Call oCombo.Select("%", psk_ByValue)
    
    '//품질검수여부
    oForm01.Items("QEYesNo").Specific.ValidValues.Add "Y", "Yes"
    oForm01.Items("QEYesNo").Specific.ValidValues.Add "N", "No"
        
    '// 불량코드1
    sQry = "SELECT b.U_MidCode, b.U_MidName From [@PS_PP002H] a Inner Join [@PS_PP002L] b On a.DocEntry = b.DocEntry Where a.U_BigCode = '1' Order by b.U_MidCode"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oMat01.Columns("BadCode1").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    
    
'    sQry = "SELECT b.U_SmalCode, b.U_SmalName From [@PS_PP003H] a Inner Join [@PS_PP003L] b On a.DocEntry = b.DocEntry Where a.U_BigCode = '1' Order by b.U_SmalCode "
'    oRecordSet01.DoQuery sQry
'    Do Until oRecordSet01.EOF
'        oMat01.Columns("BadCode2").ValidValues.Add Trim(oRecordSet01.Fields(0).Value), Trim(oRecordSet01.Fields(1).Value)
'        oRecordSet01.MoveNext
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
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    '//아이디별 사업장 세팅
    Set oCombo = oForm01.Items("BPLId").Specific
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
    '//아이디별 사번 세팅
    oForm01.Items("CntcCode").Specific.VALUE = MDC_PS_Common.User_MSTCOD
    
    '//품질검수여부
    sQry = "Select dept From [OHEM] Where U_MSTCOD = '" & Trim(oForm01.Items("CntcCode").Specific.VALUE) & "' "
    oRecordSet01.DoQuery sQry
    If Trim(oRecordSet01.Fields(0).VALUE) = "4" Or Trim(oRecordSet01.Fields(0).VALUE) = "12" Then
        oForm01.Items("QEYesNo").Specific.Select "Y"
    Else
        oForm01.Items("QEYesNo").Specific.Select "N"
    End If
    
    '//아이디별 부서 세팅
'    Set oCombo = oForm01.Items("DeptCode").Specific
'    oCombo.Select MDC_PS_Common.User_DeptCode, psk_ByValue
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
        oForm01.Items("DocNum").Enabled = False
        oForm01.Items("CardCode").Enabled = True
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("QEYesNo").Enabled = True
        oForm01.Items("Purchase").Enabled = True
        oForm01.Items("DocDate").Enabled = True
        oForm01.Items("DueDate").Enabled = True
        
        oMat01.Columns("GADocLin").Editable = True
        oMat01.Columns("ItemCode").Editable = True
        oMat01.Columns("ItemGpCd").Editable = True
        oMat01.Columns("BatchYN").Editable = True
        oMat01.Columns("BatchNum").Editable = True
        oMat01.Columns("InQty").Editable = True
        oMat01.Columns("Qty").Editable = True
        oMat01.Columns("BadQty").Editable = False
        oMat01.Columns("Weight").Editable = True
        oMat01.Columns("RealWt").Editable = True
        oMat01.Columns("BadCode1").Editable = True
        oMat01.Columns("BadCode2").Editable = True
        oMat01.Columns("UnWeight").Editable = True
        oMat01.Columns("Price").Editable = True
        oMat01.Columns("LinTotal").Editable = True
        oMat01.Columns("WhsCode").Editable = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        oForm01.Items("DocNum").Enabled = True
        oForm01.Items("CardCode").Enabled = True
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("QEYesNo").Enabled = True
        oForm01.Items("Purchase").Enabled = True
        oForm01.Items("DocDate").Enabled = True
        oForm01.Items("DueDate").Enabled = True
        
        oMat01.Columns("GADocLin").Editable = False
        oMat01.Columns("ItemCode").Editable = False
        oMat01.Columns("ItemGpCd").Editable = False
        oMat01.Columns("BatchYN").Editable = False
        oMat01.Columns("BatchNum").Editable = False
        oMat01.Columns("InQty").Editable = False
        oMat01.Columns("Qty").Editable = False
        oMat01.Columns("BadQty").Editable = False
        oMat01.Columns("Weight").Editable = False
        oMat01.Columns("RealWt").Editable = False
        oMat01.Columns("BadCode1").Editable = False
        oMat01.Columns("BadCode2").Editable = False
        oMat01.Columns("UnWeight").Editable = False
        oMat01.Columns("Price").Editable = False
        oMat01.Columns("LinTotal").Editable = False
        oMat01.Columns("WhsCode").Editable = False
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        oForm01.Items("DocNum").Enabled = False
        oForm01.Items("CardCode").Enabled = False
        oForm01.Items("BPLId").Enabled = False
        oForm01.Items("QEYesNo").Enabled = False
        oForm01.Items("Purchase").Enabled = False
        oForm01.Items("DocDate").Enabled = False
        oForm01.Items("DueDate").Enabled = False
        
        oMat01.Columns("GADocLin").Editable = False
        oMat01.Columns("ItemCode").Editable = False
        oMat01.Columns("ItemGpCd").Editable = False
        oMat01.Columns("BatchYN").Editable = False
        oMat01.Columns("BatchNum").Editable = False
        oMat01.Columns("InQty").Editable = False
        oMat01.Columns("Qty").Editable = False
        oMat01.Columns("BadQty").Editable = False
        oMat01.Columns("Weight").Editable = False
        oMat01.Columns("RealWt").Editable = False
        oMat01.Columns("BadCode1").Editable = False
        oMat01.Columns("BadCode2").Editable = False
        oMat01.Columns("UnWeight").Editable = False
        oMat01.Columns("Price").Editable = False
        oMat01.Columns("LinTotal").Editable = False
        oMat01.Columns("WhsCode").Editable = False
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub FormClear()
On Error GoTo FormClear_Error
    Dim DocNum As String
    DocNum = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM070'", "")
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
        oDS_PS_MM070L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_MM070L.offset = oRow
    oDS_PS_MM070L.setValue "U_LineNum", oRow, oRow + 1
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
    Dim ItemCode$, Qty As Long, InQty As Long, BadQty As Long, GAQty As Long, GARealWt As Currency, GAUnWt As Currency
    Dim RealWt As Currency, Price As Currency, Weight As Currency
    Dim Calculate_Weight As Currency
    Dim Calculate_Qty As Currency
    Dim BadWt As Long '불량중량(2011.11.03 송명규 추가)
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sRow = oRow
    
    Select Case oUID
    
        Case "CardCode"
        
            sQry = "Select CardName From OCRD Where CardCode = '" & Trim(oDS_PS_MM070H.GetValue("U_CardCode", 0)) & "'"
            oRecordSet01.DoQuery sQry
            
            Call oForm01.Freeze(True)
            
            Call oDS_PS_MM070H.setValue("U_CardName", 0, Trim(oRecordSet01.Fields(0).VALUE))
            
            Call oForm01.Freeze(False)
            
        Case "CntcCode"
        
            sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" & Trim(oDS_PS_MM070H.GetValue("U_CntcCode", 0)) & "'"
            oRecordSet01.DoQuery sQry
            
            Call oForm01.Freeze(True)
            
            Call oDS_PS_MM070H.setValue("U_CntcName", 0, Trim(oRecordSet01.Fields(0).VALUE))
            
            Call oForm01.Freeze(False)
            
        Case "Mat01"
        
            If oCol = "GADocLin" Then
            
                Dim GADocLin$
                
                
                If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 0) And Trim(oMat01.Columns("GADocLin").Cells(oRow).Specific.VALUE) <> "" Then
                    oMat01.FlushToDataSource
                    Call Add_MatrixRow(oMat01.RowCount, False)
                    oMat01.Columns("GADocLin").Cells(oRow).Click ct_Regular
                End If
                
                
                GADocLin = Trim(oDS_PS_MM070L.GetValue("U_GADocLin", oRow - 1))
                sQry = "Select a.U_CardCode, b.U_ItemCode, b.U_ItemName, b.U_FrgnName, b.U_ItemGpCd, b.U_QTy, b.U_Weight, b.U_RealWt, b.U_UnWeight, "
                sQry = sQry & "b.U_Price, b.U_LinTotal, b.U_WhsCode, b.U_WhsName, b.U_BDocNum, b.U_BLineNum, a.U_DocDate, b.U_OutSize, b.U_OutUnit, b.U_Auto "
                sQry = sQry & "From [@PS_MM050H] a Inner Join [@PS_MM050L] b On a.DocEntry = b.DocEntry "
                sQry = sQry & "Where a.DocNum = Left('" & GADocLin & "', CharIndex('-', '" & GADocLin & "') - 1) And b.LineId = Right('" & GADocLin & "', Len('" & GADocLin & "') - CharIndex('-', '" & GADocLin & "'))" 'U_LineNum을 LineId로 수정(2012.07.24 송명규)
                oRecordSet01.DoQuery sQry
                
                If Trim(oDS_PS_MM070H.GetValue("U_CardCode", 0)) = "" Then
                    oForm01.Items("CardCode").Specific.VALUE = Trim(oRecordSet01.Fields("U_CardCode").VALUE)
'                    oDS_PS_MM070H.setValue "U_CardCode", oRow - 1, Trim(oRecordSet01.Fields("U_CardCode").Value)
                Else
                    If Trim(oDS_PS_MM070H.GetValue("U_CardCode", 0)) <> Trim(oRecordSet01.Fields("U_CardCode").VALUE) Then
                        oErrNum = 1
                        GoTo FlushToItemValue_Error
                    End If
                End If
                
                oForm01.Freeze True
                oMat01.FlushToDataSource
                
                If oRecordSet01.RecordCount = 0 Then
                
'                    oMat01.Columns("ItemCode").Cells(oRow).Specific.VALUE = ""
'                    oMat01.Columns("ItemName").Cells(oRow).Specific.VALUE = ""
'                    oMat01.Columns("OutSize").Cells(oRow).Specific.VALUE = ""
'                    oMat01.Columns("OutUnit").Cells(oRow).Specific.VALUE = ""
''                    oMat01.Columns("FrgnName").Cells(oRow).Specific.Value = ""
''                    oMat01.Columns("ItemGpCd").Cells(oRow).Specific.Select ""
'                    oMat01.Columns("InQty").Cells(oRow).Specific.VALUE = 0
'                    oMat01.Columns("Qty").Cells(oRow).Specific.VALUE = 0
'                    oMat01.Columns("Weight").Cells(oRow).Specific.VALUE = 0
'                    oMat01.Columns("RealWt").Cells(oRow).Specific.VALUE = 0
'                    oMat01.Columns("Auto").Cells(oRow).Specific.Select "N"
'
'                    oMat01.Columns("OPORNum").Cells(oRow).Specific.VALUE = ""
'                    oMat01.Columns("POR1Num").Cells(oRow).Specific.VALUE = ""

                    '매트릭스에 데이터를 직접 바인딩하면 이벤트가 실행되기 때문에 DataSource로 바인딩하는 방식으로 수정(2011.11.22 송명규)
                    Call oDS_PS_MM070L.setValue("U_ItemCode", oRow - 1, "")
                    Call oDS_PS_MM070L.setValue("U_ItemName", oRow - 1, "")
                    Call oDS_PS_MM070L.setValue("U_OutSize", oRow - 1, "")
                    Call oDS_PS_MM070L.setValue("U_OutUnit", oRow - 1, "")
                    Call oDS_PS_MM070L.setValue("U_InQty", oRow - 1, 0)
                    Call oDS_PS_MM070L.setValue("U_Qty", oRow - 1, 0)
                    Call oDS_PS_MM070L.setValue("U_Weight", oRow - 1, 0)
                    Call oDS_PS_MM070L.setValue("U_RealWt", oRow - 1, 0)
                    Call oDS_PS_MM070L.setValue("U_Auto", oRow - 1, "N")
                    
                    Call oDS_PS_MM070L.setValue("U_OPORNum", oRow - 1, "")
                    Call oDS_PS_MM070L.setValue("U_POR1Num", oRow - 1, "")
                    
                    MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
                    
                    Set oRecordSet01 = Nothing
                    oMat01.LoadFromDataSource
                    oForm01.Freeze False
                    Exit Sub
                    
                End If
                
                oForm01.Items("DueDate").Specific.VALUE = Format(Trim(oRecordSet01.Fields("U_DocDate").VALUE), "YYYYMMDD") '납품일(헤더)
                
'                oMat01.Columns("ItemCode").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_ItemCode").VALUE) '품목코드
'                oMat01.Columns("ItemName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_ItemName").VALUE) '품목이름
'                oMat01.Columns("OutSize").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_OutSize").VALUE) '규격
'                oMat01.Columns("OutUnit").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_OutUnit").VALUE) '단위
''                oMat01.Columns("FrgnName").Cells(oRow).Specific.Value = Trim(oRecordSet01.Fields("U_FrgnName").Value)
'                oMat01.Columns("ItemGpCd").Cells(oRow).Specific.Select Trim(oRecordSet01.Fields("U_ItemGpCd").VALUE) '구매구분
'                oMat01.Columns("GAQty").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_Qty").VALUE) '가입고수량
'                oMat01.Columns("GARealWt").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_RealWt").VALUE) '가입고실중량
'                oMat01.Columns("InQty").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_Qty").VALUE) '입고수량
'                oMat01.Columns("Qty").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_Qty").VALUE) '검수수량
'                oMat01.Columns("Weight").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_Weight").VALUE) '이론중량
'                oMat01.Columns("RealWt").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_RealWt").VALUE) '실중량
'                oMat01.Columns("UnWeight").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_UnWeight").VALUE) '단중
'                oMat01.Columns("Price").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_Price").VALUE) '단가
'                oMat01.Columns("LinTotal").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_LinTotal").VALUE) '금액
'                oMat01.Columns("WhsCode").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_WhsCode").VALUE) '창고코드
'                oMat01.Columns("Auto").Cells(oRow).Specific.Select Trim(oRecordSet01.Fields("U_Auto").VALUE) '자동불출여부
'
'                oMat01.Columns("OPORNum").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_BDocNum").VALUE) '구매오더번호
'                oMat01.Columns("POR1Num").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_BLineNum").VALUE) '오더행
                
                '매트릭스에 데이터를 직접 바인딩하면 이벤트가 실행되기 때문에 DataSource로 바인딩하는 방식으로 수정(2011.11.22 송명규)
                Call oDS_PS_MM070L.setValue("U_ItemCode", oRow - 1, Trim(oRecordSet01.Fields("U_ItemCode").VALUE)) '품목코드
                Call oDS_PS_MM070L.setValue("U_ItemName", oRow - 1, Trim(oRecordSet01.Fields("U_ItemName").VALUE)) '품목이름
                Call oDS_PS_MM070L.setValue("U_OutSize", oRow - 1, Trim(oRecordSet01.Fields("U_OutSize").VALUE)) '규격
                Call oDS_PS_MM070L.setValue("U_OutUnit", oRow - 1, Trim(oRecordSet01.Fields("U_OutUnit").VALUE)) '단위
                Call oDS_PS_MM070L.setValue("U_ItemGpCd", oRow - 1, Trim(oRecordSet01.Fields("U_ItemGpCd").VALUE)) '구매구분
                Call oDS_PS_MM070L.setValue("U_GAQty", oRow - 1, Trim(oRecordSet01.Fields("U_Qty").VALUE)) '가입고수량
                Call oDS_PS_MM070L.setValue("U_GARealWt", oRow - 1, Trim(oRecordSet01.Fields("U_RealWt").VALUE)) '가입고실중량
                Call oDS_PS_MM070L.setValue("U_InQty", oRow - 1, Trim(oRecordSet01.Fields("U_Qty").VALUE)) '입고수량
                Call oDS_PS_MM070L.setValue("U_Qty", oRow - 1, Trim(oRecordSet01.Fields("U_Qty").VALUE)) '검수수량
                Call oDS_PS_MM070L.setValue("U_Weight", oRow - 1, Trim(oRecordSet01.Fields("U_Weight").VALUE)) '이론중량
                Call oDS_PS_MM070L.setValue("U_RealWt", oRow - 1, Trim(oRecordSet01.Fields("U_RealWt").VALUE)) '실중량
                Call oDS_PS_MM070L.setValue("U_BadQty", oRow - 1, 0) '불량수량
                Call oDS_PS_MM070L.setValue("U_UnWeight", oRow - 1, Trim(oRecordSet01.Fields("U_UnWeight").VALUE)) '단중
                Call oDS_PS_MM070L.setValue("U_Price", oRow - 1, Trim(oRecordSet01.Fields("U_Price").VALUE)) '단가
                Call oDS_PS_MM070L.setValue("U_LinTotal", oRow - 1, Trim(oRecordSet01.Fields("U_LinTotal").VALUE)) '금액
                Call oDS_PS_MM070L.setValue("U_WhsCode", oRow - 1, Trim(oRecordSet01.Fields("U_WhsCode").VALUE)) '창고코드
                Call oDS_PS_MM070L.setValue("U_WhsName", oRow - 1, Trim(oRecordSet01.Fields("U_WhsName").VALUE)) '창고명(2011.11.22 송명규 추가)
                Call oDS_PS_MM070L.setValue("U_Auto", oRow - 1, Trim(oRecordSet01.Fields("U_Auto").VALUE)) '자동불출여부
                Call oDS_PS_MM070L.setValue("U_OPORNum", oRow - 1, Trim(oRecordSet01.Fields("U_BDocNum").VALUE)) '구매오더번호
                Call oDS_PS_MM070L.setValue("U_POR1Num", oRow - 1, Trim(oRecordSet01.Fields("U_BLineNum").VALUE)) ''오더행
                
                oMat01.LoadFromDataSource
                oMat01.AutoResizeColumns
                oForm01.Freeze False
                
            ElseIf oCol = "ItemCode" Then
            
'                oForm01.Freeze True
                If oForm01.Mode = fm_ADD_MODE Then
                    If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 0) And Trim(oMat01.Columns("ItemCode").Cells(oRow).Specific.VALUE) <> "" Then
                        oMat01.FlushToDataSource
                        Add_MatrixRow oMat01.RowCount, False
                    End If
                End If
                
                sQry = "Select ItemName, FrgnName From OITM Where ItemCode = '" & Trim(oMat01.Columns("ItemCode").Cells(oRow).Specific.VALUE) & "'"
                oRecordSet01.DoQuery sQry
                
'                oMat01.Columns("ItemName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
'                oMat01.Columns("FrgnName").Cells(oRow).Specific.Value = Trim(oRecordSet01.Fields(1).Value)
                
                '매트릭스에 데이터를 직접 바인딩하면 이벤트가 실행되기 때문에 DataSource로 바인딩하는 방식으로 수정(2011.11.22 송명규)
                Call oForm01.Freeze(True)
                oMat01.FlushToDataSource
                
                Call oDS_PS_MM070L.setValue("U_ItemName", oRow - 1, Trim(oRecordSet01.Fields(0).VALUE))
                
                oMat01.LoadFromDataSource
                oMat01.Columns("ItemCode").Cells(oRow).Click ct_Regular
                Call oForm01.Freeze(False)
                
            ElseIf oCol = "WhsCode" Then
            
                sQry = "Select WhsName From [OWHS] Where WhsCode = '" & Trim(oMat01.Columns("WhsCode").Cells(oRow).Specific.VALUE) & "'"
                oRecordSet01.DoQuery sQry
            
                'oMat01.Columns("WhsName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                
                '매트릭스에 데이터를 직접 바인딩하면 이벤트가 실행되기 때문에 DataSource로 바인딩하는 방식으로 수정(2011.11.22 송명규)
                Call oForm01.Freeze(True)
                oMat01.FlushToDataSource
                
                Call oDS_PS_MM070L.setValue("U_WhsName", oRow - 1, Trim(oRecordSet01.Fields(0).VALUE))
                
                oMat01.LoadFromDataSource
                oMat01.Columns("WhsCode").Cells(oRow).Click ct_Regular
                Call oForm01.Freeze(False)
                
            ElseIf oCol = "InQty" Then
            
                oForm01.Freeze True
                oMat01.FlushToDataSource
                
                If Trim(oMat01.Columns("Qty").Cells(oRow).Specific.VALUE) <> "" Then
                    'BadQty = Trim(oMat01.Columns("InQty").Cells(oRow).Specific.VALUE) - Trim(oMat01.Columns("Qty").Cells(oRow).Specific.VALUE)
                    'oMat01.Columns("BadQty").Cells(oRow).Specific.VALUE = BadQty
                    BadQty = oDS_PS_MM070L.GetValue("U_InQty", oRow - 1) - oDS_PS_MM070L.GetValue("U_Qty", oRow - 1)
                    Call oDS_PS_MM070L.setValue("U_BadQty", oRow - 1, BadQty)
                Else
                    'oMat01.Columns("BadQty").Cells(oRow).Specific.VALUE = 0
                    Call oDS_PS_MM070L.setValue("U_BadQty", oRow - 1, 0)
                End If
                
                oMat01.LoadFromDataSource
                oMat01.Columns("InQty").Cells(oRow).Click
                oForm01.Freeze False
                
            ElseIf oCol = "Qty" Then
            
                oForm01.Freeze True
                oMat01.FlushToDataSource
                
                '//중량계산
                '서브작번이 추가된 작번을 넘기면 안되므로 Main 작번만 넘김(2011.10.25 송명규 수정)
                ItemCode = Left(Trim(oDS_PS_MM070L.GetValue("U_ItemCode", oRow - 1)), 11) 'Left(Trim(oMat01.Columns("ItemCode").Cells(oRow).Specific.VALUE), 11)
                Qty = oDS_PS_MM070L.GetValue("U_Qty", oRow - 1) 'Trim(oMat01.Columns("Qty").Cells(oRow).Specific.VALUE)

                Calculate_Weight = MDC_PS_Common.Calculate_Weight(ItemCode, Qty, Trim(oForm01.Items("BPLId").Specific.VALUE))
                
                'oMat01.Columns("Weight").Cells(oRow).Specific.VALUE = Calculate_Weight 'oDS_PS_MM070L.setValue "U_Weight", oRow - 1, Calculate_Weight
                Call oDS_PS_MM070L.setValue("U_Weight", oRow - 1, Calculate_Weight)
                
                '//불량수량계산
                If oDS_PS_MM070L.GetValue("U_InQty", oRow - 1) <> "" Then 'Trim(oMat01.Columns("InQty").Cells(oRow).Specific.VALUE) <> "" Then
                    BadQty = oDS_PS_MM070L.GetValue("U_InQty", oRow - 1) - oDS_PS_MM070L.GetValue("U_Qty", oRow - 1) 'Trim(oMat01.Columns("InQty").Cells(oRow).Specific.VALUE) - Trim(oMat01.Columns("Qty").Cells(oRow).Specific.VALUE)
                    Call oDS_PS_MM070L.setValue("U_BadQty", oRow - 1, BadQty) 'oMat01.Columns("BadQty").Cells(oRow).Specific.VALUE = BadQty
                    '불량중량 = 가입고중량 - 이론중량 (2011.11.03 송명규 추가)
                    BadWt = oDS_PS_MM070L.GetValue("U_GARealWt", oRow - 1) - oDS_PS_MM070L.GetValue("U_Weight", oRow - 1) 'Trim(oMat01.Columns("GARealWt").Cells(oRow).Specific.VALUE) - Trim(oMat01.Columns("Weight").Cells(oRow).Specific.VALUE)
                    Call oDS_PS_MM070L.setValue("U_BadWt", oRow - 1, BadWt) 'oMat01.Columns("BadWt").Cells(oRow).Specific.VALUE = BadWt
                Else
                    Call oDS_PS_MM070L.setValue("U_BadQty", oRow - 1, 0)
                    Call oDS_PS_MM070L.setValue("U_BadWt", oRow - 1, 0)
'                    oMat01.Columns("BadQty").Cells(oRow).Specific.VALUE = 0
'                    oMat01.Columns("BadWt").Cells(oRow).Specific.VALUE = 0
                End If
                
                '//실중량계산
                GAQty = oDS_PS_MM070L.GetValue("U_GAQty", oRow - 1)
                GARealWt = oDS_PS_MM070L.GetValue("U_GARealWt", oRow - 1)
                
'                GAQty = Trim(oMat01.Columns("GAQty").Cells(oRow).Specific.VALUE)
'                GARealWt = Trim(oMat01.Columns("GARealWt").Cells(oRow).Specific.VALUE)
                
                'If Trim(oMat01.Columns("GAQty").Cells(oRow).Specific.VALUE) <> "" And GAQty <> 0 Then
                If oDS_PS_MM070L.GetValue("U_GAQty", oRow - 1) <> "" And GAQty <> 0 Then
                    GAUnWt = Round(GARealWt / GAQty, 2)
                Else
                    GAUnWt = 0
                End If
                RealWt = GAUnWt * oDS_PS_MM070L.GetValue("U_Qty", oRow - 1) 'Trim(oMat01.Columns("Qty").Cells(oRow).Specific.VALUE)
                
                sQry = "Select ItmsGrpCod From OITM Where ItemCode = '" & ItemCode & "'"
                oRecordSet01.DoQuery sQry
                If Trim(oRecordSet01.Fields(0).VALUE) = "105" Then
                
                    Call oDS_PS_MM070L.setValue("U_RealWt", oRow - 1, Calculate_Weight)
                    'oMat01.Columns("RealWt").Cells(oRow).Specific.VALUE = Calculate_Weight
                    
                Else
                    Call oDS_PS_MM070L.setValue("U_RealWt", oRow - 1, RealWt)
                    'oMat01.Columns("RealWt").Cells(oRow).Specific.VALUE = RealWt
                End If
                
                '금액 계산(2012.05.03 송명규 추가)
                '단가X이론중량
                Price = oDS_PS_MM070L.GetValue("U_Price", oRow - 1)
                Weight = oDS_PS_MM070L.GetValue("U_Weight", oRow - 1)
                Call oDS_PS_MM070L.setValue("U_LinTotal", oRow - 1, Weight * Price)
                
                oMat01.LoadFromDataSource
                oMat01.Columns("BadQty").Cells(oRow).Click
                oForm01.Freeze False

            ElseIf oCol = "Weight" Then
            
                If Trim(oMat01.Columns("Weight").Cells(oRow).Specific.VALUE) <> "" Then 'And Trim(oMat01.Columns("Weight").Cells(oRow).Specific.Value) <> 0 Then
                
'                    Price = Trim(oMat01.Columns("Price").Cells(oRow).Specific.VALUE)
'                    Weight = Trim(oMat01.Columns("Weight").Cells(oRow).Specific.VALUE)
'                    oMat01.Columns("LinTotal").Cells(oRow).Specific.VALUE = Weight * Price
                    Call oForm01.Freeze(True)
                    oMat01.FlushToDataSource
                    
                    Price = oDS_PS_MM070L.GetValue("U_Price", oRow - 1)
                    Weight = oDS_PS_MM070L.GetValue("U_Weight", oRow - 1)
                    Call oDS_PS_MM070L.setValue("U_LinTotal", oRow - 1, Weight * Price)
                    
                    oMat01.LoadFromDataSource
                    oMat01.Columns("Weight").Cells(oRow).Click
                    Call oForm01.Freeze(False)
                End If
                
            End If
    End Select

    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FlushToItemValue_Error:
    Set oRecordSet01 = Nothing
    oForm01.Freeze False
    If oErrNum = 1 Then
        MDC_Com.MDC_GF_Message "다른 거래처의 데이터를 한문서에서 처리할 수 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "E"
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
        Case oDS_PS_MM070H.GetValue("U_BPLId", 0) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_MM070H.GetValue("U_CardCode", 0) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_MM070H.GetValue("U_CntcCode", 0) = ""
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
'        Case oDS_PS_MM070H.GetValue("U_POType", 0) = ""
'            ErrNum = 4
'            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_MM070H.GetValue("U_Purchase", 0) = "%"
            ErrNum = 5
            GoTo HeaderSpaceLineDel_Error
    End Select

'    '마감상태 체크_S(2017.11.23 송명규 추가)
'    If MDC_PS_Common.Check_Finish_Status(oForm01.Items("BPLId").Specific.VALUE, oForm01.Items("DocDate").Specific.VALUE) = False Then
'        ErrNum = 6
'        GoTo HeaderSpaceLineDel_Error
'    End If
'    '마감상태 체크_E(2017.11.23 송명규 추가)

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "사업장은 필수입력 사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "거래처코드는 필수입력 사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "담당자코드는 필수입력 사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "품의형태는 필수입력 사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "구매구분은 필수입력 사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "마감상태가 잠금입니다. 해당 일자로 등록할 수 없습니다. 전기일자를 확인하고, 회계부서로 문의하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function

Private Function MatrixSpaceLineDel() As Boolean
On Error GoTo MatrixSpaceLineDel_Error

    Dim i           As Long
    Dim ErrNum      As Integer
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim sQry        As String

    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim Entry()     As String '가입고문서번호를 저장할 배열변수
    Dim BaseEntry   As String
    Dim BaseLine    As String
    
    ErrNum = 0
    
    oMat01.FlushToDataSource

    '// 라인
    If oMat01.VisualRowCount = 0 Then
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    End If
    
    For i = 0 To oMat01.VisualRowCount - 2
        Select Case True
            Case oDS_PS_MM070L.GetValue("U_GADocLin", i) = ""
                ErrNum = 2
                GoTo MatrixSpaceLineDel_Error
            Case oDS_PS_MM070L.GetValue("U_ItemCode", i) = ""
                ErrNum = 3
                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_MM070L.GetValue("U_InQty", i) = "" Or oDS_PS_MM070L.GetValue("U_InQty", i) = 0
'                ErrNum = 4
'                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_MM070L.GetValue("U_Qty", i) = "" Or oDS_PS_MM070L.GetValue("U_Qty", i) = 0
'                ErrNum = 5
'                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_MM070L.GetValue("U_Weight", i) = 0
'                ErrNum = 6
'                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_MM070L.GetValue("U_RealWt", i) = 0
'                ErrNum = 7
'                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_MM070L.GetValue("U_Price", i) = 0
'                ErrNum = 9
'                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_MM070L.GetValue("U_LinTotal", i) = 0
'                ErrNum = 10
'                GoTo MatrixSpaceLineDel_Error
            Case Trim(oDS_PS_MM070H.GetValue("U_Purchase", 0)) <> "30" And Trim(oDS_PS_MM070H.GetValue("U_Purchase", 0)) <> "40" And Trim(oDS_PS_MM070H.GetValue("U_Purchase", 0)) <> "60" And Trim(oDS_PS_MM070L.GetValue("U_WhsCode", i)) = ""
                ErrNum = 11
                GoTo MatrixSpaceLineDel_Error
        End Select
        
        If oDS_PS_MM070L.GetValue("U_InQty", i) = oDS_PS_MM070L.GetValue("U_BadQty", i) Then
        Else
        
            If oDS_PS_MM070L.GetValue("U_Weight", i) = 0 Then
                ErrNum = 6
                GoTo MatrixSpaceLineDel_Error
            End If
            If oDS_PS_MM070L.GetValue("U_RealWt", i) = 0 Then
                ErrNum = 7
                GoTo MatrixSpaceLineDel_Error
            End If
            If oDS_PS_MM070L.GetValue("U_Price", i) = 0 Then
                ErrNum = 9
                GoTo MatrixSpaceLineDel_Error
            End If
            If oDS_PS_MM070L.GetValue("U_LinTotal", i) = 0 Then
                ErrNum = 10
                GoTo MatrixSpaceLineDel_Error
            End If
        End If
        
        If oDS_PS_MM070L.GetValue("U_BadQty", i) > 0 And oDS_PS_MM070L.GetValue("U_BadQty", i) <> "" Then
            If oDS_PS_MM070L.GetValue("U_BadCode1", i) = "" Or oDS_PS_MM070L.GetValue("U_BadCode2", i) = "" Then
                ErrNum = 8
                GoTo MatrixSpaceLineDel_Error
            End If
        End If
        
        Entry = Split(oMat01.Columns("GADocLin").Cells(i + 1).Specific.VALUE, "-")
        BaseEntry = Entry(0)
        BaseLine = Entry(1)
        
        If PS_MM070_CheckDate(BaseEntry) = False Then '가입고와 일자 체크
            
            ErrNum = 12
            GoTo MatrixSpaceLineDel_Error
        
        End If
        
    Next
    
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
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 가입고문서-행이 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 품목코드가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 입고수량이 없거나 0입니다. 확인하세요.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 검수수량이 없거나 0입니다. 확인하세요.", "E"
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 이론중량이 0입니다. 확인하세요.", "E"
    ElseIf ErrNum = 7 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 실중량이 0입니다. 확인하세요.", "E"
    ElseIf ErrNum = 8 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인에 불량수량이 있으니 불량중분류와 불량상세를 입력해야 합니다. 확인하세요.", "E"
    ElseIf ErrNum = 9 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 단가는 0보다 커야 합니다. 확인하세요.", "E"
    ElseIf ErrNum = 10 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 금액은 0보다 커야 합니다. 확인하세요.", "E"
    ElseIf ErrNum = 11 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 창고코드가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 12 Then
        Call Sbo_Application.MessageBox(i + 1 & "행 [" & oMat01.Columns("ItemCode").Cells(i + 1).Specific.VALUE & "]의 검수입고일은 가입고일과 같거나 늦어야합니다. 확인하십시오." & Chr(13) & "해당 검수입고는 전체가 등록되지 않습니다.", 1)
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
        If Trim(oDS_PS_MM070L.GetValue("U_ItemCode", i)) = "" Then
            oDS_PS_MM070L.RemoveRecord i   '// Mat01에 마지막라인(빈라인) 삭제
        End If
    Next i
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Delete_EmptyRow_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Function Add_oPurchaseDeliveryNotes(ChkType As Integer) As Boolean
On Error GoTo Add_oPurchaseDeliveryNotes_Error
    Dim DI_oPurchaseDeliveryNotes As SAPbobsCOM.Documents         '//'입고PO 문서객체
    Dim DI_oInventoryGenExit As SAPbobsCOM.Documents              '//출고 문서객체
        
    Dim i&, j&
    Dim ErrNum&, ErrCode As Long, ErrMsg$, RetVal&
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    Dim sQry            As String
    
    Dim Line_Count&
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocDate$, DueDate$, DocEntry$
    Dim ECVatGroup As String
    
    '//////////////////////////////////////////////////////////////////////
    '//애드온이 비정상적으로 DI만 실행하고 애드온이 저장이 안되었을때 수행하기 위해 2011.09.06 노근용
'    Set DI_oPurchaseDeliveryNotes = Nothing
'    Set DI_oInventoryGenExit = Nothing
'    Add_oPurchaseDeliveryNotes = True
'    Set oRecordSet01 = Nothing
'    Exit Function
    '//////////////////////////////////////////////////////////////////////
    
    oMat01.FlushToDataSource
    
    DocDate = Format(oDS_PS_MM070H.GetValue("U_DocDate", 0), "0000-00-00")
    DueDate = Format(oDS_PS_MM070H.GetValue("U_DueDate", 0), "0000-00-00")
    
    If Trim(oDS_PS_MM070H.GetValue("U_GRDocNum", 0)) = "" Then
        Sbo_Company.StartTransaction
        Set DI_oPurchaseDeliveryNotes = Nothing
        Set DI_oPurchaseDeliveryNotes = Sbo_Company.GetBusinessObject(oPurchaseDeliveryNotes) '입고PO
                
        With DI_oPurchaseDeliveryNotes
        .CardCode = Trim(oDS_PS_MM070H.GetValue("U_CardCode", 0))
        .BPL_IDAssignedToInvoice = Trim(oDS_PS_MM070H.GetValue("U_BPLId", 0))
        .DocDate = DocDate
        .DocDueDate = DueDate
        .Comments = "검수입고(" & Trim(oDS_PS_MM070H.GetValue("DocNum", 0)) & ")입고PO : " & Trim(oDS_PS_MM070H.GetValue("U_Comments", 0))
        
        sQry = "Select ECVatGroup From [OCRD] Where CardCode = '" & Trim(oDS_PS_MM070H.GetValue("U_CardCode", 0)) & "'"
        oRecordSet01.DoQuery sQry
        ECVatGroup = oRecordSet01.Fields("ECVatGroup").VALUE
        
        Line_Count = 0
        If Trim(oDS_PS_MM070H.GetValue("U_Purchase", 0)) = "30" Or Trim(oDS_PS_MM070H.GetValue("U_Purchase", 0)) = "40" Or Trim(oDS_PS_MM070H.GetValue("U_Purchase", 0)) = "60" Then '가공비, 외주제작, 고정자산 품의 일때는 Service로 생성
'        If Trim(oDS_PS_MM070H.GetValue("U_Purchase", 0)) = "30" Or "40" Or "60" Then
            .DocType = dDocument_Service
            For i = 0 To oMat01.VisualRowCount - 2
                If Trim(oDS_PS_MM070L.GetValue("U_Weight", i)) <= 0 Then
                    GoTo Sikp_Line_Service
                End If
                
                If i > 0 Then .Lines.Add
                .Lines.SetCurrentLine i
                
                .Lines.ItemDescription = Trim(oDS_PS_MM070L.GetValue("U_ItemName", i)) + "-" + Trim(oDS_PS_MM070L.GetValue("U_OutSize", i)) + "-" + Trim(oDS_PS_MM070L.GetValue("U_OutUnit", i))
                .Lines.LineTotal = Trim(oDS_PS_MM070L.GetValue("U_LinTotal", i))
                .Lines.VATGroup = ECVatGroup
                
                '품의구분에 따른 GL계정 분류(2012.07.09 송명규 수정)
                If Trim(oDS_PS_MM070H.GetValue("U_Purchase", 0)) = "60" Then '고정자산품의
                    .Lines.AccountCode = "59109020"
                ElseIf Trim(oDS_PS_MM070H.GetValue("U_Purchase", 0)) = "30" Then '가공비품의
                    .Lines.AccountCode = "55302010" '외주가공비(제) 계정
                ElseIf Trim(oDS_PS_MM070H.GetValue("U_Purchase", 0)) = "40" Then '외주제작품의
                    .Lines.AccountCode = "11506100" '원재료 계정
                End If
                '품의구분에 따른 GL계정 분류(2012.07.09 송명규 수정)
                
'                .Lines.UserFields("PQDocNum").Value = oDS_PS_MM030H.getValue("U_PQDocNum", 0)
                .Lines.UserFields("U_sItemCode").VALUE = Trim(oDS_PS_MM070L.GetValue("U_ItemCode", i))
                .Lines.UserFields("U_sItemName").VALUE = Trim(oDS_PS_MM070L.GetValue("U_ItemName", i))
                .Lines.UserFields("U_sSize").VALUE = Trim(oDS_PS_MM070L.GetValue("U_OutSize", i))
                .Lines.UserFields("U_sUnit").VALUE = Trim(oDS_PS_MM070L.GetValue("U_OutUnit", i))
                .Lines.UserFields("U_sQty").VALUE = Trim(oDS_PS_MM070L.GetValue("U_Qty", i))
                .Lines.UserFields("U_sWeight").VALUE = Trim(oDS_PS_MM070L.GetValue("U_Weight", i))
                .Lines.UserFields("U_sPrice").VALUE = Trim(oDS_PS_MM070L.GetValue("U_Price", i)) '단가 추가(2011.10.25 송명규 추가)
                
                .Lines.BaseType = "22"
                .Lines.BaseEntry = Trim(oDS_PS_MM070L.GetValue("U_OPORNum", i))
                .Lines.BaseLine = Trim(oDS_PS_MM070L.GetValue("U_POR1Num", i))
                
                sQry = "Select b.U_Payment From OPOR a inner Join POR1 b On a.DocEntry = b.DocEntry Where a.DocEntry = '" & Trim(oDS_PS_MM070L.GetValue("U_OPORNum", i)) & "' and b.LineNum = '" & Trim(oDS_PS_MM070L.GetValue("U_POR1Num", i)) & "'"
                oRecordSet01.DoQuery sQry
                
                .Lines.UserFields("U_Payment").VALUE = Trim(oRecordSet01.Fields("U_Payment").VALUE)
                
                Line_Count = Line_Count + 1
Sikp_Line_Service:
            Next i
        Else
            .DocType = dDocument_Items
            For i = 0 To oMat01.VisualRowCount - 2
                If Trim(oDS_PS_MM070L.GetValue("U_RealWt", i)) <= 0 Then
                    GoTo Sikp_Line_Item
                End If
                If i > 0 Then .Lines.Add
                .Lines.SetCurrentLine i
                
'                .Lines.UserFields("PQDocNum").Value = oDS_PS_MM070H.getValue("U_PQDocNum", 0)
                .Lines.ItemCode = Trim(oDS_PS_MM070L.GetValue("U_ItemCode", i))
                .Lines.Quantity = Trim(oDS_PS_MM070L.GetValue("U_RealWt", i))
                .Lines.UnitPrice = Trim(oDS_PS_MM070L.GetValue("U_Price", i))
                .Lines.DiscountPercent = 100 - (Trim(oDS_PS_MM070L.GetValue("U_LinTotal", i)) / (Trim(oDS_PS_MM070L.GetValue("U_RealWt", i)) * Trim(oDS_PS_MM070L.GetValue("U_Price", i))) * 100)
                .Lines.LineTotal = Trim(oDS_PS_MM070L.GetValue("U_LinTotal", i))
                .Lines.WarehouseCode = Trim(oDS_PS_MM070L.GetValue("U_WhsCode", i))
                .Lines.UserFields("U_Qty").VALUE = Trim(oDS_PS_MM070L.GetValue("U_Qty", i))
                .Lines.UserFields("U_Weight").VALUE = Trim(oDS_PS_MM070L.GetValue("U_Weight", i))
                .Lines.UserFields("U_sPrice").VALUE = Trim(oDS_PS_MM070L.GetValue("U_Price", i)) '단가 추가(2011.11.28 송명규 추가)
                
                .Lines.BaseType = "22"
                .Lines.BaseEntry = Trim(oDS_PS_MM070L.GetValue("U_OPORNum", i))
                .Lines.BaseLine = Trim(oDS_PS_MM070L.GetValue("U_POR1Num", i))
                
                If Trim(oDS_PS_MM070L.GetValue("U_BatchYN", i)) = "Y" Then
                    .Lines.BatchNumbers.BatchNumber = Trim(oDS_PS_MM070L.GetValue("U_BatchNum", i))
                    .Lines.BatchNumbers.Quantity = Trim(oDS_PS_MM070L.GetValue("U_Weight", i))
                End If
                
                sQry = "Select b.U_Payment From OPOR a inner Join POR1 b On a.DocEntry = b.DocEntry Where a.DocEntry = '" & Trim(oDS_PS_MM070L.GetValue("U_OPORNum", i)) & "' and b.LineNum = '" & Trim(oDS_PS_MM070L.GetValue("U_POR1Num", i)) & "'"
                oRecordSet01.DoQuery sQry
                
                .Lines.UserFields("U_Payment").VALUE = Trim(oRecordSet01.Fields("U_Payment").VALUE)
                Line_Count = Line_Count + 1
Sikp_Line_Item:
            Next i
        End If
        End With
        
        If Line_Count > 0 Then
            j = 0
            If Trim(oDS_PS_MM070H.GetValue("U_Purchase", 0)) = "30" Or Trim(oDS_PS_MM070H.GetValue("U_Purchase", 0)) = "40" Or Trim(oDS_PS_MM070H.GetValue("U_Purchase", 0)) = "60" Then
'            If Trim(oDS_PS_MM070H.GetValue("U_Purchase", 0)) = "30" Or "40" Or "60" Then
            Else
                If Trim(oDS_PS_MM070H.GetValue("U_OIGENum", 0)) = "" Then
                    For i = 0 To oMat01.VisualRowCount - 2
                        If Trim(oDS_PS_MM070L.GetValue("U_Auto", i)) = "Y" And Trim(oDS_PS_MM070L.GetValue("U_RealWt", i)) > 0 Then
                            If j = 0 Then
                                Set DI_oInventoryGenExit = Nothing
                                Set DI_oInventoryGenExit = Sbo_Company.GetBusinessObject(oInventoryGenExit)
                                
                                DI_oInventoryGenExit.BPL_IDAssignedToInvoice = Trim(oDS_PS_MM070H.GetValue("U_BPLId", 0))
                                DI_oInventoryGenExit.DocDate = DocDate
                                DI_oInventoryGenExit.TaxDate = DocDate
                                DI_oInventoryGenExit.Comments = "검수입고(" & Trim(oDS_PS_MM070H.GetValue("DocNum", 0)) & ")출고 : " & Trim(oDS_PS_MM070H.GetValue("U_Comments", 0))
                            End If
    
                            If j > 0 Then DI_oInventoryGenExit.Lines.Add
                            
                            sQry = "EXEC [PS_MM070_01] '" & Trim(oDS_PS_MM070L.GetValue("U_GADocLin", i)) & "'"
                            oRecordSet01.DoQuery sQry
                                                        
                            DI_oInventoryGenExit.Lines.SetCurrentLine j
                            DI_oInventoryGenExit.Lines.ItemCode = Trim(oDS_PS_MM070L.GetValue("U_ItemCode", i))
                            DI_oInventoryGenExit.Lines.WarehouseCode = Trim(oDS_PS_MM070L.GetValue("U_WhsCode", i))
                            DI_oInventoryGenExit.Lines.Quantity = Trim(oDS_PS_MM070L.GetValue("U_RealWt", i))
                            DI_oInventoryGenExit.Lines.CostingCode = Trim(oRecordSet01.Fields(0).VALUE) '배부규칙
                            
                            '금액, 단가 지정하여 입력(2014.01.29 송명규 추가)
                            DI_oInventoryGenExit.Lines.LineTotal = Trim(oDS_PS_MM070L.GetValue("U_LinTotal", i)) '금액
                            DI_oInventoryGenExit.Lines.Price = CDbl(oDS_PS_MM070L.GetValue("U_Price", i)) '단가
                            
                            j = j + 1
                        End If
                    Next i
                End If
            End If
        End If
        
        If Line_Count <> 0 Then
            '//완료
            RetVal = DI_oPurchaseDeliveryNotes.Add
            If (0 <> RetVal) Then
                Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
                ErrNum = 1
                GoTo Add_oPurchaseDeliveryNotes_Error
            End If
                
            If ChkType = 1 Then
                If DI_oInventoryGenExit Is Nothing Then
                Else
                    RetVal = DI_oInventoryGenExit.Add
                    If (0 <> RetVal) Then
                        Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
                        ErrNum = 1
                        GoTo Add_oPurchaseDeliveryNotes_Error
                    End If
                End If
                Sbo_Company.EndTransaction wf_RollBack
            ElseIf ChkType = 2 Then
                Sbo_Company.GetNewObjectCode DocEntry
                oDS_PS_MM070H.setValue "U_GRDocNum", 0, DocEntry
                
                If DI_oInventoryGenExit Is Nothing Then
                Else
                    RetVal = DI_oInventoryGenExit.Add
                    If (0 <> RetVal) Then
                        Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
                        ErrNum = 1
                        GoTo Add_oPurchaseDeliveryNotes_Error
                    End If
                    
                    Sbo_Company.GetNewObjectCode DocEntry
                    oDS_PS_MM070H.setValue "U_OIGENum", 0, DocEntry
                End If
                           
                Sbo_Company.EndTransaction wf_Commit
            End If
        Else
            Sbo_Company.EndTransaction wf_RollBack
        End If
    End If
    
    Set DI_oPurchaseDeliveryNotes = Nothing
    Set DI_oInventoryGenExit = Nothing
    Add_oPurchaseDeliveryNotes = True
    Set oRecordSet01 = Nothing
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_oPurchaseDeliveryNotes_Error:
    Set DI_oPurchaseDeliveryNotes = Nothing
    Set DI_oInventoryGenExit = Nothing
    Set oRecordSet01 = Nothing
    If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_RollBack
    Add_oPurchaseDeliveryNotes = False
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "Add_oPurchaseDeliveryNotes_Error:" & ErrCode & " - " & ErrMsg, "E"
    Else
        MDC_Com.MDC_GF_Message "Add_oPurchaseDeliveryNotes_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Function

Private Function Add_oPurchaseReturns(ChkType As Integer) As Boolean
On Error GoTo Add_oPurchaseReturns_Error
    Dim DI_oPurchaseReturns As SAPbobsCOM.Documents         '//구매오더 문서객체
    Dim DI_oInventoryGenEntry As SAPbobsCOM.Documents             '//입고 문서객체
        
    Dim i&, j&
    Dim ErrNum&, ErrCode As Long, ErrMsg$, RetVal&
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    Dim sQry            As String

    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocDate$, DueDate$, DocEntry$, ECVatGroup$
    
    oMat01.FlushToDataSource
    
    DocDate = Format(oDS_PS_MM070H.GetValue("U_DocDate", 0), "0000-00-00")
    DueDate = Format(oDS_PS_MM070H.GetValue("U_DueDate", 0), "0000-00-00")
    
    If Trim(oDS_PS_MM070H.GetValue("U_GRDocNum", 0)) <> "" And Trim(oDS_PS_MM070H.GetValue("U_ORPDNum", 0)) = "" Then
        Sbo_Company.StartTransaction
        Set DI_oPurchaseReturns = Nothing
        Set DI_oPurchaseReturns = Sbo_Company.GetBusinessObject(oPurchaseReturns)
        
        With DI_oPurchaseReturns
        .CardCode = Trim(oDS_PS_MM070H.GetValue("U_CardCode", 0))
        .BPL_IDAssignedToInvoice = Trim(oDS_PS_MM070H.GetValue("U_BPLId", 0))
        .DocDate = DocDate
        .DocDueDate = DueDate
        .Comments = "검수입고(" & Trim(oDS_PS_MM070H.GetValue("DocNum", 0)) & ")반품 : " & oDS_PS_MM070H.GetValue("U_Comments", 0)
        .UserFields("U_CancDoc").VALUE = Trim(oDS_PS_MM070H.GetValue("U_GRDocNum", 0))
        
        sQry = "Select T1.ItemCode, T1.Dscription, T1.WhsCode, T1.Quantity, T1.Price, T1.DiscPrcnt, T1.LineTotal, T1.U_Weight, T0.ObjType, "
        sQry = sQry & "T1.DocEntry, T1.LineNum, T1.U_sItemCode, T1.U_SItemName, T1.U_sSize, T1.U_sUnit, T1.U_sQty, T1.U_sWeight, T1.VatGroup "
        sQry = sQry & "From OPDN T0 Inner Join PDN1 T1 On T1.DocEntry = T0.DocEntry "
        sQry = sQry & "Where T0.DocNum = '" & Trim(oDS_PS_MM070H.GetValue("U_GRDocNum", 0)) & "' Order by T1.LineNum"
        oRecordSet01.DoQuery sQry
        
        'If Trim(oDS_PS_MM070H.GetValue("U_Purchase", i)) = "30" Or Trim(oDS_PS_MM070H.GetValue("U_Purchase", i)) = "40" Then
        If Trim(oDS_PS_MM070H.GetValue("U_Purchase", i)) = "30" Or Trim(oDS_PS_MM070H.GetValue("U_Purchase", i)) = "40" Or Trim(oDS_PS_MM070H.GetValue("U_Purchase", i)) = "60" Then
            .DocType = dDocument_Service
            
            Do Until oRecordSet01.EOF
                If i > 0 Then .Lines.Add
                .Lines.SetCurrentLine i
                
                .Lines.ItemDescription = Trim(oRecordSet01.Fields("Dscription").VALUE)
                .Lines.LineTotal = Trim(oRecordSet01.Fields("LineTotal").VALUE)
'                    .Lines.WarehouseCode = Trim(oDS_PS_MM030L.GetValue("U_WhsCode", i))
                .Lines.VATGroup = Trim(oRecordSet01.Fields("VatGroup").VALUE)
'                    .Lines.AccountCode
'                .Lines.UserFields("PQDocNum").Value = oDS_PS_MM030H.getValue("U_PQDocNum", 0)
                .Lines.UserFields("U_sItemCode").VALUE = Trim(oRecordSet01.Fields("U_sItemCode").VALUE)
                .Lines.UserFields("U_sItemName").VALUE = Trim(oRecordSet01.Fields("U_sItemName").VALUE)
                .Lines.UserFields("U_sSize").VALUE = Trim(oRecordSet01.Fields("U_sSize").VALUE)
                .Lines.UserFields("U_sUnit").VALUE = Trim(oRecordSet01.Fields("U_sUnit").VALUE)
                .Lines.UserFields("U_sQty").VALUE = Trim(oRecordSet01.Fields("U_sQty").VALUE)
                .Lines.UserFields("U_sWeight").VALUE = Trim(oRecordSet01.Fields("U_sWeight").VALUE)
                
                .Lines.BaseType = Trim(oRecordSet01.Fields("ObjType").VALUE)
                .Lines.BaseEntry = Trim(oRecordSet01.Fields("DocEntry").VALUE)
                .Lines.BaseLine = Trim(oRecordSet01.Fields("LineNum").VALUE)
                
                i = i + 1
                oRecordSet01.MoveNext
            Loop
        Else
            .DocType = dDocument_Items
            Do Until oRecordSet01.EOF
                If i > 0 Then .Lines.Add
                .Lines.SetCurrentLine i
                
'                .Lines.UserFields("PQDocNum").Value = oDS_PS_MM070H.getValue("U_PQDocNum", 0)
                .Lines.ItemCode = Trim(oRecordSet01.Fields("ItemCode").VALUE)
                .Lines.Quantity = Trim(oRecordSet01.Fields("Quantity").VALUE)
                .Lines.Price = Trim(oRecordSet01.Fields("Price").VALUE)
                .Lines.DiscountPercent = Trim(oRecordSet01.Fields("DiscPrcnt").VALUE)
                .Lines.LineTotal = Trim(oRecordSet01.Fields("LineTotal").VALUE)
                .Lines.WarehouseCode = Trim(oRecordSet01.Fields("WhsCode").VALUE)
                .Lines.UserFields("U_Qty").VALUE = Trim(oDS_PS_MM070L.GetValue("U_Qty", i))
                .Lines.UserFields("U_Weight").VALUE = Trim(oRecordSet01.Fields("U_Weight").VALUE)
                
                .Lines.BaseType = Trim(oRecordSet01.Fields("ObjType").VALUE)
                .Lines.BaseEntry = Trim(oRecordSet01.Fields("DocEntry").VALUE)
                .Lines.BaseLine = Trim(oRecordSet01.Fields("LineNum").VALUE)
                
'                        If Trim(oDS_PS_MM070L.GetValue("U_BatchYN", i)) = "Y" Then
'                            .Lines.BatchNumbers.BatchNumber = Trim(oDS_PS_MM070L.GetValue("U_BatchNum", i))
'                            .Lines.BatchNumbers.Quantity = Trim(oDS_PS_MM070L.GetValue("U_Weight", i))
'                        End If
                i = i + 1
                oRecordSet01.MoveNext
            Loop
        End If
        End With
        
        j = 0
        'If (Trim(oDS_PS_MM070H.GetValue("U_Purchase", 0)) = "30" Or Trim(oDS_PS_MM070H.GetValue("U_Purchase", 0)) = "40") Then
        If Trim(oDS_PS_MM070H.GetValue("U_Purchase", 0)) = "30" Or Trim(oDS_PS_MM070H.GetValue("U_Purchase", 0)) = "40" Or Trim(oDS_PS_MM070H.GetValue("U_Purchase", 0)) = "60" Then
        Else
            If Trim(oDS_PS_MM070H.GetValue("U_OIGENum", 0)) <> "" And Trim(oDS_PS_MM070H.GetValue("U_OIGNNum", 0)) = "" Then
                For i = 0 To oMat01.VisualRowCount - 2
                    If Trim(oDS_PS_MM070L.GetValue("U_Auto", i)) = "Y" And Trim(oDS_PS_MM070L.GetValue("U_RealWt", i)) > 0 Then
                        If j = 0 Then
                            Set DI_oInventoryGenEntry = Nothing
                            Set DI_oInventoryGenEntry = Sbo_Company.GetBusinessObject(oInventoryGenEntry)
                            
                            DI_oInventoryGenEntry.BPL_IDAssignedToInvoice = Trim(oDS_PS_MM070H.GetValue("U_BPLId", 0))
                            DI_oInventoryGenEntry.DocDate = DocDate
                            DI_oInventoryGenEntry.TaxDate = DocDate
                            DI_oInventoryGenEntry.Comments = "검수입고(" & Trim(oDS_PS_MM070H.GetValue("DocNum", 0)) & ")출고 : " & Trim(oDS_PS_MM070H.GetValue("U_Comments", 0))
                        End If

                        If j > 0 Then DI_oInventoryGenEntry.Lines.Add
                        
                        sQry = "EXEC [PS_MM070_01] '" & Trim(oDS_PS_MM070L.GetValue("U_GADocLin", i)) & "'"
                        oRecordSet01.DoQuery sQry
                        
                        DI_oInventoryGenEntry.Lines.SetCurrentLine j
                        DI_oInventoryGenEntry.Lines.ItemCode = Trim(oDS_PS_MM070L.GetValue("U_ItemCode", i))
                        DI_oInventoryGenEntry.Lines.WarehouseCode = Trim(oDS_PS_MM070L.GetValue("U_WhsCode", i))
                        DI_oInventoryGenEntry.Lines.Quantity = Trim(oDS_PS_MM070L.GetValue("U_RealWt", i))
                        DI_oInventoryGenEntry.Lines.CostingCode = Trim(oRecordSet01.Fields(0).VALUE)
                        
                        j = j + 1
                    End If
                Next i
            End If
        End If
        
        '//완료
        If ChkType = 1 Then
            If DI_oInventoryGenEntry Is Nothing Then
            Else
                RetVal = DI_oInventoryGenEntry.Add
                If (0 <> RetVal) Then
                    Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
                    ErrNum = 1
                    GoTo Add_oPurchaseReturns_Error
                End If
            End If
            
            RetVal = DI_oPurchaseReturns.Add
            If (0 <> RetVal) Then
                Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
                ErrNum = 1
                GoTo Add_oPurchaseReturns_Error
            End If
            
            Sbo_Company.EndTransaction wf_RollBack
        ElseIf ChkType = 2 Then
            If DI_oInventoryGenEntry Is Nothing Then
            Else
                RetVal = DI_oInventoryGenEntry.Add
                If (0 <> RetVal) Then
                    Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
                    ErrNum = 1
                    GoTo Add_oPurchaseReturns_Error
                End If
                
                Sbo_Company.GetNewObjectCode DocEntry
                oDS_PS_MM070H.setValue "U_OIGNNum", 0, DocEntry
            End If
            
            RetVal = DI_oPurchaseReturns.Add
            If (0 <> RetVal) Then
                Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
                ErrNum = 1
                GoTo Add_oPurchaseReturns_Error
            End If
            Sbo_Company.GetNewObjectCode DocEntry
            oDS_PS_MM070H.setValue "U_ORPDNum", 0, DocEntry
            
            Sbo_Company.EndTransaction wf_Commit
        End If
    End If
    
    Set DI_oPurchaseReturns = Nothing
    Set DI_oInventoryGenEntry = Nothing
    Add_oPurchaseReturns = True
    Set oRecordSet01 = Nothing
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_oPurchaseReturns_Error:
    Set DI_oPurchaseReturns = Nothing
    Set DI_oInventoryGenEntry = Nothing
    Set oRecordSet01 = Nothing
    If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_RollBack
    Add_oPurchaseReturns = False
    MDC_Com.MDC_GF_Message "Add_oPurchaseReturns_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Private Function Close_oPurchaseDeliveryNotes(ChkType As Integer) As Boolean
On Error GoTo Close_oPurchaseDeliveryNotes_Error
    Dim DI_oPurchaseDeliveryNotes As SAPbobsCOM.Documents         '//구매오더 문서객체
        
    Dim i&
    Dim ErrNum&, ErrCode As Long, ErrMsg$, RetVal&
    
    Dim DocEntry$
    
    DocEntry = Trim(oDS_PS_MM070H.GetValue("U_PODocNum", 0))
    
    If Trim(oDS_PS_MM070H.GetValue("U_PODocNum", 0)) <> "" Then
        Sbo_Company.StartTransaction
        Set DI_oPurchaseDeliveryNotes = Nothing
        Set DI_oPurchaseDeliveryNotes = Sbo_Company.GetBusinessObject(oPurchaseDeliveryNotes)
                
        '//완료
        If (DI_oPurchaseDeliveryNotes.GetByKey(DocEntry) = False) Then
            Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
            GoTo Close_oPurchaseDeliveryNotes_Error
        End If
        RetVal = DI_oPurchaseDeliveryNotes.Close
        If (0 <> RetVal) Then
            Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
            GoTo Close_oPurchaseDeliveryNotes_Error
        End If
        
        If ChkType = 1 Then
            Sbo_Company.EndTransaction wf_RollBack
        ElseIf ChkType = 2 Then
            Sbo_Company.EndTransaction wf_Commit
        End If
    End If
    
    Set DI_oPurchaseDeliveryNotes = Nothing
    Close_oPurchaseDeliveryNotes = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Close_oPurchaseDeliveryNotes_Error:
    Set DI_oPurchaseDeliveryNotes = Nothing
    If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_RollBack
    Close_oPurchaseDeliveryNotes = False
    MDC_Com.MDC_GF_Message "Close_oPurchaseDeliveryNotes_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Private Function PS_MM070_CheckDate(ByVal pBaseEntry As String) As Boolean
'******************************************************************************
'Function ID : PS_MM070_CheckDate()
'해당모듈    : PS_MM070
'기능        : 선행프로세스와 일자 비교
'인수        : pBaseEntry : 기준문서번호
'반환값      : True-선행프로세스보다 일자가 같거나 느릴 경우, False-선행프로세스보다 일자가 빠를 경우
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_MM070_CheckDate_Error

    Dim Query01      As String
    Dim loopCount    As Integer
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BaseEntry    As String
    Dim BaseLine     As String
    Dim DocType      As String
    Dim CurDocDate   As String
    
    BaseEntry = pBaseEntry
    BaseLine = ""
    DocType = "PS_MM070"
    CurDocDate = oForm01.Items("DocDate").Specific.VALUE

    Query01 = "         EXEC PS_Z_CHECK_DATE '"
    Query01 = Query01 & BaseEntry & "','"
    Query01 = Query01 & BaseLine & "','"
    Query01 = Query01 & DocType & "','"
    Query01 = Query01 & CurDocDate & "'"

    Call oRecordSet01.DoQuery(Query01)

    If oRecordSet01.Fields("ReturnValue").VALUE = "False" Then
        PS_MM070_CheckDate = False
    Else
        PS_MM070_CheckDate = True
    End If
    
    Set oRecordSet01 = Nothing

    Exit Function
PS_MM070_CheckDate_Error:
    PS_MM070_CheckDate = False
    Call Sbo_Application.SetStatusBarMessage("PS_MM070_CheckDate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)
End Function
