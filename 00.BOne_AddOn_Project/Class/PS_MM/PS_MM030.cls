VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_MM030"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_MM030.cls
'//  Module         : MM
'//  Description    : 품의(발주)
'//  FormType       : PS_MM030
'//  Create Date    : 2010.08.24
'//  Modified Date  :
'//  Creator        : Ryu Yung Jo
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01         As SAPbouiCOM.Form
Public oMat01          As SAPbouiCOM.Matrix
Private oDS_PS_MM030H  As SAPbouiCOM.DBDataSource    '등록헤더
Private oDS_PS_MM030L  As SAPbouiCOM.DBDataSource    '등록라인

Private oLast_Item_UID As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID  As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row  As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode     As Long
Private oDocNum        As String

Private Type deletedDatas  '문서에서 삭제된 자료를 저장할 구조체
    LineNum As Integer '삭제한 행의 행번호
End Type
Private deletedData() As deletedDatas
Private g_deletedCount As Integer

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm(Optional sDocNum As String)
On Error GoTo LoadForm_Error

    Dim i           As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01   As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_MM030.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)

    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_MM030_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
  
    oForm01.SupportedModes = -1
    'oForm01.Mode = fm_ADD_MODE
    
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
    Call FormClear
    Call Add_MatrixRow(0, True)
'    Call FormItemEnabled
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1286"), True         '// 닫기
    oForm01.EnableMenu ("1285"), False        '// 복원
    oForm01.EnableMenu ("1284"), True         '// 취소
    oForm01.EnableMenu ("1293"), True         '// 행삭제
    
    g_deletedCount = 0
    
    oForm01.Update
    
    oForm01.Freeze False
    
    oDocNum = sDocNum
    If oDocNum <> "" Then
    
        Call FormItemEnabled
    
        oForm01.Items("DocNum").Specific.VALUE = sDocNum
        Call oForm01.Items("1").Click
        
    Else
    
        Call Initialization
        Call FormItemEnabled
        
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

    Dim ItemCode    As String
    Dim ItemName    As String
    Dim Size        As String
    Dim Qty         As Long
    Dim Weight      As Currency
    Dim Unit        As String
    Dim RequestDate As String
    Dim DueDate     As String
    Dim ItemType    As String
    Dim RequestNo   As String
    Dim i           As Integer
    'Dim i&, DocTotal     As Currency, SumQty As Long, SumWeight As Currency
    Dim sQry        As String
    Dim ErrNum      As String
    
    Dim Calculate_Weight As Double
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ChildForm01 As Variant
    Set ChildForm01 = New PS_SM010
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE Or oForm01.Mode = fm_UPDATE_MODE Then
                        
'                        '저장할 때 한번만 최종구매정보를 입력(2012.06.25 송명규 수정)
'                        For i = 0 To oMat01.VisualRowCount - 2
'
'                            sQry = "EXEC PS_MM030_03 '" & Trim(oForm01.Items("BPLId").Specific.VALUE) & "', '" & Trim(oMat01.Columns("ItemCode").Cells(i + 1).Specific.VALUE) & "', '" & Trim(oMat01.Columns("ProcCode").Cells(i + 1).Specific.VALUE) & "', '" & Trim(oForm01.Items("DocNum").Specific.VALUE) & "', '" & Trim(oForm01.Items("DocDate").Specific.VALUE) & "'"
'                            oRecordset02.DoQuery sQry
'
'                            oMat01.Columns("LastCardCd").Cells(i + 1).Specific.VALUE = Trim(oRecordset02.Fields("U_CardName").VALUE)
'                            If Format(Trim(oRecordset02.Fields("U_DocDate").VALUE), "YYYYMMDD") < "19000102" Then
'                                oMat01.Columns("LastDocDt").Cells(i + 1).Specific.VALUE = ""
'                            Else
'                                oMat01.Columns("LastDocDt").Cells(i + 1).Specific.VALUE = Format(Trim(oRecordset02.Fields("U_DocDate").VALUE), "YYYYMMDD")
'                            End If
'                            oMat01.Columns("LastPrice").Cells(i + 1).Specific.VALUE = Trim(oRecordset02.Fields("U_Price").VALUE)
'
'                        Next
'                        '저장할 때 한번만 최종데이터를 입력(2012.06.25 송명규 수정)
                    
                        If HeaderSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        If MatrixSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        If PS_MM030_TotalAmount_Compare = False Then '대비견적금액 비교
                        
                            If Sbo_Application.MessageBox("대비견적금액이 견적금액보다 적습니다. 정말로 저장하시겠습니까?", "1", "예", "아니오") = "1" Then
                            Else
                        
                                BubbleEvent = False
                                Exit Sub
                                
                            End If
                        End If
                        
                        If oForm01.Mode = fm_ADD_MODE Then
                            If PS_MM030_oPurchaseOrders_Add(1) = False Then
                                BubbleEvent = False
                                Exit Sub
                            End If
                        ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                            
                            If g_deletedCount > 0 Then '행삭제를 한 경우
                            
                                For i = 0 To UBound(deletedData)
                                
                                    Call PS_MM030_oPurchaseOrders_LineDelete(Trim(deletedData(i).LineNum))
                                
                                Next
                            
                                '삭제행자료 초기화_S
                                For i = 0 To UBound(deletedData)

                                    deletedData(i).LineNum = 0

                                Next
                                ReDim deletedData(0)
                                '삭제행자료 초기화_E
                                
                                g_deletedCount = 0 '삭제행 카운트 초기화
                                
                            Else '자료를 수정한 경우
                            
                                Sbo_Company.StartTransaction
                                
                                If PS_MM030_oPurchaseOrders_Update() = False Then
                                    Call Sbo_Company.EndTransaction(wf_RollBack)
                                    BubbleEvent = False
                                    Exit Sub
                                Else
                                    Call Sbo_Company.EndTransaction(wf_Commit)
                                End If
                            End If
                            
                        End If
                    End If
                    oLast_Mode = oForm01.Mode
                ElseIf oForm01.Mode = fm_FIND_MODE Then
                    oLast_Mode = oForm01.Mode
                End If
                
'et_KEY_DOWN ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "CardCode" Then
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
                        If pval.ColUID = "PQDocNum" Then
                            If oMat01.Columns("PQDocNum").Cells(pval.Row).Specific.VALUE = "" Then
                                If Trim(oForm01.Items("BPLId").Specific.VALUE) = "" Or Trim(oForm01.Items("POType").Specific.VALUE) = "" Or Trim(oForm01.Items("Purchase").Specific.VALUE) = "" Then
                                    MDC_Com.MDC_GF_Message "사업장, 품의형태 또는 구매방식을 먼저 선택하세요.", "E"
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
            
                If pval.ItemUID = "Mat01" Then
                
                    If pval.ColUID = "PQDocNum" Then
                    
                        Set ChildForm01 = New PS_MM010
                        
                        Call ChildForm01.LoadForm(oMat01.Columns("PQDocNum").Cells(pval.Row).Specific.VALUE, oForm01.Items("BPLId").Specific.Selected.VALUE)
                    
                    End If
                
                End If
            
'et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_VALIDATE: '//10
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
                    If oForm01.Mode = fm_OK_MODE Then
                        If oLast_Mode = fm_UPDATE_MODE Then
                            Add_MatrixRow oMat01.RowCount, False
                            oLast_Mode = 100
                        ElseIf oLast_Mode = fm_FIND_MODE Then
                            Add_MatrixRow oMat01.RowCount, False
                            FormItemEnabled
                            
                            '//최종구매정보
'                        For i = 0 To oMat01.VisualRowCount - 2
''                            sQry = "Select IsNull(a.U_CardName, '') As U_CardName, MAX(a.U_DocDate) As U_DocDate, IsNull(b.U_Price, 0) As U_Price "
''                            sQry = sQry & "From [@PS_MM030H] a Inner Join [@PS_MM030L] b On a.DocEntry = b.DocEntry "
''                            sQry = sQry & "Where b.U_ItemCode = '" & Trim(oMat01.Columns("ItemCode").Cells(i + 1).Specific.Value) & "' "
''                            sQry = sQry & "And a.U_BPLId = '" & Trim(oForm01.Items("BPLId").Specific.Value) & "' "
''                            sQry = sQry & "And a.DocNum <> '" & Trim(oForm01.Items("DocNum").Specific.Value) & "' "
''                            sQry = sQry & "And a.Status = 'O' "
''                            sQry = sQry & "Group by IsNull(a.U_CardName, ''), IsNull(b.U_Price, 0)"
'                            sQry = "EXEC PS_MM030_03 '" & Trim(oForm01.Items("BPLId").Specific.VALUE) & "', '" & Trim(oMat01.Columns("ItemCode").Cells(i + 1).Specific.VALUE) & "', '" & Trim(oMat01.Columns("ProcCode").Cells(i + 1).Specific.VALUE) & "', '" & Trim(oForm01.Items("DocNum").Specific.VALUE) & "', '" & Trim(oForm01.Items("DocDate").Specific.VALUE) & "'"
'                            oRecordSet01.DoQuery sQry
'
'                            oMat01.Columns("LastCardCd").Cells(i + 1).Specific.VALUE = Trim(oRecordSet01.Fields("U_CardName").VALUE)
'                            If Format(Trim(oRecordSet01.Fields("U_DocDate").VALUE), "YYYYMMDD") < "19000102" Then
'                                oMat01.Columns("LastDocDt").Cells(i + 1).Specific.VALUE = ""
'                            Else
'                                oMat01.Columns("LastDocDt").Cells(i + 1).Specific.VALUE = Format(Trim(oRecordSet01.Fields("U_DocDate").VALUE), "YYYYMMDD")
'                            End If
'                            oMat01.Columns("LastPrice").Cells(i + 1).Specific.VALUE = Trim(oRecordSet01.Fields("U_Price").VALUE)
'                        Next i
                
                            oLast_Mode = 100
                        End If
                    ElseIf oForm01.Mode = fm_ADD_MODE And pval.Action_Success = True Then
                        oForm01.Mode = fm_OK_MODE
                        Call Sbo_Application.ActivateMenuItem("1291") '//이동(최종데이타)
'                        Call Sbo_Application.ActivateMenuItem("1282") '//추가모드
                    End If
                ElseIf pval.ItemUID = "Btn01" Then
'                    If HeaderSpaceLineDel = False Then
'                        BubbleEvent = False
'                        Exit Sub
'                    End If
                
                    Print_Report02
                ElseIf pval.ItemUID = "Btn02" Then
'                    If HeaderSpaceLineDel = False Then
'                        BubbleEvent = False
'                        Exit Sub
'                    End If
                
'                    '메일발송_S
'                    If oForm01.Items("POType").Specific.VALUE = "10" Then '자체구매일 때만 메일발송
'
'                        Call Report_Export 'Print_Report에 Export를 함께 구현할려고 하였으나 Free License로는 한번에 2개 이상의 리포트를 호출 하는것이 불가능하여 별도로 구현
'
'                        Dim ReturnValue As String
'
'                        ReturnValue = PS_MM030_SendMail
'
'                        If ReturnValue <> "" Then
'                            If ReturnValue = "EMailStringEmpty" Then
'                                Call Sbo_Application.MessageBox(oForm01.Items("CardName").Specific.VALUE & "는 BP마스터에 E-Mail 주소가 없습니다." & Chr(13) & "E-Mail 주소 등록 후 주문서 버튼을 다시 클릭하십시오.")
'                            Else
'                                Call Sbo_Application.MessageBox(ReturnValue)
'                            End If
'                        End If
'
'                    End If
'                    '메일발송_E
                    
                    Call Print_Report01 '리포트 미리보기
                    
                End If
            Case et_KEY_DOWN: '//2
'et_COMBO_SELECT ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_COMBO_SELECT: '//5
                If pval.ItemUID = "Purchase" Or pval.ItemUID = "BPLId" Then
                    oForm01.Freeze True
                    oMat01.Clear
                    oDS_PS_MM030L.Clear
                    If oForm01.Mode = fm_ADD_MODE Then
                        Call Add_MatrixRow(0, False)
                    End If
                    
                    If Trim(oForm01.Items("Purchase").Specific.VALUE) = "30" Or Trim(oForm01.Items("Purchase").Specific.VALUE) = "40" Or Trim(oForm01.Items("Purchase").Specific.VALUE) = "60" Then
                        oMat01.Columns("ItemName").Editable = True
                        oMat01.Columns("OutSize").Editable = True
                        oMat01.Columns("OutUnit").Editable = True
                        oMat01.Columns("WhsCode").Editable = False
                    Else
                        oMat01.Columns("ItemName").Editable = False
                        oMat01.Columns("OutSize").Editable = False
                        oMat01.Columns("OutUnit").Editable = False
                        oMat01.Columns("WhsCode").Editable = True
                    End If
                    oForm01.Freeze False
                End If
            Case et_CLICK: '//6
            
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
'et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    oForm01.Freeze True
                    If pval.ItemUID = "CardCode" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "CntcCode" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "PQDocNum" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        ElseIf pval.ColUID = "ItemCode" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        ElseIf pval.ColUID = "Qty" Then
                            oMat01.FlushToDataSource
                            ItemCode = Trim(oDS_PS_MM030L.GetValue("U_ItemCode", pval.Row - 1))
                            Qty = oDS_PS_MM030L.GetValue("U_Qty", pval.Row - 1)
                                            
                            Calculate_Weight = MDC_PS_Common.Calculate_Weight(ItemCode, Qty, Trim(oForm01.Items("BPLId").Specific.VALUE))
                                                                                  
                            oDS_PS_MM030L.setValue "U_Weight", pval.Row - 1, Calculate_Weight
                            oMat01.LoadFromDataSource
                            
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        ElseIf pval.ColUID = "Price" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        ElseIf pval.ColUID = "LinTotal" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        ElseIf pval.ColUID = "WhsCode" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        End If
                    End If
                    oForm01.Freeze False
                End If
            Case et_MATRIX_LOAD: '//11
                '//최종구매정보
'                For i = 0 To oMat01.VisualRowCount - 1
''                    sQry = "Select IsNull(a.U_CardName, '') As U_CardName, MAX(a.U_DocDate) As U_DocDate, IsNull(b.U_Price, 0) As U_Price "
''                    sQry = sQry & "From [@PS_MM030H] a Inner Join [@PS_MM030L] b On a.DocEntry = b.DocEntry "
''                    sQry = sQry & "Where b.U_ItemCode = '" & Trim(oMat01.Columns("ItemCode").Cells(i + 1).Specific.Value) & "' "
''                    sQry = sQry & "And a.U_BPLId = '" & Trim(oForm01.Items("BPLId").Specific.Value) & "' "
''                    sQry = sQry & "And a.DocNum <> '" & Trim(oForm01.Items("DocNum").Specific.Value) & "' "
''                    sQry = sQry & "And a.Status = 'O' "
''                    sQry = sQry & "Group by IsNull(a.U_CardName, ''), IsNull(b.U_Price, 0)"
'                    sQry = "EXEC PS_MM030_03 '" & Trim(oForm01.Items("BPLId").Specific.VALUE) & "', '" & Trim(oMat01.Columns("ItemCode").Cells(i + 1).Specific.VALUE) & "', '" & Trim(oMat01.Columns("ProcCode").Cells(i + 1).Specific.VALUE) & "', '" & Trim(oForm01.Items("DocNum").Specific.VALUE) & "', '" & Trim(oForm01.Items("DocDate").Specific.VALUE) & "'"
'                    oRecordSet01.DoQuery sQry
'
'                    oMat01.Columns("LastCardCd").Cells(i + 1).Specific.VALUE = Trim(oRecordSet01.Fields("U_CardName").VALUE)
'                    If Format(Trim(oRecordSet01.Fields("U_DocDate").VALUE), "YYYYMMDD") < "19000102" Then
'                        oMat01.Columns("LastDocDt").Cells(i + 1).Specific.VALUE = ""
'                    Else
'                        oMat01.Columns("LastDocDt").Cells(i + 1).Specific.VALUE = Format(Trim(oRecordSet01.Fields("U_DocDate").VALUE), "YYYYMMDD")
'                    End If
'                    oMat01.Columns("LastPrice").Cells(i + 1).Specific.VALUE = Trim(oRecordSet01.Fields("U_Price").VALUE)
'                Next i
                
'                For i = 0 To oMat01.VisualRowCount - 1
'                    DocTotal = DocTotal + oMat01.Columns("LinTotal").Cells(i + 1).Specific.VALUE
'                    If oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE = "" Then
'                        SumQty = SumQty
'                    Else
'                        SumQty = SumQty + oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE
'                    End If
'                    SumWeight = SumWeight + oMat01.Columns("Weight").Cells(i + 1).Specific.VALUE
'                Next i
'                oForm01.Items("DocTotal").Specific.VALUE = DocTotal
'                oForm01.Items("SumQty").Specific.VALUE = SumQty
'                oForm01.Items("SumWeight").Specific.VALUE = SumWeight
                
                Call PS_MM030_TotalAmount_Calculate
                
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

    Set oRecordSet01 = Nothing
    oForm01.Freeze False
    Call MDC_Com.MDC_GF_Message("Raise_ItemEvent_Error:" & Err.Number & " - " & Err.Description, "E")

End Sub

Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_MenuEvent_Error

    Dim i            As Integer
    Dim sQry         As String
    Dim DocNum       As String
    Dim LineNum      As String
    Dim BPLID        As String

    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소

                For i = 0 To oMat01.VisualRowCount - 2
                
                    LineNum = Trim(oMat01.Columns("LineId").Cells(i + 1).Specific.VALUE)
                    DocNum = Trim(oForm01.Items("DocNum").Specific.VALUE)
                
                    If PS_MM030_LineDelete_Possible(DocNum, LineNum) = False Then
                
                        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 품의(발주)가 가입고등록 되었습니다. 취소할 수 없습니다.", "E"
                        BubbleEvent = False
                        Exit Sub
                        
                    End If

                Next i
                
                If PS_MM030_oPurchaseOrders_Cancel() = False Then
                    BubbleEvent = False
                    Exit Sub
                Else
                    oLast_Mode = 101
End If
            Case "1286": '닫기
                                BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
                                If BPLID = 2 Then  ' 부산사업장만 해당되도록 품의취소시 가입고가 남아있을경우 닫기 처리 안됨.
                    For i = 0 To oMat01.VisualRowCount - 2
                    LineNum = Trim(oMat01.Columns("LineId").Cells(i + 1).Specific.VALUE)
                    DocNum = Trim(oForm01.Items("DocNum").Specific.VALUE)
                        If PS_MM030_LineDelete_Possible(DocNum, LineNum) = False Then
                        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 품의(발주)가 가입고등록 되었습니다. 취소할 수 없습니다.", "E"
                        BubbleEvent = False
                        Exit Sub
                        
                    End If
                Next i

                If PS_MM030_oPurchaseOrders_Close = False Then
                    BubbleEvent = False
                    Exit Sub
                Else
                    oLast_Mode = 102
                End If

                Else
                
                End If

            Case "1293": '행삭제
            
                If PS_MM030_LineDelete_Possible(oForm01.Items("DocNum").Specific.VALUE, oMat01.Columns("LineId").Cells(oLast_Col_Row).Specific.VALUE) = False Then
                
                    Call MDC_Com.MDC_GF_Message("해당 품의는 가입고가 등록되어 행삭제 할 수 없습니다.", "E")
                    BubbleEvent = False
                    Exit Sub
                
                Else

                    '구매오더 행삭제를 위해 삭제된 행번호 저장
                    ReDim Preserve deletedData(g_deletedCount)
                    deletedData(g_deletedCount).LineNum = oMat01.Columns("LineNum").Cells(oLast_Col_Row).Specific.VALUE 'DI API 상에서 삭제 로직을 적용할 경우는 라인의 순서번호로 구동이 되기때문에 행번호를 저장
                    g_deletedCount = g_deletedCount + 1
                    
                End If
            
'                If Trim(oForm01.Items("PODocNum").Specific.VALUE) <> "" Then
'                    MDC_Com.MDC_GF_Message "작성된 품의는 행삭제를 할 수 없습니다.", "E"
'                    BubbleEvent = False
'                    Exit Sub
'                End If
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
                oForm01.Freeze True
                Call FormItemEnabled
                oForm01.Items("DocNum").Click ct_Regular
                oForm01.Freeze False
            Case "1286": '닫기
            Case "1293": '행삭제
                oForm01.Freeze True
                If oMat01.RowCount <> oMat01.VisualRowCount Then
                    For i = 0 To oMat01.VisualRowCount - 1
                        oMat01.Columns("LineNum").Cells(i + 1).Specific.VALUE = i + 1
                    Next i
                    
                    oMat01.FlushToDataSource
                    oDS_PS_MM030L.RemoveRecord oDS_PS_MM030L.Size - 1       '// Mat01에 마지막라인(빈라인) 삭제
                    oMat01.Clear
                    oMat01.LoadFromDataSource
                    
                    If oMat01.Columns("PQDocNum").Cells(oMat01.RowCount).Specific.VALUE <> "" Then
                        Call Add_MatrixRow(oMat01.RowCount, False)
                    End If
                    
                    '//문서번호 15647 이전은 최종구매정보를 보여준다.
                    If oForm01.Items("DocNum").Specific.VALUE < "15647" Then
                        '//최종구매정보
                        For i = 0 To oMat01.VisualRowCount - 2
    '                        sQry = "Select IsNull(a.U_CardName, '') As U_CardName, MAX(a.U_DocDate) As U_DocDate, IsNull(b.U_Price, 0) As U_Price "
    '                        sQry = sQry & "From [@PS_MM030H] a Inner Join [@PS_MM030L] b On a.DocEntry = b.DocEntry "
    '                        sQry = sQry & "Where b.U_ItemCode = '" & Trim(oMat01.Columns("ItemCode").Cells(i + 1).Specific.Value) & "' "
    '                        sQry = sQry & "And a.U_BPLId = '" & Trim(oForm01.Items("BPLId").Specific.Value) & "' "
    '                        sQry = sQry & "And a.Status = 'O' "
    '                        sQry = sQry & "Group by IsNull(a.U_CardName, ''), IsNull(b.U_Price, 0)"
                            sQry = "EXEC PS_MM030_03 '" & Trim(oForm01.Items("BPLId").Specific.VALUE) & "', '" & Trim(oMat01.Columns("ItemCode").Cells(i + 1).Specific.VALUE) & "', '" & Trim(oMat01.Columns("ProcCode").Cells(i + 1).Specific.VALUE) & "', '" & Trim(oForm01.Items("DocNum").Specific.VALUE) & "', '" & Trim(oForm01.Items("DocDate").Specific.VALUE) & "'"
                            oRecordSet01.DoQuery sQry
                            
                            oMat01.Columns("LCardCd").Cells(i + 1).Specific.VALUE = Trim(oRecordSet01.Fields("CardCode").VALUE)
                            oMat01.Columns("LCardNm").Cells(i + 1).Specific.VALUE = Trim(oRecordSet01.Fields("U_CardName").VALUE)
                            If Format(Trim(oRecordSet01.Fields("U_DocDate").VALUE), "YYYYMMDD") < "19000102" Then
                                oMat01.Columns("LDocDate").Cells(i + 1).Specific.VALUE = ""
                            Else
                                oMat01.Columns("LDocDate").Cells(i + 1).Specific.VALUE = Format(Trim(oRecordSet01.Fields("U_DocDate").VALUE), "YYYYMMDD")
                            End If
                            oMat01.Columns("LPrice").Cells(i + 1).Specific.VALUE = Trim(oRecordSet01.Fields("U_Price").VALUE)
                        Next i
                    End If
                    
                    Call PS_MM030_TotalAmount_Calculate '전체금액 계산
                    
                End If
                
                oForm01.Freeze False
            Case "1281": '찾기
                oForm01.Freeze True
                Call FormItemEnabled
                oForm01.Items("DocNum").Click ct_Regular
                oForm01.Items("SumWeight").Specific.VALUE = 0
                
                If oDocNum = "" Then
                    '//아이디별 사업장 세팅
                    oForm01.Items("BPLId").Specific.Select MDC_PS_Common.User_BPLId, psk_ByValue
                    '//아이디별 사번 세팅
                    If MDC_PS_Common.User_SuperUserYN = "N" Then '수퍼유저인 경우는 사번 미표기(2016.01.08 송명규)
                        oForm01.Items("CntcCode").Specific.VALUE = MDC_PS_Common.User_MSTCOD
                    End If

'                    oForm01.Items("DocDate").Specific.VALUE = Format(Now, "YYYYMMDD")
                End If
                oForm01.Freeze False
            Case "1282": '추가
                oForm01.Freeze True
                
                oDS_PS_MM030H.setValue "U_DocDate", 0, Format(Now, "YYYYMMDD")
                oDS_PS_MM030H.setValue "U_DueDate", 0, Format(Now, "YYYYMMDD")
                oDS_PS_MM030H.setValue "U_POStatus", 0, "N"
                oDS_PS_MM030H.setValue "U_POFinish", 0, "N"
                oDS_PS_MM030H.setValue "U_PurType", 0, "0"
                oDS_PS_MM030H.setValue "U_Payment", 0, "10"
                Call Add_MatrixRow(0, True)
'                oForm01.Items("BPLId").Click ct_Collapsed
                oForm01.Items("SumWeight").Specific.VALUE = 0
                Call oDS_PS_MM030H.setValue("U_Payment", 0, "20") '지불조건 기본값을 현금으로 변경(2014.07.23 송명규, 류석균 요청)
                
                Call Initialization
                Call FormItemEnabled
                Call FormClear
                
                oForm01.Freeze False
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                oForm01.Freeze True
                Call FormItemEnabled
                Call oMat01.AutoResizeColumns
                If oMat01.VisualRowCount > 0 Then
                    If oMat01.Columns("ItemCode").Cells(oMat01.VisualRowCount).Specific.VALUE <> "" Then
                        If oDS_PS_MM030H.GetValue("Status", 0) = "O" Then
                            Add_MatrixRow oMat01.RowCount, False
                        End If
                    End If
                End If
                
                '//최종구매정보
'                For i = 0 To oMat01.VisualRowCount - 2
''                    sQry = "Select IsNull(a.U_CardName, '') As U_CardName, MAX(a.U_DocDate) As U_DocDate, IsNull(b.U_Price, 0) As U_Price "
''                    sQry = sQry & "From [@PS_MM030H] a Inner Join [@PS_MM030L] b On a.DocEntry = b.DocEntry "
''                    sQry = sQry & "Where b.U_ItemCode = '" & Trim(oMat01.Columns("ItemCode").Cells(i + 1).Specific.Value) & "' "
''                    sQry = sQry & "And a.U_BPLId = '" & Trim(oForm01.Items("BPLId").Specific.Value) & "' "
''                    sQry = sQry & "And a.DocNum <> '" & Trim(oForm01.Items("DocNum").Specific.Value) & "' "
''                    sQry = sQry & "And a.Status = 'O' "
''                    sQry = sQry & "Group by IsNull(a.U_CardName, ''), IsNull(b.U_Price, 0)"
'                    sQry = "EXEC PS_MM030_03 '" & Trim(oForm01.Items("BPLId").Specific.VALUE) & "', '" & Trim(oMat01.Columns("ItemCode").Cells(i + 1).Specific.VALUE) & "', '" & Trim(oMat01.Columns("ProcCode").Cells(i + 1).Specific.VALUE) & "', '" & Trim(oForm01.Items("DocNum").Specific.VALUE) & "', '" & Trim(oForm01.Items("DocDate").Specific.VALUE) & "'"
'                    oRecordSet01.DoQuery sQry
'
'
''                    oMat01.Columns("LastCardCd").Cells(i + 1).Specific.VALUE = Trim(oRecordSet01.Fields("U_CardName").VALUE)
'                    oMat01.Columns("LCardCd").Cells(i + 1).Specific.VALUE = Trim(oRecordSet01.Fields("CardCode").VALUE)
'                    oMat01.Columns("LCardNm").Cells(i + 1).Specific.VALUE = Trim(oRecordSet01.Fields("U_CardName").VALUE)
'                    If Format(Trim(oRecordSet01.Fields("U_DocDate").VALUE), "YYYYMMDD") < "19000102" Then
''                        oMat01.Columns("LastDocDt").Cells(i + 1).Specific.VALUE = ""
'                        oMat01.Columns("LDocDate").Cells(i + 1).Specific.VALUE = ""
'                    Else
''                        oMat01.Columns("LastDocDt").Cells(i + 1).Specific.VALUE = Format(Trim(oRecordSet01.Fields("U_DocDate").VALUE), "YYYYMMDD")
'                        oMat01.Columns("LDocDate").Cells(i + 1).Specific.VALUE = Format(Trim(oRecordSet01.Fields("U_DocDate").VALUE), "YYYYMMDD")
'                    End If
''                    oMat01.Columns("LastPrice").Cells(i + 1).Specific.VALUE = Trim(oRecordSet01.Fields("U_Price").VALUE)
'                    oMat01.Columns("LPrice").Cells(i + 1).Specific.VALUE = Trim(oRecordSet01.Fields("U_Price").VALUE)
'                Next i
                oForm01.Freeze False
        End Select
    End If
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_MenuEvent_Error:
    Set oRecordSet01 = Nothing
    oForm01.Freeze False
    MDC_Com.MDC_GF_Message "Raise_MenuEvent_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub Raise_FormDataEvent(ByRef FormUID As String, ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_FormDataEvent_Error
    If (BusinessObjectInfo.BeforeAction = True) Then '//BeforeAction = True
        Select Case BusinessObjectInfo.EventType
            Case et_FORM_DATA_LOAD: '//33
            Case et_FORM_DATA_ADD: '//34
                If PS_MM030_oPurchaseOrders_Add(2) = False Then
                    BubbleEvent = False
                    Exit Sub
                Else
                    Call Delete_EmptyRow
                End If
                
            Case et_FORM_DATA_UPDATE: '//35
                If oLast_Mode = fm_UPDATE_MODE Then
'                    If PS_MM030_oPurchaseOrders_Update(2) = False Then
'                        oLast_Mode = 0
'                        BubbleEvent = False
'                        Exit Sub
'                    Else
'                        oLast_Mode = 0
                        Call Delete_EmptyRow
'                    End If
'                ElseIf oLast_Mode = 101 Then '//취소
'                    If PS_MM030_oPurchaseOrders_Cancel(2) = False Then
'                        oLast_Mode = 0
'                        BubbleEvent = False
'                        Exit Sub
'                    Else
'                        oLast_Mode = 0
'                    End If
'                ElseIf oLast_Mode = 102 Then '//닫기
'                    If PS_MM030_oPurchaseOrders_Close(2) = False Then
'                        oLast_Mode = 0
'                        BubbleEvent = False
'                        Exit Sub
'                    Else
'                        oLast_Mode = 0
'                    End If
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
    Set oDS_PS_MM030H = oForm01.DataSources.DBDataSources("@PS_MM030H")
    Set oDS_PS_MM030L = oForm01.DataSources.DBDataSources("@PS_MM030L")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    
    Call oForm01.DataSources.UserDataSources.Add("DocTotal", dt_SUM)
    oForm01.Items("DocTotal").Specific.DataBind.SetBound True, "", "DocTotal"
'    oForm01.DataSources.UserDataSources.Item("DocDateFr").Value = 0

    Call oForm01.DataSources.UserDataSources.Add("SumQty", dt_SUM)
    oForm01.Items("SumQty").Specific.DataBind.SetBound True, "", "SumQty"
    
    Call oForm01.DataSources.UserDataSources.Add("SumWeight", dt_QUANTITY)
    oForm01.Items("SumWeight").Specific.DataBind.SetBound True, "", "SumWeight"

'    Call oForm01.DataSources.UserDataSources.Add("LastCardCd", dt_SHORT_TEXT, 50)
'    oMat01.Columns("LastCardCd").DataBind.SetBound True, "", "LastCardCd"
''
'    Call oForm01.DataSources.UserDataSources.Add("LastDocDt", dt_DATE)
'    oMat01.Columns("LastDocDt").DataBind.SetBound True, , "LastDocDt"
'
'    Call oForm01.DataSources.UserDataSources.Add("LastPrice", dt_SUM)
'    oMat01.Columns("LastPrice").DataBind.SetBound True, , "LastPrice"
'
    oDS_PS_MM030H.setValue "U_DocDate", 0, Format(Now, "YYYYMMDD")
    oDS_PS_MM030H.setValue "U_DueDate", 0, Format(Now, "YYYYMMDD")
    
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
    Set oCombo = oForm01.Items("POType").Specific
    sQry = "SELECT Code, Name From [@PSH_RETYPE]"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
        
    '//품의상태
    Set oCombo = oForm01.Items("POStatus").Specific
    oCombo.ValidValues.Add "Y", "승인"
    oCombo.ValidValues.Add "N", "미승인"
    Call oDS_PS_MM030H.setValue("U_POStatus", 0, "N")
    
    '//품의종결 추가 20170807
    Set oCombo = oForm01.Items("POFinish").Specific
    oCombo.ValidValues.Add "Y", "종결"
    oCombo.ValidValues.Add "N", "미종결"
    Call oDS_PS_MM030H.setValue("U_POFinish", 0, "N")
    
    
'    oCombo.Select "N", psk_ByValue
    oForm01.Items("CardCode").Click ct_Regular
    
    '//구매방식
    Set oCombo = oForm01.Items("Purchase").Specific
    sQry = "SELECT Code, Name From [@PSH_ORDTYP] Order by Code"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    Call oDS_PS_MM030H.setValue("U_PurType", 0, "0")
'    oForm01.Items("PurType").Specific.Select "0", psk_ByValue
    
    '//지불조건
    Set oCombo = oForm01.Items("Payment").Specific
    sQry = "Select U_Minor, U_CdName From [@PS_SY001L] Where Code = 'M006' Order by U_Minor"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    Call oDS_PS_MM030H.setValue("U_Payment", 0, "20") '기본값을 현금으로 변경(2014.07.23 송명규, 류석균 요청)
'    oForm01.Items("Payment").Specific.Select "10", psk_ByValue
    
    '//외주사유코드
    sQry = "        SELECT      T1.U_Minor AS [Code],"
    sQry = sQry & "             T1.U_CdName AS [Value]"
    sQry = sQry & " FROM        [@PS_SY001H] AS T0"
    sQry = sQry & "             INNER JOIN"
    sQry = sQry & "             [@PS_SY001L] AS T1"
    sQry = sQry & "                 ON T0.Code = T1.Code"
    sQry = sQry & " WHERE       T0.Code = 'P201'"
    sQry = sQry & "             AND T1.U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    T1.U_Seq"
    
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oMat01.Columns("OutCode").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    '통화
    sQry = "        SELECT      T0.U_Minor,"
    sQry = sQry & "             T1.CurrName + '(' + T0.U_CdName + ')'"
    sQry = sQry & " FROM        [@PS_SY001L] AS T0"
    sQry = sQry & "             LEFT JOIN"
    sQry = sQry & "             [OCRN] AS T1"
    sQry = sQry & "                 ON T0.U_Minor = T1.CurrCode"
    sQry = sQry & " WHERE       T0.Code = 'F004'"
    sQry = sQry & " ORDER BY    T0.U_Seq"
    
'    Call oForm01.Items("DocCur").Specific.ValidValues.Add("%", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("DocCur").Specific, sQry, "", False, False)
    Call oForm01.Items("DocCur").Specific.Select(0, psk_Index)
    
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
'    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    Call oDS_PS_MM030H.setValue("U_BPLId", 0, MDC_PS_Common.User_BPLId)
    
    '//아이디별 사번 세팅
    oForm01.Items("CntcCode").Specific.VALUE = MDC_PS_Common.User_MSTCOD
    
    '//아이디별 부서 세팅
'    Set oCombo = oForm01.Items("DeptCode").Specific
'    oCombo.Select MDC_PS_Common.User_DeptCode, psk_ByValue
    
    Call oForm01.Items("DocCur").Specific.Select(0, psk_Index)
    oForm01.Items("DocRate").Specific.VALUE = 1

    Call oForm01.Items("CardCode").Click
    
    g_deletedCount = 0 '삭제행 카운트 초기화
    
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
        
        oForm01.Items("CardCode").Enabled = True
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("CntcCode").Enabled = True
        oForm01.Items("POType").Enabled = True
        oForm01.Items("Purchase").Enabled = True
        oForm01.Items("DocDate").Enabled = True
        oForm01.Items("DueDate").Enabled = True
        
        oForm01.Items("POStatus").Enabled = False
       ' oForm01.Items("POFinish").Enabled = False
        oForm01.Items("AdmsDate").Enabled = False
                
        oForm01.Items("Mat01").Enabled = True
        oMat01.Columns("PQDocNum").Editable = True
        oMat01.Columns("PQLinNum").Editable = True
        oMat01.Columns("ItemCode").Editable = True
        oMat01.Columns("ItemName").Editable = True
        oMat01.Columns("OutSize").Editable = True
        oMat01.Columns("OutUnit").Editable = True
        oMat01.Columns("Qty").Editable = True
        oMat01.Columns("UnWeight").Editable = True
        oMat01.Columns("Price").Editable = True
        oMat01.Columns("LinTotal").Editable = True
        oMat01.Columns("WhsCode").Editable = True
        oMat01.Columns("Comments").Editable = True
        
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        oForm01.Items("DocNum").Enabled = True
        
        oForm01.Items("CardCode").Enabled = True
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("CntcCode").Enabled = True
        oForm01.Items("POType").Enabled = True
        oForm01.Items("Purchase").Enabled = True
        oForm01.Items("DocDate").Enabled = True
        oForm01.Items("DueDate").Enabled = True
        
        oForm01.Items("POStatus").Enabled = False
      '  oForm01.Items("POFinish").Enabled = False
        oForm01.Items("AdmsDate").Enabled = False
        
        oForm01.Items("Mat01").Enabled = False
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        oForm01.Items("DocNum").Enabled = False
        oForm01.Items("CardCode").Enabled = False
        oForm01.Items("BPLId").Enabled = False
        oForm01.Items("CntcCode").Enabled = True
        oForm01.Items("POType").Enabled = False
        oForm01.Items("Purchase").Enabled = False
        oForm01.Items("DocDate").Enabled = True
        oForm01.Items("DueDate").Enabled = True
        
        oForm01.Items("POStatus").Enabled = False
       ' oForm01.Items("POFinish").Enabled = False
        oForm01.Items("AdmsDate").Enabled = False
        If Trim(oDS_PS_MM030H.GetValue("Canceled", 0)) = "Y" Then
            oForm01.Items("Mat01").Enabled = False
        Else
            oForm01.Items("Mat01").Enabled = True
            If Trim(oDS_PS_MM030H.GetValue("U_POStatus", 0)) = "Y" Then
                oMat01.Columns("PQDocNum").Editable = False
                oMat01.Columns("PQLinNum").Editable = False
                oMat01.Columns("ItemCode").Editable = False
                oMat01.Columns("ItemName").Editable = False
                oMat01.Columns("OutSize").Editable = False
                oMat01.Columns("OutUnit").Editable = False
                oMat01.Columns("Qty").Editable = False
                oMat01.Columns("UnWeight").Editable = False
                oMat01.Columns("Price").Editable = False
                oMat01.Columns("LinTotal").Editable = False
                oMat01.Columns("WhsCode").Editable = False
                oMat01.Columns("Comments").Editable = False
                
            Else
                oMat01.Columns("PQDocNum").Editable = True
                oMat01.Columns("PQLinNum").Editable = True
                oMat01.Columns("ItemCode").Editable = True
                oMat01.Columns("ItemName").Editable = True
                oMat01.Columns("OutSize").Editable = True
                oMat01.Columns("OutUnit").Editable = True
                oMat01.Columns("Qty").Editable = True
                oMat01.Columns("UnWeight").Editable = True
                oMat01.Columns("Price").Editable = True
                oMat01.Columns("LinTotal").Editable = True
                oMat01.Columns("WhsCode").Editable = True
                oMat01.Columns("Comments").Editable = True
            End If
        End If
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub FormClear()
On Error GoTo FormClear_Error

    Dim DocNum As String
    DocNum = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM030'", "")
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
        oDS_PS_MM030L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_MM030L.offset = oRow
    oDS_PS_MM030L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error

    Dim i         As Integer
    Dim ErrNum    As Integer
    Dim sQry      As String
    Dim sRow      As Long
    Dim sSeq      As String
    Dim WhsCode   As String
    Dim WhsName   As String
    Dim DocTotal  As Currency
    Dim SumQty    As Long
    Dim SumWeight As Currency
        
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim oRecordset02 As SAPbobsCOM.Recordset
    Dim oRecordSet03 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Set oRecordset02 = Sbo_Company.GetBusinessObject(BoRecordset)
    Set oRecordSet03 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sRow = oRow
    
    Select Case oUID
        Case "CardCode"
            sQry = "Select CardName From OCRD Where CardCode = '" & Trim(oDS_PS_MM030H.GetValue("U_CardCode", 0)) & "'"
            oRecordSet01.DoQuery sQry
            
            oDS_PS_MM030H.setValue "U_CardName", 0, Trim(oRecordSet01.Fields(0).VALUE)
            
            sQry = "Select U_payment From OCRD Where CardCode = '" & Trim(oDS_PS_MM030H.GetValue("U_CardCode", 0)) & "'"
            oRecordSet01.DoQuery sQry
            
            oDS_PS_MM030H.setValue "U_payment", 0, Trim(oRecordSet01.Fields(0).VALUE)
            
        Case "CntcCode"
            sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" & Trim(oDS_PS_MM030H.GetValue("U_CntcCode", 0)) & "'"
            oRecordSet01.DoQuery sQry
            
            oDS_PS_MM030H.setValue "U_CntcName", 0, Trim(oRecordSet01.Fields(0).VALUE)
            
        Case "Mat01"
            If oCol = "PQDocNum" Then
                oForm01.Freeze True
                
                oMat01.FlushToDataSource
                
                WhsCode = User_WhsCode("1")
                WhsCode = Left(WhsCode, 2) + oForm01.Items("BPLId").Specific.VALUE
                
                '프로시저로 수정(2012.03.26 송명규)
                sQry = "      EXEC PS_MM030_50 '"
                sQry = sQry & WhsCode & "','"
                sQry = sQry & Trim(oDS_PS_MM030L.GetValue("U_PQDocNum", oRow - 1)) & "','"
                sQry = sQry & Trim(oForm01.Items("DocCur").Specific.Selected.VALUE) & "'"
                
                oRecordSet01.DoQuery sQry
                
                If oRecordSet01.RecordCount = 0 Then
                    MDC_Com.MDC_GF_Message "구매견적문서가 취소되었거나 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
    '                ErrNum = 1
    '                GoTo FlushToItemValue_Error
                Else
                    If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 0) And Trim(oMat01.Columns("PQDocNum").Cells(oRow).Specific.VALUE) <> "" Then
                        oMat01.FlushToDataSource
                        Call Add_MatrixRow(oMat01.RowCount, False)
                        oMat01.Columns("PQDocNum").Cells(oRow).Click ct_Regular
                    End If
                
                    Do Until oRecordSet01.EOF
                        sSeq = "Y"
                        For i = 0 To oMat01.VisualRowCount - 2
                            If Trim(oDS_PS_MM030L.GetValue("U_PQDocNum", i)) = Trim(oRecordSet01.Fields(0).VALUE) And _
                               Trim(oDS_PS_MM030L.GetValue("U_PQLinNum", i)) = Trim(oRecordSet01.Fields(1).VALUE) Then
                                sSeq = "N"
                            End If
                        Next i
                        If sSeq = "Y" Then
                            oDS_PS_MM030L.setValue "U_PQDocNum", sRow - 1, Trim(oRecordSet01.Fields(0).VALUE)
                            oDS_PS_MM030L.setValue "U_PQLinNum", sRow - 1, Trim(oRecordSet01.Fields(1).VALUE)
                            oDS_PS_MM030L.setValue "U_ItemCode", sRow - 1, Trim(oRecordSet01.Fields("U_ItemCode").VALUE)
                            oDS_PS_MM030L.setValue "U_ItemName", sRow - 1, Trim(oRecordSet01.Fields("U_ItemName").VALUE)
'                            oDS_PS_MM030L.setValue "U_FrgnName", sRow - 1, Trim(oRecordSet01.Fields("U_FrgnName").Value)
                            oDS_PS_MM030L.setValue "U_Qty", sRow - 1, Trim(oRecordSet01.Fields("U_Qty").VALUE)
                            oDS_PS_MM030L.setValue "U_Weight", sRow - 1, Trim(oRecordSet01.Fields("U_Weight").VALUE)
                            If Trim(oRecordSet01.Fields("U_Qty").VALUE) = 0 Then
                                oDS_PS_MM030L.setValue "U_UnWeight", sRow - 1, 0
                            Else
                                oDS_PS_MM030L.setValue "U_UnWeight", sRow - 1, Trim(oRecordSet01.Fields("U_Weight").VALUE) / Trim(oRecordSet01.Fields("U_Qty").VALUE)
                            End If
                            oDS_PS_MM030L.setValue "U_WhsCode", sRow - 1, WhsCode
                            oDS_PS_MM030L.setValue "U_WhsName", sRow - 1, Trim(oRecordSet01.Fields("WhsName").VALUE)
'                            oDS_PS_MM030L.setValue "U_OutItmCd", sRow - 1, Trim(oRecordSet01.Fields("U_OutItmCd").Value)
                            oDS_PS_MM030L.setValue "U_OutSize", sRow - 1, Trim(oRecordSet01.Fields("U_OutSize").VALUE)
                            oDS_PS_MM030L.setValue "U_OutUnit", sRow - 1, Trim(oRecordSet01.Fields("U_OutUnit").VALUE)
                            oDS_PS_MM030L.setValue "U_Auto", sRow - 1, Trim(oRecordSet01.Fields("U_Auto").VALUE)
                            oDS_PS_MM030L.setValue "U_Comments", sRow - 1, Trim(oRecordSet01.Fields("U_Comments").VALUE)
                            oDS_PS_MM030L.setValue "U_ProcCode", sRow - 1, Trim(oRecordSet01.Fields("U_ProcCode").VALUE)
                            oDS_PS_MM030L.setValue "U_ProcName", sRow - 1, Trim(oRecordSet01.Fields("U_ProcName").VALUE)
                            '이론중량 추가(2012.03.26 송명규)
                            oDS_PS_MM030L.setValue "U_TWeight", sRow - 1, Trim(oRecordSet01.Fields("U_TWeight").VALUE)
                            
                            If Trim(oForm01.Items("BPLId").Specific.VALUE) = "1" And Trim(oForm01.Items("Purchase").Specific.VALUE) = "30" Then
                                '창원 외주 단가
                                sQry = "        Select      T0.U_eCardCod,"
                                sQry = sQry & "             T0.U_ItemCode,"
                                sQry = sQry & "             IsNull(T0.U_Cprice, 0) As U_Cprice,"
                                sQry = sQry & "             T0.U_CtrDate "
                                sQry = sQry & " From        [@PS_PP006H] T0"
                                sQry = sQry & "             Inner Join"
                                sQry = sQry & "             ("
                                sQry = sQry & "                 Select      U_eCardCod,"
                                sQry = sQry & "                             U_ItemCode,"
                                sQry = sQry & "                             U_CpCode,"
                                sQry = sQry & "                             MAX(U_CtrDate) As U_CtrDate"
                                sQry = sQry & "                 From        [@PS_PP006H]"
                                sQry = sQry & "                 Group by    U_eCardCod,"
                                sQry = sQry & "                             U_ItemCode,"
                                sQry = sQry & "                             U_CpCode"
                                sQry = sQry & "             ) T1 "
                                sQry = sQry & "                 On T1.U_eCardCod = T0.U_eCardCod"
                                sQry = sQry & "                 And T1.U_ItemCode = T0.U_ItemCode"
                                sQry = sQry & "                 And T1.U_CpCode = T0.U_CpCode"
                                sQry = sQry & "                 And T1.U_CtrDate = T0.U_CtrDate"
                                sQry = sQry & " Where       T0.U_eCardCod = '" & Trim(oForm01.Items("CardCode").Specific.VALUE) & "' "
                                sQry = sQry & "             And T0.U_ItemCode = '" & Trim(oRecordSet01.Fields("U_ItemCode").VALUE) & "' "
                                sQry = sQry & "             And T0.U_CpCode = '" & Trim(oRecordSet01.Fields("U_ProcCode").VALUE) & "' "
                                
                                Call oRecordset02.DoQuery(sQry)
                                
'                                oMat01.Columns("Price").Cells(sRow).Specific.Value = Trim(oRecordset02.Fields("U_Cprice").Value)
                                oDS_PS_MM030L.setValue "U_Price", sRow - 1, Trim(oRecordset02.Fields("U_Cprice").VALUE)
                                oDS_PS_MM030L.setValue "U_LinTotal", sRow - 1, Trim(oRecordset02.Fields("U_Cprice").VALUE) * Trim(oRecordSet01.Fields("U_Weight").VALUE)
                            End If
                            
                            '//구매실적단가 가져오기
                            sQry = "    EXEC PS_MM030_03 '"
                            sQry = sQry & Trim(oForm01.Items("BPLId").Specific.VALUE) & "','"
                            sQry = sQry & Trim(oRecordSet01.Fields("U_ItemCode").VALUE) & "','"
                            sQry = sQry & Trim(oRecordSet01.Fields("U_ProcCode").VALUE) & "','"
                            sQry = sQry & Trim(oForm01.Items("DocNum").Specific.VALUE) & "','"
                            sQry = sQry & Trim(oForm01.Items("DocDate").Specific.VALUE) & "'"
                            
                            Call oRecordSet03.DoQuery(sQry)
                            
                            oDS_PS_MM030L.setValue "U_LCardCd", sRow - 1, Trim(oRecordSet03.Fields("CardCode").VALUE)
                            oDS_PS_MM030L.setValue "U_LCardNm", sRow - 1, Trim(oRecordSet03.Fields("U_CardName").VALUE)
                            
                            If Format(Trim(oRecordSet03.Fields("U_DocDate").VALUE), "YYYYMMDD") < "19000102" Then
                                oDS_PS_MM030L.setValue "U_LDocDate", sRow - 1, ""
                            Else
                                oDS_PS_MM030L.setValue "U_LDocDate", sRow - 1, Format(Trim(oRecordSet03.Fields("U_DocDate").VALUE), "YYYYMMDD")
                            End If
                            oDS_PS_MM030L.setValue "U_LPrice", sRow - 1, Trim(oRecordSet03.Fields("U_Price").VALUE)
                            '//구매실적단가 가져오기 끝
                            
                            '외주사유관련 내용 시작(2013.08.13 송명규 수정)
                            If Trim(oRecordSet01.Fields("U_OutCode").VALUE) <> "" Then
                                oDS_PS_MM030L.setValue "U_OutCode", sRow - 1, Trim(oRecordSet01.Fields("U_OutCode").VALUE) '외주사유코드
                            End If
                            oDS_PS_MM030L.setValue "U_OutNote", sRow - 1, Trim(oRecordSet01.Fields("U_OutNote").VALUE) '외주사유내용
                            oDS_PS_MM030L.setValue "U_MComment", sRow - 1, Trim(oRecordSet01.Fields("U_MComment").VALUE) '자재담당의견
                            '외주사유관련 내용 종료(2013.08.13 송명규 수정)
                            
                            '환율 세팅 시작(2016.06.22 송명규 수정)
                            oForm01.Items("DocRate").Specific.VALUE = Trim(oRecordSet01.Fields("DocRate").VALUE) '환율
                            '환율 세팅 종료(2016.06.22 송명규 수정)
                            
                            Call Add_MatrixRow(sRow, False)
                            sRow = sRow + 1
                        End If
                        oRecordSet01.MoveNext
                    Loop
           
                    If oMat01.VisualRowCount > 0 Then
                        If Trim(oDS_PS_MM030L.GetValue("U_ItemCode", oMat01.VisualRowCount - 1)) = "" Then
                            oDS_PS_MM030L.RemoveRecord oMat01.VisualRowCount - 1
                        End If
                    End If
                    
                    oMat01.LoadFromDataSource
                    
'                    '//최종구매정보
'                    For i = 0 To oMat01.VisualRowCount - 2
''                        sQry = "Select IsNull(a.U_CardName, '') As U_CardName, MAX(a.U_DocDate) As U_DocDate, IsNull(b.U_Price, 0) As U_Price "
''                        sQry = sQry & "From [@PS_MM030H] a Inner Join [@PS_MM030L] b On a.DocEntry = b.DocEntry "
''                        sQry = sQry & "Where b.U_ItemCode = '" & Trim(oMat01.Columns("ItemCode").Cells(i + 1).Specific.Value) & "' "
''                        sQry = sQry & "And a.U_BPLId = '" & Trim(oForm01.Items("BPLId").Specific.Value) & "' "
''                        sQry = sQry & "And a.Status = 'O' "
''                        sQry = sQry & "Group by IsNull(a.U_CardName, ''), IsNull(b.U_Price, 0)"
''                        sQry = "EXEC PS_MM030_03 '" & Trim(oForm01.Items("BPLId").Specific.VALUE) & "', '" & Trim(oMat01.Columns("ItemCode").Cells(i + 1).Specific.VALUE) & "', '" & Trim(oMat01.Columns("ProcCode").Cells(i + 1).Specific.VALUE) & "', '" & Trim(oForm01.Items("DocNum").Specific.VALUE) & "', '" & Trim(oForm01.Items("DocDate").Specific.VALUE) & "'"
''                        oRecordset02.DoQuery sQry
''                        oMat01.Columns("LCardCd").Cells(i + 1).Specific.VALUE = Trim(oRecordset02.Fields("CardCode").VALUE)
''                        oMat01.Columns("LCardNm").Cells(i + 1).Specific.VALUE = Trim(oRecordset02.Fields("U_CardName").VALUE)
''
''                        'oMat01.Columns("LastCardCd").Cells(i + 1).Specific.VALUE = Trim(oRecordset02.Fields("U_CardName").VALUE)
''                        If Format(Trim(oRecordset02.Fields("U_DocDate").VALUE), "YYYYMMDD") < "19000102" Then
''                            'oMat01.Columns("LastDocDt").Cells(i + 1).Specific.VALUE = ""
''                            oMat01.Columns("LDocDate").Cells(i + 1).Specific.VALUE = ""
''                        Else
''                            'oMat01.Columns("LastDocDt").Cells(i + 1).Specific.VALUE = Format(Trim(oRecordset02.Fields("U_DocDate").VALUE), "YYYYMMDD")
''                            oMat01.Columns("LDocDate").Cells(i + 1).Specific.VALUE = Format(Trim(oRecordset02.Fields("U_DocDate").VALUE), "YYYYMMDD")
''                        End If
''                        'oMat01.Columns("LastPrice").Cells(i + 1).Specific.VALUE = Trim(oRecordset02.Fields("U_Price").VALUE)
''                        oMat01.Columns("LPrice").Cells(i + 1).Specific.VALUE = Trim(oRecordset02.Fields("U_Price").VALUE)
'
'                        DocTotal = DocTotal + oMat01.Columns("LinTotal").Cells(i + 1).Specific.VALUE
'                        If oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE = "" Then
'                            SumQty = SumQty
'                        Else
'                            SumQty = SumQty + oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE
'                        End If
'                        SumWeight = SumWeight + oMat01.Columns("Weight").Cells(i + 1).Specific.VALUE
'
'                    Next i
'                    oForm01.Items("DocTotal").Specific.VALUE = DocTotal
'                    oForm01.Items("SumQty").Specific.VALUE = SumQty
'                    oForm01.Items("SumWeight").Specific.VALUE = SumWeight
                    
                    Call PS_MM030_TotalAmount_Calculate
                    
                End If
                oForm01.Freeze False
            ElseIf oCol = "ItemCode" Then
                oForm01.Freeze True
                If oForm01.Mode = fm_ADD_MODE Then
                    If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 0) And Trim(oMat01.Columns("ItemCode").Cells(oRow).Specific.VALUE) <> "" Then
                        oMat01.FlushToDataSource
                        Add_MatrixRow oMat01.RowCount, False
                    End If
                End If
                
                sQry = "Select ItemName, FrgnName From OITM Where ItemCode = '" & Trim(oMat01.Columns("ItemCode").Cells(oRow).Specific.VALUE) & "'"
                oRecordSet01.DoQuery sQry
                oMat01.Columns("ItemName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                 
                '//최종구매정보
'                For i = 0 To oMat01.VisualRowCount - 2
''                    sQry = "Select IsNull(a.U_CardName, '') As U_CardName, MAX(a.U_DocDate) As U_DocDate, IsNull(b.U_Price, 0) As U_Price "
''                    sQry = sQry & "From [@PS_MM030H] a Inner Join [@PS_MM030L] b On a.DocEntry = b.DocEntry "
''                    sQry = sQry & "Where b.U_ItemCode = '" & Trim(oMat01.Columns("ItemCode").Cells(i + 1).Specific.Value) & "' "
''                    sQry = sQry & "And a.U_BPLId = '" & Trim(oForm01.Items("BPLId").Specific.Value) & "' "
''                    sQry = sQry & "And a.Status = 'O' "
''                    sQry = sQry & "Group by IsNull(a.U_CardName, ''), IsNull(b.U_Price, 0)"
'                    sQry = "EXEC PS_MM030_03 '" & Trim(oForm01.Items("BPLId").Specific.VALUE) & "', '" & Trim(oMat01.Columns("ItemCode").Cells(i + 1).Specific.VALUE) & "', '" & Trim(oMat01.Columns("ProcCode").Cells(i + 1).Specific.VALUE) & "', '" & Trim(oForm01.Items("DocNum").Specific.VALUE) & "', '" & Trim(oForm01.Items("DocDate").Specific.VALUE) & "'"
'                    oRecordset02.DoQuery sQry
'
'                    oMat01.Columns("LastCardCd").Cells(i + 1).Specific.VALUE = Trim(oRecordset02.Fields("U_CardName").VALUE)
'                    If Format(Trim(oRecordset02.Fields("U_DocDate").VALUE), "YYYYMMDD") < "19000102" Then
'                        oMat01.Columns("LastDocDt").Cells(i + 1).Specific.VALUE = ""
'                    Else
'                        oMat01.Columns("LastDocDt").Cells(i + 1).Specific.VALUE = Format(Trim(oRecordset02.Fields("U_DocDate").VALUE), "YYYYMMDD")
'                    End If
'                    oMat01.Columns("LastPrice").Cells(i + 1).Specific.VALUE = Trim(oRecordset02.Fields("U_Price").VALUE)
'                Next i
                
                oMat01.Columns("ItemCode").Cells(oRow).Click ct_Regular
                oForm01.Freeze False
            ElseIf oCol = "Qty" Then
                oForm01.Freeze True
                oMat01.FlushToDataSource
                If Trim(oMat01.Columns("Price").Cells(oRow).Specific.VALUE) = 0 Then
                    oDS_PS_MM030L.setValue "U_LinTotal", oRow - 1, 0
                Else
                    oDS_PS_MM030L.setValue "U_LinTotal", oRow - 1, Round(Trim(oMat01.Columns("Weight").Cells(oRow).Specific.VALUE) * Trim(oMat01.Columns("Price").Cells(oRow).Specific.VALUE), 0)
                End If
                oMat01.LoadFromDataSource
                
'                '//최종구매정보
'                For i = 0 To oMat01.VisualRowCount - 2
'''                    sQry = "Select IsNull(a.U_CardName, '') As U_CardName, MAX(a.U_DocDate) As U_DocDate, IsNull(b.U_Price, 0) As U_Price "
'''                    sQry = sQry & "From [@PS_MM030H] a Inner Join [@PS_MM030L] b On a.DocEntry = b.DocEntry "
'''                    sQry = sQry & "Where b.U_ItemCode = '" & Trim(oMat01.Columns("ItemCode").Cells(i + 1).Specific.Value) & "' "
'''                    sQry = sQry & "And a.U_BPLId = '" & Trim(oForm01.Items("BPLId").Specific.Value) & "' "
'''                    sQry = sQry & "And a.Status = 'O' "
'''                    sQry = sQry & "Group by IsNull(a.U_CardName, ''), IsNull(b.U_Price, 0)"
''                    sQry = "EXEC PS_MM030_03 '" & Trim(oForm01.Items("BPLId").Specific.VALUE) & "', '" & Trim(oMat01.Columns("ItemCode").Cells(i + 1).Specific.VALUE) & "', '" & Trim(oMat01.Columns("ProcCode").Cells(i + 1).Specific.VALUE) & "', '" & Trim(oForm01.Items("DocNum").Specific.VALUE) & "', '" & Trim(oForm01.Items("DocDate").Specific.VALUE) & "'"
''                    oRecordset02.DoQuery sQry
''
''                    oMat01.Columns("LastCardCd").Cells(i + 1).Specific.VALUE = Trim(oRecordset02.Fields("U_CardName").VALUE)
''                    If Format(Trim(oRecordset02.Fields("U_DocDate").VALUE), "YYYYMMDD") < "19000102" Then
''                        oMat01.Columns("LastDocDt").Cells(i + 1).Specific.VALUE = ""
''                    Else
''                        oMat01.Columns("LastDocDt").Cells(i + 1).Specific.VALUE = Format(Trim(oRecordset02.Fields("U_DocDate").VALUE), "YYYYMMDD")
''                    End If
''                    oMat01.Columns("LastPrice").Cells(i + 1).Specific.VALUE = Trim(oRecordset02.Fields("U_Price").VALUE)
'
'                    DocTotal = DocTotal + oMat01.Columns("LinTotal").Cells(i + 1).Specific.VALUE
'                    If oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE = "" Then
'                        SumQty = SumQty
'                    Else
'                        SumQty = SumQty + oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE
'                    End If
'                    SumWeight = SumWeight + oMat01.Columns("Weight").Cells(i + 1).Specific.VALUE
'                Next i
'                oForm01.Items("DocTotal").Specific.VALUE = DocTotal
'                oForm01.Items("SumQty").Specific.VALUE = SumQty
'                oForm01.Items("SumWeight").Specific.VALUE = SumWeight
                
                Call PS_MM030_TotalAmount_Calculate
                
                oMat01.Columns("UnWeight").Cells(oRow).Click
                oForm01.Freeze False
                
            ElseIf oCol = "Price" Then
            
                Dim LineTotal As Double
                Dim TAmt As Double
                Dim FCPrice As Double
                Dim FCAmount As Double
                Dim DocRate As Double
                            
                oForm01.Freeze True
                
                oMat01.FlushToDataSource
                If Trim(oMat01.Columns("Price").Cells(oRow).Specific.VALUE) = 0 Then

                    Call oDS_PS_MM030L.setValue("U_LinTotal", oRow - 1, 0)

                Else

                    '품목분류에 따라 이론중량에 따른 이론금액을 계산(2012.03.26 송명규 추가)
                    sQry = "SELECT U_ItmMsort FROM OITM WHERE ItemCode = '" & Trim(oMat01.Columns("ItemCode").Cells(oRow).Specific.VALUE) & "'"
                    Call oRecordSet01.DoQuery(sQry)

                    If oRecordSet01.Fields("U_ItmMsort").VALUE = "30603" Then '품목의 분류가 원재_소재류 중 봉인 경우
                        LineTotal = Round(Trim(oMat01.Columns("TWeight").Cells(oRow).Specific.VALUE) * Trim(oMat01.Columns("Price").Cells(oRow).Specific.VALUE), 0)
                        TAmt = Round(Trim(oMat01.Columns("TWeight").Cells(oRow).Specific.VALUE) * Trim(oMat01.Columns("Price").Cells(oRow).Specific.VALUE), 0)
                    Else '품목의 분류가 원재_소재류 중 봉을 제외한 경우
                        LineTotal = Round(Trim(oMat01.Columns("Weight").Cells(oRow).Specific.VALUE) * Trim(oMat01.Columns("Price").Cells(oRow).Specific.VALUE), 0)
                        TAmt = Round(Trim(oMat01.Columns("Weight").Cells(oRow).Specific.VALUE) * Trim(oMat01.Columns("Price").Cells(oRow).Specific.VALUE), 0)
                    End If

                    '외화단가 계산
                    If oForm01.Items("DocCur").Specific.VALUE = "JPY" Then '일본 엔화는 100원 기준으로 관리하기 때문에 100으로 나눔
                        DocRate = oForm01.Items("DocRate").Specific.VALUE / 100
                    Else
                        DocRate = oForm01.Items("DocRate").Specific.VALUE
                    End If
                    
                    If oForm01.Items("DocCur").Specific.VALUE = "KRW" Then
                        FCPrice = 0
                        FCAmount = 0
                    Else
                        FCPrice = Trim(oMat01.Columns("Price").Cells(oRow).Specific.VALUE) * DocRate
                        FCAmount = Trim(oMat01.Columns("Weight").Cells(oRow).Specific.VALUE) * FCPrice
                    End If
                    
                    Call oDS_PS_MM030L.setValue("U_LinTotal", oRow - 1, LineTotal) '금액
                    Call oDS_PS_MM030L.setValue("U_TAmt", oRow - 1, LineTotal) '이론금액
                    Call oDS_PS_MM030L.setValue("U_FCPrice", oRow - 1, FCPrice) '외화단가
                    Call oDS_PS_MM030L.setValue("U_FCAmount", oRow - 1, FCAmount) '외화금액
'                    oMat01.Columns("LinTotal").Cells(oRow).Specific.VALUE = LineTotal

                End If
                oMat01.LoadFromDataSource
                
'                '//최종구매정보
'                For i = 0 To oMat01.VisualRowCount - 2
''                    sQry = "Select IsNull(a.U_CardName, '') As U_CardName, MAX(a.U_DocDate) As U_DocDate, IsNull(b.U_Price, 0) As U_Price "
''                    sQry = sQry & "From [@PS_MM030H] a Inner Join [@PS_MM030L] b On a.DocEntry = b.DocEntry "
''                    sQry = sQry & "Where b.U_ItemCode = '" & Trim(oMat01.Columns("ItemCode").Cells(i + 1).Specific.Value) & "' "
''                    sQry = sQry & "And a.U_BPLId = '" & Trim(oForm01.Items("BPLId").Specific.Value) & "' "
''                    sQry = sQry & "And a.Status = 'O' "
''                    sQry = sQry & "Group by IsNull(a.U_CardName, ''), IsNull(b.U_Price, 0)"
'
'                    '//속도 문제 때문에 단가 변경 시에는 최종구매정보를 보여주지 않게 주석처리(2012.06.25 송명규)
''                    sQry = "EXEC PS_MM030_03 '" & Trim(oForm01.Items("BPLId").Specific.VALUE) & "', '" & Trim(oMat01.Columns("ItemCode").Cells(i + 1).Specific.VALUE) & "', '" & Trim(oMat01.Columns("ProcCode").Cells(i + 1).Specific.VALUE) & "', '" & Trim(oForm01.Items("DocNum").Specific.VALUE) & "', '" & Trim(oForm01.Items("DocDate").Specific.VALUE) & "'"
''                    oRecordset02.DoQuery sQry
''
''                    oMat01.Columns("LastCardCd").Cells(i + 1).Specific.VALUE = Trim(oRecordset02.Fields("U_CardName").VALUE)
''                    If Format(Trim(oRecordset02.Fields("U_DocDate").VALUE), "YYYYMMDD") < "19000102" Then
''                        oMat01.Columns("LastDocDt").Cells(i + 1).Specific.VALUE = ""
''                    Else
''                        oMat01.Columns("LastDocDt").Cells(i + 1).Specific.VALUE = Format(Trim(oRecordset02.Fields("U_DocDate").VALUE), "YYYYMMDD")
''                    End If
''                    oMat01.Columns("LastPrice").Cells(i + 1).Specific.VALUE = Trim(oRecordset02.Fields("U_Price").VALUE)
'                    '//속도 문제 때문에 단가 변경 시에는 최종구매정보를 보여주지 않게 주석처리(2012.06.25 송명규)
'
'                    DocTotal = DocTotal + oMat01.Columns("LinTotal").Cells(i + 1).Specific.VALUE
'                    If oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE = "" Then
'                        SumQty = SumQty
'                    Else
'                        SumQty = SumQty + oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE
'                    End If
'                    SumWeight = SumWeight + oMat01.Columns("Weight").Cells(i + 1).Specific.VALUE
'                Next i
'
'                oForm01.Items("DocTotal").Specific.VALUE = DocTotal
'                oForm01.Items("SumQty").Specific.VALUE = SumQty
'                oForm01.Items("SumWeight").Specific.VALUE = SumWeight
                
                Call PS_MM030_TotalAmount_Calculate
                
                oMat01.Columns("Price").Cells(oRow).Click
                oForm01.Freeze False
            ElseIf oCol = "LinTotal" Then
                oForm01.Freeze True
'                oMat01.FlushToDataSource
'                If Trim(oMat01.Columns("LinTotal").Cells(oRow).Specific.Value) = 0 Then
'                    oDS_PS_MM030L.setValue "U_Price", oRow - 1, 0
'                Else
'                    If Trim(oMat01.Columns("Weight").Cells(oRow).Specific.Value) = 0 Then
'                        oDS_PS_MM030L.setValue "U_Price", oRow - 1, 0
'                    Else
'                        oDS_PS_MM030L.setValue "U_Price", oRow - 1, Trim(oMat01.Columns("LinTotal").Cells(oRow).Specific.Value) / Trim(oMat01.Columns("Weight").Cells(oRow).Specific.Value)
'                    End If
'                End If
'                oMat01.LoadFromDataSource
                
'                '//최종구매정보
'                For i = 0 To oMat01.VisualRowCount - 2
''                    sQry = "EXEC PS_MM030_03 '" & Trim(oForm01.Items("BPLId").Specific.VALUE) & "', '" & Trim(oMat01.Columns("ItemCode").Cells(i + 1).Specific.VALUE) & "', '" & Trim(oMat01.Columns("ProcCode").Cells(i + 1).Specific.VALUE) & "', '" & Trim(oForm01.Items("DocNum").Specific.VALUE) & "', '" & Trim(oForm01.Items("DocDate").Specific.VALUE) & "'"
''                    oRecordset02.DoQuery sQry
''
''                    oMat01.Columns("LastCardCd").Cells(i + 1).Specific.VALUE = Trim(oRecordset02.Fields("U_CardName").VALUE)
''                    If Format(Trim(oRecordset02.Fields("U_DocDate").VALUE), "YYYYMMDD") < "19000102" Then
''                        oMat01.Columns("LastDocDt").Cells(i + 1).Specific.VALUE = ""
''                    Else
''                        oMat01.Columns("LastDocDt").Cells(i + 1).Specific.VALUE = Format(Trim(oRecordset02.Fields("U_DocDate").VALUE), "YYYYMMDD")
''                    End If
''                    oMat01.Columns("LastPrice").Cells(i + 1).Specific.VALUE = Trim(oRecordset02.Fields("U_Price").VALUE)
'
'                    DocTotal = DocTotal + oMat01.Columns("LinTotal").Cells(i + 1).Specific.VALUE
'                    If oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE = "" Then
'                        SumQty = SumQty
'                    Else
'                        SumQty = SumQty + oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE
'                    End If
'                    SumWeight = SumWeight + oMat01.Columns("Weight").Cells(i + 1).Specific.VALUE
'                Next i
'                oForm01.Items("DocTotal").Specific.VALUE = DocTotal
'                oForm01.Items("SumQty").Specific.VALUE = SumQty
'                oForm01.Items("SumWeight").Specific.VALUE = SumWeight
                
                Call PS_MM030_TotalAmount_Calculate
                
                oMat01.Columns("Price").Cells(oRow).Click
                oForm01.Freeze False
            ElseIf oCol = "WhsCode" Then
                sQry = "Select WhsName From [OWHS] Where WhsCode = '" & Trim(oMat01.Columns("WhsCode").Cells(oRow).Specific.VALUE) & "'"
                oRecordSet01.DoQuery sQry
            
                oMat01.Columns("WhsName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            End If
            
            Call oMat01.AutoResizeColumns
            
    End Select

    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FlushToItemValue_Error:
    Set oRecordSet01 = Nothing
    oForm01.Freeze False
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "구매견적문서가 취소되었거나 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
    Else
        MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error

    Dim ErrNum As Integer
    Dim DocNum As String

    ErrNum = 0

    '// Check
    Select Case True
        Case oDS_PS_MM030H.GetValue("U_CardCode", 0) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_MM030H.GetValue("U_BPLId", 0) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_MM030H.GetValue("U_CntcCode", 0) = ""
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_MM030H.GetValue("U_POType", 0) = ""
            ErrNum = 4
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_MM030H.GetValue("U_Purchase", 0) = ""
            ErrNum = 5
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_MM030H.GetValue("U_DocDate", 0) = ""
            ErrNum = 6
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_MM030H.GetValue("U_DueDate", 0) = ""
            ErrNum = 7
            GoTo HeaderSpaceLineDel_Error
    End Select

'    '마감상태 체크_S(2017.11.23 송명규 추가)
'    If MDC_PS_Common.Check_Finish_Status(oForm01.Items("BPLId").Specific.VALUE, oForm01.Items("DocDate").Specific.VALUE) = False Then
'        ErrNum = 8
'        GoTo HeaderSpaceLineDel_Error
'    End If
'    '마감상태 체크_E(2017.11.23 송명규 추가)

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "거래처는 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "사업장은 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "담당자는 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "품의형태는 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "구매방식은 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "전기일은 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 7 Then
        MDC_Com.MDC_GF_Message "납품일은 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 8 Then
        MDC_Com.MDC_GF_Message "마감상태가 잠금입니다. 해당 일자로 등록할 수 없습니다. 전기일을 확인하고, 회계부서로 문의하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function

Private Function MatrixSpaceLineDel() As Boolean
On Error GoTo MatrixSpaceLineDel_Error

    Dim i               As Long
    Dim ErrNum          As Integer
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String

    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    ErrNum = 0
    
    oMat01.FlushToDataSource

    '// 라인
    If oMat01.VisualRowCount = 0 Then
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    End If
    
    For i = 0 To oMat01.VisualRowCount - 2
        Select Case True
            Case oDS_PS_MM030L.GetValue("U_ItemCode", i) = ""
                ErrNum = 2
                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_MM030L.GetValue("U_Qty", i) = "" Or oDS_PS_MM030L.GetValue("U_Qty", i) = 0
'                ErrNum = 3
'                GoTo MatrixSpaceLineDel_Error
            Case oDS_PS_MM030L.GetValue("U_Weight", i) = ""
                ErrNum = 4
                GoTo MatrixSpaceLineDel_Error
            Case oDS_PS_MM030L.GetValue("U_Price", i) = 0
                ErrNum = 5
                GoTo MatrixSpaceLineDel_Error
            Case oDS_PS_MM030L.GetValue("U_LinTotal", i) = 0
                ErrNum = 6
                GoTo MatrixSpaceLineDel_Error
'            Case Check_Purchase_Type(oDS_PS_MM030L.GetValue("U_PQDocNum", i), oForm01.Items("Purchase").Specific.Selected.VALUE) = False '구매견적의 품의구분과 불일치할 때는 입력 불가(2012.05.01 송명규 추가), OverFlow 오류 확인 필요(2016.04.01 송명규)
'                ErrNum = 7
'                GoTo MatrixSpaceLineDel_Error
            Case oDS_PS_MM030L.GetValue("U_WhsCode", i) = ""
                ErrNum = 8
                GoTo MatrixSpaceLineDel_Error
                
            Case PS_MM030_CheckDate(oMat01.Columns("PQDocNum").Cells(i + 1).Specific.VALUE) = False '구매견적과 일자 체크
        
                ErrNum = 9
                GoTo MatrixSpaceLineDel_Error
        
        End Select
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
        MDC_Com.MDC_GF_Message "품목코드는 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "수량은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "중량은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "단가는 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "금액은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 7 Then
        MDC_Com.MDC_GF_Message i + 1 & "행의 견적과 품의구분이 일치하지 않습니다. 확인하세요.", "E"
    ElseIf ErrNum = 8 Then
        MDC_Com.MDC_GF_Message "창고코드는 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 9 Then
        Call Sbo_Application.MessageBox(i + 1 & "행 [" & oMat01.Columns("ItemCode").Cells(i + 1).Specific.VALUE & "]의 구매품의일은 구매견적일과 같거나 늦어야합니다. 확인하십시오." & Chr(13) & "해당 품의는 전체가 등록되지 않습니다.", 1)
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function

Sub Delete_EmptyRow()
On Error GoTo Delete_EmptyRow_Error

    Dim i As Integer
    
    oMat01.FlushToDataSource
    
    For i = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PS_MM030L.GetValue("U_ItemCode", i)) = "" Then
            oDS_PS_MM030L.RemoveRecord i   '// Mat01에 마지막라인(빈라인) 삭제
        End If
    Next i
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Delete_EmptyRow_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Function PS_MM030_oPurchaseOrders_Add(ChkType As Integer) As Boolean
On Error GoTo PS_MM030_oPurchaseOrders_Add_Error

    Dim DI_oPurchaseOrders As SAPbobsCOM.Documents         '//구매오더 문서객체
        
    Dim i           As Integer
    Dim ErrNum      As Integer
    Dim ErrCode     As Long
    Dim ErrMsg      As String
    Dim RetVal      As Integer
    Dim DocDate     As String
    Dim DueDate     As String
    Dim DocEntry    As String
    Dim VATGroup    As String
    Dim sQry        As String
        
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim oRecordset02 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Set oRecordset02 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    oMat01.FlushToDataSource
    
    DocDate = Format(oDS_PS_MM030H.GetValue("U_DocDate", 0), "0000-00-00")
    DueDate = Format(oDS_PS_MM030H.GetValue("U_DueDate", 0), "0000-00-00")
    
    '지불조건명 조회 끝
    
    If Trim(oDS_PS_MM030H.GetValue("U_PODocNum", 0)) = "" Then
        Sbo_Company.StartTransaction
        Set DI_oPurchaseOrders = Nothing
        Set DI_oPurchaseOrders = Sbo_Company.GetBusinessObject(oPurchaseOrders)
        
        With DI_oPurchaseOrders
            .CardCode = Trim(oDS_PS_MM030H.GetValue("U_CardCode", 0))
            .BPL_IDAssignedToInvoice = Trim(oDS_PS_MM030H.GetValue("U_BPLId", 0))
'            .ContactPersonCode = Trim(oDS_PS_MM030H.getValue("U_CntcCode", 0))
            .DocDate = DocDate
            .DocDueDate = DueDate
            .Comments = oDS_PS_MM030H.GetValue("U_Comments", 0)
            .UserFields("U_reType").VALUE = Trim(oDS_PS_MM030H.GetValue("U_POType", 0))
            .UserFields("U_okYN").VALUE = Trim(oDS_PS_MM030H.GetValue("U_POStatus", 0))
            .UserFields("U_OrdTyp").VALUE = Trim(oDS_PS_MM030H.GetValue("U_Purchase", 0))
           
            
            sQry = "Select ECVatGroup From [OCRD] Where CardCode = '" & Trim(oDS_PS_MM030H.GetValue("U_CardCode", 0)) & "'"
            oRecordSet01.DoQuery sQry
            
            If Trim(oDS_PS_MM030H.GetValue("U_Purchase", 0)) = "30" Or Trim(oDS_PS_MM030H.GetValue("U_Purchase", 0)) = "40" Or Trim(oDS_PS_MM030H.GetValue("U_Purchase", 0)) = "60" Then
                .DocType = dDocument_Service
                For i = 0 To oMat01.VisualRowCount - 2
                    If i > 0 Then .Lines.Add
                    .Lines.SetCurrentLine i
                    
                    .Lines.ItemDescription = Trim(oDS_PS_MM030L.GetValue("U_ItemName", i)) + "-" + Trim(oDS_PS_MM030L.GetValue("U_OutSize", i)) + "-" + Trim(oDS_PS_MM030L.GetValue("U_OutUnit", i))
                    .Lines.LineTotal = Trim(oDS_PS_MM030L.GetValue("U_LinTotal", i))
'                    .Lines.WarehouseCode = Trim(oDS_PS_MM030L.GetValue("U_WhsCode", i))
                    .Lines.VATGroup = oRecordSet01.Fields("ECVatGroup").VALUE
'                    .Lines.AccountCode
    '                .Lines.UserFields("PQDocNum").Value = oDS_PS_MM030H.getValue("U_PQDocNum", 0)
                    .Lines.UserFields("U_sItemCode").VALUE = Trim(oDS_PS_MM030L.GetValue("U_ItemCode", i))
                    .Lines.UserFields("U_sItemName").VALUE = Trim(oDS_PS_MM030L.GetValue("U_ItemName", i))
                    .Lines.UserFields("U_sSize").VALUE = Trim(oDS_PS_MM030L.GetValue("U_OutSize", i))
                    .Lines.UserFields("U_sUnit").VALUE = Trim(oDS_PS_MM030L.GetValue("U_OutUnit", i))
                    If Trim(oDS_PS_MM030L.GetValue("U_Qty", i)) <> "" Or Trim(oDS_PS_MM030L.GetValue("U_Qty", i)) <> 0 Then
                        .Lines.UserFields("U_sQty").VALUE = Trim(oDS_PS_MM030L.GetValue("U_Qty", i))
                    End If
                    .Lines.UserFields("U_sWeight").VALUE = Trim(oDS_PS_MM030L.GetValue("U_Weight", i))
                    
                    .Lines.UserFields("U_MM010Doc").VALUE = Trim(oDS_PS_MM030L.GetValue("U_PQDocNum", i)) '구매견적문서번호(2017.04.13 송명규 추가)
                    .Lines.UserFields("U_MM010Lin").VALUE = Trim(oDS_PS_MM030L.GetValue("U_PQLinNum", i)) '구매견적라인번호(2017.04.13 송명규 추가)
                    
                    '//작번을 입력
                    'If Trim(oDS_PS_MM030H.GetValue("U_Purchase", 0)) = "10" Then
                     sQry = "EXEC PS_MM030_04 '" & Trim(oDS_PS_MM030L.GetValue("U_PQDocNum", i)) & "', '" & Trim(oDS_PS_MM030L.GetValue("U_PQLinNum", i)) & "'"
                     oRecordset02.DoQuery sQry

                    '//작번
                    .Lines.UserFields("U_OrdNum").VALUE = Trim(oRecordset02.Fields(0).VALUE)
                    .Lines.UserFields("U_OrdSub1").VALUE = Trim(oRecordset02.Fields(1).VALUE)
                    .Lines.UserFields("U_OrdSub2").VALUE = Trim(oRecordset02.Fields(2).VALUE)
                    .Lines.UserFields("U_Payment").VALUE = Trim(oDS_PS_MM030H.GetValue("U_Payment", 0))
                Next i
            Else
                .DocType = dDocument_Items
                For i = 0 To oMat01.VisualRowCount - 2
                    If i > 0 Then .Lines.Add
                    .Lines.SetCurrentLine i
                    
    '                .Lines.UserFields("PQDocNum").Value = oDS_PS_MM030H.getValue("U_PQDocNum", 0)
                    .Lines.ItemCode = Trim(oDS_PS_MM030L.GetValue("U_ItemCode", i))
                    .Lines.Quantity = Trim(oDS_PS_MM030L.GetValue("U_Weight", i))
                    .Lines.Price = Trim(oDS_PS_MM030L.GetValue("U_Price", i))
                    .Lines.LineTotal = Trim(oDS_PS_MM030L.GetValue("U_LinTotal", i))
                    .Lines.WarehouseCode = Trim(oDS_PS_MM030L.GetValue("U_WhsCode", i))
                    .Lines.VATGroup = oRecordSet01.Fields("ECVatGroup").VALUE
                    If Trim(oDS_PS_MM030L.GetValue("U_Qty", i)) <> "" Or Trim(oDS_PS_MM030L.GetValue("U_Qty", i)) <> 0 Then
                        .Lines.UserFields("U_Qty").VALUE = Trim(oDS_PS_MM030L.GetValue("U_Qty", i))
                    End If
                    .Lines.UserFields("U_Weight").VALUE = Trim(oDS_PS_MM030L.GetValue("U_Weight", i))
                    
                    .Lines.UserFields("U_MM010Doc").VALUE = Trim(oDS_PS_MM030L.GetValue("U_PQDocNum", i)) '구매견적문서번호(2017.04.13 송명규 추가)
                    .Lines.UserFields("U_MM010Lin").VALUE = Trim(oDS_PS_MM030L.GetValue("U_PQLinNum", i)) '구매견적라인번호(2017.04.13 송명규 추가)
                    
                    'If Trim(oDS_PS_MM030H.GetValue("U_Purchase", 0)) = "10" Then
                     sQry = "EXEC PS_MM030_04 '" & Trim(oDS_PS_MM030L.GetValue("U_PQDocNum", i)) & "', '" & Trim(oDS_PS_MM030L.GetValue("U_PQLinNum", i)) & "'"
                     oRecordset02.DoQuery sQry

                    '//작번
                    .Lines.UserFields("U_OrdNum").VALUE = Trim(oRecordset02.Fields(0).VALUE)
                    .Lines.UserFields("U_OrdSub1").VALUE = Trim(oRecordset02.Fields(1).VALUE)
                    .Lines.UserFields("U_OrdSub2").VALUE = Trim(oRecordset02.Fields(2).VALUE)
                    .Lines.UserFields("U_Payment").VALUE = Trim(oDS_PS_MM030H.GetValue("U_Payment", 0))
                Next i
            End If
        End With
        
        '//완료
        RetVal = DI_oPurchaseOrders.Add
        If (0 <> RetVal) Then
            Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
            ErrNum = 1
            GoTo PS_MM030_oPurchaseOrders_Add_Error
        End If
        
        If ChkType = 1 Then
            Sbo_Company.EndTransaction wf_RollBack
        ElseIf ChkType = 2 Then
            Sbo_Company.GetNewObjectCode DocEntry
'            sQry = "Select DocNum From [OIGN] Where DocEntry = '" & sDocEntry & "'"
'            oRecordset.DoQuery sQry
'            sDocNum = Trim(oRecordset.Fields(0).Value)
            oDS_PS_MM030H.setValue "U_PODocNum", 0, DocEntry
                       
            Sbo_Company.EndTransaction wf_Commit
        End If
    End If
    
    Set oRecordSet01 = Nothing
    Set oRecordset02 = Nothing
    Set DI_oPurchaseOrders = Nothing
    PS_MM030_oPurchaseOrders_Add = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_MM030_oPurchaseOrders_Add_Error:
    Set oRecordSet01 = Nothing
    Set oRecordset02 = Nothing
    Set DI_oPurchaseOrders = Nothing
    If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_RollBack
    PS_MM030_oPurchaseOrders_Add = False
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "PS_MM030_oPurchaseOrders_Add_Error:" & ErrCode & " - " & ErrMsg, "E"
    Else
        MDC_Com.MDC_GF_Message "PS_MM030_oPurchaseOrders_Add_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Function

Private Function PS_MM030_oPurchaseOrders_Update() As Boolean
On Error GoTo PS_MM030_oPurchaseOrders_Update_Error

    Dim DI_oPurchaseOrders As SAPbobsCOM.Documents         '//구매오더 문서객체
        
    Dim i        As Integer
    Dim ErrNum   As Integer
    Dim ErrCode  As Long
    Dim ErrMsg   As String
    Dim RetVal   As Integer
    Dim DocDate  As String
    Dim DueDate  As String
    Dim DocEntry As String
    Dim sQry     As String
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
        
    oMat01.FlushToDataSource
    
    DocEntry = Trim(oDS_PS_MM030H.GetValue("U_PODocNum", 0))
    DocDate = Format(oDS_PS_MM030H.GetValue("U_DocDate", 0), "0000-00-00")
    DueDate = Format(oDS_PS_MM030H.GetValue("U_DueDate", 0), "0000-00-00")
    
    If Trim(oDS_PS_MM030H.GetValue("U_PODocNum", 0)) <> "" Then
        
        Set DI_oPurchaseOrders = Nothing
        Set DI_oPurchaseOrders = Sbo_Company.GetBusinessObject(oPurchaseOrders)

        If (DI_oPurchaseOrders.GetByKey(DocEntry) = False) Then
            Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
            GoTo PS_MM030_oPurchaseOrders_Update_Error
        End If
        
        With DI_oPurchaseOrders
            .CardCode = Trim(oDS_PS_MM030H.GetValue("U_CardCode", 0))
            .BPL_IDAssignedToInvoice = Trim(oDS_PS_MM030H.GetValue("U_BPLId", 0))
'            .ContactPersonCode = Trim(oDS_PS_MM030H.getValue("U_CntcCode", 0))
            .DocDate = DocDate
            .DocDueDate = DueDate
            .Comments = oDS_PS_MM030H.GetValue("U_Comments", 0)
            .UserFields("U_reType").VALUE = Trim(oDS_PS_MM030H.GetValue("U_POType", 0))
            .UserFields("U_okYN").VALUE = Trim(oDS_PS_MM030H.GetValue("U_POStatus", 0))
            .UserFields("U_OrdTyp").VALUE = Trim(oDS_PS_MM030H.GetValue("U_Purchase", 0))
            
'            .Lines.Delete
            
            sQry = "Select ECVatGroup From [OCRD] Where CardCode = '" & Trim(oDS_PS_MM030H.GetValue("U_CardCode", 0)) & "'"
            oRecordSet01.DoQuery sQry
            
            If Trim(oDS_PS_MM030H.GetValue("U_Purchase", 0)) = "30" Or Trim(oDS_PS_MM030H.GetValue("U_Purchase", 0)) = "40" Or Trim(oDS_PS_MM030H.GetValue("U_Purchase", 0)) = "60" Then
                .DocType = dDocument_Service
                For i = 0 To oMat01.VisualRowCount - 2
                    If i > 0 Then .Lines.Add
                    .Lines.SetCurrentLine i
                    
                    .Lines.ItemDescription = Trim(oDS_PS_MM030L.GetValue("U_ItemName", i)) + "-" + Trim(oDS_PS_MM030L.GetValue("U_OutSize", i)) + "-" + Trim(oDS_PS_MM030L.GetValue("U_OutUnit", i))
                    .Lines.LineTotal = Trim(oDS_PS_MM030L.GetValue("U_LinTotal", i))
                    .Lines.ShipDate = DueDate
'                    .Lines.WarehouseCode = Trim(oDS_PS_MM030L.GetValue("U_WhsCode", i))
'                    .Lines.VATGroup = oRecordSet01.Fields("ECVatGroup").Value
'                    .Lines.AccountCode
    '                .Lines.UserFields("PQDocNum").Value = oDS_PS_MM030H.getValue("U_PQDocNum", 0)
                    .Lines.UserFields("U_sItemCode").VALUE = Trim(oDS_PS_MM030L.GetValue("U_ItemCode", i))
                    .Lines.UserFields("U_sItemName").VALUE = Trim(oDS_PS_MM030L.GetValue("U_ItemName", i))
                    .Lines.UserFields("U_sSize").VALUE = Trim(oDS_PS_MM030L.GetValue("U_OutSize", i))
                    .Lines.UserFields("U_sUnit").VALUE = Trim(oDS_PS_MM030L.GetValue("U_OutUnit", i))
                    .Lines.UserFields("U_sQty").VALUE = Trim(oDS_PS_MM030L.GetValue("U_Qty", i))
                    .Lines.UserFields("U_sWeight").VALUE = Trim(oDS_PS_MM030L.GetValue("U_Weight", i))
                    .Lines.UserFields("U_Payment").VALUE = Trim(oDS_PS_MM030H.GetValue("U_Payment", 0))
                Next i
            Else
                .DocType = dDocument_Items
'                .Lines.Delete
                For i = 0 To oMat01.VisualRowCount - 2
                    If i > 0 Then .Lines.Add
                    .Lines.SetCurrentLine i
                    
    '                .Lines.UserFields("PQDocNum").Value = oDS_PS_MM030H.getValue("U_PQDocNum", 0)
                    .Lines.ItemCode = Trim(oDS_PS_MM030L.GetValue("U_ItemCode", i))
                    .Lines.Quantity = Trim(oDS_PS_MM030L.GetValue("U_Weight", i))
                    
'                    .Lines.Price = Trim(oDS_PS_MM030L.GetValue("U_Price", i))
                    .Lines.UnitPrice = Trim(oDS_PS_MM030L.GetValue("U_Price", i))
'                    .Lines.PriceAfterVAT = 0 'Trim(oDS_PS_MM030L.GetValue("U_Price", i)) + (Trim(oDS_PS_MM030L.GetValue("U_Price", i)) * 0.1)
                    .Lines.DiscountPercent = 0
                    .Lines.LineTotal = Trim(oDS_PS_MM030L.GetValue("U_LinTotal", i))
                    
                    .Lines.WarehouseCode = Trim(oDS_PS_MM030L.GetValue("U_WhsCode", i))
                    .Lines.VATGroup = Trim(oRecordSet01.Fields("ECVatGroup").VALUE)
                    .Lines.ShipDate = DueDate
                    If Trim(oDS_PS_MM030L.GetValue("U_Qty", i)) <> "" Or Trim(oDS_PS_MM030L.GetValue("U_Qty", i)) <> 0 Then
                        .Lines.UserFields("U_Qty").VALUE = Trim(oDS_PS_MM030L.GetValue("U_Qty", i))
                    End If
                    .Lines.UserFields("U_Weight").VALUE = Trim(oDS_PS_MM030L.GetValue("U_Weight", i))
                    .Lines.UserFields("U_Payment").VALUE = Trim(oDS_PS_MM030H.GetValue("U_Payment", 0))
                Next i
            End If
        End With
        
        '//완료
        RetVal = DI_oPurchaseOrders.Update
        
        If (0 <> RetVal) Then
            Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
            GoTo PS_MM030_oPurchaseOrders_Update_Error
        End If
        
    End If
    
    Set oRecordSet01 = Nothing
    Set DI_oPurchaseOrders = Nothing
    PS_MM030_oPurchaseOrders_Update = True
    
    Exit Function

PS_MM030_oPurchaseOrders_Update_Error:

    Set oRecordSet01 = Nothing
    Set DI_oPurchaseOrders = Nothing
    PS_MM030_oPurchaseOrders_Update = False
    If RetVal <> 0 Then
        Call MDC_Com.MDC_GF_Message("PS_MM030_oPurchaseOrders_Update_Error:" & ErrCode & " - " & ErrMsg, "E")
    Else
        Call MDC_Com.MDC_GF_Message("PS_MM030_oPurchaseOrders_Update_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
    
End Function

Private Function PS_MM030_oPurchaseOrders_Cancel() As Boolean
On Error GoTo PS_MM030_oPurchaseOrders_Cancel_Error

    Dim DI_oPurchaseOrders As SAPbobsCOM.Documents         '//구매오더 문서객체
        
    Dim i        As Integer
    Dim ErrNum   As Integer
    Dim ErrCode  As Long
    Dim ErrMsg   As String
    Dim RetVal   As Integer
    Dim DocEntry As String

    DocEntry = Trim(oDS_PS_MM030H.GetValue("U_PODocNum", 0))
    
    If Trim(oDS_PS_MM030H.GetValue("U_PODocNum", 0)) <> "" Then

        Set DI_oPurchaseOrders = Nothing
        Set DI_oPurchaseOrders = Sbo_Company.GetBusinessObject(oPurchaseOrders)
                
        '//완료
        If (DI_oPurchaseOrders.GetByKey(DocEntry) = False) Then
            Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
            GoTo PS_MM030_oPurchaseOrders_Cancel_Error
        End If
        RetVal = DI_oPurchaseOrders.Cancel
        If (0 <> RetVal) Then
            Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
            ErrNum = 1
            GoTo PS_MM030_oPurchaseOrders_Cancel_Error
        End If

    End If
    
    Set DI_oPurchaseOrders = Nothing
    PS_MM030_oPurchaseOrders_Cancel = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_MM030_oPurchaseOrders_Cancel_Error:

    Set DI_oPurchaseOrders = Nothing
    PS_MM030_oPurchaseOrders_Cancel = False
    
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "PS_MM030_oPurchaseOrders_Cancel_Error:" & ErrCode & " - " & ErrMsg, "E"
    Else
        MDC_Com.MDC_GF_Message "PS_MM030_oPurchaseOrders_Cancel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    
End Function

Private Function PS_MM030_oPurchaseOrders_Close() As Boolean
On Error GoTo PS_MM030_oPurchaseOrders_Close_Error

    Dim DI_oPurchaseOrders As SAPbobsCOM.Documents         '//구매오더 문서객체
    
    Dim i        As Integer
    Dim ErrNum   As Integer
    Dim ErrCode  As Long
    Dim ErrMsg   As String
    Dim RetVal   As Integer
    Dim DocEntry As String
    
    DocEntry = Trim(oDS_PS_MM030H.GetValue("U_PODocNum", 0))
    
    If Trim(oDS_PS_MM030H.GetValue("U_PODocNum", 0)) <> "" Then
        
        Set DI_oPurchaseOrders = Nothing
        Set DI_oPurchaseOrders = Sbo_Company.GetBusinessObject(oPurchaseOrders)
                
        '//완료
        If (DI_oPurchaseOrders.GetByKey(DocEntry) = False) Then
            Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
            GoTo PS_MM030_oPurchaseOrders_Close_Error
        End If
        RetVal = DI_oPurchaseOrders.Close
        If (0 <> RetVal) Then
            Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
            GoTo PS_MM030_oPurchaseOrders_Close_Error
        End If
        
    End If
    
    Set DI_oPurchaseOrders = Nothing
    PS_MM030_oPurchaseOrders_Close = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_MM030_oPurchaseOrders_Close_Error:

    Set DI_oPurchaseOrders = Nothing
    PS_MM030_oPurchaseOrders_Close = False
    MDC_Com.MDC_GF_Message "PS_MM030_oPurchaseOrders_Close_Error:" & Err.Number & " - " & Err.Description, "E"
    
End Function

Private Function PS_MM030_oPurchaseOrders_LineDelete(ByVal pLineNum As Integer) As Boolean
On Error GoTo PS_MM030_oPurchaseOrders_LineDelete_Error

    Dim DI_oPurchaseOrders As SAPbobsCOM.Documents         '//구매오더 문서객체
    Dim i        As Integer
    Dim ErrNum   As Integer
    Dim ErrCode  As Long
    Dim ErrMsg   As String
    Dim RetVal   As Integer
    Dim DocEntry As String
    
    DocEntry = Trim(oDS_PS_MM030H.GetValue("U_PODocNum", 0))
    
    'If Trim(oDS_PS_MM030H.GetValue("U_PODocNum", 0)) <> "" Then
        
        Set DI_oPurchaseOrders = Nothing
        Set DI_oPurchaseOrders = Sbo_Company.GetBusinessObject(oPurchaseOrders)
                
        '//완료
        If (DI_oPurchaseOrders.GetByKey(DocEntry) = False) Then
            Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
            GoTo PS_MM030_oPurchaseOrders_LineDelete_Error
        Else
        
            Call DI_oPurchaseOrders.Lines.SetCurrentLine(pLineNum - 1) 'LineID에서 1을 빼야 구매오더 행번호와 일치
            Call DI_oPurchaseOrders.Lines.Delete
        
        End If
        
        RetVal = DI_oPurchaseOrders.Update
        
        If (0 <> RetVal) Then
            Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
            GoTo PS_MM030_oPurchaseOrders_LineDelete_Error
        End If
        
    'End If
    
    Set DI_oPurchaseOrders = Nothing
    PS_MM030_oPurchaseOrders_LineDelete = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_MM030_oPurchaseOrders_LineDelete_Error:
    Set DI_oPurchaseOrders = Nothing
    If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_RollBack
    PS_MM030_oPurchaseOrders_LineDelete = False
    MDC_Com.MDC_GF_Message "PS_MM030_oPurchaseOrders_LineDelete_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Private Sub Report_Export()
On Error GoTo Report_Export_Error

    Dim ErrNum       As Integer
    Dim DocNum       As String
    Dim WinTitle     As String
    Dim ReportName   As String
    Dim sQry         As String
    Dim BPLID        As String
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    Call ConnectODBC
    
'    If Trim(oForm01.Items("DocNum").Specific.Value) = "" Then
'        ErrNum = 1
'        GoTo Report_Export_Error
'    End If
    
    DocNum = Trim(oForm01.Items("DocNum").Specific.VALUE)
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
        
    WinTitle = "주문서 [PS_MM030_01]"
        
    If BPLID = "2" Then
        ReportName = "PS_MM030_03.rpt" '기계사업부(2015.08.11 추가)
    Else
        ReportName = "PS_MM030_01.rpt"
    End If
       
    '//Formula 수식필드
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    
    '//SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    '//조회조건문
    sQry = "EXEC [PS_MM030_01] '" & DocNum & "'"
'    oRecordSet01.DoQuery sQry
'    If oRecordSet01.RecordCount = 0 Then
'        ErrNum = 2
'        GoTo Report_Export_Error
'    End If
    
    '리포트 Export(2017.05.17 송명규)
    Dim ExportString As String
    If oForm01.Items("POType").Specific.VALUE = "10" Then '자체구매일 때만 Export함
    
        ExportString = "C:\ReportExport\PS_MM030_주문서_" & oForm01.Items("DocNum").Specific.VALUE & ".pdf"
    
    Else
    
        ExportString = ""
        
    End If
    '리포트 Export(2017.05.17 송명규)
    
    'PDF 파일 Export
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "1", "N", "V", ExportString) = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Report_Export_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다.확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "Report_Export_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Private Sub Print_Report01()
'******************************************************************************
'Function ID : Print_Report01()
'해당모듈    : PS_MM030
'기능        : 주문서 출력
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo Print_Report01_Error

    Dim ErrNum       As Integer
    Dim DocNum       As String
    Dim WinTitle     As String
    Dim ReportName   As String
    Dim sQry         As String
    Dim BPLID        As String
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    Call ConnectODBC
    
'    If Trim(oForm01.Items("DocNum").Specific.Value) = "" Then
'        ErrNum = 1
'        GoTo Print_Report01_Error
'    End If
    
    DocNum = Trim(oForm01.Items("DocNum").Specific.VALUE)
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
        
    WinTitle = "주문서 [PS_MM030_01]"
        
    If BPLID = "2" Then
        ReportName = "PS_MM030_03.rpt" '기계사업부(2015.08.11 추가)
    Else
        ReportName = "PS_MM030_01.rpt"
    End If
       
    '//Formula 수식필드
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    
    '//SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    '//조회조건문
    sQry = "EXEC [PS_MM030_01] '" & DocNum & "'"
'    oRecordSet01.DoQuery sQry
'    If oRecordSet01.RecordCount = 0 Then
'        ErrNum = 2
'        GoTo Print_Report01_Error
'    End If
    
    '리포트 미리보기 호출
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "1", "N", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Report01_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다.확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "Print_Report01_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Private Sub Print_Report02()
'******************************************************************************
'Function ID : Print_Report02()
'해당모듈    : PS_MM030
'기능        : 품의서 출력
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo Print_Report02_Error

    Dim ErrNum&
    Dim WinTitle   As String
    Dim ReportName As String
    Dim sQry       As String
    Dim oText(1)   As String
    Dim i          As Integer
    Dim DocNum     As String
    Dim Purchase   As String
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    Call ConnectODBC
    
    DocNum = Trim(oForm01.Items("DocNum").Specific.VALUE)
    Purchase = Trim(oForm01.Items("Purchase").Specific.VALUE)
    
'    If Trim(oForm01.Items("DocNum").Specific.Value) = "" Then
'        ErrNum = 1
'        GoTo Print_Report02_Error
'    End If
        
    WinTitle = "품의서 [PS_MM030_02]"
    ReportName = "PS_MM030_02.rpt"
       
    '//Formula 수식필드
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    
    If Purchase = "10" Then
        oText(1) = "구 매 품 의 서 (원자재)"
    ElseIf Purchase = "20" Then
        oText(1) = "구 매 품 의 서 (부재료)"
    ElseIf Purchase = "30" Then
        oText(1) = "구 매 품 의 서 (외주가공비)"
    ElseIf Purchase = "40" Then
        oText(1) = "구 매 품 의 서 (외주제작)"
    ElseIf Purchase = "50" Then
        oText(1) = "구 매 품 의 서 (상품)"
    ElseIf Purchase = "60" Then
        oText(1) = "구 매 품 의 서 (고정자산)"
    End If
    
    For i = 1 To 1
        If Len("" & i & "") = 1 Then
            gRpt_Formula(i) = "F0" & i & ""
        Else
            gRpt_Formula(i) = "F" & i & ""
        End If
        gRpt_Formula_Value(i) = oText(i)
    Next i
    
    '//SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    '//조회조건문
    sQry = "EXEC [PS_MM030_02] '" & DocNum & "'"
'    oRecordSet01.DoQuery sQry
'    If oRecordSet01.RecordCount = 0 Then
'        ErrNum = 2
'        GoTo Print_Report02_Error
'    End If
    
    '//CR Action
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "1", "N", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Report02_Error:
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다.확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "Print_Report02_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Public Function Check_Purchase_Type(ByVal prmPQDocNum As Integer, ByVal prmPOType As String) As Boolean
'******************************************************************************
'Function ID : Check_Purchase_Type()
'해당모듈    : PS_MM030
'기능        : 품의 등록 시 구매견적과 품의구분이 일치하는 여부 검사
'인수        : prmPQDocNum(견적문서번호), prmPOType(품의구분)
'반환값      : 품의구분의 일치여부(True, False)
'특이사항    : 없음
'******************************************************************************
On Error GoTo Check_Purchase_Type_Error

    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    sQry = "         SELECT     U_Purchase"
    sQry = sQry & "  FROM       [@PS_MM010H]"
    sQry = sQry & "  WHERE      DocEntry = " & prmPQDocNum
    
    Call oRecordSet01.DoQuery(sQry)
    
    If oRecordSet01.Fields("U_Purchase").VALUE = prmPOType Then
        Check_Purchase_Type = True
    Else
        Check_Purchase_Type = False
    End If

    Set oRecordSet01 = Nothing
    Exit Function

Check_Purchase_Type_Error:
    Set oRecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "Check_Purchase_Type_Error_Error " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Function PS_MM030_CheckDate(ByVal pBaseEntry As String) As Boolean
'******************************************************************************
'Function ID : PS_MM030_CheckDate()
'해당모듈    : PS_MM030
'기능        : 선행프로세스와 일자 비교
'인수        : pBaseEntry : 기준문서번호
'반환값      : True-선행프로세스보다 일자가 같거나 느릴 경우, False-선행프로세스보다 일자가 빠를 경우
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_MM030_CheckDate_Error

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
    DocType = "PS_MM030"
    CurDocDate = oForm01.Items("DocDate").Specific.VALUE

    Query01 = "         EXEC PS_Z_CHECK_DATE '"
    Query01 = Query01 & BaseEntry & "','"
    Query01 = Query01 & BaseLine & "','"
    Query01 = Query01 & DocType & "','"
    Query01 = Query01 & CurDocDate & "'"

    Call oRecordSet01.DoQuery(Query01)

    If oRecordSet01.Fields("ReturnValue").VALUE = "False" Then
        PS_MM030_CheckDate = False
    Else
        PS_MM030_CheckDate = True
    End If
    
    Set oRecordSet01 = Nothing

    Exit Function
PS_MM030_CheckDate_Error:
    PS_MM030_CheckDate = False
    Call Sbo_Application.SetStatusBarMessage("PS_MM030_CheckDate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)
End Function

Private Function PS_MM030_LineDelete_Possible(ByVal pDocEntry As String, ByVal pLineID As String) As Boolean
'******************************************************************************
'Function ID : PS_MM030_LineDelete_Possible()
'해당모듈    : PS_MM030
'기능        : 행삭제 가능 여부 확인
'인수        : pDocEntry : 문서번호, pLineID : 라인ID
'반환값      : True-가입고가 등록되지 않은 경우(삭제가능), False-가입고가 등록된 경우(삭제불가능)
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_MM030_LineDeletPossbile_Error

    Dim QueryStr01   As String
    Dim loopCount    As Integer
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    QueryStr01 = "              SELECT      COUNT(*) AS [Cnt]"
    QueryStr01 = QueryStr01 & " FROM        [@PS_MM050H] A"
    QueryStr01 = QueryStr01 & "             INNER JOIN"
    QueryStr01 = QueryStr01 & "             [@PS_MM050L] B"
    QueryStr01 = QueryStr01 & "                 ON A.DocEntry = B.DocEntry"
    QueryStr01 = QueryStr01 & " WHERE       B.U_PODocNum = '" & pDocEntry & "'"
    QueryStr01 = QueryStr01 & "             AND B.U_POLinNum = '" & pLineID & "'"
    QueryStr01 = QueryStr01 & "             AND A.Status = 'O'"

    Call oRecordSet01.DoQuery(QueryStr01)

    If oRecordSet01.Fields("Cnt").VALUE > 0 Then
        PS_MM030_LineDelete_Possible = False '삭제 불가
    Else
        PS_MM030_LineDelete_Possible = True '삭제 가능
    End If

    Exit Function
    
PS_MM030_LineDeletPossbile_Error:
    PS_MM030_LineDelete_Possible = False
    Call Sbo_Application.SetStatusBarMessage("PS_MM030_LineDeletPossbile_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)
End Function

Private Sub PS_MM030_TotalAmount_Calculate()
'******************************************************************************
'Function ID : PS_MM030_TotalAmount_Calculate()
'해당모듈    : PS_MM030
'기능        : 금액합계 계산
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_MM030_TotalAmount_Calculate_Error

    Dim i         As Integer
    Dim DocTotal  As Currency
    Dim SumQty    As Long
    Dim SumWeight As Currency

    Call oMat01.FlushToDataSource

    For i = 0 To oMat01.VisualRowCount - 1
        DocTotal = DocTotal + oMat01.Columns("LinTotal").Cells(i + 1).Specific.VALUE
        If oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE = "" Then
            SumQty = SumQty
        Else
            SumQty = SumQty + oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE
        End If
        SumWeight = SumWeight + oMat01.Columns("Weight").Cells(i + 1).Specific.VALUE
    Next i
    
    oForm01.Items("DocTotal").Specific.VALUE = DocTotal
    oForm01.Items("SumQty").Specific.VALUE = SumQty
    oForm01.Items("SumWeight").Specific.VALUE = SumWeight

    Call oMat01.LoadFromDataSource

    Exit Sub
    
PS_MM030_TotalAmount_Calculate_Error:
    Call Sbo_Application.SetStatusBarMessage("PS_MM030_TotalAmount_Calculate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)
End Sub

Private Function PS_MM030_TotalAmount_Compare() As Boolean
'******************************************************************************
'Function ID : PS_MM030_TotalAmount_Compare()
'해당모듈    : PS_MM030
'기능        : 견적금액 비교
'인수        : 없음
'반환값      : 최소견적금액보다 합계금액이 작으면 False, 같거나 크면 True
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_MM030_TotalAmount_Compare_Error

    Dim DocTotal   As Currency
    Dim EAmount1   As Currency
    Dim EAmount2   As Currency
    Dim EAmount3   As Currency
    Dim MinAmount1 As Currency '비교후 최소값 저장용
    Dim MinAmount2 As Currency '비교후 최소값 저장용
    
    DocTotal = oForm01.Items("DocTotal").Specific.VALUE
    EAmount1 = oForm01.Items("EAmount1").Specific.VALUE
    EAmount2 = oForm01.Items("EAmount2").Specific.VALUE
    EAmount3 = oForm01.Items("EAmount3").Specific.VALUE
    
    If EAmount1 + EAmount2 + EAmount3 > 0 Then
    
        If (EAmount1 <> 0 And EAmount2 <> 0) Then
            If EAmount1 < EAmount2 Then
                MinAmount1 = EAmount1
            Else
                MinAmount1 = EAmount2
            End If
        ElseIf EAmount1 = 0 And EAmount2 <> 0 Then
            MinAmount1 = EAmount2
        ElseIf EAmount1 <> 0 And EAmount2 = 0 Then
            MinAmount1 = EAmount1
        ElseIf EAmount1 = 0 And EAmount2 = 0 Then
            MinAmount1 = EAmount3
        End If
    
        If EAmount3 <> 0 Then
            If MinAmount1 < EAmount3 Then
                MinAmount2 = MinAmount1
            Else
                MinAmount2 = EAmount3
            End If
        Else
            MinAmount2 = MinAmount1
        End If
    
    Else
    
        MinAmount2 = DocTotal
    
    End If

    If DocTotal <= MinAmount2 Then
        PS_MM030_TotalAmount_Compare = True
    Else
        PS_MM030_TotalAmount_Compare = False
    End If

    Exit Function

PS_MM030_TotalAmount_Compare_Error:
    PS_MM030_TotalAmount_Compare = False
    Call Sbo_Application.SetStatusBarMessage("PS_MM030_TotalAmount_Compare_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)
End Function

Private Function PS_MM030_SendMail() As String
'******************************************************************************
'Function ID : PS_MM030_SendMail()
'해당모듈    : PS_MM030
'기능        : Outlook을 이용한 메일발송
'인수        : 없음
'반환값      : 성공:"", 실패:Err.Description
'특이사항    : 사용안됨(2017.05.23), 추후 필요할 것 같아서 남겨둠
'******************************************************************************
On Error GoTo PS_MM030_SendMail_Error

    Dim objOutlook As Outlook.Application
    Dim objMail As MailItem

    Dim strToAddress As String
    Dim strSubject As String
    Dim strBody As String

    Set objOutlook = New Outlook.Application
    Set objMail = Outlook.CreateItem(olMailItem)

    strToAddress = MDC_GetData.Get_ReData("ISNULL(E_Mail, '')", "CardCode", "OCRD", "'" & oForm01.Items("CardCode").Specific.VALUE & "'")
    If strToAddress = "" Then 'BP마스터에 E-Mail 주소가 없는 경우
        GoTo PS_MM030_SendMail_EMailStringEmpty
    End If
    
    strSubject = "풍산홀딩스 주문서 첨부 (주문번호:" & oForm01.Items("DocNum").Specific.VALUE & ")"
    strBody = "주문서를 첨부합니다." & Chr(13) & "첨부파일을 참조하십시오."

    objMail.To = strToAddress
    objMail.Subject = strSubject
    objMail.BodyFormat = olFormatPlain
    objMail.Body = strBody
    Call objMail.Attachments.Add("C:\ReportExport\PS_MM030_주문서_" & oForm01.Items("DocNum").Specific.VALUE & ".pdf")

    Call objMail.send
    
    Set objMail = Nothing
    Set objOutlook = Nothing
    
    PS_MM030_SendMail = ""

    Exit Function
    
PS_MM030_SendMail_EMailStringEmpty:

    PS_MM030_SendMail = "EMailStringEmpty"
    
    Set objMail = Nothing
    Set objOutlook = Nothing
    
    Exit Function
    
PS_MM030_SendMail_Error:

    PS_MM030_SendMail = Err.Description

    Set objMail = Nothing
    Set objOutlook = Nothing

End Function

