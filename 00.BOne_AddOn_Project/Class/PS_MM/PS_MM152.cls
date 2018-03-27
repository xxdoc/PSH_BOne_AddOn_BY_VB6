VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_MM152"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_MM152.cls
'//  Module         : MM
'//  Description    : 가공품 출고
'//  FormType       : PS_MM152
'//  Create Date    : 2010.11.17
'//  Modified Date  :
'//  Creator        : Ryu Yung Jo
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01             As SAPbouiCOM.Form
Public oMat01              As SAPbouiCOM.Matrix
Private oDS_PS_MM152H As SAPbouiCOM.DBDataSource    '등록헤더
Private oDS_PS_MM152L As SAPbouiCOM.DBDataSource    '등록라인

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

Private oLast_Item_UID      As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID       As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row       As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&

Private oDocEntryNext As Long
Private oCheck As String

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_MM152.srf")
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
    
    oFormUniqueID01 = "PS_MM152_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
  
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    '화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
    oForm01.DataBrowser.BrowseBy = "DocEntry"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting
    Call Initialization
    Call FormClear
    Call Add_MatrixRow(0, True)
    Call FormItemEnabled
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1285"), False        '// 복원
    oForm01.EnableMenu ("1287"), True         '// 복제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1284"), True         '// 취소
    oForm01.EnableMenu ("1293"), False        '// 행삭제
        
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
    Dim ItemCode$, ItemName$, Size$, Qty&, Weight As Currency, Unit$, RequestDate$, DueDate$, ItemType$, RequestNo$
    Dim Calculate_Weight As Double
    
    Dim DocEntry As Long
    Dim DocEntryMax As Long
    Dim DocEntryNext As Long
    
    Dim ScrapWt As Currency
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE Then
'                        If Trim(Sbo_Company.UserName) = "66302" Or Trim(Sbo_Company.UserName) = "71090" Then
'                        Else
'                            MDC_Com.MDC_GF_Message "해당 아이디는 해당 문서를 추가할 수 없습니다.", "E"
'                            BubbleEvent = False
'                            Exit Sub
'                        End If

                        If HeaderSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        If MatrixSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        '//스크랩 중량 계산
                        oMat01.FlushToDataSource
                        For i = 0 To oMat01.RowCount - 2
                            If Trim(oDS_PS_MM152L.GetValue("U_ReStatus", i)) = "Y" Then
                                sQry = "Select U_ItmBSort From [OITM] Where ItemCode = '" & Trim(oDS_PS_MM152L.GetValue("U_ItemCode", i)) & "'"
                                oRecordSet01.DoQuery sQry
                                If Trim(oRecordSet01.Fields(0).VALUE) = "101" Then '//휘팅일때만
'                                    sQry = "Select  b.U_OutDoc, b.U_OutLine, IsNull(d.U_OutWt, 0) - Sum(IsNull(b.U_OutWt, 0)) - IsNull(d.U_ReWt, 0) "
'                                    sQry = sQry & "From [@PS_MM152H] a, "
'                                    sQry = sQry & "[@PS_MM152L] b, "
'                                    sQry = sQry & "[@PS_MM130H] c, "
'                                    sQry = sQry & "[@PS_MM130L] d "
'                                    sQry = sQry & "where a.DocEntry = b.DocEntry "
'                                    sQry = sQry & "and c.DocEntry = d.DocEntry "
'                                    sQry = sQry & "and b.U_OutDoc = c.U_OutDoc "
'                                    sQry = sQry & "and b.U_OutLine = d.U_LineNum "
'                                    sQry = sQry & "and b.U_OutDoc = '" & Trim(oDS_PS_MM152L.GetValue("U_OutDoc", i)) & "' "
'                                    sQry = sQry & "and b.U_OutLine = '" & Trim(oDS_PS_MM152L.GetValue("U_OutLine", i)) & "' "
'                                    sQry = sQry & "and a.Canceled = 'N' "
'                                    sQry = sQry & "and a.U_OKYNC <> 'C' "
'                                    sQry = sQry & "Group by b.U_OutDoc, b.U_OutLine, IsNull(d.U_OutWt, 0), IsNull(d.U_ReWt, 0) "
                                    
                                    sQry = "Select a.U_OutDoc, b.U_LineNum, b.U_OutWt - Isnull(c.U_OutWt,0) - Isnull(b.U_ReWt,0) "
                                    sQry = sQry & "from [@PS_MM130H] a Inner Join [@PS_MM130L] b On a.DocEntry = b.DocEntry "
                                    sQry = sQry & "Inner Join (Select t1.U_OutDoc, "
                                    sQry = sQry & "t1.U_OutLine, "
                                    sQry = sQry & "Isnull(Sum(t1.U_OutWt),0) As U_OutWt "
                                    sQry = sQry & "From [@PS_MM152H] t Inner Join [@PS_MM152L] t1 On t.DocEntry = t1.DocEntry "
                                    sQry = sQry & "Where t.Canceled = 'N' "
                                    sQry = sQry & "and t.U_OKYNC <> 'C' "
                                    sQry = sQry & "Group by t1.U_OutDoc, "
                                    sQry = sQry & "t1.U_OutLine ) c On a.U_OutDoc = c.U_OutDoc And b.U_LineNum = c.U_OutLine "
                                    sQry = sQry & "where a.U_OutDoc = '" & Trim(oDS_PS_MM152L.GetValue("U_OutDoc", i)) & "' "
                                    sQry = sQry & "and b.U_LineNum = '" & Trim(oDS_PS_MM152L.GetValue("U_OutLine", i)) & "' "

'                                    sQry = "Select  a.U_OutDoc, a.U_OutLine, IsNull(b.U_OutWt, 0) - Sum(IsNull(a.U_OutWt, 0)) - IsNull(b.U_ReWt, 0) "
'                                    sQry = sQry & "From [@PS_MM152L] a Inner Join [@PS_MM130L] b On a.U_OutDoc = b.U_OutDoc And a.U_OutLine = b.U_LineNum "
'                                    sQry = sQry & "Inner Join [@PS_MM152H] c On a.DocEntry = c.DocEntry "
'                                    sQry = sQry & "Where a.U_OutDoc = '" & Trim(oDS_PS_MM152L.GetValue("U_OutDoc", i)) & "' "
'                                    sQry = sQry & "And a.U_OutLine = '" & Trim(oDS_PS_MM152L.GetValue("U_OutLine", i)) & "' "
'                                    sQry = sQry & "And c.CanCeled = 'N' "
'                                    sQry = sQry & "And c.U_OKYNC <> 'C' "
'                                    sQry = sQry & "Group by a.U_OutDoc, a.U_OutLine, b.U_OutWt, b.U_ReWt "
                                    oRecordSet01.DoQuery sQry
                                    
                                    '//스크랩 중량 = 미입고 잔량 - 현재 입고 중량
                                    ScrapWt = oRecordSet01.Fields(2).VALUE - Trim(oDS_PS_MM152L.GetValue("U_OutWt", i))
                                    
                                    If oRecordSet01.RecordCount = 0 Then
                                        sQry = "Select U_OutDoc, U_LineNum, IsNull(U_OutWt, 0) - IsNull(U_ReWt, 0) "
                                        sQry = sQry & "From [@PS_MM130L] "
                                        sQry = sQry & "Where U_OutDoc = '" & Trim(oDS_PS_MM152L.GetValue("U_OutDoc", i)) & "' "
                                        sQry = sQry & "And U_LineNum = '" & Trim(oDS_PS_MM152L.GetValue("U_OutLine", i)) & "' "
                                        oRecordSet01.DoQuery sQry
                                        
                                        ScrapWt = oRecordSet01.Fields(2).VALUE - Trim(oDS_PS_MM152L.GetValue("U_OutWt", i))
                                    End If
                                    
                                    Call oDS_PS_MM152L.setValue("U_ScrapWt", i, ScrapWt)
                               End If
                            Else
                                'Call oDS_PS_MM152L.setValue("U_ScrapWt", i, 0)
                            End If
                        Next i
                        oMat01.LoadFromDataSource
                        
                        Call Delete_EmptyRow
                    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                        If HeaderSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        If Trim(oDS_PS_MM152H.GetValue("U_OKYNC", 0)) <> "C" Then
                            If MatrixSpaceLineDel = False Then
                                BubbleEvent = False
                                Exit Sub
                            End If
                        
                        
                            '//스크랩 중량 계산
                            oMat01.FlushToDataSource
                            For i = 0 To oMat01.RowCount - 2
                                If Trim(oDS_PS_MM152L.GetValue("U_ReStatus", i)) = "Y" Then
                                    sQry = "Select U_ItmBSort From [OITM] Where ItemCode = '" & Trim(oDS_PS_MM152L.GetValue("U_ItemCode", i)) & "'"
                                    oRecordSet01.DoQuery sQry
                                    If Trim(oRecordSet01.Fields(0).VALUE) = "101" Then '//휘팅일때만
                                        sQry = "Select  b.U_OutDoc, b.U_OutLine, IsNull(d.U_OutWt, 0) - Sum(IsNull(b.U_OutWt, 0)) - IsNull(d.U_ReWt, 0) "
                                        sQry = sQry & "From [@PS_MM152H] a, "
                                        sQry = sQry & "[@PS_MM152L] b, "
                                        sQry = sQry & "[@PS_MM130H] c, "
                                        sQry = sQry & "[@PS_MM130L] d "
                                        sQry = sQry & "where a.DocEntry = b.DocEntry "
                                        sQry = sQry & "and c.DocEntry = d.DocEntry "
                                        sQry = sQry & "and b.U_OutDoc = c.U_OutDoc "
                                        sQry = sQry & "and b.U_OutLine = d.U_LineNum "
                                        sQry = sQry & "and b.U_OutDoc = '" & Trim(oDS_PS_MM152L.GetValue("U_OutDoc", i)) & "' "
                                        sQry = sQry & "and b.U_OutLine = '" & Trim(oDS_PS_MM152L.GetValue("U_OutLine", i)) & "' "
                                        sQry = sQry & "and a.Canceled = 'N' "
                                        sQry = sQry & "and a.U_OKYNC <> 'C' "
                                        sQry = sQry & "Group by b.U_OutDoc, b.U_OutLine, IsNull(d.U_OutWt, 0), IsNull(d.U_ReWt, 0) "
                                        oRecordSet01.DoQuery sQry
                                        
                                        
                                        '//스크랩 중량 = 미입고 잔량 - 현재 입고 중량
                                        ScrapWt = oRecordSet01.Fields(2).VALUE '- Trim(oDS_PS_MM152L.GetValue("U_OutWt", i))
                                        
                                        '//수정전 입고중량
                                        sQry = "Select U_OutWt From [@PS_MM152L] Where DocEntry = '" & Trim(oDS_PS_MM152L.GetValue("DocEntry", i)) & "' "
                                        sQry = sQry & "And U_OutDoc = '" & Trim(oDS_PS_MM152L.GetValue("U_OutDoc", i)) & "' "
                                        sQry = sQry & "And U_OutLine = '" & Trim(oDS_PS_MM152L.GetValue("U_OutLine", i)) & "' "
                                        oRecordSet01.DoQuery sQry
                                        
                                        ScrapWt = ScrapWt - oRecordSet01.Fields(0).VALUE + Trim(oDS_PS_MM152L.GetValue("U_OutWt", i))
                                        
                                        
                                        If oRecordSet01.RecordCount = 0 Then
                                            sQry = "Select U_OutDoc, U_LineNum, IsNull(U_OutWt, 0) - IsNull(U_ReWt, 0) "
                                            sQry = sQry & "From [@PS_MM130L] "
                                            sQry = sQry & "Where U_OutDoc = '" & Trim(oDS_PS_MM152L.GetValue("U_OutDoc", i)) & "' "
                                            sQry = sQry & "And U_LineNum = '" & Trim(oDS_PS_MM152L.GetValue("U_OutLine", i)) & "' "
                                            oRecordSet01.DoQuery sQry
                                            
                                            ScrapWt = oRecordSet01.Fields(2).VALUE - Trim(oDS_PS_MM152L.GetValue("U_OutWt", i))
                                        End If
                                        
                                        Call oDS_PS_MM152L.setValue("U_ScrapWt", i, ScrapWt)
                                   End If
                                Else
                                    'Call oDS_PS_MM152L.setValue("U_ScrapWt", i, 0)
                                End If
                            Next i
                            oMat01.LoadFromDataSource
                        End If
                        If Add_PS_PP040(pval) = False Then
                            BubbleEvent = False
                            Exit Sub
                        Else
                            If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_Commit
                        End If
                        
                        Call Delete_EmptyRow
                    ElseIf oForm01.Mode = fm_FIND_MODE Then
                        If Trim(oForm01.Items("DocEntry").Specific.VALUE) = "" Then
                            
                        Else
                            DocEntryNext = Trim(oForm01.Items("DocEntry").Specific.VALUE)
    
                            sQry = "Select U_CardCode From [@PS_MM152H] Where DocEntry = '" & DocEntryNext & "'"
                            oRecordSet01.DoQuery sQry
    
                            If Trim(Sbo_Company.UserName) <> Trim(oRecordSet01.Fields(0).VALUE) Then
                                If Left(Trim(Sbo_Company.UserName), 1) = "6" Or Left(Trim(Sbo_Company.UserName), 1) = "7" Then
                                    MDC_Com.MDC_GF_Message "해당 문서 번호는 없는 다른 거래처의 문서입니다. 확실한 문서번호를 입력하세요.", "E"
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End If
'et_KEY_DOWN ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "CardCode" Then
                        If oForm01.Items(pval.ItemUID).Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "OtDocLin" Then
                            If oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE = "" Then
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
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE And pval.Action_Success = True Then
                        oForm01.Mode = fm_OK_MODE
                        Call Sbo_Application.ActivateMenuItem("1282")
                    End If
                ElseIf pval.ItemUID = "Btn01" Then
                    Print_Report01
                End If
            Case et_KEY_DOWN: '//2
'et_COMBO_SELECT ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_COMBO_SELECT: '//5
                If pval.ItemUID = "Mat01" And pval.ColUID = "ReStatus" Then
                    If Trim(oMat01.Columns("OutGbn").Cells(pval.Row).Specific.VALUE) = "10" Then
                        If Trim(oMat01.Columns("ReStatus").Cells(pval.Row).Specific.VALUE) = "Y" Then
                            oMat01.Columns("MUseQty").Cells(pval.Row).Specific.VALUE = oMat01.Columns("MDUseQty").Cells(pval.Row).Specific.VALUE
                            oMat01.Columns("MUseWt").Cells(pval.Row).Specific.VALUE = oMat01.Columns("MDUseWt").Cells(pval.Row).Specific.VALUE
                        ElseIf Trim(oMat01.Columns("ReStatus").Cells(pval.Row).Specific.VALUE) = "N" Then
                            oMat01.Columns("MUseQty").Cells(pval.Row).Specific.VALUE = 0
                            oMat01.Columns("MUseWt").Cells(pval.Row).Specific.VALUE = 0
                        End If
                    ElseIf Trim(oMat01.Columns("OutGbn").Cells(pval.Row).Specific.VALUE) = "20" Then
                        oMat01.Columns("MUseQty").Cells(pval.Row).Specific.VALUE = 0
                        oMat01.Columns("MUseWt").Cells(pval.Row).Specific.VALUE = 0
                    End If
                End If
            Case et_CLICK: '//6
            Case et_MATRIX_LINK_PRESSED '//8
'et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "CardCode" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "OtDocLin" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        ElseIf pval.ColUID = "MUseQty" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        ElseIf pval.ColUID = "OutQty" Then
                            sQry = "Select U_ItmBSort From OITM Where ItemCode = '" & Trim(oMat01.Columns("ItemCode").Cells(pval.Row).Specific.VALUE) & "'"
                            oRecordSet01.DoQuery sQry
                            If Trim(oRecordSet01.Fields(0).VALUE) = "102" Then
                                oMat01.Columns("OutWt").Cells(pval.Row).Specific.VALUE = oMat01.Columns("OutQty").Cells(pval.Row).Specific.VALUE
                            End If
                        ElseIf pval.ColUID = "NQty" Then
                            sQry = "Select U_ItmBSort From OITM Where ItemCode = '" & Trim(oMat01.Columns("ItemCode").Cells(pval.Row).Specific.VALUE) & "'"
                            oRecordSet01.DoQuery sQry
                            If Trim(oRecordSet01.Fields(0).VALUE) = "102" Then
                                oMat01.Columns("NWeight").Cells(pval.Row).Specific.VALUE = oMat01.Columns("NQty").Cells(pval.Row).Specific.VALUE
                            End If
'                            oMat01.FlushToDataSource
'                            ItemCode = Trim(oDS_PS_MM010L.GetValue("U_ItemCode", pval.Row - 1))
'                            Qty = oDS_PS_MM010L.GetValue("U_Qty", pval.Row - 1)
'
'                            Calculate_Weight = MDC_PS_Common.Calculate_Weight(ItemCode, Qty,Trim(oForm01.Items("BPLId").Specific.Value))
'
'                            oDS_PS_MM010L.setValue "U_Weight", pval.Row - 1, Calculate_Weight
'                            oMat01.LoadFromDataSource
'                            oMat01.Columns("Weight").Cells(pval.Row).Click ct_Regular
                        End If
                    End If
                End If
'et_MATRIX_LOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_MATRIX_LOAD: '//11
'                DocEntryNext = Trim(oForm01.Items("DocEntry").Specific.Value)
'
'                sQry = "Select U_CardCode From [@PS_MM152H] Where DocEntry = '" & DocEntryNext & "'"
'                oRecordSet01.DoQuery sQry
'
'                If Trim(Sbo_Company.UserName) <> Trim(oRecordSet01.Fields(0).Value) Then
'                    If Left(Trim(Sbo_Company.UserName), 1) = "6" Then
'                        MDC_Com.MDC_GF_Message "해당 문서 번호는 없는 다른 거래처의 문서입니다. 확실한 문서번호를 입력하세요.", "E"
'                        BubbleEvent = False
'                        Exit Sub
'                    End If
'                End If
                            
                Add_MatrixRow oMat01.RowCount, False
                FormItemEnabled
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
                Set oDS_PS_MM152H = Nothing
                Set oDS_PS_MM152L = Nothing
        End Select
    End If
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
    Set oRecordSet01 = Nothing
    oForm01.Freeze False
    MDC_Com.MDC_GF_Message "Raise_ItemEvent_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_MenuEvent_Error
    Dim i&
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Dim DocEntry As Long
    Dim DocEntryMax As Long
    Dim DocEntryNext As Long
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
                If oForm01.Mode = fm_ADD_MODE Then
                    BubbleEvent = False
                    Exit Sub
                End If
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                If Trim(oForm01.Items("DocEntry").Specific.VALUE) = "" Then
                    DocEntry = 0
                Else
                    DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
                End If
                sQry = "Select Max(DocEntry) From [@PS_MM152H]"
                oRecordSet01.DoQuery sQry
                DocEntryMax = Trim(oRecordSet01.Fields(0).VALUE)
                
                If pval.MenuUID = "1288" Then '//다음
                    If oForm01.Mode = fm_ADD_MODE Then
                        Sbo_Application.ActivateMenuItem ("1290")
                        BubbleEvent = False
                        Exit Sub
                    ElseIf oForm01.Mode = fm_FIND_MODE Then
                        If (oForm01.Items("DocEntry").Specific.VALUE = "") Then
                            Sbo_Application.ActivateMenuItem ("1290")
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
One_More_Check_1288:
                    DocEntryNext = DocEntry + 1
                    If DocEntryNext > DocEntryMax Then
                        DocEntry = 0
                        GoTo One_More_Check_1288
                    End If
                        
                    sQry = "Select U_CardCode From [@PS_MM152H] Where DocEntry = '" & DocEntryNext & "'"
                    oRecordSet01.DoQuery sQry
                    
                    If Trim(Sbo_Company.UserName) <> Trim(oRecordSet01.Fields(0).VALUE) Then
                        If Left(Trim(Sbo_Company.UserName), 1) = "6" Or Left(Trim(Sbo_Company.UserName), 1) = "7" Then
                            DocEntry = DocEntry + 1
                            GoTo One_More_Check_1288
                        Else
                            oCheck = "False"
                            Exit Sub
                        End If
                    Else
                        oCheck = "True"
                        oDocEntryNext = DocEntryNext
                    End If
                ElseIf pval.MenuUID = "1289" Then '//이전
                    If oForm01.Mode = fm_ADD_MODE Then
                        Sbo_Application.ActivateMenuItem ("1291")
                        BubbleEvent = False
                        Exit Sub
                    ElseIf oForm01.Mode = fm_FIND_MODE Then
                        If (oForm01.Items("DocEntry").Specific.VALUE = "") Then
                            Sbo_Application.ActivateMenuItem ("1291")
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
One_More_Check_1289:
                    DocEntryNext = DocEntry - 1
                    If DocEntryNext < 1 Then
                        DocEntry = DocEntryMax + 1
                        GoTo One_More_Check_1289
                    End If
                        
                    sQry = "Select U_CardCode From [@PS_MM152H] Where DocEntry = '" & DocEntryNext & "'"
                    oRecordSet01.DoQuery sQry
                    
                    If Trim(Sbo_Company.UserName) <> Trim(oRecordSet01.Fields(0).VALUE) Then
                        If Left(Trim(Sbo_Company.UserName), 1) = "6" Or Left(Trim(Sbo_Company.UserName), 1) = "7" Then
                            DocEntry = DocEntry - 1
                            GoTo One_More_Check_1289
                        Else
                            oCheck = "False"
                            Exit Sub
                        End If
                    Else
                        oCheck = "True"
                        oDocEntryNext = DocEntryNext
                    End If
                ElseIf pval.MenuUID = "1290" Then '//맨첨
                    DocEntryNext = 0
One_More_Check_1290:
                    DocEntryNext = DocEntryNext + 1
                    
                    sQry = "Select U_CardCode From [@PS_MM152H] Where DocEntry = '" & DocEntryNext & "'"
                    oRecordSet01.DoQuery sQry
                    
                    If Trim(Sbo_Company.UserName) <> Trim(oRecordSet01.Fields(0).VALUE) Then
                        If Left(Trim(Sbo_Company.UserName), 1) = "6" Or Left(Trim(Sbo_Company.UserName), 1) = "7" Then
                            GoTo One_More_Check_1290
                        Else
                            oCheck = "False"
                            Exit Sub
                        End If
                    Else
                        oCheck = "True"
                        oDocEntryNext = DocEntryNext
                    End If
                ElseIf pval.MenuUID = "1291" Then '//맨뒤
                    DocEntryNext = DocEntryMax + 1
One_More_Check_1291:
                    DocEntryNext = DocEntryNext - 1
                    
                    sQry = "Select U_CardCode From [@PS_MM152H] Where DocEntry = '" & DocEntryNext & "'"
                    oRecordSet01.DoQuery sQry
                    
                    If Trim(Sbo_Company.UserName) <> Trim(oRecordSet01.Fields(0).VALUE) Then
                        If Left(Trim(Sbo_Company.UserName), 1) = "6" Or Left(Trim(Sbo_Company.UserName), 1) = "7" Then
                            GoTo One_More_Check_1291
                        Else
                            oCheck = "False"
                            Exit Sub
                        End If
                    Else
                        oCheck = "True"
                        oDocEntryNext = DocEntryNext
                    End If
                End If
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
'                oForm01.Freeze True
'                If oMat01.RowCount <> oMat01.VisualRowCount Then
'                    For i = 0 To oMat01.VisualRowCount - 1
'                        oMat01.Columns("LineNum").Cells(i + 1).Specific.Value = i + 1
'                    Next i
'
'                    oMat01.FlushToDataSource
'                    oDS_PS_MM152L.RemoveRecord oDS_PS_MM152L.Size - 1       '// Mat01에 마지막라인(빈라인) 삭제
'                    oMat01.Clear
'                    oMat01.LoadFromDataSource
'
'                    If oMat01.Columns("OrdMgNum").Cells(oMat01.RowCount).Specific.Value <> "" Then
'                        Call Add_MatrixRow(oMat01.RowCount, False)
'                    End If
'                End If
'                oForm01.Freeze False
            Case "1281": '찾기
                oForm01.Freeze True
                oDS_PS_MM152H.setValue "U_BPLId", 0, "1"
                Call FormItemEnabled
                If Left(Trim(Sbo_Company.UserName), 1) = "6" Or Left(Trim(Sbo_Company.UserName), 1) = "7" Then
            '        oForm01.Items("CardCode").Specific.Value = Trim(Sbo_Company.UserName)
                    oForm01.Items("CardCode").Specific.VALUE = Trim(Sbo_Company.UserName)
'                    oDS_PS_MM152H.setValue "U_CardCode", 0, Trim(Sbo_Company.UserName)
                    sQry = "Select CardName From OCRD Where CardCode = '" & Trim(oForm01.Items("CardCode").Specific.VALUE) & "'"
                    oRecordSet01.DoQuery sQry
                    oForm01.Items("CardName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
'                    oDS_PS_MM152H.setValue "U_CardName", 0, Trim(oRecordSet01.Fields(0).Value)
                End If
                oForm01.Items("DocEntry").Click ct_Regular
                oForm01.Items("CardCode").Enabled = False
                oForm01.Freeze False
            Case "1282": '추가
                oForm01.Freeze True
                Call Initialization
                Call FormClear
                Call FormItemEnabled
                Add_MatrixRow 0, True
                oForm01.Freeze False
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                oForm01.Freeze True
                If oCheck = "True" Then
                    Sbo_Application.ActivateMenuItem ("1281")
                    oForm01.Items("DocEntry").Specific.VALUE = oDocEntryNext
                    oForm01.Items("1").Click ct_Regular
                    
                    oCheck = "False"
                    oDocEntryNext = 0
                End If
                
                Call FormItemEnabled
                oForm01.Freeze False
            Case "1287": '// 복제
                oForm01.Freeze True
                Call FormClear
                oDS_PS_MM152H.setValue "Status", 0, "O"
                oDS_PS_MM152H.setValue "Canceled", 0, "N"
                
                Call FormItemEnabled

'                For i = 0 To oMat01.VisualRowCount - 1
'                    oMat01.FlushToDataSource
'                    oDS_PS_MM152L.setValue "DocEntry", i, ""
'                    oMat01.LoadFromDataSource
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
    Set oDS_PS_MM152H = oForm01.DataSources.DBDataSources("@PS_MM152H")
    Set oDS_PS_MM152L = oForm01.DataSources.DBDataSources("@PS_MM152L")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    
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
    
    oMat01.Columns("OutGbn").ValidValues.Add "10", "원재료"
    oMat01.Columns("OutGbn").ValidValues.Add "20", "재공"
    
    oMat01.Columns("ReStatus").ValidValues.Add "Y", "완료"
    oMat01.Columns("ReStatus").ValidValues.Add "N", "미완료"
    
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
'    Set oCombo = oForm01.Items("BPLId").Specific
'    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
'    oCombo.Select "1", psk_ByValue
    oDS_PS_MM152H.setValue "U_BPLId", 0, "1"
    
    oDS_PS_MM152H.setValue "U_DocDate", 0, Format(Now, "yyyymmdd")
    
'    oForm01.Items("OKYNC").Specific.Select "N", psk_ByValue
    oDS_PS_MM152H.setValue "U_OKYNC", 0, "N"

    If Left(Trim(Sbo_Company.UserName), 1) = "6" Or Left(Trim(Sbo_Company.UserName), 1) = "7" Then
'        oForm01.Items("CardCode").Specific.Value = Trim(Sbo_Company.UserName)
        oDS_PS_MM152H.setValue "U_CardCode", 0, Trim(Sbo_Company.UserName)
        sQry = "Select CardName From OCRD Where CardCode = '" & Trim(oDS_PS_MM152H.GetValue("U_CardCode", 0)) & "'"
        oRecordSet01.DoQuery sQry
        oDS_PS_MM152H.setValue "U_CardName", 0, Trim(oRecordSet01.Fields(0).VALUE)
    End If
    
    '//아이디별 사번 세팅
'    oForm01.Items("CntcCode").Specific.Value = MDC_PS_Common.User_MSTCOD
    
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
    oForm01.Items("Comments").Click ct_Regular
    If (oForm01.Mode = fm_ADD_MODE) Then
        If Trim(Sbo_Company.UserName) = "66302" Or Trim(Sbo_Company.UserName) = "71090" Then
            oForm01.Items("BPLId").Enabled = False
            oForm01.Items("CardCode").Enabled = False
            oForm01.Items("DocDate").Enabled = True
            oForm01.Items("OKYNC").Enabled = False
            oForm01.Items("DocEntry").Enabled = False

            oMat01.Columns("OtDocLin").Editable = True
            oMat01.Columns("OutQty").Editable = True
            oMat01.Columns("OutWt").Editable = True
            oMat01.Columns("NQty").Editable = True
            oMat01.Columns("NWeight").Editable = True
            oMat01.Columns("CPWt").Editable = True
            oMat01.Columns("CPWtName").Editable = True
            oMat01.Columns("MUseQty").Editable = True
            oMat01.Columns("MUseWt").Editable = True
            oMat01.Columns("ReStatus").Editable = True

            oForm01.Items("DocDate").Click ct_Regular
        Else
            oForm01.Items("BPLId").Enabled = True
            oForm01.Items("CardCode").Enabled = True
            oForm01.Items("DocDate").Enabled = True
            oForm01.Items("OKYNC").Enabled = False
            oForm01.Items("DocEntry").Enabled = False

            oMat01.Columns("OtDocLin").Editable = True
            oMat01.Columns("OutQty").Editable = True
            oMat01.Columns("OutWt").Editable = True
            oMat01.Columns("NQty").Editable = True
            oMat01.Columns("NWeight").Editable = True
            oMat01.Columns("CPWt").Editable = True
            oMat01.Columns("CPWtName").Editable = True
            oMat01.Columns("MUseQty").Editable = True
            oMat01.Columns("MUseWt").Editable = True
            oMat01.Columns("ReStatus").Editable = True
        End If
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        oForm01.Items("BPLId").Enabled = False
        oForm01.Items("CardCode").Enabled = True
        oForm01.Items("DocDate").Enabled = True
        oForm01.Items("OKYNC").Enabled = True
        oForm01.Items("DocEntry").Enabled = True

        oMat01.Columns("OtDocLin").Editable = False
        oMat01.Columns("OutQty").Editable = False
        oMat01.Columns("OutWt").Editable = False
        oMat01.Columns("NQty").Editable = False
        oMat01.Columns("NWeight").Editable = False
        oMat01.Columns("CPWt").Editable = False
        oMat01.Columns("CPWtName").Editable = False
        oMat01.Columns("MUseQty").Editable = False
        oMat01.Columns("MUseWt").Editable = False
        oMat01.Columns("ReStatus").Editable = False

        oForm01.Items("DocEntry").Click ct_Regular
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        If Trim(Sbo_Company.UserName) = Trim(oForm01.Items("CardCode").Specific.VALUE) Then
            If Trim(oDS_PS_MM152H.GetValue("U_OKYNC", 0)) <> "Y" Then
                oForm01.Items("BPLId").Enabled = False
                oForm01.Items("CardCode").Enabled = False
                oForm01.Items("DocDate").Enabled = True
                oForm01.Items("OKYNC").Enabled = False
                oForm01.Items("DocEntry").Enabled = False
    
                oMat01.Columns("OtDocLin").Editable = False
                oMat01.Columns("OutQty").Editable = False
                oMat01.Columns("OutWt").Editable = False
                oMat01.Columns("NQty").Editable = False
                oMat01.Columns("NWeight").Editable = False
                oMat01.Columns("CPWt").Editable = True
                oMat01.Columns("CPWtName").Editable = True
                oMat01.Columns("MUseQty").Editable = True
                oMat01.Columns("MUseWt").Editable = True
                oMat01.Columns("ReStatus").Editable = True
                                
                oForm01.EnableMenu ("1284"), True         '// 취소
    
                oForm01.Items("DocDate").Click ct_Regular
            Else
                oForm01.Items("BPLId").Enabled = False
                oForm01.Items("CardCode").Enabled = False
                oForm01.Items("DocDate").Enabled = False
                oForm01.Items("OKYNC").Enabled = False
                oForm01.Items("DocEntry").Enabled = False
    
                oMat01.Columns("OtDocLin").Editable = False
                oMat01.Columns("OutQty").Editable = False
                oMat01.Columns("OutWt").Editable = False
                oMat01.Columns("NQty").Editable = False
                oMat01.Columns("NWeight").Editable = False
                oMat01.Columns("CPWt").Editable = False
                oMat01.Columns("CPWtName").Editable = False
                oMat01.Columns("MUseQty").Editable = False
                oMat01.Columns("MUseWt").Editable = False
                oMat01.Columns("ReStatus").Editable = False
                
                oForm01.EnableMenu ("1284"), False         '// 취소
            End If
        Else
            If Trim(oDS_PS_MM152H.GetValue("U_OKYNC", 0)) <> "Y" Then
                oForm01.Items("BPLId").Enabled = False
                oForm01.Items("CardCode").Enabled = False
                oForm01.Items("DocDate").Enabled = True
                If Left(Trim(Sbo_Company.UserName), 1) = "6" Or Left(Trim(Sbo_Company.UserName), 1) = "7" Then
                    oForm01.Items("OKYNC").Enabled = False
                Else
                    If Trim(oDS_PS_MM152H.GetValue("CanCeled", 0)) = "Y" Then
                        oForm01.Items("OKYNC").Enabled = False
                    Else
                        oForm01.Items("OKYNC").Enabled = True
                    End If
                End If
                
                oForm01.Items("DocEntry").Enabled = False
    
                oMat01.Columns("OtDocLin").Editable = True
                oMat01.Columns("OutQty").Editable = True
                oMat01.Columns("OutWt").Editable = True
                oMat01.Columns("NQty").Editable = True
                oMat01.Columns("NWeight").Editable = True
                oMat01.Columns("CPWt").Editable = True
                oMat01.Columns("CPWtName").Editable = True
                oMat01.Columns("MUseQty").Editable = True
                oMat01.Columns("MUseWt").Editable = True
                oMat01.Columns("ReStatus").Editable = True
                
                oForm01.EnableMenu ("1284"), True         '// 취소
            Else
                oForm01.Items("BPLId").Enabled = False
                oForm01.Items("CardCode").Enabled = False
                oForm01.Items("DocDate").Enabled = False
                If Left(Trim(Sbo_Company.UserName), 1) = "6" Or Left(Trim(Sbo_Company.UserName), 1) = "7" Then
                    oForm01.Items("OKYNC").Enabled = False
                Else
                    If Trim(oDS_PS_MM152H.GetValue("CanCeled", 0)) = "Y" Then
                        oForm01.Items("OKYNC").Enabled = False
                    Else
                        oForm01.Items("OKYNC").Enabled = True
                    End If
                End If
                
                oForm01.Items("DocEntry").Enabled = False
    
                oMat01.Columns("OtDocLin").Editable = False
                oMat01.Columns("OutQty").Editable = False
                oMat01.Columns("OutWt").Editable = False
                oMat01.Columns("NQty").Editable = False
                oMat01.Columns("NWeight").Editable = False
                oMat01.Columns("CPWt").Editable = False
                oMat01.Columns("CPWtName").Editable = False
                oMat01.Columns("MUseQty").Editable = False
                oMat01.Columns("MUseWt").Editable = False
                oMat01.Columns("ReStatus").Editable = False
                
                oForm01.EnableMenu ("1284"), False         '// 취소
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
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM152'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormClear_Error:
    MDC_Com.MDC_GF_Message "FormClear_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_MM152L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_MM152L.Offset = oRow
    oDS_PS_MM152L.setValue "U_LineNum", oRow, oRow + 1
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
    Dim MDUseQty As Long, MDUseWt As Currency, MDUseUnWt As Currency, MUseWt As Currency
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sRow = oRow
    
    Select Case oUID
        Case "CardCode"
            sQry = "Select CardName From OCRD Where CardCode = '" & Trim(oDS_PS_MM152H.GetValue("U_CardCode", 0)) & "'"
            oRecordSet01.DoQuery sQry
            
            oDS_PS_MM152H.setValue "U_CardName", 0, Trim(oRecordSet01.Fields(0).VALUE)
        Case "Mat01"
            If oCol = "OtDocLin" Then
                oForm01.Freeze True
                oMat01.FlushToDataSource

                If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 0) And Trim(oMat01.Columns("OtDocLin").Cells(oRow).Specific.VALUE) <> "" Then
                    oMat01.FlushToDataSource
                    Call Add_MatrixRow(oMat01.RowCount, False)
                    oMat01.Columns("OtDocLin").Cells(oRow).Click ct_Regular
                End If
                
                sQry = "EXEC [PS_MM152_01] '" & Trim(oDS_PS_MM152H.GetValue("U_BPLId", 0)) & "', '" & Trim(oDS_PS_MM152H.GetValue("U_CardCode", 0)) & "', "
                sQry = sQry & "'" & Trim(oDS_PS_MM152L.GetValue("U_OtDocLin", oRow - 1)) & "' , '2'"
                oRecordSet01.DoQuery sQry
                
                oMat01.Columns("ItemCode").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_ItemCode").VALUE)
                oMat01.Columns("ItemName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_ItemName").VALUE)
                oMat01.Columns("Size").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_Size").VALUE)
                oMat01.Columns("Mark").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_Mark").VALUE)
                oMat01.Columns("OutItmCd").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_OutItmCd").VALUE)
                oMat01.Columns("OutItmNm").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_OutItmNm").VALUE)
                oMat01.Columns("OutGbn").Cells(oRow).Specific.Select Trim(oRecordSet01.Fields("U_OutGbn").VALUE)
                oMat01.Columns("InQty").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_InQty").VALUE)
                oMat01.Columns("InWt").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_InWt").VALUE)
                oMat01.Columns("MTUseQty").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_MTUseQty").VALUE)
                oMat01.Columns("MTUseWt").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_MTUseWt").VALUE)
                oMat01.Columns("MDUseQty").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_MDUseQty").VALUE)
                oMat01.Columns("MDUseWt").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_MDUseWt").VALUE)
                oMat01.Columns("OutDoc").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_OutDoc").VALUE)
                oMat01.Columns("OutLine").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_OutLine").VALUE)
                oMat01.Columns("JakNum").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_OrdNum").VALUE)
                oMat01.Columns("ReStatus").Cells(oRow).Specific.Select "N"
                oMat01.Columns("CpCode").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_CpCode").VALUE)
                oMat01.Columns("CpName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_CpName").VALUE)
                oMat01.Columns("TCpCode").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_TCpCode").VALUE)
                oMat01.Columns("TCpName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_TCpName").VALUE)
'                sQry = "Select ItemName, FrgnName From OITM Where ItemCode = '" & Trim(oMat01.Columns("ItemCode").Cells(oRow).Specific.Value) & "'"
'                oRecordSet01.DoQuery sQry
'                oMat01.Columns("ItemName").Cells(oRow).Specific.Value = Trim(oRecordSet01.Fields(0).Value)
'                oMat01.Columns("FrgnName").Cells(oRow).Specific.Value = Trim(oRecordSet01.Fields(1).Value)
'
'                oMat01.Columns("ItemCode").Cells(oRow).Click ct_Regular
                oForm01.Freeze False
            ElseIf oCol = "MUseQty" Then
                oForm01.Freeze True
                MDUseQty = oMat01.Columns("MDUseQty").Cells(oRow).Specific.VALUE
                MDUseWt = oMat01.Columns("MDUseWt").Cells(oRow).Specific.VALUE
                If MDUseQty <> 0 Then
                    MDUseUnWt = MDUseWt / MDUseQty
                Else
                    MDUseUnWt = 0
                End If
                MUseWt = MDUseUnWt * oMat01.Columns("MUseQty").Cells(oRow).Specific.VALUE
                If MUseWt > MDUseWt Then
                    oMat01.Columns("MUseWt").Cells(oRow).Specific.VALUE = MDUseWt
                Else
                    oMat01.Columns("MUseWt").Cells(oRow).Specific.VALUE = MUseWt
                End If
                oMat01.Columns("MUseQty").Cells(oRow).Click ct_Regular
                oForm01.Freeze False
            End If
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
    Dim ErrNum          As Integer
    Dim DocNum          As String

    ErrNum = 0

    '// Check
    Select Case True
        Case oDS_PS_MM152H.GetValue("U_BPLId", 0) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_MM152H.GetValue("U_CardCode", 0) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_MM152H.GetValue("U_DocDate", 0) = ""
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "사업장은 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "외주거래처는 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "전기일자는 필수입력사항입니다. 확인하세요.", "E"
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
    Dim MUseWt As Currency, MDUseWt As Currency, MUseQty As Currency, MDUseQty As Currency
    
    Dim OutDoc As String
    Dim OutLine As String
    Dim ItemCode As String
    
    Dim MM130Wt As Currency
    Dim MM152Wt As Currency
    Dim MM132Wt As Currency
    Dim OutWt   As Currency
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    ErrNum = 0
    
    oMat01.FlushToDataSource

    '// 라인
    If oMat01.VisualRowCount = 0 Then
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    ElseIf oMat01.VisualRowCount = 1 Then
        If oDS_PS_MM152L.GetValue("U_OtDocLin", 0) = "" Then
            ErrNum = 2
            GoTo MatrixSpaceLineDel_Error
        End If
    End If
    
    For i = 0 To oMat01.VisualRowCount - 2
        Select Case True
            Case Trim(oDS_PS_MM152L.GetValue("U_OtDocLin", i)) = ""
                ErrNum = 3
                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_MM152L.GetValue("U_Qty", i) = "" Or oDS_PS_MM152L.GetValue("U_Qty", i) = 0
'                ErrNum = 4
'                GoTo MatrixSpaceLineDel_Error
 '           Case oDS_PS_MM152L.GetValue("U_OutWt", i) = 0 And oDS_PS_MM152L.GetValue("U_ScrapWt", i) = 0 And oDS_PS_MM152L.GetValue("U_MUseWt", i) = 0 And oDS_PS_MM152L.GetValue("U_Loss", i) = 0 '2017.11.04 Loss만 (이영진과장요청) - 반납받은 스크랩등록
            Case oDS_PS_MM152L.GetValue("U_OutWt", i) = 0 And oDS_PS_MM152L.GetValue("U_NWeight", i) = 0 And oDS_PS_MM152L.GetValue("U_ScrapWt", i) = 0 And oDS_PS_MM152L.GetValue("U_MUseWt", i) = 0 And oDS_PS_MM152L.GetValue("U_Loss", i) = 0 '2017.11.04 Loss만 (이영진과장요청) - 반납받은 스크랩등록
                ErrNum = 5
                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_MM152L.GetValue("U_Price", i) = 0
'                ErrNum = 6
'                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_MM152L.GetValue("U_LinTotal", i) = 0
'                ErrNum = 7
'                GoTo MatrixSpaceLineDel_Error
        End Select
        
        MUseQty = Trim(oDS_PS_MM152L.GetValue("U_MUseQty", i))
        MUseWt = Trim(oDS_PS_MM152L.GetValue("U_MUseWt", i))
        MDUseQty = Trim(oDS_PS_MM152L.GetValue("U_MDUseQty", i))
        MDUseWt = Trim(oDS_PS_MM152L.GetValue("U_MDUseWt", i))
        If MUseWt > MDUseWt Then
            ErrNum = 4
            GoTo MatrixSpaceLineDel_Error
        End If
        
'        If Trim(oDS_PS_MM152L.GetValue("U_OutGbn", i)) = "10" Then
'            If MUseQty > MDUseQty Then
'                ErrNum = 5
'                GoTo MatrixSpaceLineDel_Error
'            End If
'        End If
        
        '//출고중량 체크
        OutDoc = Trim(oDS_PS_MM152L.GetValue("U_OutDoc", i))
        OutLine = Trim(oDS_PS_MM152L.GetValue("U_OutLine", i))
        
        ItemCode = Trim(oDS_PS_MM152L.GetValue("U_ItemCode", i))
        
        sQry = "Select U_ItmBsort from OITM Where ItemCode = '" & ItemCode & "'"
        oRecordSet.DoQuery sQry
        
        If Trim(oRecordSet.Fields(0).VALUE) = "101" Then
            '//휘팅일경우 출고중량과 반납중량을 check
            
            '//반출중량
            sQry = "Select b.U_OutWt from [@PS_MM130H] a, [@PS_MM130L] b Where a.DocEntry = b.DocEntry and a.Canceled = 'N' and a.U_OKYNC <> 'C' and a.U_OutDoc = '" & OutDoc & "' and b.U_LineNum = '" & OutLine & "'"
            oRecordSet.DoQuery sQry
            MM130Wt = Trim(oRecordSet.Fields(0).VALUE)
            
            '//외주입고중량
            sQry = "Select Isnull(Sum(b.U_OutWt + b.U_ScrapWt),0) from [@PS_MM152H] a, [@PS_MM152L] b Where a.DocEntry = b.DocEntry and a.Canceled = 'N' and a.U_OKYNC <> 'C' And b.U_OutDoc = '" & OutDoc & "' and b.U_OutLine = '" & OutLine & "'"
            oRecordSet.DoQuery sQry
            
            '// 기입고량 + 현재입력량(스크랩포함)
            MM152Wt = Trim(oRecordSet.Fields(0).VALUE)
            
            If oForm01.Mode = fm_UPDATE_MODE Then
                '//현재입력되어있는 출고량
                sQry = "Select Isnull(b.U_OutWt,0) from [@PS_MM152H] a, [@PS_MM152L] b "
                sQry = sQry + "Where a.DocEntry = b.DocEntry and a.Canceled = 'N' and a.U_OKYNC <> 'C' And b.U_OutDoc = '" & OutDoc & "' and b.U_OutLine = '" & OutLine & "' "
                sQry = sQry + "And a.DocEntry = '" & Trim(oDS_PS_MM152L.GetValue("DocEntry", i)) & "'"
                oRecordSet.DoQuery sQry
                
                OutWt = Trim(oRecordSet.Fields(0).VALUE)
            Else
                OutWt = 0
            End If
            
            MM152Wt = MM152Wt - OutWt + Trim(oDS_PS_MM152L.GetValue("U_OutWt", i))
            
            '//반품량
            sQry = "Select Isnull(Sum(b.U_ReWt),0) from [@PS_MM132H] a, [@PS_MM132L] b Where a.DocEntry = b.DocEntry and a.Canceled = 'N' and a.U_OKYNC <> 'C' And b.U_OutDoc = '" & OutDoc & "' and b.U_OutLine = '" & OutLine & "'"
            oRecordSet.DoQuery sQry

            MM132Wt = Trim(oRecordSet.Fields(0).VALUE)
            
            '//출고수량 보다 입고수량이 많으면 안된다.
            If MM130Wt < (MM152Wt + MM132Wt) Then
                ErrNum = 6
                GoTo MatrixSpaceLineDel_Error
            End If
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
        MDC_Com.MDC_GF_Message "첫라인에 반출문서-행 번호가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "반출문서-행 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 원재료사용중량은 원재료잔량중량보다 클수 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 출고중량은 0보다 커야만 합니다. 확인하세요.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "단가는 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "반출중량보다 입고중량이 큽니다. 확인하세요.", "E"
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
        If Trim(oDS_PS_MM152L.GetValue("U_OtDocLin", i)) = "" Then
            oDS_PS_MM152L.RemoveRecord i
        End If
    Next i
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Delete_EmptyRow_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Function Add_PS_PP040(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
On Error GoTo Add_PS_PP040_Error
    Dim i&, j&
    Dim ErrNum&, RowNum&
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim lTypeCount As Integer
    
    Dim BPLID$, JakNum$, OutQty As Long, OutWt As Double
    Dim NQty As Long, NWeight As Double '불량수량, 중량
    Dim PP040H_DocEntry$, PP030M_LineId$, PP030M_U_Sequence$, PP030M_U_CpCode$, PP030M_U_CpName$, DocType$, AutoKey&, DocDate$, OutGbn$, CpCode$, TCpCode$
    Dim iTemp01&, sTemp01$, ScrapWt As Currency
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Sbo_Company.StartTransaction
    oMat01.FlushToDataSource
    
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
    
'    '//스크랩 중량 계산
'    For i = 0 To oMat01.RowCount - 2
'        If Trim(oDS_PS_MM152L.GetValue("U_ReStatus", i)) = "Y" Then
''            sQry = "Select U_ItmBSort From [OITM] Where ItemCode = '" & Trim(oDS_PS_MM152L.GetValue("U_ItemCode", i)) & "'"
''            oRecordSet01.DoQuery sQry
''            If Trim(oRecordSet01.Fields(0).Value) = "102" Then
''            If Trim(oDS_PS_MM152L.GetValue("U_OutItmCd", i)) = "" Then
''                Call oDS_PS_MM152L.setValue("U_ScrapWt", i, 0)
''            Else
'                sQry = "Select  a.U_OutDoc, a.U_OutLine, IsNull(b.U_OutWt, 0) - Sum(IsNull(a.U_OutWt, 0)) - IsNull(b.U_ReWt, 0) "
'                sQry = sQry & "From [@PS_MM152L] a Inner Join [@PS_MM130L] b On a.U_OutDoc = b.U_OutDoc And a.U_OutLine = b.U_LineNum "
'                sQry = sQry & "Inner Join [@PS_MM152H] c On a.DocEntry = c.DocEntry "
'                sQry = sQry & "Where a.U_OutDoc = '" & Trim(oDS_PS_MM152L.GetValue("U_OutDoc", i)) & "' "
'                sQry = sQry & "And a.U_OutLine = '" & Trim(oDS_PS_MM152L.GetValue("U_OutLine", i)) & "' "
'                sQry = sQry & "And c.CanCeled = 'N' "
'                sQry = sQry & "Group by a.U_OutDoc, a.U_OutLine, b.U_OutWt, b.U_ReWt "
'                oRecordSet01.DoQuery sQry
'                ScrapWt = oRecordSet01.Fields(2).VALUE
'
'                Call oDS_PS_MM152L.setValue("U_ScrapWt", i, ScrapWt)
''            End If
'        ElseIf Trim(oDS_PS_MM152L.GetValue("U_ReStatus", i)) = "N" Then
'            Call oDS_PS_MM152L.setValue("U_ScrapWt", i, 0)
'        End If
'    Next i
    
    If Trim(oDS_PS_MM152H.GetValue("U_OKYNC", 0)) = "C" Then
        For i = 0 To oMat01.RowCount - 2
            If Trim(oDS_PS_MM152L.GetValue("U_OtDocLin", i)) <> "" And Trim(oDS_PS_MM152L.GetValue("U_PP040Doc", i)) <> "" Then
                OutQty = Trim(oDS_PS_MM152L.GetValue("U_OutQty", i)) + Trim(oDS_PS_MM152L.GetValue("U_NQty", i)) '출고수량 + 불량수량
                OutWt = Trim(oDS_PS_MM152L.GetValue("U_OutWt", i)) + Val(oDS_PS_MM152L.GetValue("U_NWeight", i)) '출고중량 + 불량중량
                
                
                sQry = "Update [@PS_PP040H] Set Status = 'C', Canceled = 'Y' Where DocEntry = '" & Trim(oDS_PS_MM152L.GetValue("U_PP040Doc", i)) & "'"
                oRecordSet01.DoQuery sQry
                
                sQry = "Update [@PS_MM130L] "
                sQry = sQry & "Set U_InQty = IsNull(U_InQty, 0) - " & OutQty & ", U_InWt = IsNull(U_InWt, 0) - " & OutWt & " "
                sQry = sQry & "From [@PS_MM130L] a Inner Join [@PS_MM130H] b On a.DocEntry = b.DocEntry "
                sQry = sQry & "Where b.U_OutDoc = '" & Trim(oDS_PS_MM152L.GetValue("U_OutDoc", i)) & "' "
                sQry = sQry & "And a.U_LineNum = '" & Trim(oDS_PS_MM152L.GetValue("U_OutLine", i)) & "' "
                oRecordSet01.DoQuery sQry
                         
                sQry = "Update [@PS_MM152L] Set U_ReStatus = 'N' From [@PS_MM152H] a Where [@PS_MM152L].DocEntry = a.DocEntry "
                sQry = sQry & "And a.Canceled = 'N' And a.U_OKYNC <> 'C' And [@PS_MM152L].DocEntry <> '" & Trim(oDS_PS_MM152L.GetValue("DocEntry", i)) & "' "
                sQry = sQry & "And [@PS_MM152L].U_OtDocLin = '" & Trim(oDS_PS_MM152L.GetValue("U_OtDocLin", i)) & "' "
                sQry = sQry & "And [@PS_MM152L].U_ReStatus = 'Y' "
                oRecordSet01.DoQuery sQry
                
                         
                Call oDS_PS_MM152L.setValue("U_PP040Doc", i, "")
            ElseIf Trim(oDS_PS_MM152L.GetValue("U_OtDocLin", i)) <> "" And Trim(oDS_PS_MM152L.GetValue("U_PP040Doc", i)) = "" Then
            End If
        Next i
    ElseIf Trim(oDS_PS_MM152H.GetValue("U_OKYNC", 0)) = "Y" Then
        For i = 0 To oMat01.RowCount - 2
            If Trim(oDS_PS_MM152L.GetValue("U_OtDocLin", i)) <> "" And Trim(oDS_PS_MM152L.GetValue("U_PP040Doc", i)) <> "" Then
            ElseIf Trim(oDS_PS_MM152L.GetValue("U_OtDocLin", i)) <> "" And Trim(oDS_PS_MM152L.GetValue("U_PP040Doc", i)) = "" And oDS_PS_MM152L.GetValue("U_OutQty", i) + oDS_PS_MM152L.GetValue("U_OutWt", i) <> 0 Then
'                DocDate = Format(Trim(oDS_PS_MM152H.GetValue("U_DocDate", 0)), "YYYYMMDD")
                DocDate = Trim(oDS_PS_MM152H.GetValue("U_DocDate", 0))
                
                JakNum = Trim(oDS_PS_MM152L.GetValue("U_JakNum", i))
                OutQty = Trim(oDS_PS_MM152L.GetValue("U_OutQty", i)) + CDbl(Trim(oDS_PS_MM152L.GetValue("U_Sample", i))) '외주수량 + 시료수량
                OutWt = Trim(oDS_PS_MM152L.GetValue("U_OutWt", i)) + CDbl(Trim(oDS_PS_MM152L.GetValue("U_Sample", i))) '외주중량 + 시료수량
                
                NQty = Trim(oDS_PS_MM152L.GetValue("U_NQty", i))
                NWeight = Trim(oDS_PS_MM152L.GetValue("U_NWeight", i))
                
                OutGbn = Trim(oDS_PS_MM152L.GetValue("U_OutGbn", i))
                CpCode = Trim(oDS_PS_MM152L.GetValue("U_CpCode", i))
                TCpCode = Trim(oDS_PS_MM152L.GetValue("U_TCpCode", i))
                
                sQry = "EXEC [PS_MM152_02] '" & JakNum & "','" & OutGbn & "', '10', '" & CpCode & "','" & TCpCode & "'"
                oRecordSet01.DoQuery sQry
                
                If oRecordSet01.RecordCount = 0 Then
                    ErrNum = 1
                    GoTo Add_PS_PP040_Error
                End If
                
                lTypeCount = 0
                Do Until oRecordSet01.EOF
                    ReDim Preserve PS_PP040(lTypeCount)
                    
                    PS_PP040(lTypeCount).PP030HNo = Trim(oRecordSet01.Fields("DocEntry").VALUE)
                    PS_PP040(lTypeCount).PP030MNo = Trim(oRecordSet01.Fields("LineId").VALUE)
                    PS_PP040(lTypeCount).OrdNum = JakNum
                    PS_PP040(lTypeCount).Sequence = Trim(oRecordSet01.Fields("U_Sequence").VALUE)
                    PS_PP040(lTypeCount).OrdGbn = Trim(oRecordSet01.Fields("U_OrdGbn").VALUE)
                    PS_PP040(lTypeCount).ItemCode = Trim(oRecordSet01.Fields("U_ItemCode").VALUE)
                    PS_PP040(lTypeCount).ItemName = Trim(oRecordSet01.Fields("U_ItemName").VALUE)
                    PS_PP040(lTypeCount).CpCode = Trim(oRecordSet01.Fields("U_CpCode").VALUE)
                    PS_PP040(lTypeCount).CpName = Trim(oRecordSet01.Fields("U_CpName").VALUE)
                    
                    PS_PP040(lTypeCount).Chk = "N"
                    
                    lTypeCount = lTypeCount + 1
                    oRecordSet01.MoveNext
                Loop
                            
                '// DocEntry
                sQry = "Select AutoKey From [ONNM] Where ObjectCode = 'PS_PP040'"
                oRecordSet01.DoQuery sQry
                PP040H_DocEntry = Trim(oRecordSet01.Fields("AutoKey").VALUE)
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
                sQry = sQry & "'" & BPLID & "',"
                sQry = sQry & "'" & PS_PP040(0).ItemCode & "',"
                sQry = sQry & "'" & PS_PP040(0).ItemName & "',"
                sQry = sQry & "'" & JakNum & "',"
                sQry = sQry & "'00',"
                sQry = sQry & "'000',"
                sQry = sQry & "'" & PS_PP040(0).PP030HNo & "',"
                sQry = sQry & "'" & DocType & "',"
                sQry = sQry & "'N',"
                sQry = sQry & "'" & DocDate & "'"
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
                        If iTemp01 > 1 Then
                            NQty = 0
                            NWeight = 0
                        End If
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
                        sQry = sQry & "'" & BPLID & "',"
                        sQry = sQry & "'" & PS_PP040(j).ItemCode & "',"
                        sQry = sQry & "'" & PS_PP040(j).ItemName & "',"
                        sQry = sQry & "'" & JakNum & "',"
                        sQry = sQry & "'000',"
                        sQry = sQry & "'00',"
                        sQry = sQry & "'" & PS_PP040(j).PP030HNo & "',"
                        sQry = sQry & "'" & PS_PP040(j).PP030MNo & "',"
                        sQry = sQry & "'" & OutQty + NQty & "',"
                        sQry = sQry & "'" & OutWt + NWeight & "',"
                        sQry = sQry & "'" & OutQty & "',"
                        sQry = sQry & "'" & OutWt & "',"
                        sQry = sQry & "'" & NQty & "',"
                        sQry = sQry & "'" & NWeight & "',"
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
                Next j
                sQry = "Update [ONNM] Set AutoKey = '" & AutoKey & "' Where ObjectCode = 'PS_PP040'"
                oRecordSet01.DoQuery sQry
                
                NQty = Trim(oDS_PS_MM152L.GetValue("U_NQty", i))
                NWeight = Trim(oDS_PS_MM152L.GetValue("U_NWeight", i))
                
                sQry = "Update [@PS_MM130L] "
                sQry = sQry & "Set U_InQty = IsNull(U_InQty, 0) + " & OutQty + NQty & ", U_InWt = IsNull(U_InWt, 0) + " & OutWt + NWeight & " "
                sQry = sQry & "From [@PS_MM130L] a Inner Join [@PS_MM130H] b On a.DocEntry = b.DocEntry "
                sQry = sQry & "Where b.U_OutDoc = '" & Trim(oDS_PS_MM152L.GetValue("U_OutDoc", i)) & "' "
                sQry = sQry & "And a.U_LineNum = '" & Trim(oDS_PS_MM152L.GetValue("U_OutLine", i)) & "' "
                oRecordSet01.DoQuery sQry
                
                Call oDS_PS_MM152L.setValue("U_PP040Doc", i, PS_PP040DocEntry(i).PP040DocEntry)
            End If
        Next i
    End If
    
    oMat01.LoadFromDataSource
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

Private Sub Print_Report01()
On Error GoTo Print_Report01_Error
    Dim ErrNum&
    Dim DocNum                  As String
    Dim WinTitle                As String
    Dim ReportName              As String
    Dim sQry                    As String
    Dim oRecordSet01            As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call ConnectODBC
    
'    If Trim(oForm01.Items("OKYNC").Specific.Value) <> "Y" Then
'        ErrNum = 3
'        GoTo Print_Report01_Error
'    End If
    
    DocNum = Trim(oForm01.Items("DocEntry").Specific.VALUE)
        
    If DocNum = "" Then
        ErrNum = 1
        GoTo Print_Report01_Error
    End If
        
    WinTitle = "거래명세표 [PS_MM152_01]"
    ReportName = "PS_MM152_01.rpt"
       
    '//Formula 수식필드
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    
    '//SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    '//조회조건문
    sQry = "EXEC [PS_MM152_03] '" & DocNum & "'"
    oRecordSet01.DoQuery sQry
    If oRecordSet01.RecordCount = 0 Then
        ErrNum = 2
        GoTo Print_Report01_Error
    End If
    
    '//CR Action
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "1", "N", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Report01_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Or ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다.확인해 주세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "승인되지 않은 문서는 거래명세표를 출력할 수 없습니다. 확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "Print_Report01_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub
