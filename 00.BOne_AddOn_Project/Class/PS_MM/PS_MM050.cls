VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_MM050"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_MM050.cls
'//  Module         : MM
'//  Description    : 가입고
'//  FormType       : PS_MM050
'//  Create Date    : 2010.08.26
'//  Modified Date  :
'//  Creator        : Ryu Yung Jo
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01   As String
Public oForm01           As SAPbouiCOM.Form
Public oMat01            As SAPbouiCOM.Matrix
Public oMat02            As SAPbouiCOM.Matrix
Private oDS_PS_MM050H    As SAPbouiCOM.DBDataSource    '등록헤더
Private oDS_PS_MM050L    As SAPbouiCOM.DBDataSource    '등록라인
Private oDS_PS_TEMPTABLE As SAPbouiCOM.DBDataSource

Private oLast_Item_UID   As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID    As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row    As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error

    Dim i           As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01   As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_MM050.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)

    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_MM050_" & GetTotalFormsCount
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
    
    Call FormResize
    
    FormItemEnabled
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1286"), True         '// 닫기
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

    Dim i                As Long
    Dim j                As Long
    Dim sQry             As String
    Dim ItemCode         As String
    Dim Qty              As Long
    Dim Calculate_Weight As Double
    Dim DocTotal         As Currency
    Dim SumQty           As Long
    Dim SumWeight        As Currency
    Dim TempForm01       As Variant
    
    Dim oRecordSet01     As SAPbobsCOM.Recordset
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
                        If oMat01.Columns("PQDocNum").Cells(pval.Row).Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    End If
                End If
                
'et_COMBO_SELECT ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            
                If pval.ItemUID = "Mat01" Then
                    If pval.Row > 0 Then
                        
                        Call oMat01.SelectRow(pval.Row, True, False)
                    
                    End If
                ElseIf pval.ItemUID = "Mat02" Then
                    If pval.Row > 0 Then
                        
                        Call oMat02.SelectRow(pval.Row, True, False)
                        
                        oLast_Item_UID = pval.ItemUID
                        oLast_Col_UID = pval.ColUID
                        oLast_Col_Row = pval.Row
                    
                    End If
                Else
                    
                    oLast_Item_UID = pval.ItemUID
                    oLast_Col_UID = ""
                    oLast_Col_Row = 0
                    
                End If
            
'et_DOUBLE_CLICK ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_DOUBLE_CLICK: '//7
                If pval.ItemUID = "Mat01" And pval.Row <> "0" Then 'And pval.ColUID = "LineNum" Then
                
                    oForm01.Freeze True
                    
                    If oMat02.VisualRowCount = 0 Then
                        oDS_PS_MM050L.Clear
                    End If
                    
                    '품의금액 VS 가입고 금액 비교(상단 메트릭스에 값을 조회 해놓은 상태에서 품의금액을 변경한 경우 체크)
                    
                    
                    j = 0
                    For i = 0 To oMat02.VisualRowCount - 1
'                        oMat02.LoadFromDataSource
                        If Trim(oDS_PS_MM050L.GetValue("U_PODocNum", i)) = Trim(oDS_PS_TEMPTABLE.GetValue("U_sField01", pval.Row - 1)) And _
                           Trim(oDS_PS_MM050L.GetValue("U_POLinNum", i)) = Trim(oDS_PS_TEMPTABLE.GetValue("U_iField02", pval.Row - 1)) Then
                            MDC_Com.MDC_GF_Message "같은 행을 두번 선택할 수 없습니다. 확인하세요." & Err.Number & " - " & Err.Description, "W"
                            j = 1
                        End If
                    Next
                                        
                    If j = 0 Then
                        oMat02.FlushToDataSource
                        If oMat02.VisualRowCount = 0 Then
                            Call Add_MatrixRow(oMat02.VisualRowCount, True, "Mat02")
                        Else
                            Call Add_MatrixRow(oMat02.VisualRowCount, False, "Mat02")
                        End If
                        
                        oMat02.Columns("PODocNum").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_TEMPTABLE.GetValue("U_sField01", pval.Row - 1))
                        oMat02.Columns("POLinNum").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_TEMPTABLE.GetValue("U_iField02", pval.Row - 1))
                        oMat02.Columns("ItemCode").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_TEMPTABLE.GetValue("U_sField03", pval.Row - 1))
                        oMat02.Columns("ItemName").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_TEMPTABLE.GetValue("U_sField04", pval.Row - 1))
'                        oMat02.Columns("FrgnName").Cells(oMat02.VisualRowCount).Specific.Value = Trim(oDS_PS_TEMPTABLE.GetValue("U_sField17", pval.Row - 1))
                        oMat02.Columns("OutSize").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_TEMPTABLE.GetValue("U_sField23", pval.Row - 1))
                        oMat02.Columns("OutUnit").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_TEMPTABLE.GetValue("U_sField24", pval.Row - 1))
                        oMat02.Columns("ItemGpCd").Cells(oMat02.VisualRowCount).Specific.Select Trim(oDS_PS_TEMPTABLE.GetValue("U_sField18", pval.Row - 1))
                        oMat02.Columns("POType").Cells(oMat02.VisualRowCount).Specific.Select Trim(oDS_PS_TEMPTABLE.GetValue("U_sField05", pval.Row - 1))
'                        oMat02.Columns("Weight").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_TEMPTABLE.GetValue("U_iField04", pval.Row - 1))
                        'oMat02.Columns("Qty").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_TEMPTABLE.GetValue("U_iField04", pval.Row - 1))
                        
                        Dim qField01 As Long
                        Dim iField04 As Long
                        
                        qField01 = Trim(oDS_PS_TEMPTABLE.GetValue("U_qField01", pval.Row - 1))
                        iField04 = Trim(oDS_PS_TEMPTABLE.GetValue("U_iField04", pval.Row - 1))
                        
                        If qField01 = iField04 Then '입고수량과 미입고 수량이 같을때 미입고 수량을 입고수량에 넣어준다
                            If Trim(oDS_PS_TEMPTABLE.GetValue("U_sField12", pval.Row - 1)) = "102" And (Trim(oDS_PS_TEMPTABLE.GetValue("U_sField18", pval.Row - 1)) = "30" Or Trim(oDS_PS_TEMPTABLE.GetValue("U_sField18", pval.Row - 1)) = "40") Then
                                '기계사업부 외주 제작과 외주가공은 입고중량을 입고수량에 넣어준다.
                                '중량이 소수점일 경우 INT로 컨버트(2014.06.05 송명규), CInt 다시 환원(2014.07.04 송명규)
'                                oMat02.Columns("Qty").Cells(oMat02.VisualRowCount).Specific.VALUE = CInt(Trim(oDS_PS_TEMPTABLE.GetValue("U_qField01", pval.Row - 1)))
                                oMat02.Columns("Qty").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_TEMPTABLE.GetValue("U_qField01", pval.Row - 1))
                            Else
                                oMat02.Columns("Qty").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_TEMPTABLE.GetValue("U_iField03", pval.Row - 1))
                            End If
                        'If Trim(oDS_PS_TEMPTABLE.GetValue("U_iField03", pval.Row - 1)) = Trim(oDS_PS_TEMPTABLE.GetValue("U_iField04", pval.Row - 1)) Then
                            
                            oMat02.Columns("Weight").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_TEMPTABLE.GetValue("U_qField01", pval.Row - 1))
                            oMat02.Columns("RealWt").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_TEMPTABLE.GetValue("U_qField01", pval.Row - 1))
                        
                        Else
'                            Calculate_Weight = MDC_PS_Common.Calculate_Weight(Trim(oDS_PS_TEMPTABLE.GetValue("U_sField03", pval.Row - 1)), Trim(oDS_PS_TEMPTABLE.GetValue("U_iField04", pval.Row - 1)), Trim(oForm01.Items("BPLId").Specific.Value))
'                            oMat02.Columns("RealWt").Cells(oMat02.VisualRowCount).Specific.Value = Calculate_Weight
                        End If
                        oMat02.Columns("UnWeight").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_TEMPTABLE.GetValue("U_qField02", pval.Row - 1))
                        oMat02.Columns("LinTotal").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_TEMPTABLE.GetValue("U_aField02", pval.Row - 1))
'                        oMat02.Columns("Price").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_TEMPTABLE.GetValue("U_aField01", pval.Row - 1))
                        
                        oMat02.Columns("WhsCode").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_TEMPTABLE.GetValue("U_sField12", pval.Row - 1))
                        oMat02.Columns("WhsName").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_TEMPTABLE.GetValue("U_sField13", pval.Row - 1))
                        oMat02.Columns("Auto").Cells(oMat02.VisualRowCount).Specific.Select Trim(oDS_PS_TEMPTABLE.GetValue("U_sField25", pval.Row - 1))
                        
                        If Trim(oDS_PS_TEMPTABLE.GetValue("U_sField18", pval.Row - 1)) = "30" Or Trim(oDS_PS_TEMPTABLE.GetValue("U_sField18", pval.Row - 1)) = "40" Then '가공비품의(30), 외주제작품의(40) (2011.10.28 송명규 수정)
                            oMat02.Columns("QEYesNo").Cells(oMat02.VisualRowCount).Specific.Select "Y"
                        Else '가공비품의(30), 외주제작품의(40) 외
                            oMat02.Columns("QEYesNo").Cells(oMat02.VisualRowCount).Specific.Select "N"
                        End If
    '                    oMat02.Columns("QEDate").Cells(oMat02.VisualRowCount).Specific.Value = Trim(oDS_PS_TEMPTABLE.getValue("U_sField01", pval.Row - 1))
                        oMat02.Columns("BDocNum").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_TEMPTABLE.GetValue("U_sField15", pval.Row - 1))
                        oMat02.Columns("BLineNum").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_TEMPTABLE.GetValue("U_sField16", pval.Row - 1))
                        
                        oMat02.FlushToDataSource
                        Call oDS_PS_MM050L.setValue("U_Price", oMat02.VisualRowCount - 1, Trim(oDS_PS_TEMPTABLE.GetValue("U_aField01", pval.Row - 1)))
                        oMat02.LoadFromDataSource
                        If qField01 <> iField04 Then
                            oMat02.Columns("Weight").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_TEMPTABLE.GetValue("U_iField04", pval.Row - 1))
                        End If
                        oMat02.Columns("Qty").Cells(oMat02.VisualRowCount).Click ct_Regular
                        j = 0
                    End If
                    oForm01.Freeze False
                    BubbleEvent = False
                    
                ElseIf pval.ItemUID = "Mat01" And pval.Row = 0 Then '컬럼 더블클릭 정렬 추가(2015.04.03 송명규, 정영욱 요청)
            
                    oMat01.Columns(pval.ColUID).TitleObject.Sortable = True
                    oMat01.FlushToDataSource
                    
                ElseIf pval.ItemUID = "Mat02" Then
                
                    If pval.ColUID = "QEYesNo" Then
                    
                        If pval.Row = 0 And oMat02.RowCount > 1 Then
                            
                            Dim SelectedValue As String
                            Dim loopCount As Integer
                            SelectedValue = oMat02.Columns("QEYesNo").Cells(1).Specific.Selected.VALUE '첫 행에 선택된 값을 저장
                            
                            For loopCount = 1 To oMat02.VisualRowCount
                            
                                Call oMat02.Columns("QEYesNo").Cells(loopCount).Specific.Select(SelectedValue)
                            
                            Next
                            
                        End If
                    
                    End If
                
                End If
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
            Case et_MATRIX_LOAD: '//11
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
'et_FORM_RESIZE//////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_FORM_RESIZE '//20
                
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
            
                If pval.ItemUID = "Mat02" Then
                    If pval.Row > 0 Then
                        oLast_Item_UID = pval.ItemUID
                        oLast_Col_UID = pval.ColUID
                        oLast_Col_Row = pval.Row
                    End If
                Else
                    oLast_Item_UID = pval.ItemUID
                    oLast_Col_UID = ""
                    oLast_Col_Row = 0
                End If
            
            Case et_LOST_FOCUS: '//4
'et_FORM_UNLOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_FORM_UNLOAD: '//17
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.EventType
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
                oForm01.Freeze True
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
                        oForm01.Mode = fm_OK_MODE
                        Call Sbo_Application.ActivateMenuItem("1282")
                    End If
                ElseIf pval.ItemUID = "Btn01" Then
                    If HeaderSpaceLineDel = False Then
                        oMat01.Clear
                        oDS_PS_TEMPTABLE.Clear
                        BubbleEvent = False
                        Exit Sub
                    End If
                    Call Display_MatrixData
                End If
                oForm01.Freeze False
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "BPLId" Or pval.ItemUID = "POType" Then
                        oMat01.Clear
                        oDS_PS_TEMPTABLE.Clear
                        oMat02.Clear
                        oDS_PS_MM050L.Clear
                    End If
                End If
            
            Case et_CLICK: '//6
                If pval.ItemUID = "RadioMat01" Then
                    oForm01.Freeze True
                    oForm01.Settings.MatrixUID = "Mat01"
                    oForm01.Settings.EnableRowFormat = True
                    oForm01.Settings.Enabled = True
                    oForm01.Freeze False
                ElseIf pval.ItemUID = "RadioMat02" Then
                    oForm01.Freeze True
                    oForm01.Settings.MatrixUID = "Mat02"
                    oForm01.Settings.EnableRowFormat = True
                    oForm01.Settings.Enabled = True
                    oForm01.Freeze False
                End If
'et_DOUBLE_CLICK ///////////'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_DOUBLE_CLICK: '//7
                If pval.ItemUID = "Mat01" And pval.Row <> "0" And pval.ColUID = "LineNum" Then
                    
                    oForm01.Freeze True
'                    oMat01.FlushToDataSource
'                    If Trim(oDS_PS_TEMPTABLE.getValue("U_sField08", 0)) = "" Or Trim(oDS_PS_TEMPTABLE.getValue("U_sField08", 0)) = "N" Then
'                        Admision = "Y"
'                    ElseIf Trim(oDS_PS_TEMPTABLE.getValue("U_sField08", 0)) = "Y" Then
'                        Admision = "N"
'                    End If
'                    For i = 0 To oMat01.VisualRowCount - 1
'                        oDS_PS_TEMPTABLE.setValue "U_sField08", i, Admision
'                    Next i
'                    oMat01.LoadFromDataSource
                    oForm01.Freeze False
                End If
            Case et_MATRIX_LINK_PRESSED '//8
                If pval.ItemUID = "Mat01" And pval.ColUID = "PODocNum" Then
                    Set TempForm01 = New PS_MM030
                    Call TempForm01.LoadForm(oMat01.Columns("PODocNum").Cells(pval.Row).Specific.VALUE)
                    Set TempForm01 = Nothing
                End If
'et_VALIDATE ///////////'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_VALIDATE: '//10
                oForm01.Freeze True
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "CardCode" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "CntcCode" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "Mat02" Then
                        If pval.ColUID = "Qty" Then
                            oMat02.FlushToDataSource
                            ItemCode = Trim(oDS_PS_MM050L.GetValue("U_ItemCode", pval.Row - 1))
                            Qty = oDS_PS_MM050L.GetValue("U_Qty", pval.Row - 1)
                            
                            Calculate_Weight = MDC_PS_Common.Calculate_Weight(ItemCode, Qty, Trim(oForm01.Items("BPLId").Specific.VALUE))
                            If oForm01.Mode = fm_ADD_MODE Then '//추가모드일때
                                If Trim(oDS_PS_TEMPTABLE.GetValue("U_sField12", pval.Row - 1)) = "102" And (Trim(oDS_PS_TEMPTABLE.GetValue("U_sField18", pval.Row - 1)) = "30" Or Trim(oDS_PS_TEMPTABLE.GetValue("U_sField18", pval.Row - 1)) = "40") Then
                                    '기계사업부 외주가공, 외주제작 중량은 수량과 동일하다.
                                    Calculate_Weight = Qty
                                    oMat02.Columns("RealWt").Cells(pval.Row).Specific.VALUE = Calculate_Weight
                                End If
                            Else
                                If oDS_PS_MM050L.GetValue("U_WhsCode", pval.Row - 1) = "102" And (oDS_PS_MM050L.GetValue("U_ItemGpCd", pval.Row - 1) = "30" Or oDS_PS_MM050L.GetValue("U_ItemGpCd", pval.Row - 1) = "40") Then
                                    Calculate_Weight = Qty
                                    oMat02.Columns("RealWt").Cells(pval.Row).Specific.VALUE = Calculate_Weight
                                End If
                            End If
                            oMat02.Columns("Weight").Cells(pval.Row).Specific.VALUE = Calculate_Weight
                            oMat02.Columns("RealWt").Cells(pval.Row).Specific.VALUE = Calculate_Weight
'                            oDS_PS_MM050L.setValue "U_Weight", pval.Row - 1, Calculate_Weight

                            sQry = "Select ItmsGrpCod From OITM Where ItemCode = '" & ItemCode & "'"
                            oRecordSet01.DoQuery sQry
                            If Trim(oRecordSet01.Fields(0).VALUE) = "105" Then '부자재
                                oMat02.Columns("RealWt").Cells(pval.Row).Specific.VALUE = Calculate_Weight
'                                oDS_PS_MM050L.setValue "U_RealWt", pval.Row - 1, Calculate_Weight
                            End If
'                            oMat02.LoadFromDataSource
'                            oMat02.Columns("Weight").Cells(pval.Row).Click ct_Regular
                        ElseIf pval.ColUID = "Price" Or pval.ColUID = "Weight" Or pval.ColUID = "RealWt" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        ElseIf pval.ColUID = "LinTotal" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        ElseIf pval.ColUID = "WhsCode" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        End If
                    End If
                End If
                oForm01.Freeze False
            Case et_MATRIX_LOAD: '//11
                For i = 0 To oMat02.VisualRowCount - 1
                    DocTotal = DocTotal + oMat02.Columns("LinTotal").Cells(i + 1).Specific.VALUE
                    If oMat02.Columns("Qty").Cells(i + 1).Specific.VALUE = "" Then
                        SumQty = SumQty
                    Else
                        SumQty = SumQty + oMat02.Columns("Qty").Cells(i + 1).Specific.VALUE
                    End If
                    SumWeight = SumWeight + oMat02.Columns("RealWt").Cells(i + 1).Specific.VALUE
                Next i
                oForm01.Items("DocTotal").Specific.VALUE = DocTotal
                oForm01.Items("SumQty").Specific.VALUE = SumQty
                oForm01.Items("SumWeight").Specific.VALUE = SumWeight
                
                Call oMat01.AutoResizeColumns
                Call oMat02.AutoResizeColumns
                
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
'et_FORM_RESIZE /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_FORM_RESIZE '//20
                Call FormResize
                
                
                
                
'                oMat01.AutoResizeColumns
'                oMat02.AutoResizeColumns
'                oForm.Items("56").Left = oForm.Items("Mat02").Left
'
'                oMat01.Columns("SpecCode").Width = 80
'                oMat01.Columns("Stext").Width = 150
'                oMat01.Columns("Sdivtext").Width = 80
'                oMat01.Columns("Min").Width = 80
'                oMat01.Columns("Max").Width = 80
'
'                oMat02.Columns("ProcCode").Width = 100
'                oMat02.Columns("Actual").Width = 35
'                oMat02.Columns("Issue").Width = 35
'                oMat02.Columns("Receipt").Width = 35
'                oMat02.Columns("Subc").Width = 35
'                oMat02.Columns("RecProc").Width = 100
'                oMat02.Columns("Comments").Width = 50
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
    If oForm01 Is Nothing Then
    Else
        oForm01.Freeze False
    End If
    MDC_Com.MDC_GF_Message "Raise_ItemEvent_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_MenuEvent_Error
    Dim i&
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
                If oForm01.Mode = fm_OK_MODE Then
                    If (PS_MM050_Validate("취소") = False) Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    If Sbo_Application.MessageBox("정말로 취소하시겠습니까?", "1", "예", "아니오") <> "1" Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                Else
                    MDC_Com.MDC_GF_Message "현재 모드에서는 취소할수 없습니다.", "W"
                    BubbleEvent = False
                    Exit Sub
                End If
            Case "1286": '닫기
                If oForm01.Mode = fm_OK_MODE Then
                    If (PS_MM050_Validate("닫기") = False) Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    If Sbo_Application.MessageBox("정말로 닫기하시겠습니까?", "1", "예", "아니오") <> "1" Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                Else
                    MDC_Com.MDC_GF_Message "현재 모드에서는 닫기할수 없습니다.", "W"
                    BubbleEvent = False
                    Exit Sub
                End If
            Case "1293": '행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
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
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            
'                If oMat02.RowCount <> oMat02.VisualRowCount Then
'                    For i = 0 To oMat02.VisualRowCount - 1
'                        oMat02.Columns("LineNum").Cells(i + 1).Specific.VALUE = i + 1
'                    Next i
'
'                    oMat01.FlushToDataSource
'                    oDS_PS_MM050L.RemoveRecord oDS_PS_MM050L.Size - 1       '// Mat01에 마지막라인(빈라인) 삭제
'                    oMat02.Clear
'                    oMat02.LoadFromDataSource
'                End If
            
            Case "1281": '찾기
                oForm01.Freeze True
                FormItemEnabled
                '//아이디별 사업장 세팅
                oForm01.Items("BPLId").Specific.Select MDC_PS_Common.User_BPLId, psk_ByValue
                '//아이디별 사번 세팅
                If MDC_PS_Common.User_SuperUserYN = "N" Then '수퍼유저인 경우는 사번 미표기
                    oForm01.Items("CntcCode").Specific.VALUE = MDC_PS_Common.User_MSTCOD
                End If
                
                oForm01.Items("POStatus").Specific.Select "Y"
                
                oForm01.Items("DocNum").Click ct_Regular
                oForm01.Freeze False
            Case "1282": '추가
                oForm01.Freeze True
                Call Initialization
                Call FormItemEnabled
                Call FormClear
                oDS_PS_MM050H.setValue "U_DocDate", 0, Format(Now, "YYYYMMDD")
                oForm01.Items("CardCode").Click ct_Collapsed
                oForm01.Freeze False
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                oForm01.Freeze True
                oMat01.FlushToDataSource
                oMat01.Clear
                oDS_PS_TEMPTABLE.Clear
                oMat01.LoadFromDataSource
                FormItemEnabled
                oForm01.Freeze False
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

Public Sub Raise_RightClickEvent(ByRef FormUID As String, ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_RightClickEvent_Error
    If (eventInfo.BeforeAction = True) Then
        '//작업
    ElseIf (eventInfo.BeforeAction = False) Then
        '//작업
    End If
    
    If eventInfo.ItemUID = "Mat02" Then
        If eventInfo.Row > 0 Then
            oLast_Item_UID = eventInfo.ItemUID
            oLast_Col_UID = eventInfo.ColUID
            oLast_Col_Row = eventInfo.Row
        End If
    Else
        oLast_Item_UID = eventInfo.ItemUID
        oLast_Col_UID = ""
        oLast_Col_Row = 0
    End If
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub CreateItems()
On Error GoTo CreateItems_Error
    '//디비데이터 소스 개체 할당
    Set oDS_PS_TEMPTABLE = oForm01.DataSources.DBDataSources("@PS_TEMPTABLE")
    Set oDS_PS_MM050H = oForm01.DataSources.DBDataSources("@PS_MM050H")
    Set oDS_PS_MM050L = oForm01.DataSources.DBDataSources("@PS_MM050L")
        
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    Set oMat02 = oForm01.Items("Mat02").Specific
    
    oMat01.SelectionMode = ms_NotSupported
    oMat02.SelectionMode = ms_NotSupported
    
    Call oForm01.DataSources.UserDataSources.Add("DueDateFr", dt_DATE, 8)
    oForm01.Items("DueDateFr").Specific.DataBind.SetBound True, "", "DueDateFr"
'    oForm01.DataSources.UserDataSources.Item("DueDateFr").Value = Format(Now, "YYYY-MM") & "-01"
    
    Call oForm01.DataSources.UserDataSources.Add("DueDateTo", dt_DATE, 8)
    oForm01.Items("DueDateTo").Specific.DataBind.SetBound True, "", "DueDateTo"
'    oForm01.DataSources.UserDataSources.Item("DueDateTo").Value = Date
    
    oDS_PS_MM050H.setValue "U_DocDate", 0, Format(Now, "yyyymmdd")
    
    Call oForm01.DataSources.UserDataSources.Add("DocTotal", dt_SUM)
    oForm01.Items("DocTotal").Specific.DataBind.SetBound True, "", "DocTotal"
'    oForm01.DataSources.UserDataSources.Item("DocDateFr").Value = 0

    Call oForm01.DataSources.UserDataSources.Add("SumQty", dt_SUM)
    oForm01.Items("SumQty").Specific.DataBind.SetBound True, "", "SumQty"
    
    Call oForm01.DataSources.UserDataSources.Add("SumWeight", dt_QUANTITY)
    oForm01.Items("SumWeight").Specific.DataBind.SetBound True, "", "SumWeight"
    
    oForm01.DataSources.UserDataSources.Add "RadioMat01", dt_SHORT_TEXT, 1
    oForm01.Items("RadioMat01").Specific.DataBind.SetBound True, "", "RadioMat01"
    
    oForm01.DataSources.UserDataSources.Add "RadioMat02", dt_SHORT_TEXT, 1
    oForm01.Items("RadioMat02").Specific.DataBind.SetBound True, "", "RadioMat02"
    
    oForm01.Items("RadioMat01").Specific.GroupWith "RadioMat02"
    
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
    sQry = "SELECT BPLId, BPLName From [OBPL] order by 1"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    '//품목구분
    sQry = "SELECT Code, Name From [@PSH_ORDTYP] Order by Code"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oMat01.Columns("ItemGpCd").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oMat02.Columns("ItemGpCd").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    '//품의형태
    Set oCombo = oForm01.Items("POType").Specific
    sQry = "SELECT Code, Name From [@PSH_RETYPE]"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oMat01.Columns("POType").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oMat02.Columns("POType").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
        
    '//품의상태
    Set oCombo = oForm01.Items("POStatus").Specific
    oCombo.ValidValues.Add "Y", "승인"
'    oCombo.ValidValues.Add "N", "미승인"
    oCombo.Select 0, psk_Index
   
    '//품질검수여부
    oMat02.Columns("QEYesNo").ValidValues.Add "Y", "Yes"
    oMat02.Columns("QEYesNo").ValidValues.Add "N", "No"
    
    '//지불조건(2013.07.25 송명규 추가)
'    sQry = "Select U_Minor, U_CdName From [@PS_SY001L] Where Code = 'M006' Order by U_Minor"
'    oRecordSet01.DoQuery sQry
'    Do Until oRecordSet01.EOF
'        oMat01.Columns("Payment").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
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
    
    '//아이디별 사업장 세팅
    Set oCombo = oForm01.Items("BPLId").Specific
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
    '//아이디별 사번 세팅
    oForm01.Items("CntcCode").Specific.VALUE = MDC_PS_Common.User_MSTCOD
    
    '//아이디별 부서 세팅
'    Set oCombo = oForm01.Items("DeptCode").Specific
'    oCombo.Select MDC_PS_Common.User_DeptCode, psk_ByValue

    Set oCombo = oForm01.Items("POStatus").Specific
    oCombo.Select 0, psk_Index     '승인(Y)고정
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
        oForm01.Items("POStatus").Enabled = True
        oForm01.Items("DocDate").Enabled = True
                
        oForm01.Items("Mat02").Enabled = True
        oMat02.Columns("Qty").Editable = True
        oMat02.Columns("Weight").Editable = True
        oMat02.Columns("RealWt").Editable = True
        oMat02.Columns("UnWeight").Editable = True
        oMat02.Columns("Price").Editable = True
        oMat02.Columns("LinTotal").Editable = True
        oMat02.Columns("WhsCode").Editable = True
        oMat02.Columns("QEYesNo").Editable = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        oForm01.Items("DocNum").Enabled = True
        
        oForm01.Items("CardCode").Enabled = True
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("CntcCode").Enabled = True
        oForm01.Items("POType").Enabled = True
        oForm01.Items("POStatus").Enabled = True
        oForm01.Items("DocDate").Enabled = True
                
        oForm01.Items("Mat02").Enabled = False
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        oForm01.Items("DocNum").Enabled = False
        
        oForm01.Items("CardCode").Enabled = True
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("CntcCode").Enabled = True
        oForm01.Items("POType").Enabled = True
        oForm01.Items("POStatus").Enabled = True
        oForm01.Items("DocDate").Enabled = True
                
        oForm01.Items("Mat02").Enabled = True
        oMat02.Columns("Qty").Editable = True
        oMat02.Columns("Weight").Editable = True
        oMat02.Columns("RealWt").Editable = True
        oMat02.Columns("UnWeight").Editable = True
        oMat02.Columns("Price").Editable = True
        oMat02.Columns("LinTotal").Editable = True
        oMat02.Columns("WhsCode").Editable = True
        oMat02.Columns("QEYesNo").Editable = True
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub FormClear()
On Error GoTo FormClear_Error
    Dim DocNum As String
    DocNum = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM050'", "")
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

Sub Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean, Optional ItemUID$)
On Error GoTo Add_MatrixRow_Error
    If ItemUID = "Mat01" Then
        If RowIserted = False Then '//행추가여부
            oDS_PS_TEMPTABLE.InsertRecord (oRow)
        End If
        oMat01.AddRow
        oDS_PS_TEMPTABLE.Offset = oRow
        oDS_PS_TEMPTABLE.setValue "U_LineNum", oRow, oRow + 1
        oMat01.LoadFromDataSource
    ElseIf ItemUID = "Mat02" Then
        If RowIserted = False Then '//행추가여부
            oDS_PS_MM050L.InsertRecord (oRow)
        End If
        oMat02.AddRow
        oDS_PS_MM050L.Offset = oRow
        oDS_PS_MM050L.setValue "U_LineNum", oRow, oRow + 1
        oMat02.LoadFromDataSource
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error

    Dim i            As Long
    Dim ErrNum       As Integer
    Dim sQry         As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim sRow         As Long
    Dim sSeq         As String
    Dim DocTotal     As Currency
    Dim SumQty       As Long
    Dim SumWeight    As Currency
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sRow = oRow
    
    Select Case oUID
        Case "CardCode"
            sQry = "Select CardName From OCRD Where CardCode = '" & Trim(oForm01.Items("CardCode").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            oForm01.Items("CardName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            
            oMat01.Clear
            oDS_PS_TEMPTABLE.Clear
            oMat02.Clear
            oDS_PS_MM050L.Clear
            
        Case "CntcCode"
            sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" & Trim(oDS_PS_MM050H.GetValue("U_CntcCode", 0)) & "'"
            oRecordSet01.DoQuery sQry
            
            oDS_PS_MM050H.setValue "U_CntcName", 0, Trim(oRecordSet01.Fields(0).VALUE)
            
        Case "Mat01"
            oForm01.Freeze True
            
            If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 0) And Trim(oMat01.Columns("PQDocNum").Cells(oRow).Specific.VALUE) <> "" Then
                oMat01.FlushToDataSource
                Call Add_MatrixRow(oMat01.RowCount, False)
                oMat01.Columns("PQDocNum").Cells(oRow).Click ct_Regular
            End If
            
            oMat01.FlushToDataSource
            
            sQry = "Select a.DocNum, b.LineId, b.U_ItemCode " 'U_LineNum를 LineId로 수정(2012.07.24 송명규)
            sQry = sQry & "From [@PS_MM010H] a Inner Join [@PS_MM010L] b On a.DocEntry = b.DocEntry "
            sQry = sQry & "Where a.DocNum = '" & Trim(oDS_PS_MM050L.GetValue("U_PQDocNum", oRow - 1)) & "' "
            sQry = sQry & "And a.Status = 'O'"
            oRecordSet01.DoQuery sQry
            
'            If oRecordSet01.RecordCount = 0 Then
''                oDS_PS_MM030L.setValue "U_PQDocNum", sRow - 1, Trim(oRecordSet01.Fields(0).Value)
'                oMat01.LoadFromDataSource
'                ErrNum = 1
'                GoTo FlushToItemValue_Error
'            End If
            Do Until oRecordSet01.EOF
                sSeq = "Y"
                For i = 0 To oMat01.VisualRowCount - 2
                    If Trim(oDS_PS_MM050L.GetValue("U_PQDocNum", i)) = Trim(oRecordSet01.Fields(0).VALUE) And _
                       Trim(oDS_PS_MM050L.GetValue("U_PQLinNum", i)) = Trim(oRecordSet01.Fields(1).VALUE) Then
                        sSeq = "N"
                    End If
                Next i
                If sSeq = "Y" Then
                    oDS_PS_MM050L.setValue "U_PQDocNum", sRow - 1, Trim(oRecordSet01.Fields(0).VALUE)
                    oDS_PS_MM050L.setValue "U_PQLinNum", sRow - 1, Trim(oRecordSet01.Fields(1).VALUE)
'                    oMat01.Columns("PQLinNum").Cells(oRow).Specific.Value = Trim(oRecordSet01.Fields(1).Value)
                    Call Add_MatrixRow(sRow, False)
                    sRow = sRow + 1
                End If
                oRecordSet01.MoveNext
            Loop
   
            If oMat01.VisualRowCount > 0 Then
                If Trim(oDS_PS_MM050L.GetValue("U_ItemCode", oMat01.VisualRowCount - 1)) = "" Then
                    oDS_PS_MM050L.RemoveRecord oMat01.VisualRowCount - 1
                End If
            End If
            
            oMat01.LoadFromDataSource
            
            oForm01.Freeze False
        Case "Mat02"
            If oCol = "Price" Or oCol = "Weight" Then
               oMat02.FlushToDataSource
               If Trim(oMat02.Columns("Price").Cells(oRow).Specific.VALUE) = 0 Then
                   oDS_PS_MM050L.setValue "U_LinTotal", oRow - 1, 0
               Else
                   '//금액 반올림 2012/04/03 노근용 수정
                   oDS_PS_MM050L.setValue "U_LinTotal", oRow - 1, Round(Trim(oMat02.Columns("Weight").Cells(oRow).Specific.VALUE) * Trim(oMat02.Columns("Price").Cells(oRow).Specific.VALUE), 0)
               End If
               oMat02.LoadFromDataSource
               oMat02.Columns("Price").Cells(oRow).Click
               
               For i = 0 To oMat02.VisualRowCount - 1
                    DocTotal = DocTotal + oMat02.Columns("LinTotal").Cells(i + 1).Specific.VALUE
                    If oMat02.Columns("Qty").Cells(i + 1).Specific.VALUE = "" Then
                        SumQty = SumQty
                    Else
                        SumQty = SumQty + oMat02.Columns("Qty").Cells(i + 1).Specific.VALUE
                    End If
                    SumWeight = SumWeight + oMat02.Columns("RealWt").Cells(i + 1).Specific.VALUE
                Next i
                oForm01.Items("DocTotal").Specific.VALUE = DocTotal
                oForm01.Items("SumQty").Specific.VALUE = SumQty
                oForm01.Items("SumWeight").Specific.VALUE = SumWeight
                
            ElseIf oCol = "LinTotal" Then
               oMat02.FlushToDataSource
               If Trim(oMat02.Columns("LinTotal").Cells(oRow).Specific.VALUE) = 0 Then
                   oDS_PS_MM050L.setValue "U_Price", oRow - 1, 0
               Else
                   If Trim(oMat02.Columns("Weight").Cells(oRow).Specific.VALUE) = 0 Then
                       oDS_PS_MM050L.setValue "U_Price", oRow - 1, 0
                   Else
                       oDS_PS_MM050L.setValue "U_Price", oRow - 1, Trim(oMat02.Columns("LinTotal").Cells(oRow).Specific.VALUE) / Trim(oMat02.Columns("Weight").Cells(oRow).Specific.VALUE)
                   End If
               End If
               oMat02.LoadFromDataSource
               
               For i = 0 To oMat02.VisualRowCount - 1
                    DocTotal = DocTotal + oMat02.Columns("LinTotal").Cells(i + 1).Specific.VALUE
                    If oMat02.Columns("Qty").Cells(i + 1).Specific.VALUE = "" Then
                        SumQty = SumQty
                    Else
                        SumQty = SumQty + oMat02.Columns("Qty").Cells(i + 1).Specific.VALUE
                    End If
                    SumWeight = SumWeight + oMat02.Columns("RealWt").Cells(i + 1).Specific.VALUE
                Next i
                oForm01.Items("DocTotal").Specific.VALUE = DocTotal
                oForm01.Items("SumQty").Specific.VALUE = SumQty
                oForm01.Items("SumWeight").Specific.VALUE = SumWeight
                
               oMat02.Columns("Price").Cells(oRow).Click
            ElseIf oCol = "RealWt" Then
                For i = 0 To oMat02.VisualRowCount - 1
                    DocTotal = DocTotal + oMat02.Columns("LinTotal").Cells(i + 1).Specific.VALUE
                    If oMat02.Columns("Qty").Cells(i + 1).Specific.VALUE = "" Then
                        SumQty = SumQty
                    Else
                        SumQty = SumQty + oMat02.Columns("Qty").Cells(i + 1).Specific.VALUE
                    End If
                    SumWeight = SumWeight + oMat02.Columns("RealWt").Cells(i + 1).Specific.VALUE
                Next i
                oForm01.Items("DocTotal").Specific.VALUE = DocTotal
                oForm01.Items("SumQty").Specific.VALUE = SumQty
                oForm01.Items("SumWeight").Specific.VALUE = SumWeight
            ElseIf oCol = "WhsCode" Then
                sQry = "Select WhsName From [OWHS] Where WhsCode = '" & Trim(oMat02.Columns("WhsCode").Cells(oRow).Specific.VALUE) & "'"
                oRecordSet01.DoQuery sQry
            
                oMat02.Columns("WhsName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            End If
    End Select

    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FlushToItemValue_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "구매견적문서가 취소되었거나 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "E"
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
        Case oDS_PS_MM050H.GetValue("U_CardCode", 0) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_MM050H.GetValue("U_BPLId", 0) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_MM050H.GetValue("U_CntcCode", 0) = ""
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_MM050H.GetValue("U_POType", 0) = ""
            ErrNum = 4
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_MM050H.GetValue("U_POStatus", 0) = ""
            ErrNum = 5
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_MM050H.GetValue("U_DocDate", 0) = ""
            ErrNum = 6
            GoTo HeaderSpaceLineDel_Error
    End Select

'    '마감상태 체크_S(2017.11.23 송명규 추가)
'    If MDC_PS_Common.Check_Finish_Status(oForm01.Items("BPLId").Specific.VALUE, oForm01.Items("DocDate").Specific.VALUE) = False Then
'        ErrNum = 7
'        GoTo HeaderSpaceLineDel_Error
'    End If
'    '마감상태 체크_E(2017.11.23 송명규 추가)

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    oForm01.Freeze False
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "거래처는 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "사업장은 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "담당자는 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "품의형태는 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "품의상태은 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "전기일은 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 7 Then
        MDC_Com.MDC_GF_Message "마감상태가 잠금입니다. 해당 일자로 등록할 수 없습니다. 전기일자를 확인하고, 회계부서로 문의하세요.", "E"
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
    Dim ErrRowCount     As Long
    
    Dim PODocNum        As String
    Dim POLinNum        As String
    
    '//수량체크
    Dim RealWt          As Currency
    Dim PendNum         As Currency
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    ErrNum = 0
    
    oMat01.FlushToDataSource

    '// 라인
    If oMat02.VisualRowCount = 0 Then
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    Else
        '//품의수량과 입고수량 Check
        
        
'        For i = 1 To oMat02.VisualRowCount
'
'            RealWt = oMat02.Columns("RealWt").Cells(i).Specific.VALUE
'
'            If RealWt <= 0 Then
'                ErrNum = 3
'                GoTo MatrixSpaceLineDel_Error
'            End If
'
'            '//창원의 원자재만 제외
'            If Trim(oMat02.Columns("WhsCode").Cells(i).Specific.VALUE) = "101" And Trim(oMat02.Columns("ItemGpCd").Cells(i).Specific.VALUE) = "10" Then
'
'            Else
'                PODocNum = Trim(oMat02.Columns("PODocNum").Cells(i).Specific.VALUE)
'                POLinNum = Trim(oMat02.Columns("POLinNum").Cells(i).Specific.VALUE)
'
'                '//품의수량
'                sQry = "Select  sum(b.U_Weight) - Sum(ISNULL(c.U_Weight, 0)) + Sum(ISNULL(d.U_BadQty, 0)) As PendNum"
'                sQry = sQry & " From    [@PS_MM030H] a"
'                sQry = sQry & " Inner Join [@PS_MM030L] b On a.DocEntry = b.DocEntry"
'                sQry = sQry & " Left  Join (Select b.U_PODocNum, b.U_POLinNum, Sum(b.U_Qty) As U_Qty, Sum(b.U_Weight) As U_Weight"
'                sQry = sQry & " From  [@PS_MM050H] a Inner Join [@PS_MM050L] b On a.DocEntry = b.DocEntry"
'                sQry = sQry & " Where  a.Canceled = 'N' Group by b.U_PODocNum, b.U_POLinNum ) c On c.U_PODocNum = a.DocNum And c.U_POLinNum = b.U_LineNum"
'                sQry = sQry & " Left  Join (Select  Left(b.U_GADocLin, CharIndex('-', b.U_GADocLin) - 1) As GADocNum"
'                sQry = sQry & " ,Right(b.U_GADocLin, Len(b.U_GADocLin) - CharIndex('-', b.U_GADocLin)) As GADocLin"
'                sQry = sQry & " , b.U_BadQty"
'                sQry = sQry & " From  [@PS_MM070H] a Inner Join [@PS_MM070L] b On a.DocEntry = b.DocEntry"
'                sQry = sQry & " Where  a.Canceled = 'N' ) d On d.GADocNum = c.DocEntry And d.GADocLin = c.U_LineNum"
'                sQry = sQry & " Where   b.DocEntry = '" & PODocNum & "'"
'                sQry = sQry & " And   b.U_LineNum = '" & POLinNum & "'"
'
'                oRecordset.DoQuery sQry
'
'                RealWt = oMat02.Columns("RealWt").Cells(i).Specific.VALUE
'                PendNum = oRecordset.Fields("PendNum").VALUE '//입고잔량
'                '//입고잔량보다 더 들어올수 없다.
'                If RealWt > PendNum Then
'                    ErrNum = 2
'                    GoTo MatrixSpaceLineDel_Error
'                End If
'            End If
'        Next

        For i = 1 To oMat02.VisualRowCount
        
            If oMat02.Columns("Qty").Cells(i).Specific.VALUE = "" Then '수량 체크(수량에 빈값이면 입력 불가)(2012.01.04 송명규 추가)
            
                ErrNum = 4
                ErrRowCount = i
                GoTo MatrixSpaceLineDel_Error
            
            ElseIf PS_MM050_CheckDate(oMat02.Columns("PODocNum").Cells(i).Specific.VALUE) = False Then '구매품의와 일자 체크
            
                ErrNum = 5
                ErrRowCount = i
                GoTo MatrixSpaceLineDel_Error
            
            End If
        
        Next

    End If
    
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
        MDC_Com.MDC_GF_Message "품의수량보다 입고수량이 초과됩니다. 확인바랍니다.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "실중량은 0이거나 0보다 작을수 없습니다. 확인바랍니다.", "E"
        
    ElseIf ErrNum = 4 Then '수량 체크(2012.01.04 송명규 추가)
    
        Call MDC_Com.MDC_GF_Message(ErrRowCount & "행의 수량에 값이 없습니다. 확인바랍니다.", "E")
        Call oMat02.Columns("Qty").Cells(ErrRowCount).Click(ct_Regular)
        
    ElseIf ErrNum = 5 Then
    
        Call Sbo_Application.MessageBox(ErrRowCount & "행 [" & oMat01.Columns("ItemCode").Cells(ErrRowCount).Specific.VALUE & "]의 가입고일은 구매품의일과 같거나 늦어야합니다. 확인하십시오." & Chr(13) & "해당 가입고는 전체가 등록되지 않습니다.", 1)

    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function

Private Function Display_MatrixData() As Boolean
On Error GoTo Display_MatrixData_Error
    Dim i               As Long
    Dim ErrNum          As Integer
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    Dim sQry            As String
    
    Dim CardCode$, BPLID$, CntcCode$, POType$, POStatus$, DueDateFr$, DueDateTo$
    Dim sCnt&
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    ErrNum = 0
    
    CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
    CntcCode = Trim(oForm01.Items("CntcCode").Specific.VALUE)
    POType = Trim(oForm01.Items("POType").Specific.VALUE)
    POStatus = Trim(oForm01.Items("POStatus").Specific.VALUE)
    DueDateFr = Trim(oForm01.Items("DueDateFr").Specific.VALUE)
    DueDateTo = Trim(oForm01.Items("DueDateTo").Specific.VALUE)
    
    If CardCode = "" Then CardCode = "%"
    If BPLID = "" Then BPLID = "%"
    If CntcCode = "" Then CntcCode = "%"
    If POType = "" Then POType = "%"
    If POStatus = "" Then POStatus = "%"
    If DueDateFr = "" Then DueDateFr = "20100101"
    If DueDateTo = "" Then DueDateTo = "20991231"
    
    sQry = "EXEC [PS_MM050_01] '" & CardCode & "', '" & BPLID & "', '" & CntcCode & "', '" & POType & "', "
    sQry = sQry & "'" & POStatus & "', '" & DueDateFr & "', '" & DueDateTo & "'"
    oRecordSet01.DoQuery sQry
    
    oMat01.Clear
    oDS_PS_TEMPTABLE.Clear
    
    If (oRecordSet01.RecordCount = 0) Then
        MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
        Set oRecordSet01 = Nothing
        Exit Function
    End If
    
    oForm01.Freeze True
    sCnt = 0
    If oRecordSet01.RecordCount > 0 Then
        Do Until oRecordSet01.EOF
            oDS_PS_TEMPTABLE.InsertRecord (sCnt)
            oDS_PS_TEMPTABLE.Offset = sCnt
            oDS_PS_TEMPTABLE.setValue "U_iField01", sCnt, sCnt + 1
            oDS_PS_TEMPTABLE.setValue "U_sField01", sCnt, Trim(oRecordSet01.Fields(0).VALUE)
            oDS_PS_TEMPTABLE.setValue "U_iField02", sCnt, Trim(oRecordSet01.Fields(1).VALUE)
            oDS_PS_TEMPTABLE.setValue "U_iField06", sCnt, Trim(oRecordSet01.Fields("EBELN").VALUE)
            
            oDS_PS_TEMPTABLE.setValue "U_sField03", sCnt, Trim(oRecordSet01.Fields(2).VALUE)
            oDS_PS_TEMPTABLE.setValue "U_sField04", sCnt, Trim(oRecordSet01.Fields(3).VALUE)
'            oDS_PS_TEMPTABLE.setValue "U_sField17", sCnt, Trim(oRecordSet01.Fields("U_FrgnName").Value)
            oDS_PS_TEMPTABLE.setValue "U_sField18", sCnt, Trim(oRecordSet01.Fields("U_ItemGpCd").VALUE)
            oDS_PS_TEMPTABLE.setValue "U_sField05", sCnt, Trim(oRecordSet01.Fields(4).VALUE)
            oDS_PS_TEMPTABLE.setValue "U_dField01", sCnt, Format(Trim(oRecordSet01.Fields(5).VALUE), "YYYYMMDD")
            oDS_PS_TEMPTABLE.setValue "U_dField02", sCnt, Format(Trim(oRecordSet01.Fields(6).VALUE), "YYYYMMDD")
            oDS_PS_TEMPTABLE.setValue "U_iField03", sCnt, Trim(oRecordSet01.Fields("U_Qty").VALUE)
            oDS_PS_TEMPTABLE.setValue "U_qField01", sCnt, Trim(oRecordSet01.Fields("U_Weight").VALUE)
            oDS_PS_TEMPTABLE.setValue "U_qField02", sCnt, Trim(oRecordSet01.Fields("U_UnWeight").VALUE)
            oDS_PS_TEMPTABLE.setValue "U_aField01", sCnt, Trim(oRecordSet01.Fields("U_Price").VALUE)
            oDS_PS_TEMPTABLE.setValue "U_aField02", sCnt, Trim(oRecordSet01.Fields("U_LinTotal").VALUE)
            oDS_PS_TEMPTABLE.setValue "U_sField12", sCnt, Trim(oRecordSet01.Fields("U_WhsCode").VALUE)
            oDS_PS_TEMPTABLE.setValue "U_sField13", sCnt, Trim(oRecordSet01.Fields("U_WhsName").VALUE)
            oDS_PS_TEMPTABLE.setValue "U_iField04", sCnt, Trim(oRecordSet01.Fields("PendNum").VALUE)
            oDS_PS_TEMPTABLE.setValue "U_sField15", sCnt, Trim(oRecordSet01.Fields("U_PODocNum").VALUE)
            oDS_PS_TEMPTABLE.setValue "U_sField16", sCnt, Trim(oRecordSet01.Fields("VisOrder").VALUE)
            
            oDS_PS_TEMPTABLE.setValue "U_sField23", sCnt, Trim(oRecordSet01.Fields("U_OutSize").VALUE)
            oDS_PS_TEMPTABLE.setValue "U_sField24", sCnt, Trim(oRecordSet01.Fields("U_OutUnit").VALUE)
            oDS_PS_TEMPTABLE.setValue "U_sField25", sCnt, Trim(oRecordSet01.Fields("U_Auto").VALUE)
            oDS_PS_TEMPTABLE.setValue "U_sField22", sCnt, Trim(oRecordSet01.Fields("U_Comments").VALUE)
            
            oDS_PS_TEMPTABLE.setValue "U_sField10", sCnt, Trim(oRecordSet01.Fields("U_ReqCntc").VALUE) '청구자
            
            oDS_PS_TEMPTABLE.setValue "U_sField07", sCnt, Trim(oRecordSet01.Fields("U_OrdNum").VALUE) '작번
            oDS_PS_TEMPTABLE.setValue "U_sField08", sCnt, Trim(oRecordSet01.Fields("U_OrdSub1").VALUE) 'Sub작번1
            oDS_PS_TEMPTABLE.setValue "U_sField09", sCnt, Trim(oRecordSet01.Fields("U_OrdSub2").VALUE) 'Sub작번2
            
            'oDS_PS_TEMPTABLE.setValue "U_sField21", sCnt, Trim(oRecordSet01.Fields("U_Payment").VALUE)
            
            sCnt = sCnt + 1
            oRecordSet01.MoveNext
        Loop
    End If
    Call oMat01.LoadFromDataSource
    Call oMat01.AutoResizeColumns
    
    oForm01.Freeze False
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet01 = Nothing
    Display_MatrixData = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Display_MatrixData_Error:
    oForm01.Freeze False
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "조회할 데이터가 없습니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "Display_MatrixData_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    Display_MatrixData = False
End Function

Private Function Update_PO() As Boolean
On Error GoTo Update_PO_Error
    Dim i               As Long
    Dim ErrNum          As Integer
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    Dim sQry            As String
    
    Dim CardCode$, BPLID$, CntcCode$, DocDateFr$, DocDateTo$, DueDateFr$, DueDateTo$, Purchase$, POStatus$
    Dim sCnt&
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    ErrNum = 0
    
    CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
    CntcCode = Trim(oForm01.Items("CntcCode").Specific.VALUE)
    DocDateFr = Trim(oForm01.Items("DocDateFr").Specific.VALUE)
    DocDateTo = Trim(oForm01.Items("DocDateTo").Specific.VALUE)
    DueDateFr = Trim(oForm01.Items("DueDateFr").Specific.VALUE)
    DueDateTo = Trim(oForm01.Items("DueDateTo").Specific.VALUE)
    Purchase = Trim(oForm01.Items("Purchase").Specific.VALUE)
    POStatus = Trim(oForm01.Items("POStatus").Specific.VALUE)
    
    If CardCode = "" Then CardCode = "%"
    If BPLID = "" Then BPLID = "%"
    If CntcCode = "" Then CntcCode = "%"
    If DocDateFr = "" Then DocDateFr = "20100101"
    If DocDateTo = "" Then DocDateTo = "20991231"
    If DueDateFr = "" Then DueDateFr = "20100101"
    If DueDateTo = "" Then DueDateTo = "20991231"
    If Purchase = "" Then Purchase = "%"
    If POStatus = "" Then POStatus = "%"
    
    sQry = "EXEC [PS_MM040_01] '" & CardCode & "', '" & BPLID & "', '" & CntcCode & "', '" & DocDateFr & "', "
    sQry = sQry & "'" & DocDateTo & "', '" & DueDateFr & "', '" & DueDateTo & "', '" & Purchase & "', '" & POStatus & "'"
    oRecordSet01.DoQuery sQry
    
    oDS_PS_TEMPTABLE.Clear
    
    sCnt = 0
    If oRecordSet01.RecordCount > 0 Then
        Do Until oRecordSet01.EOF
            oDS_PS_TEMPTABLE.InsertRecord (sCnt)
            oDS_PS_TEMPTABLE.Offset = sCnt
            oDS_PS_TEMPTABLE.setValue "U_iField01", sCnt, sCnt + 1
            oDS_PS_TEMPTABLE.setValue "U_sField01", sCnt, Trim(oRecordSet01.Fields(0).VALUE)
            oDS_PS_TEMPTABLE.setValue "U_sField02", sCnt, Trim(oRecordSet01.Fields(1).VALUE)
            oDS_PS_TEMPTABLE.setValue "U_sField03", sCnt, Trim(oRecordSet01.Fields(2).VALUE)
            oDS_PS_TEMPTABLE.setValue "U_dField01", sCnt, Format(Trim(oRecordSet01.Fields(3).VALUE), "YYYYMMDD")
            oDS_PS_TEMPTABLE.setValue "U_dField02", sCnt, Format(Trim(oRecordSet01.Fields(4).VALUE), "YYYYMMDD")
            oDS_PS_TEMPTABLE.setValue "U_sField04", sCnt, Trim(oRecordSet01.Fields(5).VALUE)
            oDS_PS_TEMPTABLE.setValue "U_sField05", sCnt, Trim(oRecordSet01.Fields(6).VALUE)
            oDS_PS_TEMPTABLE.setValue "U_sField06", sCnt, Trim(oRecordSet01.Fields(7).VALUE)
            oDS_PS_TEMPTABLE.setValue "U_sField07", sCnt, Trim(oRecordSet01.Fields(8).VALUE)
            oDS_PS_TEMPTABLE.setValue "U_sField08", sCnt, Trim(oRecordSet01.Fields(9).VALUE)
            oDS_PS_TEMPTABLE.setValue "U_dField03", sCnt, Format(Trim(oRecordSet01.Fields(10).VALUE), "YYYYMMDD")
            
            sCnt = sCnt + 1
            oRecordSet01.MoveNext
        Loop
    End If
    oMat01.LoadFromDataSource
    
    Set oRecordSet01 = Nothing
    Update_PO = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Update_PO_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "조회할 데이터가 없습니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "Display_MatrixDate_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    Update_PO = False
End Function

Private Sub FormResize()
On Error GoTo FormResize_Error
    oForm01.Items("Mat01").Top = 82
    oForm01.Items("Mat01").Left = 6
    oForm01.Items("Mat01").Width = oForm01.Width - 18
    oForm01.Items("Mat01").Height = (oForm01.Height - oForm01.Items("Mat01").Top - (oForm01.Height - oForm01.Items("Comments").Top)) / 2 - 10

    oForm01.Items("Mat02").Top = oForm01.Items("Mat01").Height + oForm01.Items("Mat01").Top
    oForm01.Items("Mat02").Left = oForm01.Items("Mat01").Left
    oForm01.Items("Mat02").Width = oForm01.Items("Mat01").Width
    oForm01.Items("Mat02").Height = oForm01.Items("Mat01").Height - 5
    
    Exit Sub
FormResize_Error:
    Sbo_Application.SetStatusBarMessage "FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_MM050_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PS_MM050_Validate_Error
    PS_MM050_Validate = True
    Dim i, j As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    If MDC_PS_Common.GetValue("SELECT Canceled FROM [@PS_MM050H] WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.Value & "'", 0, 1) = "Y" Then
'        Call MDC_Com.MDC_GF_Message("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", "W")
'        PS_MM050_Validate = False
'        GoTo PS_MM050_Validate_Exit
'    End If
'
    If ValidateType = "수정" Then
'        '//삭제된 행을 찾아서 삭제가능성 검사 , 만약 입력된행이 수정이 불가능하도록 변경이 필요하다면 삭제된행 찾는구문 제거
'        Dim Exist As Boolean
'        Exist = False
'        Query01 = "SELECT DocEntry,LineNum,ItemCode FROM [RDR1] WHERE DocEntry = '" & oForm01.Items("8").Specific.Value & "'"
'        RecordSet01.DoQuery Query01
'        For i = 0 To RecordSet01.RecordCount - 1
'            Exist = False
'            For j = 1 To oMat01.RowCount - 1
'                '//라인번호가 같고, 품목코드가 같으면 존재하는행 , LineNum에 값이 존재하는지 확인필요(행삭제된행인경우 LineNum이 존재하지않음)
'                If Val(RecordSet01.Fields(1).Value) = Val(oMat01.Columns("U_LineNum").Cells(j).Specific.Value) And RecordSet01.Fields(2).Value = oMat01.Columns("1").Cells(j).Specific.Value And oMat01.Columns("U_LineNum").Cells(j).Specific.Value <> "" Then
'                    Exist = True
'                End If
'            Next
'            If (Exist = False) Then '//삭제된 행중
'                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD030L] WHERE U_ORDRNum = '" & Val(RecordSet01.Fields(0).Value) & "' AND U_RDR1Num = '" & Val(RecordSet01.Fields(1).Value) & "'", 0, 1)) > 0 Then
'                    MDC_Com.MDC_GF_Message "삭제된행이 다른사용자에 의해 출하,선출요청되었습니다. 적용할수 없습니다.", "W"
'                    PS_MM050_Validate = False
'                    GoTo PS_MM050_Validate_Exit
'                End If
'            End If
'            RecordSet01.MoveNext
'        Next
    ElseIf ValidateType = "행삭제" Then
        '//행삭제전 행삭제가능여부검사
        If oForm01.Mode = fm_OK_MODE Or oForm01.Mode = fm_UPDATE_MODE Then '//추가,수정모드일때행삭제가능검사
            If (oMat02.Columns("LineNum").Cells(oLast_Col_Row).Specific.VALUE = "") Then '//새로추가된 행인경우, 삭제하여도 무방하다
            Else
       
                Query01 = "                SELECT      COUNT(*) AS [Count],"
                Query01 = Query01 & "                T0.DocEntry AS [DocEntry],"
                Query01 = Query01 & "                T1.U_LineNum AS [LineNum]"
                Query01 = Query01 & " FROM       [@PS_MM070H] AS T0"
                Query01 = Query01 & "                INNER JOIN"
                Query01 = Query01 & "                [@PS_MM070L] AS T1"
                Query01 = Query01 & "                   ON T0.DocEntry = T1.DocEntry"
                Query01 = Query01 & " WHERE      T1.U_GADocLin = '" & oForm01.Items("DocNum").Specific.VALUE & "-" & oMat02.Columns("LineID").Cells(oLast_Col_Row).Specific.VALUE & "'"
    '            Query01 = Query01 & "                AND T1.U_POR1Num = " & oMat02.Columns("BLineNum").Cells(i).Specific.VALUE
                Query01 = Query01 & "                AND T0.Status = 'O'"
                Query01 = Query01 & " GROUP BY T0.DocEntry,"
                Query01 = Query01 & "                T1.U_LineNum"
           
                Call RecordSet01.DoQuery(Query01)
            
                If RecordSet01.Fields("Count").VALUE > 0 Then
            
                    Call MDC_Com.MDC_GF_Message("선택한 품목은 이미 검수입고 되었습니다(검수입고 문서번호 : " & RecordSet01.Fields("DocEntry").VALUE & "-" & RecordSet01.Fields("LineNum").VALUE & "). 해당 가입고문서를 행삭제할 수 없습니다.", "E")
                    PS_MM050_Validate = False
                    GoTo PS_MM050_Validate_Exit
            
                End If
                    

'                    Call MDC_Com.MDC_GF_Message("해당 내용은 행삭제할수 없습니다.", "W")
'                    PS_MM050_Validate = False
'                    GoTo PS_MM050_Validate_Exit
'                End If
            End If
        End If
    ElseIf ValidateType = "취소" Then
    
        For i = 1 To oMat02.VisualRowCount
        
            Query01 = "                SELECT      COUNT(*) AS [Count],"
            Query01 = Query01 & "                T0.DocEntry AS [DocEntry],"
            Query01 = Query01 & "                T1.U_LineNum AS [LineNum]"
            Query01 = Query01 & " FROM       [@PS_MM070H] AS T0"
            Query01 = Query01 & "                INNER JOIN"
            Query01 = Query01 & "                [@PS_MM070L] AS T1"
            Query01 = Query01 & "                   ON T0.DocEntry = T1.DocEntry"
            Query01 = Query01 & " WHERE      T1.U_GADocLin = '" & oForm01.Items("DocNum").Specific.VALUE & "-" & oMat02.Columns("LineID").Cells(i).Specific.VALUE & "'"
'            Query01 = Query01 & "                AND T1.U_POR1Num = " & oMat02.Columns("BLineNum").Cells(i).Specific.VALUE
            Query01 = Query01 & "                AND T0.Status = 'O'"
            Query01 = Query01 & " GROUP BY T0.DocEntry,"
            Query01 = Query01 & "                T1.U_LineNum"
       
            Call RecordSet01.DoQuery(Query01)
        
            If RecordSet01.Fields("Count").VALUE > 0 Then
        
                Call MDC_Com.MDC_GF_Message(i & "행 품목이 이미 검수입고 되었습니다(검수입고 문서번호 : " & RecordSet01.Fields("DocEntry").VALUE & "-" & RecordSet01.Fields("LineNum").VALUE & "). 해당 가입고문서를 취소할 수 없습니다.", "E")
                PS_MM050_Validate = False
                GoTo PS_MM050_Validate_Exit
        
            End If
        
        Next
    
    ElseIf ValidateType = "닫기" Then
    
        For i = 1 To oMat02.VisualRowCount
        
            Query01 = "                SELECT      COUNT(*) AS [Count],"
            Query01 = Query01 & "                T0.DocEntry AS [DocEntry],"
            Query01 = Query01 & "                T1.U_LineNum AS [LineNum]"
            Query01 = Query01 & " FROM       [@PS_MM070H] AS T0"
            Query01 = Query01 & "                INNER JOIN"
            Query01 = Query01 & "                [@PS_MM070L] AS T1"
            Query01 = Query01 & "                   ON T0.DocEntry = T1.DocEntry"
            Query01 = Query01 & " WHERE      T1.U_GADocLin = '" & oForm01.Items("DocNum").Specific.VALUE & "-" & oMat02.Columns("LineID").Cells(i).Specific.VALUE & "'"
'            Query01 = Query01 & "                AND T1.U_POR1Num = " & oMat02.Columns("BLineNum").Cells(i).Specific.VALUE
            Query01 = Query01 & "                AND T0.Status = 'O'"
            Query01 = Query01 & " GROUP BY T0.DocEntry,"
            Query01 = Query01 & "                T1.U_LineNum"
       
            Call RecordSet01.DoQuery(Query01)
        
            If RecordSet01.Fields("Count").VALUE > 0 Then
        
                Call MDC_Com.MDC_GF_Message(i & "행 품목이 이미 검수입고 되었습니다(검수입고 문서번호 : " & RecordSet01.Fields("DocEntry").VALUE & "-" & RecordSet01.Fields("LineNum").VALUE & "). 해당 가입고문서를 닫기할 수 없습니다.", "E")
                PS_MM050_Validate = False
                GoTo PS_MM050_Validate_Exit
        
            End If
        
        Next

    End If
    
    Set RecordSet01 = Nothing
    Exit Function
PS_MM050_Validate_Exit:
    Set RecordSet01 = Nothing
    Exit Function
PS_MM050_Validate_Error:
    PS_MM050_Validate = False
    Sbo_Application.SetStatusBarMessage "PS_MM050_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub Raise_EVENT_ROW_DELETE(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_ROW_DELETE_Error
    Dim i As Long
    If (oLast_Col_Row > 0) Then
        If pval.BeforeAction = True Then
            If (PS_MM050_Validate("행삭제") = False) Then
                BubbleEvent = False
                Exit Sub
            End If
        '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
        
            For i = 1 To oMat02.VisualRowCount
                oMat02.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            
            oMat02.FlushToDataSource
            Call oDS_PS_MM050L.RemoveRecord(oDS_PS_MM050L.Size - 1)
            oMat02.LoadFromDataSource
            Call oForm01.Update
            
            If oMat02.RowCount = 0 Then
                Call Add_MatrixRow(0)
            Else
            
                If Trim(oDS_PS_MM050L.GetValue("U_PODocNum", oMat02.RowCount - 1)) <> "" Then
                    Call Add_MatrixRow(oMat02.RowCount)
                End If
                
            End If
        
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_MM050_CheckDate(ByVal pBaseEntry As String) As Boolean
'******************************************************************************
'Function ID : PS_MM050_CheckDate()
'해당모듈    : PS_MM050
'기능        : 선행프로세스와 일자 비교
'인수        : pBaseEntry : 기준문서번호
'반환값      : True-선행프로세스보다 일자가 같거나 느릴 경우, False-선행프로세스보다 일자가 빠를 경우
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_MM050_CheckDate_Error

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
    DocType = "PS_MM050"
    CurDocDate = Trim(oForm01.Items("DocDate").Specific.VALUE)

    Query01 = "         EXEC PS_Z_CHECK_DATE '"
    Query01 = Query01 & BaseEntry & "','"
    Query01 = Query01 & BaseLine & "','"
    Query01 = Query01 & DocType & "','"
    Query01 = Query01 & CurDocDate & "'"

    Call oRecordSet01.DoQuery(Query01)

    If oRecordSet01.Fields("ReturnValue").VALUE = "False" Then
        PS_MM050_CheckDate = False
    Else
        PS_MM050_CheckDate = True
    End If
    
    Set oRecordSet01 = Nothing

    Exit Function
PS_MM050_CheckDate_Error:
    PS_MM050_CheckDate = False
    Call Sbo_Application.SetStatusBarMessage("PS_MM050_CheckDate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)
End Function
