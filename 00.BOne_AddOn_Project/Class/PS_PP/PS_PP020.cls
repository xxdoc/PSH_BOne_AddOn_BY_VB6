VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_PP020.cls
'//  Module         : PP
'//  Description    : 작번등록
'//  FormType       : PS_PP020
'//  Create Date    : 2010.10.05
'//  Modified Date  :
'//  Creator        : Ryu Yung Jo
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01   As String
Public oForm01           As SAPbouiCOM.Form
Public oMat01            As SAPbouiCOM.Matrix
Public oMat02            As SAPbouiCOM.Matrix
Private oDS_PS_PP020H    As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_TEMPTABLE As SAPbouiCOM.DBDataSource

Private oLast_ItemUID    As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_ColUID     As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_ColRow     As Long   '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&
Private oLast_RightClick_DocEntry$

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error

    Dim i           As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01   As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP020.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)

    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_PP020_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting
    Call Initialization
    Call Add_MatrixRow(0, True)
    Call LoadCaption
    Call FormItemEnabled
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1285"), False        '// 복원
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

    Dim i&, j&
    Dim ErrNum&
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    
    Dim ItemCode$, ItemName$, Size$, Qty&, Weight As Currency, Unit$, RequestDate$, DueDate$, ItemType$, RequestNo$
    Dim RFC_Sender$
    Dim Seq&
    Dim ChildForm01  As Variant
    Set ChildForm01 = New PS_SM010
    Dim sQry         As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "Btn01" Then
                    If oForm01.Mode = fm_ADD_MODE Then
                        If HeaderSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        If MatrixSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
'                        If Add_JakNum(pval) = False Then
'                            BubbleEvent = False
'                            Exit Sub
'                        End If
                        
                        If PS_PP020_AddData = False Then
                        
                            BubbleEvent = False
                            Exit Sub
                        
                        End If
                        
                        oMat01.Clear
                        oMat01.FlushToDataSource
                        oMat01.LoadFromDataSource
                        Add_MatrixRow 0, True
                        
'                        Call Delete_EmptyRow
                        oLast_Mode = oForm01.Mode
                    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                        If MatrixSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        If PS_PP020_AddData = False Then
                        
                            BubbleEvent = False
                            Exit Sub
                        
                        End If
                        
'                        If Update_JakNum(pval) = False Then
'                            BubbleEvent = False
'                            Exit Sub
'                        End If
                        
                        oForm01.Mode = fm_OK_MODE
                        Call LoadCaption
                    ElseIf oForm01.Mode = fm_FIND_MODE Then
                        oLast_Mode = oForm01.Mode
                    End If
                ElseIf pval.ItemUID = "Btn02" Then
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    oForm01.Mode = fm_OK_MODE
                    Call LoadCaption
                    Call LoadData
                ElseIf pval.ItemUID = "Btn03" Then
                    If Add_SubJakNum(pval) = False Then
                        BubbleEvent = False
                        Exit Sub
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
                    ElseIf pval.ItemUID = "ItemCode" Then
                        If oForm01.Items("ItemCode").Specific.VALUE = "" Then
                            Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row)
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "JakName" Then
                            If oMat01.Columns("JakName").Cells(pval.Row).Specific.VALUE = "" Then
                                Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                            End If
                        ElseIf pval.ColUID = "CardCode" Then
                            If oMat01.Columns("CardCode").Cells(pval.Row).Specific.VALUE = "" Then
                                Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                            End If
                        ElseIf pval.ColUID = "ShipCode" Then
                            If oMat01.Columns("ShipCode").Cells(pval.Row).Specific.VALUE = "" Then
                                Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                            End If
                        End If
                    End If
                End If
            Case et_COMBO_SELECT: '//5
                If pval.ItemUID = "ReWork" Then
                    oDS_PS_TEMPTABLE.Clear
                    oMat02.Clear
                End If
'et_CLICK ///////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
                
                If pval.ItemUID = "Mat01" Then
                    If pval.Row > 0 Then
                        Call oMat01.SelectRow(pval.Row, True, False)
                    End If
                End If
'et_DOUBLE_CLICK ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_DOUBLE_CLICK: '//7
                If pval.ItemUID = "Mat01" And pval.Row <> "0" Then 'And pval.ColUID = "DocNum" Then
                    j = 0
                    If oMat02.VisualRowCount = 0 Then
                        oDS_PS_TEMPTABLE.Clear
                    End If
                    
                    Dim JakName$, SubNo1$, SubNo2$
                    JakName = Trim(oDS_PS_PP020H.GetValue("U_JakName", pval.Row - 1))
                    If Trim(oDS_PS_PP020H.GetValue("U_SubNo1", pval.Row - 1)) = "00" Then
                        If oForm01.Items("ReWork").Specific.VALUE = "10" Then '//정상
                            If Trim(oDS_PS_PP020H.GetValue("U_SubNo2", pval.Row - 1)) = "000" Then
                                sQry = "        SELECT  MAX(ISNULL(U_SubNo1, '00')) "
                                sQry = sQry & " FROM    [@PS_PP020H] "
                                sQry = sQry & " WHERE   U_JakName = '" & JakName & "' "
                                sQry = sQry & "         AND U_SubNo2 = '" & Trim(oDS_PS_PP020H.GetValue("U_SubNo2", pval.Row - 1)) & "'"
                                oRecordSet01.DoQuery sQry
                                
                                SubNo1 = Format(Trim(oRecordSet01.Fields(0).VALUE) + 1, "00")
                                SubNo2 = "000"
                            Else
                                MDC_Com.MDC_GF_Message "해당 작번은 서브작번을 만들 수 없습니다. 확인하세요." & Err.Number & " - " & Err.Description, "W"
                                j = 1
                            End If
                        Else
                            If Trim(oDS_PS_PP020H.GetValue("U_SubNo2", pval.Row - 1)) = "000" Then
                                sQry = "        SELECT  MAX(ISNULL(U_SubNo1, '00')) "
                                sQry = sQry & " FROM    [@PS_PP020H] "
                                sQry = sQry & " WHERE   U_JakName = '" & JakName & "' "
                                sQry = sQry & "         AND U_SubNo1 >= '80' "
                                sQry = sQry & "         AND U_SubNo2 = '" & Trim(oDS_PS_PP020H.GetValue("U_SubNo2", pval.Row - 1)) & "'"
                                oRecordSet01.DoQuery sQry
                                
                                If Trim(oRecordSet01.Fields(0).VALUE) = "" Then
                                    SubNo1 = "80"
                                Else
                                    SubNo1 = Format(Trim(oRecordSet01.Fields(0).VALUE) + 1, "00")
                                End If
                                SubNo2 = "000"
                            Else
                                MDC_Com.MDC_GF_Message "해당 작번은 서브작번을 만들 수 없습니다. 확인하세요." & Err.Number & " - " & Err.Description, "W"
                                j = 1
                            End If
                            
                        End If
                    ElseIf Trim(oDS_PS_PP020H.GetValue("U_SubNo1", pval.Row - 1)) <> "00" Then
                        If Trim(oDS_PS_PP020H.GetValue("U_SubNo2", pval.Row - 1)) = "000" Then
                            sQry = "        SELECT  MAX(ISNULL(U_SubNo2, '000')) "
                            sQry = sQry & " FROM    [@PS_PP020H] Where U_JakName = '" & JakName & "' "
                            sQry = sQry & "         AND U_SubNo1 = '" & Trim(oDS_PS_PP020H.GetValue("U_SubNo1", pval.Row - 1)) & "' "

                            oRecordSet01.DoQuery sQry
                            
                            SubNo1 = Format(Trim(oDS_PS_PP020H.GetValue("U_SubNo1", pval.Row - 1)), "00")
                            SubNo2 = Format(Trim(oRecordSet01.Fields(0).VALUE) + 1, "000")
                        Else
                            MDC_Com.MDC_GF_Message "해당 작번은 서브작번을 만들 수 없습니다. 확인하세요." & Err.Number & " - " & Err.Description, "W"
                            j = 1
                        End If
                    End If
                    
                    For i = 0 To oMat02.VisualRowCount - 1
                        oMat02.LoadFromDataSource
                        If Trim(oDS_PS_TEMPTABLE.GetValue("U_sField01", i)) = JakName And _
                           Trim(oDS_PS_TEMPTABLE.GetValue("U_sField02", i)) = SubNo1 And _
                           Trim(oDS_PS_TEMPTABLE.GetValue("U_sField03", i)) = SubNo2 Then
                            MDC_Com.MDC_GF_Message "같은 행을 두번 선택할 수 없습니다. 확인하세요." & Err.Number & " - " & Err.Description, "W"
                            j = 1
                        End If
                    Next

                    If j = 0 Then
                        oForm01.Freeze True
                        Call Add_MatrixRow(oMat02.VisualRowCount, False, "Mat02")
                        oMat02.Columns("JakName").Cells(oMat02.VisualRowCount).Specific.VALUE = JakName
                        oMat02.Columns("SubNo1").Cells(oMat02.VisualRowCount).Specific.VALUE = Format(SubNo1, "00")
                        oMat02.Columns("SubNo2").Cells(oMat02.VisualRowCount).Specific.VALUE = Format(SubNo2, "000")
'                        oMat02.Columns("JakMyung").Cells(oMat02.VisualRowCount).Specific.Value = Format(SubNo2, "000")
'                        oMat02.Columns("JakSize").Cells(oMat02.VisualRowCount).Specific.Value = Format(SubNo2, "000")
'                        oMat02.Columns("JakUnit").Cells(oMat02.VisualRowCount).Specific.Value = Format(SubNo2, "000")
                        oMat02.Columns("ItemCode").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_PP020H.GetValue("U_ItemCode", pval.Row - 1))
                        oMat02.Columns("ItemName").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_PP020H.GetValue("U_ItemName", pval.Row - 1))
                        oMat02.Columns("Material").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_PP020H.GetValue("U_Material", pval.Row - 1))
                        oMat02.Columns("Unit").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_PP020H.GetValue("U_Unit", pval.Row - 1))
                        oMat02.Columns("Size").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_PP020H.GetValue("U_Size", pval.Row - 1))
                        oMat02.Columns("ItmBSort").Cells(oMat02.VisualRowCount).Specific.Select Trim(oDS_PS_PP020H.GetValue("U_ItmBSort", pval.Row - 1)), psk_ByValue
                        oMat02.Columns("SjDocNum").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_PP020H.GetValue("U_SjDocNum", pval.Row - 1))
                        oMat02.Columns("SjLinNum").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_PP020H.GetValue("U_SjLinNum", pval.Row - 1))
                        oMat02.Columns("CardCode").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_PP020H.GetValue("U_CardCode", pval.Row - 1))
                        oMat02.Columns("CardName").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_PP020H.GetValue("U_CardName", pval.Row - 1))
                        oMat02.Columns("ShipCode").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_PP020H.GetValue("U_ShipCode", pval.Row - 1))
                        oMat02.Columns("ShipName").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_PP020H.GetValue("U_ShipName", pval.Row - 1))
                        If Trim(oDS_PS_PP020H.GetValue("U_InOutGbn", pval.Row - 1)) = "" Then
                        
                        Else
                            oMat02.Columns("InOutGbn").Cells(oMat02.VisualRowCount).Specific.Select Trim(oDS_PS_PP020H.GetValue("U_InOutGbn", pval.Row - 1)), psk_ByValue
                        End If
                        oMat02.Columns("JakDate").Cells(oMat02.VisualRowCount).Specific.VALUE = Format(Now, "YYYYMMDD")
                        oMat02.Columns("ProDate").Cells(oMat02.VisualRowCount).Specific.VALUE = Format(Now, "YYYYMMDD")
                        oMat02.Columns("ReDate").Cells(oMat02.VisualRowCount).Specific.VALUE = Format(Now, "YYYYMMDD")
'                        oMat02.Columns("SjWeight").Cells(oMat02.VisualRowCount).Specific.Value = Trim(oDS_PS_PP020H.GetValue("U_SjWeight", pval.Row - 1))
'                        oMat02.Columns("WrWeight").Cells(oMat02.VisualRowCount).Specific.Value = Trim(oDS_PS_PP020H.GetValue("U_WrWeight", pval.Row - 1))
                        oMat02.Columns("SjDcDate").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_PP020H.GetValue("U_SjDcDate", pval.Row - 1))
                        oMat02.Columns("SjDuDate").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_PP020H.GetValue("U_SjDuDate", pval.Row - 1))
                        oMat02.Columns("SlePrice").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_PP020H.GetValue("U_SlePrice", pval.Row - 1))
                        oMat02.Columns("WorkGbn").Cells(oMat02.VisualRowCount).Specific.Select Trim(oDS_PS_PP020H.GetValue("U_WorkGbn", pval.Row - 1))
                        
                        oMat02.Columns("OrderAmt").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_PP020H.GetValue("U_OrderAmt", pval.Row - 1)) '수주금액
                        oMat02.Columns("NegoAmt").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_PP020H.GetValue("U_NegoAmt", pval.Row - 1)) 'Nego금액
                        oMat02.Columns("TrgtAmt").Cells(oMat02.VisualRowCount).Specific.VALUE = Trim(oDS_PS_PP020H.GetValue("U_TrgtAmt", pval.Row - 1)) '목표금액
                        
'                        oMat02.Columns("Comments").Cells(oMat02.VisualRowCount).Specific.Value = Trim(oDS_PS_PP020H.GetValue("U_Comments", pval.Row - 1))

                        oMat02.FlushToDataSource
                        oMat02.LoadFromDataSource
                        oMat02.AutoResizeColumns
                        
                        oForm01.Freeze False
                        j = 0
                    End If
                        BubbleEvent = False

                '외주구분의 첫행을 선택한 후 컬럼 타이틀을 더블클릭하면 첫행의 값으로 첫행 이외의 값을 자동으로 선택하는 기능 추가(2012.08.20 송명규 수정)
                ElseIf pval.ItemUID = "Mat01" And pval.Row = "0" And pval.ColUID = "InOutGbn" Then
                
                    Dim FirstInOutGbn As String
                
                    oMat01.FlushToDataSource
                    FirstInOutGbn = oDS_PS_PP020H.GetValue("U_InOutGbn", 0)
                    For i = 1 To oMat01.VisualRowCount - 2
                        
                        Call oDS_PS_PP020H.setValue("U_InOutgbn", i, FirstInOutGbn)
                        
                    Next
                    oMat01.LoadFromDataSource
                    
                End If
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
                If pval.ItemUID = "Mat01" Then
                    If oForm01.Mode = fm_ADD_MODE Then
                    Else
                        oForm01.Mode = fm_UPDATE_MODE
                        Call LoadCaption
                    End If
                End If
                
                If pval.ItemUID = "WorkGbn" And oForm01.Mode = fm_ADD_MODE Then
                    oForm01.Freeze True
                    oMat01.Clear
                    oMat01.FlushToDataSource
                    oMat01.LoadFromDataSource
                    
                    Call Add_MatrixRow(0, True)
                    If oForm01.Items("WorkGbn").Specific.Selected.VALUE = "10" Then
                        oMat01.Columns("RegNum").Editable = True
                        oMat01.Columns("ItemCode").Editable = False
                    Else
                        oMat01.Columns("RegNum").Editable = False
                        oMat01.Columns("ItemCode").Editable = True
                    End If
                    oForm01.Freeze False
                End If
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
'et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "CntcCode" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "CardCode" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "ItemCode" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "JakName" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        ElseIf pval.ColUID = "ItemCode" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        ElseIf pval.ColUID = "ShipCode" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        End If
                        
                        If oForm01.Mode = fm_ADD_MODE Then
                        Else
                            oForm01.Mode = fm_UPDATE_MODE
                            Call LoadCaption
                        End If
                    End If
                End If
            Case et_MATRIX_LOAD: '//11
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
'et_FORM_RESIZE /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_FORM_RESIZE '//20
            
                oForm01.Items("Mat01").Top = 82
                oForm01.Items("Mat01").Left = 6
                oForm01.Items("Mat01").Width = oForm01.Width - 18
                oForm01.Items("Mat01").Height = (oForm01.Height - oForm01.Items("Mat01").Top - (oForm01.Height - oForm01.Items("Btn01").Top)) / 3 * 2 - 25
                
                oForm01.Items("RadioMat02").Top = oForm01.Items("Mat01").Height + oForm01.Items("Mat01").Top - 4
                oForm01.Items("RadioMat02").Left = 6
                oForm01.Items("RadioMat02").Height = 20
                
                oForm01.Items("33").Top = oForm01.Items("RadioMat02").Top
                oForm01.Items("ReWork").Top = oForm01.Items("RadioMat02").Top
                
                oForm01.Items("Mat02").Top = oForm01.Items("Mat01").Height + oForm01.Items("Mat01").Top + 15
                oForm01.Items("Mat02").Left = oForm01.Items("Mat01").Left
                oForm01.Items("Mat02").Width = oForm01.Items("Mat01").Width
                oForm01.Items("Mat02").Height = oForm01.Items("Mat01").Height / 2 + 5
                
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
            Case et_LOST_FOCUS: '//4
'et_FORM_UNLOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_FORM_UNLOAD: '//17
                RemoveForms oFormUniqueID01
                Set oForm01 = Nothing
                Set oMat01 = Nothing
                Set oDS_PS_PP020H = Nothing
        End Select
    End If
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
    Set ProgressBar01 = Nothing
    Set oRecordSet01 = Nothing
    If ErrNum = 101 Then
        ErrNum = 0
        MDC_Com.MDC_GF_Message "Raise_ItemEvent_Error:" & Err.Number & " - " & Err.Description, "E"
        BubbleEvent = False
    Else
        MDC_Com.MDC_GF_Message "Raise_ItemEvent_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_MenuEvent_Error
    Dim i&
    Dim ReturnValue&
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    Dim ErrNum As String
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
'[1293:행삭제] //////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1293": '행삭제
                If oLast_ItemUID = "Mat01" Then
                    ReturnValue = Sbo_Application.MessageBox("해당 라인의 작번을 삭제합니다. 삭제 후 복원할 수 없습니다. 삭제하시겠습니까?", 1, "&확인", "&취소")
                    Select Case ReturnValue
                        Case 1:
                            '작업지시등록여부 체크
                            sQry = "        SELECT  COUNT(*) AS [Cnt]"
                            sQry = sQry & " FROM    [@PS_PP030H]"
                            sQry = sQry & " WHERE   U_BaseNum = '" & oLast_RightClick_DocEntry & "'"
                            sQry = sQry & "         AND [Status] = 'O'"
                            sQry = sQry & "         AND U_OrdGbn IN ('105','106')"
                            
                            Call oRecordSet01.DoQuery(sQry)
                            
                            If oRecordSet01.Fields("Cnt").VALUE > 0 Then '작업지시등록이 존재하면
                                ErrNum = "1"
                                BubbleEvent = False
                                GoTo Raise_MenuEvent_Error
                            End If
                            
                            sQry = "        DELETE "
                            sQry = sQry & " FROM    [@PS_PP020H] "
                            sQry = sQry & " WHERE   DocEntry = '" & oLast_RightClick_DocEntry & "'"
                            Call oRecordSet01.DoQuery(sQry)

                            oLast_RightClick_DocEntry = 0
                            Call MDC_Com.MDC_GF_Message("작번이 삭제되었습니다.", "S")
                        Case 2:
                            Call MDC_Com.MDC_GF_Message("실행이 취소되었습니다.", "S")
                            BubbleEvent = False
                            Exit Sub
                    End Select
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
                        oMat01.Columns("DocNum").Cells(i + 1).Specific.VALUE = i + 1
                    Next i

                    oMat01.FlushToDataSource
                    oDS_PS_PP020H.RemoveRecord oDS_PS_PP020H.Size - 1       '// Mat01에 마지막라인(빈라인) 삭제
                    oMat01.Clear
                    oMat01.LoadFromDataSource
                    oForm01.Freeze False
                End If
                
                If oMat02.RowCount <> oMat02.VisualRowCount - 1 Then
                    oForm01.Freeze True
                    For i = 0 To oMat02.VisualRowCount
                        oMat01.Columns("DocNum").Cells(i + 1).Specific.VALUE = i + 1
                    Next i

                    oMat02.FlushToDataSource
                    oDS_PS_TEMPTABLE.RemoveRecord oDS_PS_TEMPTABLE.Size - 1       '// Mat02에 마지막라인(빈라인) 삭제
                    oMat02.Clear
                    oMat02.LoadFromDataSource
                    oForm01.Freeze False
                End If
'[1281:찾기] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1281": '찾기
                FormItemEnabled
                oForm01.Mode = fm_OK_MODE
                Call LoadCaption
'                oForm01.Items("OrdType").Click ct_Regular
'[1282:추가] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1282": '추가
                oForm01.Freeze True
                FormItemEnabled
                
                If Trim(oForm01.Items("WorkGbn").Specific.VALUE) = "10" Then
                    oMat01.Columns("RegNum").Editable = True
                    oMat01.Columns("ItemCode").Editable = False
                Else
                    oMat01.Columns("RegNum").Editable = False
                    oMat01.Columns("ItemCode").Editable = True
                End If
                
                oForm01.Mode = fm_ADD_MODE
                oMat01.Clear
                oMat01.FlushToDataSource
                oMat01.LoadFromDataSource
                Add_MatrixRow 0, True
                Call LoadCaption
                oForm01.Freeze False
'[1288~1291:네비게이션] /////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                FormItemEnabled
                If oMat01.VisualRowCount > 0 Then
                    If oMat01.Columns("CGNo").Cells(oMat01.VisualRowCount).Specific.VALUE <> "" Then
                        If oDS_PS_PP020H.GetValue("Status", 0) = "O" Then
                            Add_MatrixRow oMat01.RowCount, False
                        End If
                    End If
                End If
  
        End Select
    End If
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_MenuEvent_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = "1" Then
        Call MDC_Com.MDC_GF_Message("작업지시등록이 존재하는 작번입니다. 삭제할 수 없습니다.", "E")
    Else
        Call MDC_Com.MDC_GF_Message("Raise_MenuEvent_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
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
            oLast_ItemUID = eventInfo.ItemUID
            If oLast_ItemUID = "Mat01" Then
                oLast_RightClick_DocEntry = Trim(oMat01.Columns("DocEntry").Cells(eventInfo.Row).Specific.VALUE)
            End If
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
    Set oDS_PS_PP020H = oForm01.DataSources.DBDataSources("@PS_PP020H")
    Set oDS_PS_TEMPTABLE = oForm01.DataSources.DBDataSources("@PS_TEMPTABLE")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    Set oMat02 = oForm01.Items("Mat02").Specific
    
    Call oForm01.DataSources.UserDataSources.Add("PuDateFr", dt_DATE, 8)
    oForm01.Items("PuDateFr").Specific.DataBind.SetBound True, "", "PuDateFr"
'    oForm01.DataSources.UserDataSources.Item("PuDateFr").Value = Format(Now, "YYYY-MM") & "-01"
    
    Call oForm01.DataSources.UserDataSources.Add("PuDateTo", dt_DATE, 8)
    oForm01.Items("PuDateTo").Specific.DataBind.SetBound True, "", "PuDateTo"
'    oForm01.DataSources.UserDataSources.Item("PuDateTo").Value = Date

    Call oForm01.DataSources.UserDataSources.Add("JakDateFr", dt_DATE, 8)
    oForm01.Items("JakDateFr").Specific.DataBind.SetBound True, "", "JakDateFr"
'    oForm01.DataSources.UserDataSources.Item("JakDateFr").Value = Format(Now, "YYYY-MM") & "-01"
    
    Call oForm01.DataSources.UserDataSources.Add("JakDateTo", dt_DATE, 8)
    oForm01.Items("JakDateTo").Specific.DataBind.SetBound True, "", "JakDateTo"
'    oForm01.DataSources.UserDataSources.Item("JakDateTo").Value = Date

    oForm01.DataSources.UserDataSources.Add "RadioMat01", dt_SHORT_TEXT, 1
    oForm01.Items("RadioMat01").Specific.DataBind.SetBound True, "", "RadioMat01"
    
    oForm01.DataSources.UserDataSources.Add "RadioMat02", dt_SHORT_TEXT, 1
    oForm01.Items("RadioMat02").Specific.DataBind.SetBound True, "", "RadioMat02"
    
    oForm01.Items("RadioMat01").Specific.GroupWith "RadioMat02"

    oMat01.Columns("OrderAmt").Visible = False '메인작번등록-견적금액 숨김
    oMat01.Columns("NegoAmt").Visible = False '메인작번등록-네고금액 숨김
    oMat01.Columns("TrgtAmt").Visible = False '메인작번등록-목표금액 숨김
    
    oMat02.Columns("OrderAmt").Visible = False '서브작번등록-견적금액 숨김
    oMat02.Columns("NegoAmt").Visible = False '서브작번등록-네고금액 숨김
    oMat02.Columns("TrgtAmt").Visible = False '서브작번등록-목표금액 숨김

    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
CreateItems_Error:
    MDC_Com.MDC_GF_Message "CreateItems_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub ComboBox_Setting()
On Error GoTo ComboBox_Setting_Error
    '//콤보에 기본값설정
    Dim oCombo       As SAPbouiCOM.ComboBox
    Dim sQry         As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    '// 사업장
    Set oCombo = oForm01.Items("BPLId").Specific
    sQry = "SELECT BPLId, BPLName From [OBPL] order by 1"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
'        oMat01.Columns("BPLId").ValidValues.Add Trim(oRecordset01.Fields(0).Value), Trim(oRecordset01.Fields(1).Value)
        oRecordSet01.MoveNext
    Loop
    
    '//품목대분류
    Set oCombo = oForm01.Items("ItmBSort").Specific
    sQry = "SELECT Code, Name From [@PSH_ITMBSORT] Where Code in ('105', '106') Order by Code"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oMat01.Columns("ItmBSort").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    oCombo.Select 0, psk_Index
    
    '//작업구분
    Set oCombo = oForm01.Items("WorkGbn").Specific
    oCombo.ValidValues.Add "10", "영업"
    oCombo.ValidValues.Add "20", "정비"
    oCombo.ValidValues.Add "30", "멀티"
    oCombo.ValidValues.Add "40", "신동"
    oCombo.ValidValues.Add "50", "R/D"
    oCombo.Select 0, psk_Index
    
    oMat01.Columns("WorkGbn").ValidValues.Add "10", "영업"
    oMat01.Columns("WorkGbn").ValidValues.Add "20", "정비"
    oMat01.Columns("WorkGbn").ValidValues.Add "30", "멀티"
    oMat01.Columns("WorkGbn").ValidValues.Add "40", "신동"
    oMat01.Columns("WorkGbn").ValidValues.Add "50", "R/D"
    
    oMat02.Columns("WorkGbn").ValidValues.Add "10", "영업"
    oMat02.Columns("WorkGbn").ValidValues.Add "20", "정비"
    oMat02.Columns("WorkGbn").ValidValues.Add "30", "멀티"
    oMat02.Columns("WorkGbn").ValidValues.Add "40", "신동"
    oMat02.Columns("WorkGbn").ValidValues.Add "50", "R/D"
    
    '//작번구분
    Set oCombo = oForm01.Items("JakGbn").Specific
    oCombo.ValidValues.Add "99", "전체"
    oCombo.ValidValues.Add "00", "메인작번"
    oCombo.ValidValues.Add "01", "서브작번"
    oCombo.Select 0, psk_Index
    
    '//재작업구분
    Set oCombo = oForm01.Items("ReWork").Specific
    oCombo.ValidValues.Add "10", "정상"
    oCombo.ValidValues.Add "20", "재작업"
    oCombo.Select 0, psk_Index
    
    '//작번구분
    oMat01.Columns("InOutGbn").ValidValues.Add "%", "선택"
    oMat01.Columns("InOutGbn").ValidValues.Add "IN", "자체"
    oMat01.Columns("InOutGbn").ValidValues.Add "OUT", "외주"
    
    oMat02.Columns("InOutGbn").ValidValues.Add "%", "선택"
    oMat02.Columns("InOutGbn").ValidValues.Add "IN", "자체"
    oMat02.Columns("InOutGbn").ValidValues.Add "OUT", "외주"
    
    '재작업 사유(Mat01)
    sQry = "        SELECT  B.U_Minor, "
    sQry = sQry & "         B.U_CdName"
    sQry = sQry & " FROM    [@PS_SY001H] AS A"
    sQry = sQry & "         INNER JOIN"
    sQry = sQry & "         [@PS_SY001L] AS B"
    sQry = sQry & "             ON A.Code = B.Code"
    sQry = sQry & "             AND A.Code = 'P202'"

    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ReWkRn"), sQry)
    
    '재작업 사유(Mat02)
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat02.Columns("ReWkRn"), sQry)

    '연간품여부
    Call oMat01.Columns("YearPdYN").ValidValues.Add("N", "N")
    Call oMat01.Columns("YearPdYN").ValidValues.Add("Y", "Y")
    
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

    Dim oCombo As SAPbouiCOM.ComboBox
    
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

Sub FormItemEnabled()
On Error GoTo FormItemEnabled_Error
    If (oForm01.Mode = fm_ADD_MODE) Then
        oForm01.Items("CardCode").Click ct_Regular
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("ItmBSort").Enabled = True
        oForm01.Items("WorkGbn").Enabled = True
        oForm01.Items("CardCode").Enabled = True
        oForm01.Items("ItemCode").Enabled = True
        oForm01.Items("PuDateFr").Enabled = True
        oForm01.Items("PuDateTo").Enabled = True
        oForm01.Items("JakDateFr").Enabled = False
        oForm01.Items("JakDateTo").Enabled = False
        oForm01.Items("JakGbn").Enabled = False
        oForm01.Items("Btn02").Enabled = False
        
        oMat01.Columns("JakName").Editable = True
        oMat01.Columns("SubNo1").Editable = False
        oMat01.Columns("SubNo2").Editable = False
        oMat01.Columns("JakMyung").Editable = True
        oMat01.Columns("JakSize").Editable = True
        oMat01.Columns("JakUnit").Editable = True
        oMat01.Columns("RegNum").Editable = False
        oMat01.Columns("ItemCode").Editable = False
        oMat01.Columns("CardCode").Editable = True
        oMat01.Columns("ShipCode").Editable = True
        oMat01.Columns("InOutGbn").Editable = True
        oMat01.Columns("ProDate").Editable = True
        oMat01.Columns("ReDate").Editable = True
        oMat01.Columns("WrWeight").Editable = True
        oMat01.Columns("Comments").Editable = True
        oMat01.Columns("Status").Editable = True
        
        oForm01.Items("RadioMat02").Visible = False
        oForm01.Items("Mat02").Visible = False
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        oForm01.Items("CardCode").Click ct_Regular
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("ItmBSort").Enabled = True
        oForm01.Items("WorkGbn").Enabled = True
        oForm01.Items("CardCode").Enabled = True
        oForm01.Items("ItemCode").Enabled = True
        oForm01.Items("PuDateFr").Enabled = False
        oForm01.Items("PuDateTo").Enabled = False
        oForm01.Items("JakDateFr").Enabled = True
        oForm01.Items("JakDateTo").Enabled = True
        oForm01.Items("JakGbn").Enabled = True
        oForm01.Items("Btn02").Enabled = True
        
        oMat01.Columns("JakName").Editable = False
        oMat01.Columns("SubNo1").Editable = False
        oMat01.Columns("SubNo2").Editable = False
        oMat01.Columns("JakMyung").Editable = True
        oMat01.Columns("JakSize").Editable = True
        oMat01.Columns("JakUnit").Editable = True
        oMat01.Columns("RegNum").Editable = False
        oMat01.Columns("ItemCode").Editable = False
        oMat01.Columns("CardCode").Editable = True
        oMat01.Columns("ShipCode").Editable = True
        oMat01.Columns("InOutGbn").Editable = True
        oMat01.Columns("ProDate").Editable = True
        oMat01.Columns("ReDate").Editable = True
        oMat01.Columns("WrWeight").Editable = True
        oMat01.Columns("Comments").Editable = True
        oMat01.Columns("Status").Editable = True
        
        oForm01.Items("RadioMat02").Visible = True
        oForm01.Items("Mat02").Visible = True
    ElseIf (oForm01.Mode = fm_OK_MODE) Then

    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub
Sub Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean, Optional ItemUID$)
On Error GoTo Add_MatrixRow_Error
    If ItemUID = "Mat02" Then
        If RowIserted = False Then '//행추가여부
            oDS_PS_TEMPTABLE.InsertRecord (oRow)
        End If
        oMat02.AddRow
        oDS_PS_TEMPTABLE.Offset = oRow
        oDS_PS_TEMPTABLE.setValue "U_iField01", oRow, oRow + 1
        oMat02.LoadFromDataSource
    Else
        If RowIserted = False Then '//행추가여부
            oDS_PS_PP020H.InsertRecord (oRow)
        End If
        oMat01.AddRow
        oDS_PS_PP020H.Offset = oRow
        oDS_PS_PP020H.setValue "DocNum", oRow, oRow + 1
        oMat01.LoadFromDataSource
    End If
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
    Dim RegNum&
    Dim Calculate_Weight As Double
    Dim oRecordSet01     As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
        Case "ShipCode"
            sQry = "Select CardName From OCRD Where CardCode = '" & Trim(oForm01.Items("ShipCode").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            oForm01.Items("ShipName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
        Case "ItemCode"
            sQry = "Select ItemName From OITM Where ItemCode = '" & Trim(oForm01.Items("ItemCode").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            oForm01.Items("ItemName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
        Case "CardCode"
            sQry = "Select CardName From OCRD Where CardCode = '" & Trim(oForm01.Items("CardCode").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            oForm01.Items("CardName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
        Case "Mat01"
            If oCol = "JakName" Then
                Dim JakName$
                oForm01.Freeze True
                If oForm01.Mode = fm_ADD_MODE Then
                    If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 0) And Trim(oMat01.Columns("JakName").Cells(oRow).Specific.VALUE) <> "" Then
                        oMat01.FlushToDataSource
                        Add_MatrixRow oMat01.RowCount, False
                    End If
                End If
                
                JakName = Trim(oMat01.Columns("JakName").Cells(oRow).Specific.VALUE)
                sQry = "        Select  a.DocEntry,"
                sQry = sQry & "         a.DocNum,"
                sQry = sQry & "         a.Period,"
                sQry = sQry & "         a.Instance,"
                sQry = sQry & "         a.Series,"
                sQry = sQry & "         a.Handwrtten,"
                sQry = sQry & "         a.Canceled,"
                sQry = sQry & "         a.Object,"
                sQry = sQry & "         a.LogInst,"
                sQry = sQry & "         a.UserSign,"
                sQry = sQry & "         a.Transfered,"
                sQry = sQry & "         a.Status,"
                sQry = sQry & "         a.CreateDate,"
                sQry = sQry & "         a.CreateTime,"
                sQry = sQry & "         a.UpdateDate,"
                sQry = sQry & "         a.UpdateTime,"
                sQry = sQry & "         a.DataSource,"
                sQry = sQry & "         a.U_BPLId,"
                sQry = sQry & "         a.U_RegNum,"
                sQry = sQry & "         a.U_ItemCode,"
                sQry = sQry & "         a.U_ItemName,"
                sQry = sQry & "         a.U_Material,"
                sQry = sQry & "         a.U_Unit,"
                sQry = sQry & "         a.U_Size,"
                sQry = sQry & "         a.U_ItmBSort,"
                sQry = sQry & "         a.U_CpName,"
                sQry = sQry & "         a.U_SjDocNum,"
                sQry = sQry & "         a.U_SjLinNum,"
                sQry = sQry & "         a.U_SjQty,"
                sQry = sQry & "         a.U_SjWeight,"
                sQry = sQry & "         a.U_SjDcDate,"
                sQry = sQry & "         a.U_SjDuDate,"
                sQry = sQry & "         a.U_SlePrice,"
                sQry = sQry & "         a.U_WorkGbn,"
                sQry = sQry & "         a.U_CardCode,"
                sQry = sQry & "         a.U_CardName,"
                sQry = sQry & "         a.U_ShipCode,"
                sQry = sQry & "         a.U_ShipName,"
                sQry = sQry & "         a.U_PuDate,"
                sQry = sQry & "         a.U_ProDate,"
                sQry = sQry & "         a.U_Comments,"
                sQry = sQry & "         a.U_JakName,"
                sQry = sQry & "         a.U_SubNo1,"
                sQry = sQry & "         a.U_SubNo2,"
                sQry = sQry & "         a.U_Status,"
                sQry = sQry & "         a.U_UseDept,"
                sQry = sQry & "         ISNULL(a.U_ReWeight, 0) AS U_ReWeight"
                sQry = sQry & " FROM    [@PS_PP010H] a "
                sQry = sQry & "         LEFT JOIN "
                sQry = sQry & "         [@PS_SD010H] b "
                sQry = sQry & "             ON a.U_RegNum = b.U_RegNum"
                sQry = sQry & " WHERE   a.U_JakName = '" & JakName & "'"
                
                Call oRecordSet01.DoQuery(sQry)
                                
                If Trim(oRecordSet01.Fields(0).VALUE) <> "" Then
                    sQry = "        Select  a.DocEntry,"
                    sQry = sQry & "         a.DocNum,"
                    sQry = sQry & "         a.Period,"
                    sQry = sQry & "         a.Instance,"
                    sQry = sQry & "         a.Series,"
                    sQry = sQry & "         a.Handwrtten,"
                    sQry = sQry & "         a.Canceled,"
                    sQry = sQry & "         a.Object,"
                    sQry = sQry & "         a.LogInst,"
                    sQry = sQry & "         a.UserSign,"
                    sQry = sQry & "         a.Transfered,"
                    sQry = sQry & "         a.Status,"
                    sQry = sQry & "         a.CreateDate,"
                    sQry = sQry & "         a.CreateTime,"
                    sQry = sQry & "         a.UpdateDate,"
                    sQry = sQry & "         a.UpdateTime,"
                    sQry = sQry & "         a.DataSource,"
                    sQry = sQry & "         a.U_BPLId,"
                    sQry = sQry & "         a.U_RegNum,"
                    sQry = sQry & "         a.U_ItemCode,"
                    sQry = sQry & "         a.U_ItemName,"
                    sQry = sQry & "         a.U_Material,"
                    sQry = sQry & "         a.U_Unit,"
                    sQry = sQry & "         a.U_Size,"
                    sQry = sQry & "         a.U_ItmBSort,"
                    sQry = sQry & "         a.U_CpName,"
                    sQry = sQry & "         a.U_SjDocNum,"
                    sQry = sQry & "         a.U_SjLinNum,"
                    sQry = sQry & "         a.U_SjQty,"
                    sQry = sQry & "         a.U_SjWeight,"
                    sQry = sQry & "         a.U_SjDcDate,"
                    sQry = sQry & "         a.U_SjDuDate,"
                    sQry = sQry & "         a.U_SlePrice,"
                    sQry = sQry & "         a.U_WorkGbn,"
                    sQry = sQry & "         a.U_CardCode,"
                    sQry = sQry & "         a.U_CardName,"
                    sQry = sQry & "         a.U_ShipCode,"
                    sQry = sQry & "         a.U_ShipName,"
                    sQry = sQry & "         a.U_PuDate,"
                    sQry = sQry & "         a.U_ProDate,"
                    sQry = sQry & "         a.U_Comments,"
                    sQry = sQry & "         a.U_JakName,"
                    sQry = sQry & "         a.U_SubNo1,"
                    sQry = sQry & "         a.U_SubNo2,"
                    sQry = sQry & "         a.U_Status,"
                    sQry = sQry & "         a.U_UseDept,"
                    sQry = sQry & "         ISNULL(a.U_ReWeight, 0) AS U_ReWeight,"
                    sQry = sQry & "         B.SalUnitMsr "
                    sQry = sQry & " FROM    [@PS_PP010H] a "
                    sQry = sQry & "         INNER JOIN "
                    sQry = sQry & "         [OITM] b "
                    sQry = sQry & "             ON a.U_ItemCode = b.ItemCode "
                    sQry = sQry & " WHERE   a.U_jakName = '" & JakName & "'"
                    
                    Call oRecordSet01.DoQuery(sQry)
                    
                End If
                                                        
                oMat01.Columns("SubNo1").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_SubNo1").VALUE)
                oMat01.Columns("SubNo2").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_SubNo2").VALUE)
                oMat01.Columns("RegNum").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_RegNum").VALUE)
                oMat01.Columns("ItemCode").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_ItemCode").VALUE)
                oMat01.Columns("ItemName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_ItemName").VALUE)
                oMat01.Columns("Material").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_Material").VALUE)
                oMat01.Columns("Unit").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("SalUnitMsr").VALUE)
                oMat01.Columns("Size").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_Size").VALUE)
                oMat01.Columns("ItmBSort").Cells(oRow).Specific.Select Trim(oRecordSet01.Fields("U_ItmBSort").VALUE), psk_ByValue
                oMat01.Columns("SjDocNum").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_SjDocNum").VALUE)
                oMat01.Columns("SjLinNum").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_SjLinNum").VALUE)
                oMat01.Columns("CardCode").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_CardCode").VALUE)
                oMat01.Columns("CardName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_CardName").VALUE)
                oMat01.Columns("ShipCode").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_ShipCode").VALUE)
                oMat01.Columns("ShipName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_ShipName").VALUE)
                Call oMat01.Columns("InOutGbn").Cells(oRow).Specific.Select("%", psk_ByValue) '외주구분을 선택하도록 유도하기 위해 수정(2012.07.12 송명규)
                oMat01.Columns("JakDate").Cells(oRow).Specific.VALUE = Format(Date, "YYYYMMDD")
                oMat01.Columns("ProDate").Cells(oRow).Specific.VALUE = Format(Trim(oRecordSet01.Fields("U_ProDate").VALUE), "YYYYMMDD")
                oMat01.Columns("ReDate").Cells(oRow).Specific.VALUE = Format(Trim(oRecordSet01.Fields("U_SjDuDate").VALUE), "YYYYMMDD")
                oMat01.Columns("SjWeight").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_SjWeight").VALUE)
                oMat01.Columns("SjDcDate").Cells(oRow).Specific.VALUE = Format(Trim(oRecordSet01.Fields("U_SjDcDate").VALUE), "YYYYMMDD")
                oMat01.Columns("SjDuDate").Cells(oRow).Specific.VALUE = Format(Trim(oRecordSet01.Fields("U_SjDuDate").VALUE), "YYYYMMDD")
                oMat01.Columns("SlePrice").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_SlePrice").VALUE)
                oMat01.Columns("WorkGbn").Cells(oRow).Specific.Select Trim(oRecordSet01.Fields("U_WorkGbn").VALUE), psk_ByValue
                oMat01.Columns("PP010Doc").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("DocEntry").VALUE)
                oMat01.Columns("WrWeight").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_ReWeight").VALUE)
                Call oMat01.Columns("YearPdYN").Cells(oRow).Specific.Select("N", psk_ByValue) '연간품여부(2015.06.15 송명규 추가)
                
                oMat01.Columns("Comments").Cells(oRow).Click ct_Regular
                oForm01.Freeze False
            ElseIf oCol = "ItemCode" Then
                oForm01.Freeze True
                If oForm01.Mode = fm_ADD_MODE Then
                    If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 0) And Trim(oMat01.Columns("ItemCode").Cells(oRow).Specific.VALUE) <> "" Then
                        oMat01.FlushToDataSource
                        Add_MatrixRow oMat01.RowCount, False
                    End If
                End If
                
                sQry = "        Select  a.ItemName, "
                sQry = sQry & "         a.U_Material, "
                sQry = sQry & "         a.SalUnitMsr, "
                sQry = sQry & "         a.U_Size, "
                sQry = sQry & "         a.U_ItmBSort, "
                sQry = sQry & "         b.U_CPNaming  "
                sQry = sQry & " From    OITM a "
                sQry = sQry & "         Left Join "
                sQry = sQry & "         [@PSH_ItmMSort] b"
                sQry = sQry & "             ON a.U_ItmBSort = b.U_rCode "
                sQry = sQry & "             And a.U_ItmMsort = b.U_Code "
                sQry = sQry & " WHERE   a.ItemCode = '" & Trim(oMat01.Columns("ItemCode").Cells(oRow).Specific.VALUE) & "'"
                oRecordSet01.DoQuery sQry
                oMat01.Columns("ItemName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
'                oMat01.Columns("PuDate").Cells(oRow).Specific.Value = Format(Date, "YYYYMMDD")
'                oMat01.Columns("ItmBSort").Cells(oRow).Specific.Select Trim(oRecordset01.Fields(1).Value), psk_ByValue
                                
                oMat01.Columns("ItemCode").Cells(oRow).Click ct_Regular
                oForm01.Freeze False
            ElseIf oCol = "ShipCode" Then
                sQry = "Select CardName From OCRD Where CardCode = '" & Trim(oMat01.Columns("ShipCode").Cells(oRow).Specific.VALUE) & "'"
                oRecordSet01.DoQuery sQry
                
                oMat01.Columns("ShipName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            End If
    End Select

'    Set oRecordset = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
FlushToItemValue_Error:
    oForm01.Freeze False
    MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error

    Dim ErrNum          As Integer
    
    ErrNum = 0

    '// Check
    Select Case True
        Case Trim(oForm01.Items("BPLId").Specific.VALUE) = ""
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

    Dim i            As Long
    Dim ErrNum       As Integer
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim sQry         As String

    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    ErrNum = 0

    oMat01.FlushToDataSource

    '// 라인
    If oMat01.VisualRowCount = 0 Then
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    End If

    For i = 1 To oMat01.VisualRowCount - 1
        Select Case True
'            Case oDS_PS_PP020H.GetValue("U_ItemCode", i) = ""
'                ErrNum = 2
'                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_PP020H.GetValue("U_Qty", i) = ""
'                ErrNum = 3
'                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_PP020H.GetValue("U_DocDate", i) = ""
'                ErrNum = 4
'                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_PP020H.GetValue("U_DueDate", i) = ""
'                ErrNum = 5
'                GoTo MatrixSpaceLineDel_Error

            '외주구분 필수 체크(2012.07.12 송명규 수정, 최수환이사 요청)
            Case Trim(oDS_PS_PP020H.GetValue("U_InOutGbn", i - 1)) = "%"
                ErrNum = 6
                GoTo MatrixSpaceLineDel_Error
            '작번중복체크(서브작번 포함) 기능 필요
            
        End Select
    Next
    oMat01.LoadFromDataSource

    Set oRecordSet01 = Nothing
    MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
MatrixSpaceLineDel_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "라인 데이터가 없습니다. 확인하세요.", "E"
'    ElseIf ErrNum = 2 Then
'        MDC_Com.MDC_GF_Message "" & i & "번 라인의 품목코드가 없습니다. 확인하세요.", "E"
'    ElseIf ErrNum = 3 Then
'        MDC_Com.MDC_GF_Message "" & i & "번 라인의 수량이 없습니다. 확인하세요.", "E"
'    ElseIf ErrNum = 4 Then
'        MDC_Com.MDC_GF_Message "" & i & "번 라인의 청구일자가 없습니다. 확인하세요.", "E"
'    ElseIf ErrNum = 5 Then
'        MDC_Com.MDC_GF_Message "" & i & "번 라인의 납기일자가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "" & i & "번 라인의 외주구분이 선택되지 않았습니다. 확인하세요.", "E"
    ElseIf ErrNum = 7 Then
        MDC_Com.MDC_GF_Message "" & i & "번 라인의 작번은 이미 등록되었습니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function

Sub Delete_EmptyRow()
On Error GoTo Delete_EmptyRow_Error
    Dim i&
    
'    oMat01.FlushToDataSource
'
'    For i = 0 To oMat01.VisualRowCount - 1
'        If Trim(oDS_PS_MM010L.GetValue("U_CGNo", i)) = "" Then
'            oDS_PS_MM010L.RemoveRecord i   '// Mat01에 마지막라인(빈라인) 삭제
'        End If
'    Next i
'
'    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Delete_EmptyRow_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub LoadCaption()
On Error GoTo LoadCaption_Error
    If oForm01.Mode = fm_ADD_MODE Then
        oForm01.Items("Btn01").Specific.Caption = "추가"
    ElseIf oForm01.Mode = fm_OK_MODE Then
        oForm01.Items("Btn01").Specific.Caption = "확인"
    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
        oForm01.Items("Btn01").Specific.Caption = "갱신"
    End If

    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadCaption_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Function PS_PP020_AddData()
'******************************************************************************
'Function ID : PS_PP020_AddData()
'해당모듈    : PS_PP020
'기능        : 데이터 INSERT, UPDATE(기존 데이터가 존재하면 UPDATE, 아니면 INSERT)
'인수        : 없음
'반환값      : 없음
'특이사항    : 선행프로세스 일자 체크
'******************************************************************************
On Error GoTo PS_PP020_AddData_Error

    Dim loopCount As Integer
    Dim ErrNum    As String
    Dim ErrOrdNum As String '선행프로세스보다 일자가 빨라서 저장되지 않는 작번을 저장
    
    Dim Query01   As String
    Dim Query02   As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim RecordSet02 As SAPbobsCOM.Recordset
    Set RecordSet02 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry    As String
    Dim BPLID       As String
    Dim JakName     As String
    Dim SubNo1      As String
    Dim SubNo2      As String
    Dim JakMyung    As String
    Dim JakSize     As String
    Dim JakUnit     As String
    Dim RegNum      As String
    Dim ItemCode    As String
    Dim ItemName    As String
    Dim Material    As String
    Dim Unit        As String
    Dim Size        As String
    Dim ItmBsort    As String
    Dim SjDocNum    As String
    Dim SjLinNum    As String
    Dim SjWeight    As String
    Dim WrWeight    As String
    Dim SjDcDate    As String
    Dim SjDuDate    As String
    Dim SlePrice    As String
    Dim WorkGbn     As String
    Dim CardCode    As String
    Dim CardName    As String
    Dim ShipCode    As String
    Dim ShipName    As String
    Dim InOutGbn    As String
    Dim JakDate     As String
    Dim ProDate     As String
    Dim ReDate      As String
    Dim Comments    As String
    Dim PP010Doc    As String
    Dim YearPdYN    As String
    Dim OrderAmt    As String
    Dim NegoAmt     As String
    Dim TrgtAmt     As String
    Dim DrawQty     As String
    Dim ReWkRn      As String
    Dim Status      As String
    
    Dim BaseEntry   As String
    Dim BaseLine    As String
    Dim DocType     As String
    Dim CurDocDate  As String

    Dim MinusNum   As Integer '화면모드에 따라 VisualRowCount에서 빼줄 수
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("저장 중...", oMat01.VisualRowCount - 1, False)

    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
    
    If oForm01.Mode = fm_ADD_MODE Then
        MinusNum = 2
    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
        MinusNum = 1
    End If

    Call oMat01.FlushToDataSource
    For loopCount = 0 To oMat01.VisualRowCount - MinusNum
    
        DocEntry = Trim(oDS_PS_PP020H.GetValue("DocEntry", loopCount)) 'DocEntry는 프로시저에서 처리
        WorkGbn = Trim(oDS_PS_PP020H.GetValue("U_WorkGbn", loopCount))
        JakName = Trim(oDS_PS_PP020H.GetValue("U_JakName", loopCount))
        SubNo1 = Trim(oDS_PS_PP020H.GetValue("U_SubNo1", loopCount))
        SubNo2 = Trim(oDS_PS_PP020H.GetValue("U_SubNo2", loopCount))
        JakMyung = Trim(oDS_PS_PP020H.GetValue("U_JakMyung", loopCount))
        JakSize = Trim(oDS_PS_PP020H.GetValue("U_JakSize", loopCount))
        JakUnit = Trim(oDS_PS_PP020H.GetValue("U_JakUnit", loopCount))
        RegNum = Trim(oDS_PS_PP020H.GetValue("U_RegNum", loopCount))
        ItemCode = Trim(oDS_PS_PP020H.GetValue("U_ItemCode", loopCount))
        ItemName = Make_ItemName(Trim(oDS_PS_PP020H.GetValue("U_ItemName", loopCount)))
        Material = Trim(oDS_PS_PP020H.GetValue("U_Material", loopCount))
        Unit = Trim(oDS_PS_PP020H.GetValue("U_Unit", loopCount))
        Size = Trim(oDS_PS_PP020H.GetValue("U_Size", loopCount))
        ItmBsort = Trim(oDS_PS_PP020H.GetValue("U_ItmBSort", loopCount))
        SjDocNum = Trim(oDS_PS_PP020H.GetValue("U_SjDocNum", loopCount))
        SjLinNum = Trim(oDS_PS_PP020H.GetValue("U_SjLinNum", loopCount))
        SjWeight = Trim(oDS_PS_PP020H.GetValue("U_SjWeight", loopCount))
        WrWeight = Trim(oDS_PS_PP020H.GetValue("U_WrWeight", loopCount))
        SjDcDate = Trim(oDS_PS_PP020H.GetValue("U_SjDcDate", loopCount))
        SjDuDate = Trim(oDS_PS_PP020H.GetValue("U_SjDuDate", loopCount))
        SlePrice = Trim(oDS_PS_PP020H.GetValue("U_SlePrice", loopCount))
        CardCode = Trim(oDS_PS_PP020H.GetValue("U_CardCode", loopCount))
        CardName = Trim(oDS_PS_PP020H.GetValue("U_CardName", loopCount))
        ShipCode = Trim(oDS_PS_PP020H.GetValue("U_ShipCode", loopCount))
        ShipName = Trim(oDS_PS_PP020H.GetValue("U_ShipName", loopCount))
        Comments = Trim(oDS_PS_PP020H.GetValue("U_Comments", loopCount))
        InOutGbn = Trim(oDS_PS_PP020H.GetValue("U_InOutGbn", loopCount))
        JakDate = Trim(oDS_PS_PP020H.GetValue("U_JakDate", loopCount))
        ProDate = Trim(oDS_PS_PP020H.GetValue("U_ProDate", loopCount))
        ReDate = Trim(oDS_PS_PP020H.GetValue("U_ReDate", loopCount))
        PP010Doc = Trim(oDS_PS_PP020H.GetValue("U_PP010Doc", loopCount))
        DrawQty = Trim(oDS_PS_PP020H.GetValue("U_DrawQty", loopCount))
        YearPdYN = Trim(oDS_PS_PP020H.GetValue("U_YearPdYN", loopCount))
        OrderAmt = Trim(oDS_PS_PP020H.GetValue("U_OrderAmt", loopCount))
        NegoAmt = Trim(oDS_PS_PP020H.GetValue("U_NegoAmt", loopCount))
        TrgtAmt = Trim(oDS_PS_PP020H.GetValue("U_TrgtAmt", loopCount))
        ReWkRn = Trim(oDS_PS_PP020H.GetValue("U_ReWkRn", loopCount))
        Status = "O"
        
        Query01 = "         EXEC PS_PP020_03 '"
        Query01 = Query01 & DocEntry & "','"
        Query01 = Query01 & BPLID & "','"
        Query01 = Query01 & JakName & "','"
        Query01 = Query01 & SubNo1 & "','"
        Query01 = Query01 & SubNo2 & "','"
        Query01 = Query01 & JakMyung & "','"
        Query01 = Query01 & JakSize & "','"
        Query01 = Query01 & JakUnit & "','"
        Query01 = Query01 & RegNum & "','"
        Query01 = Query01 & ItemCode & "','"
        Query01 = Query01 & ItemName & "','"
        Query01 = Query01 & Material & "','"
        Query01 = Query01 & Unit & "','"
        Query01 = Query01 & Size & "','"
        Query01 = Query01 & ItmBsort & "','"
        Query01 = Query01 & SjDocNum & "','"
        Query01 = Query01 & SjLinNum & "','"
        Query01 = Query01 & SjWeight & "','"
        Query01 = Query01 & WrWeight & "','"
        Query01 = Query01 & SjDcDate & "','"
        Query01 = Query01 & SjDuDate & "','"
        Query01 = Query01 & SlePrice & "','"
        Query01 = Query01 & WorkGbn & "','"
        Query01 = Query01 & CardCode & "','"
        Query01 = Query01 & CardName & "','"
        Query01 = Query01 & ShipCode & "','"
        Query01 = Query01 & ShipName & "','"
        Query01 = Query01 & InOutGbn & "','"
        Query01 = Query01 & JakDate & "','"
        Query01 = Query01 & ProDate & "','"
        Query01 = Query01 & ReDate & "','"
        Query01 = Query01 & Comments & "','"
        Query01 = Query01 & PP010Doc & "','"
        Query01 = Query01 & YearPdYN & "','"
        Query01 = Query01 & OrderAmt & "','"
        Query01 = Query01 & NegoAmt & "','"
        Query01 = Query01 & TrgtAmt & "','"
        Query01 = Query01 & DrawQty & "','"
        Query01 = Query01 & ReWkRn & "','"
        Query01 = Query01 & Status & "'"

        '선행프로세스 대비 일자체크_S
        'Entry = Split(SjDocLin, "-")
        BaseEntry = PP010Doc
        BaseLine = "0"
        DocType = "PS_PP020"
        CurDocDate = JakDate
     
        Query02 = "         EXEC PS_Z_CHECK_DATE '"
        Query02 = Query02 & BaseEntry & "','"
        Query02 = Query02 & BaseLine & "','"
        Query02 = Query02 & DocType & "','"
        Query02 = Query02 & CurDocDate & "'"
    
        Call RecordSet02.DoQuery(Query02)
        '선행프로세스 대비 일자체크_E
    
        If RecordSet02.Fields("ReturnValue").VALUE = "True" Then
         
            Call RecordSet01.DoQuery(Query01) '등록
         
        Else
         
            ErrOrdNum = ErrOrdNum & " [" & ItemCode & "]"
         
        End If
        
    Next

    If ErrOrdNum <> "" Then '하나라도 선행프로세스 일자가 빠른 작번이 있으면
        ErrNum = "1"
        GoTo PS_PP020_AddData_Error
    End If

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Call MDC_Com.MDC_GF_Message("전체 저장 완료!", "S")
    
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    
    PS_PP020_AddData = True

    Exit Function

PS_PP020_AddData_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    PS_PP020_AddData = False
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    
    If ErrNum = "1" Then
    
        Call Sbo_Application.MessageBox("작번등록일은 생산의뢰접수일과 같거나 늦어야합니다. 확인하십시오." & Chr(13) & ErrOrdNum, 1)
        
        '등록되지 않은 작번이 있어도 화면 Clear_S
        Call oMat01.Clear
        Call oMat01.FlushToDataSource
        Call oMat01.LoadFromDataSource
        Call Add_MatrixRow(0, True)
        '등록되지 않은 작번이 있어도 화면 Clear_E
        
    Else
    
        Call MDC_Com.MDC_GF_Message("PS_PP020_AddData_Error:" & Err.Number & " - " & Err.Description, "E")
    
    End If
    
End Function

Public Function Add_JakNum(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
'******************************************************************************
'Function ID : Add_JakNum()
'해당모듈    : PS_PP020
'기능        : 데이터 INSERT
'인수        : 없음
'반환값      : 없음
'특이사항    : 2017.02.17 패치 이후 사용 안함. PS_SD020_AddData 로 대체 (2017.02.27 송명규)
'******************************************************************************
On Error GoTo Add_JakNum_Error
    Dim i As Integer
    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry$, BPLID$, RegNum$, ItemCode$, ItemName$, Material$, Unit$, Size$, ItmBsort$, CpName$, SjDocNum$, SjLinNum$, PP010Doc$
    Dim SjQty&, SjWeight As Currency, SjDcDate$, SjDuDate$, SlePrice As Currency, WorkGbn$, CardCode$, CardName$, ShipCode$, ShipName$
    Dim Comments$, JakName$, SubNo1$, SubNo2$, Temp01$, InOutGbn$, JakDate$, ProDate$, ReDate$, WrWeight As Currency, JakMyung$, JakSize$, JakUnit$
    
    Dim DrawQty As Double
    Dim YearPdYN As String '연간품여부(2015.06.15 송명규)
    
    Dim OrderAmt As Currency '수주금액(2015.10.12 송명규)
    Dim NegoAmt As Currency 'Nego금액(2015.10.12 송명규)
    Dim TrgtAmt As Currency '목표금액(생산)(2015.10.12 송명규)
    
    oMat01.FlushToDataSource
    
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
    For i = 0 To oMat01.RowCount - 2
        WorkGbn = Trim(oDS_PS_PP020H.GetValue("U_WorkGbn", i))
        JakName = Trim(oDS_PS_PP020H.GetValue("U_JakName", i))
        SubNo1 = Trim(oDS_PS_PP020H.GetValue("U_SubNo1", i))
        SubNo2 = Trim(oDS_PS_PP020H.GetValue("U_SubNo2", i))
        JakMyung = Trim(oDS_PS_PP020H.GetValue("U_JakMyung", i))
        JakSize = Trim(oDS_PS_PP020H.GetValue("U_JakSize", i))
        JakUnit = Trim(oDS_PS_PP020H.GetValue("U_JakUnit", i))
        RegNum = Trim(oDS_PS_PP020H.GetValue("U_RegNum", i))
        ItemCode = Trim(oDS_PS_PP020H.GetValue("U_ItemCode", i))
        ItemName = Make_ItemName(Trim(oDS_PS_PP020H.GetValue("U_ItemName", i)))
        Material = Trim(oDS_PS_PP020H.GetValue("U_Material", i))
        Unit = Trim(oDS_PS_PP020H.GetValue("U_Unit", i))
        Size = Trim(oDS_PS_PP020H.GetValue("U_Size", i))
        ItmBsort = Trim(oDS_PS_PP020H.GetValue("U_ItmBSort", i))
        SjDocNum = Trim(oDS_PS_PP020H.GetValue("U_SjDocNum", i))
        SjLinNum = Trim(oDS_PS_PP020H.GetValue("U_SjLinNum", i))
'        If Trim(oDS_PS_PP020H.GetValue("U_SjQty", i)) = "" Then
'            SjQty = 0
'        Else
'            SjQty = Trim(oDS_PS_PP020H.GetValue("U_SjQty", i))
'        End If
        SjWeight = Trim(oDS_PS_PP020H.GetValue("U_SjWeight", i))
        WrWeight = Trim(oDS_PS_PP020H.GetValue("U_WrWeight", i))
        SjDcDate = Trim(oDS_PS_PP020H.GetValue("U_SjDcDate", i))
        SjDuDate = Trim(oDS_PS_PP020H.GetValue("U_SjDuDate", i))
        SlePrice = Trim(oDS_PS_PP020H.GetValue("U_SlePrice", i))
'        WorkGbn = Trim(oDS_PS_PP020H.GetValue("U_WorkGbn", i))
        CardCode = Trim(oDS_PS_PP020H.GetValue("U_CardCode", i))
        CardName = Trim(oDS_PS_PP020H.GetValue("U_CardName", i))
        ShipCode = Trim(oDS_PS_PP020H.GetValue("U_ShipCode", i))
        ShipName = Trim(oDS_PS_PP020H.GetValue("U_ShipName", i))
        Comments = Trim(oDS_PS_PP020H.GetValue("U_Comments", i))
        InOutGbn = Trim(oDS_PS_PP020H.GetValue("U_InOutGbn", i))
        JakDate = Trim(oDS_PS_PP020H.GetValue("U_JakDate", i))
        ProDate = Trim(oDS_PS_PP020H.GetValue("U_ProDate", i))
        ReDate = Trim(oDS_PS_PP020H.GetValue("U_ReDate", i))
        PP010Doc = Trim(oDS_PS_PP020H.GetValue("U_PP010Doc", i))
        
        DrawQty = Trim(oDS_PS_PP020H.GetValue("U_DrawQty", i))
        
        YearPdYN = Trim(oDS_PS_PP020H.GetValue("U_YearPdYN", i))
        
        OrderAmt = Trim(oDS_PS_PP020H.GetValue("U_OrderAmt", i))
        NegoAmt = Trim(oDS_PS_PP020H.GetValue("U_NegoAmt", i))
        TrgtAmt = Trim(oDS_PS_PP020H.GetValue("U_TrgtAmt", i))
        
        '// DocEntry
        sQry = "Select IsNULL(Max(DocEntry), 0) From [@PS_PP020H] "
        RecordSet01.DoQuery sQry
        If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
            DocEntry = 1
        Else
            DocEntry = Trim(RecordSet01.Fields(0).VALUE) + 1
        End If
        
        sQry = "INSERT INTO [@PS_PP020H]"
        sQry = sQry & " ("
        sQry = sQry & " DocEntry,"
        sQry = sQry & " DocNum,"
        sQry = sQry & " U_BPLId,"
        sQry = sQry & " U_JakName,"
        sQry = sQry & " U_SubNo1,"
        sQry = sQry & " U_SubNo2,"
        sQry = sQry & " U_JakMyung,"
        sQry = sQry & " U_JakSize,"
        sQry = sQry & " U_JakUnit,"
        sQry = sQry & " U_RegNum,"
        sQry = sQry & " U_ItemCode,"
        sQry = sQry & " U_ItemName,"
        sQry = sQry & " U_Material,"
        sQry = sQry & " U_Unit,"
        sQry = sQry & " U_Size,"
        sQry = sQry & " U_ItmBSort,"
'        sQry = sQry & " U_CpName,"
        sQry = sQry & " U_SjDocNum,"
        sQry = sQry & " U_SjLinNum,"
'        sQry = sQry & " U_SjQty,"
        sQry = sQry & " U_SjWeight,"
        sQry = sQry & " U_WrWeight,"
        sQry = sQry & " U_SjDcDate,"
        sQry = sQry & " U_SjDuDate,"
        sQry = sQry & " U_SlePrice,"
        sQry = sQry & " U_WorkGbn,"
        sQry = sQry & " U_CardCode,"
        sQry = sQry & " U_CardName,"
        sQry = sQry & " U_ShipCode,"
        sQry = sQry & " U_ShipName,"
        sQry = sQry & " U_InOutGbn,"
        sQry = sQry & " U_JakDate,"
        sQry = sQry & " U_ProDate,"
        sQry = sQry & " U_ReDate,"
        sQry = sQry & " U_Comments,"
        sQry = sQry & " U_PP010Doc,"
        sQry = sQry & " U_YearPdYN,"
        sQry = sQry & " U_OrderAmt," '수주금액
        sQry = sQry & " U_NegoAmt," 'Nego금액
        sQry = sQry & " U_TrgtAmt," '목표금액
        sQry = sQry & " U_Status,"
        sQry = sQry & " U_DrawQty,"
        sQry = sQry & " CreateDate"
        sQry = sQry & " ) "
        sQry = sQry & "VALUES("
        sQry = sQry & DocEntry & ","
        sQry = sQry & DocEntry & ","
        sQry = sQry & "'" & BPLID & "',"
        sQry = sQry & "'" & JakName & "',"
        sQry = sQry & "'" & SubNo1 & "',"
        sQry = sQry & "'" & SubNo2 & "',"
        sQry = sQry & "'" & JakMyung & "',"
        sQry = sQry & "'" & JakSize & "',"
        sQry = sQry & "'" & JakUnit & "',"
        sQry = sQry & "'" & RegNum & "',"
        sQry = sQry & "'" & ItemCode & "',"
        sQry = sQry & "'" & ItemName & "',"
        sQry = sQry & "'" & Material & "',"
        sQry = sQry & "'" & Unit & "',"
        sQry = sQry & "'" & Size & "',"
        sQry = sQry & "'" & ItmBsort & "',"
'        sQry = sQry & "'" & CpName & "',"
        sQry = sQry & "'" & SjDocNum & "',"
        sQry = sQry & "'" & SjLinNum & "',"
'        sQry = sQry & "'" & SjQty & "',"
        sQry = sQry & "'" & SjWeight & "',"
        sQry = sQry & "'" & WrWeight & "',"
        sQry = sQry & "'" & SjDcDate & "',"
        sQry = sQry & "'" & SjDuDate & "',"
        sQry = sQry & "'" & SlePrice & "',"
        sQry = sQry & "'" & WorkGbn & "',"
        sQry = sQry & "'" & CardCode & "',"
        sQry = sQry & "'" & CardName & "',"
        sQry = sQry & "'" & ShipCode & "',"
        sQry = sQry & "'" & ShipName & "',"
        sQry = sQry & "'" & InOutGbn & "',"
        sQry = sQry & "'" & JakDate & "',"
        sQry = sQry & "'" & ProDate & "',"
        sQry = sQry & "'" & ReDate & "',"
        sQry = sQry & "'" & Comments & "',"
        sQry = sQry & "'" & PP010Doc & "',"
        sQry = sQry & "'" & YearPdYN & "',"
        sQry = sQry & "'" & OrderAmt & "'," '수주금액
        sQry = sQry & "'" & NegoAmt & "'," 'Nego금액
        sQry = sQry & "'" & TrgtAmt & "'," '목표금액
        sQry = sQry & "'O',"
        sQry = sQry & "'" & DrawQty & "',"
        sQry = sQry & " GETDATE()"
        sQry = sQry & ")"
        RecordSet01.DoQuery sQry
        
        sQry = "Update [@PS_PP010H] Set U_Status = 'C' Where DocEntry = '" & PP010Doc & "'"
        RecordSet01.DoQuery sQry
    Next
    
    MDC_Com.MDC_GF_Message "작번등록 완료!", "S"
    
    Set RecordSet01 = Nothing
    Add_JakNum = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_JakNum_Error:
    Add_JakNum = False
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "Add_JakNum_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Function Update_JakNum(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
'******************************************************************************
'Function ID : Update_JakNum()
'해당모듈    : PS_PP020
'기능        : 데이터 UPDATE
'인수        : 없음
'반환값      : 없음
'특이사항    : 2017.02.17 패치 이후 사용 안함. PS_SD020_AddData 로 대체 (2017.02.17 송명규)
'******************************************************************************
On Error GoTo Update_JakNum_Error
    Dim i As Integer
    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim RegNum$, ItemCode$, ItemName$, Material$, Unit$, Size$, ItmBsort$, CpName$, SjDocNum$, SjLinNum$
    Dim SjQty&, SjWeight As Currency, SjDcDate$, SjDuDate$, SlePrice As Currency, WorkGbn$, CardCode$, CardName$, ShipCode$, ShipName$
    Dim PuDate$, ProDate$, Comments$, JakName$, SubNo1$, SubNo2$, Temp01$, Status$, InOutGbn$, WrWeight As Currency, ReDate$, JakMyung$, JakSize$, JakUnit$
    
    Dim DrawQty As Double
    Dim ReWkRn As String '재작업사유(2013.09.10 송명규 추가)
    Dim YearPdYN As String '연간품여부(2015.06.15 송명규 추가)
    
    Dim OrderAmt As Currency '수주금액(2015.10.12 송명규)
    Dim NegoAmt As Currency 'Nego금액(2015.10.12 송명규)
    Dim TrgtAmt As Currency '목표금액(생산)(2015.10.12 송명규)
    
    oMat01.FlushToDataSource
    
    For i = 0 To oMat01.RowCount - 1
'        ItemCode = Trim(oDS_PS_PP020H.GetValue("U_ItemCode", i))
'        ItemName = Trim(oDS_PS_PP020H.GetValue("U_ItemName", i))
'        Material = Trim(oDS_PS_PP020H.GetValue("U_Material", i))
'        Unit = Trim(oDS_PS_PP020H.GetValue("U_Unit", i))
'        Size = Trim(oDS_PS_PP020H.GetValue("U_Size", i))
'        ItmBSort = Trim(oDS_PS_PP020H.GetValue("U_ItmBSort", i))
'        CpName = Trim(oDS_PS_PP020H.GetValue("U_CpName", i))
'        SjDocNum = Trim(oDS_PS_PP020H.GetValue("U_SjDocNum", i))
'        SjLinNum = Trim(oDS_PS_PP020H.GetValue("U_SjLinNum", i))
'        If Trim(oDS_PS_PP020H.GetValue("U_SjQty", i)) = "" Then
'            SjQty = 0
'        Else
'            SjQty = Trim(oDS_PS_PP020H.GetValue("U_SjQty", i))
'        End If
'        SjWeight = Trim(oDS_PS_PP020H.GetValue("U_SjWeight", i))
'        SjDcDate = Trim(oDS_PS_PP020H.GetValue("U_SjDcDate", i))
'        SjDuDate = Trim(oDS_PS_PP020H.GetValue("U_SjDuDate", i))
'        SlePrice = Trim(oDS_PS_PP020H.GetValue("U_SlePrice", i))
'        WorkGbn = Trim(oDS_PS_PP020H.GetValue("U_WorkGbn", i))
        JakMyung = Trim(oDS_PS_PP020H.GetValue("U_JakMyung", i))
        JakSize = Trim(oDS_PS_PP020H.GetValue("U_JakSize", i))
        JakUnit = Trim(oDS_PS_PP020H.GetValue("U_JakUnit", i))
        CardCode = Trim(oDS_PS_PP020H.GetValue("U_CardCode", i))
        CardName = Trim(oDS_PS_PP020H.GetValue("U_CardName", i))
        ShipCode = Trim(oDS_PS_PP020H.GetValue("U_ShipCode", i))
        ShipName = Trim(oDS_PS_PP020H.GetValue("U_ShipName", i))
        InOutGbn = Trim(oDS_PS_PP020H.GetValue("U_InOutGbn", i))
        ProDate = Trim(oDS_PS_PP020H.GetValue("U_ProDate", i))
        ReDate = Trim(oDS_PS_PP020H.GetValue("U_ReDate", i))
        Comments = Trim(oDS_PS_PP020H.GetValue("U_Comments", i))
        JakName = Trim(oDS_PS_PP020H.GetValue("U_JakName", i))
        SubNo1 = Trim(oDS_PS_PP020H.GetValue("U_SubNo1", i))
        SubNo2 = Trim(oDS_PS_PP020H.GetValue("U_SubNo2", i))
        Status = Trim(oDS_PS_PP020H.GetValue("U_Status", i))
        WrWeight = Trim(oDS_PS_PP020H.GetValue("U_WrWeight", i))
        DrawQty = Trim(oDS_PS_PP020H.GetValue("U_DrawQty", i))
        ReWkRn = Trim(oDS_PS_PP020H.GetValue("U_ReWkRn", i))
        YearPdYN = Trim(oDS_PS_PP020H.GetValue("U_YearPdYN", i))
        
        OrderAmt = Trim(oDS_PS_PP020H.GetValue("U_OrderAmt", i))
        NegoAmt = Trim(oDS_PS_PP020H.GetValue("U_NegoAmt", i))
        TrgtAmt = Trim(oDS_PS_PP020H.GetValue("U_TrgtAmt", i))
        
        sQry = "UPDATE [@PS_PP020H] "
        sQry = sQry & "SET "
        sQry = sQry & "U_JakMyung = '" & JakMyung & "', "
        sQry = sQry & "U_JakSize = '" & JakSize & "', "
        sQry = sQry & "U_JakUnit = '" & JakUnit & "', "
        sQry = sQry & "U_CardCode = '" & CardCode & "', "
        sQry = sQry & "U_CardName = '" & CardName & "', "
        sQry = sQry & "U_ShipCode = '" & ShipCode & "', "
        sQry = sQry & "U_ShipName = '" & ShipName & "', "
        sQry = sQry & "U_InOutGbn = '" & InOutGbn & "', "
        sQry = sQry & "U_ProDate = '" & ProDate & "', "
        sQry = sQry & "U_ReDate = '" & ReDate & "', "
        sQry = sQry & "U_WrWeight = '" & WrWeight & "', "
        sQry = sQry & "U_Comments = '" & Comments & "', "
        sQry = sQry & "U_Status = '" & Status & "', "
        sQry = sQry & "U_DrawQty = '" & DrawQty & "', "
        sQry = sQry & "U_YearPdYN = '" & YearPdYN & "', "
        sQry = sQry & "U_OrderAmt = '" & OrderAmt & "', " '수주금액
        sQry = sQry & "U_NegoAmt = '" & NegoAmt & "', " 'Nego금액
        sQry = sQry & "U_TrgtAmt = '" & TrgtAmt & "', " '목표금액
        sQry = sQry & "U_ReWkRn = '" & ReWkRn & "' "
        sQry = sQry & "Where U_JakName = '" & JakName & "' And U_SubNo1 = '" & SubNo1 & "' And U_SubNo2 = '" & SubNo2 & "'"
    
        RecordSet01.DoQuery sQry
    Next

    MDC_Com.MDC_GF_Message "작번등록 갱신 완료!", "S"
    
    Set RecordSet01 = Nothing
    Update_JakNum = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Update_JakNum_Error:
    Update_JakNum = False
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "Update_JakNum_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Sub LoadData()
On Error GoTo LoadData_Error

    Dim i As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim RegNum$, ItemCode$, ItemName$, Material$, Unit$, Size$, ItmBsort$, CpName$, SjDocNum$, SjLinNum$
    Dim SjQty&, SjWeight As Currency, SjDcDate$, SjDuDate$, SlePrice As Currency, WorkGbn$, CardCode$, CardName$, ShipCode$, ShipName$
    Dim PuDate$, ProDate$, Comments$, JakName$, SubNo1$, SubNo2$, Temp01$, Status$, BPLID$, JakDateFr$, JakDateTo$, JakGbn$
        
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
    ItmBsort = Trim(oForm01.Items("ItmBSort").Specific.VALUE)
    WorkGbn = Trim(oForm01.Items("WorkGbn").Specific.VALUE)
    CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
    ItemCode = Trim(oForm01.Items("ItemCode").Specific.VALUE)
    JakDateFr = Trim(oForm01.Items("JakDateFr").Specific.VALUE)
    JakDateTo = Trim(oForm01.Items("JakDateTo").Specific.VALUE)
    JakGbn = Trim(oForm01.Items("JakGbn").Specific.VALUE)
    
    If BPLID = "" Then BPLID = "%"
    If CardCode = "" Then CardCode = "%"
    If ItemCode = "" Then ItemCode = "%"
    If JakDateFr = "" Then JakDateFr = "19000101"
    If JakDateTo = "" Then JakDateTo = "20991231"
    If JakGbn = "" Then JakGbn = "%"
    
    sQry = "      EXEC [PS_PP020_01] '"
    sQry = sQry & BPLID & "','"
    sQry = sQry & ItmBsort & "','"
    sQry = sQry & WorkGbn & "','"
    sQry = sQry & CardCode & "','"
    sQry = sQry & ItemCode & "','"
    sQry = sQry & JakDateFr & "','"
    sQry = sQry & JakDateTo & "','"
    sQry = sQry & JakGbn & "'"
    
    Call oRecordSet01.DoQuery(sQry)
    
    oMat01.Clear
    oDS_PS_PP020H.Clear
    
    If (oRecordSet01.RecordCount = 0) Then
        MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
        Set oRecordSet01 = Nothing
        Exit Sub
    End If
    
    oForm01.Freeze True
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_PP020H.Size Then
            oDS_PS_PP020H.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_PP020H.Offset = i
        oDS_PS_PP020H.setValue "DocNum", i, i + 1
        oDS_PS_PP020H.setValue "DocEntry", i, Trim(oRecordSet01.Fields("DocEntry").VALUE)
        oDS_PS_PP020H.setValue "U_JakName", i, Trim(oRecordSet01.Fields("U_JakName").VALUE)
        oDS_PS_PP020H.setValue "U_SubNo1", i, Trim(oRecordSet01.Fields("U_SubNo1").VALUE)
        oDS_PS_PP020H.setValue "U_SubNo2", i, Trim(oRecordSet01.Fields("U_SubNo2").VALUE)
        oDS_PS_PP020H.setValue "U_JakMyung", i, Trim(oRecordSet01.Fields("U_JakMyung").VALUE)
        oDS_PS_PP020H.setValue "U_JakSize", i, Trim(oRecordSet01.Fields("U_JakSize").VALUE)
        oDS_PS_PP020H.setValue "U_JakUnit", i, Trim(oRecordSet01.Fields("U_JakUnit").VALUE)
        oDS_PS_PP020H.setValue "U_RegNum", i, Trim(oRecordSet01.Fields("U_RegNum").VALUE)
        oDS_PS_PP020H.setValue "U_ItemCode", i, Trim(oRecordSet01.Fields("U_ItemCode").VALUE)
        oDS_PS_PP020H.setValue "U_ItemName", i, Trim(oRecordSet01.Fields("U_ItemName").VALUE)
        oDS_PS_PP020H.setValue "U_Material", i, Trim(oRecordSet01.Fields("U_Material").VALUE)
        oDS_PS_PP020H.setValue "U_Unit", i, Trim(oRecordSet01.Fields("U_Unit").VALUE)
        oDS_PS_PP020H.setValue "U_Size", i, Trim(oRecordSet01.Fields("U_Size").VALUE)
        oDS_PS_PP020H.setValue "U_ItmBSort", i, Trim(oRecordSet01.Fields("U_ItmBSort").VALUE)
'        oDS_PS_PP020H.setValue "U_CpName", i, Trim(oRecordset01.Fields("U_CpName").Value)
        oDS_PS_PP020H.setValue "U_SjDocNum", i, Trim(oRecordSet01.Fields("U_SjDocNum").VALUE)
        oDS_PS_PP020H.setValue "U_SjLinNum", i, Trim(oRecordSet01.Fields("U_SjLinNum").VALUE)
'        oDS_PS_PP020H.setValue "U_SjQty", i, Trim(oRecordset01.Fields("U_SjQty").Value)
        oDS_PS_PP020H.setValue "U_SjWeight", i, Trim(oRecordSet01.Fields("U_SjWeight").VALUE)
        oDS_PS_PP020H.setValue "U_WrWeight", i, Trim(oRecordSet01.Fields("U_WrWeight").VALUE)
        oDS_PS_PP020H.setValue "U_SjDcDate", i, Format(Trim(oRecordSet01.Fields("U_SjDcDate").VALUE), "YYYYMMDD")
        oDS_PS_PP020H.setValue "U_SjDuDate", i, Format(Trim(oRecordSet01.Fields("U_SjDuDate").VALUE), "YYYYMMDD")
        oDS_PS_PP020H.setValue "U_SlePrice", i, Trim(oRecordSet01.Fields("U_SlePrice").VALUE)
        oDS_PS_PP020H.setValue "U_WorkGbn", i, Trim(oRecordSet01.Fields("U_WorkGbn").VALUE)
        oDS_PS_PP020H.setValue "U_CardCode", i, Trim(oRecordSet01.Fields("U_CardCode").VALUE)
        oDS_PS_PP020H.setValue "U_CardName", i, Trim(oRecordSet01.Fields("U_CardName").VALUE)
        oDS_PS_PP020H.setValue "U_ShipCode", i, Trim(oRecordSet01.Fields("U_ShipCode").VALUE)
        oDS_PS_PP020H.setValue "U_ShipName", i, Trim(oRecordSet01.Fields("U_ShipName").VALUE)
        oDS_PS_PP020H.setValue "U_JakDate", i, Format(Trim(oRecordSet01.Fields("U_JakDate").VALUE), "YYYYMMDD")
        oDS_PS_PP020H.setValue "U_ProDate", i, Format(Trim(oRecordSet01.Fields("U_ProDate").VALUE), "YYYYMMDD")
        oDS_PS_PP020H.setValue "U_ReDate", i, Format(Trim(oRecordSet01.Fields("U_ReDate").VALUE), "YYYYMMDD")
        oDS_PS_PP020H.setValue "U_Comments", i, Trim(oRecordSet01.Fields("U_Comments").VALUE)
        oDS_PS_PP020H.setValue "U_InOutGbn", i, Trim(oRecordSet01.Fields("U_InOutGbn").VALUE)
        oDS_PS_PP020H.setValue "U_Status", i, Trim(oRecordSet01.Fields("U_Status").VALUE)
        oDS_PS_PP020H.setValue "U_PP010Doc", i, Trim(oRecordSet01.Fields("U_PP010Doc").VALUE)
        oDS_PS_PP020H.setValue "U_DrawQty", i, Trim(oRecordSet01.Fields("U_DrawQty").VALUE)
        oDS_PS_PP020H.setValue "U_ReWkRn", i, Trim(oRecordSet01.Fields("U_ReWkRn").VALUE)
        oDS_PS_PP020H.setValue "U_YearPdYN", i, Trim(oRecordSet01.Fields("U_YearPdYN").VALUE)
        
        oDS_PS_PP020H.setValue "U_OrderAmt", i, Trim(oRecordSet01.Fields("U_OrderAmt").VALUE) '수주금액
        oDS_PS_PP020H.setValue "U_NegoAmt", i, Trim(oRecordSet01.Fields("U_NegoAmt").VALUE) 'Nego금액
        oDS_PS_PP020H.setValue "U_TrgtAmt", i, Trim(oRecordSet01.Fields("U_TrgtAmt").VALUE) '목표금액
        
        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
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

Public Function Add_SubJakNum(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
On Error GoTo Add_SubJakNum_Error

    Dim i As Integer
    Dim ErrNum&
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry$, BPLID$, RegNum$, ItemCode$, ItemName$, Material$, Unit$, Size$, ItmBsort$, CpName$, SjDocNum$, SjLinNum$
    Dim SjQty&, SjWeight As Currency, SjDcDate$, SjDuDate$, SlePrice As Currency, WorkGbn$, CardCode$, CardName$, ShipCode$, ShipName$
    Dim Comments$, JakName$, SubNo1$, SubNo2$, Temp01$, InOutGbn$, JakDate$, ProDate$, ReDate$, WrWeight As Currency, JakMyung$, JakSize$, JakUnit$
    Dim Status As String
    
    Dim ReWkRn As String '재작업 사유
    Dim ReWork As String '재작업 구분
    
    Dim OrderAmt As Currency '수주금액(2015.10.12 송명규)
    Dim NegoAmt  As Currency 'Nego금액(2015.10.12 송명규)
    Dim TrgtAmt  As Currency '목표금액(생산)(2015.10.12 송명규)
    
    Dim ErrOrdNum   As String '선행프로세스보다 일자가 빨라서 저장되지 않는 작번을 저장
    Dim BaseEntry   As String
    Dim BaseLine    As String
    Dim DocType     As String
    Dim CurDocDate  As String
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("저장 중...", oMat01.VisualRowCount - 1, False)
    
    oMat01.FlushToDataSource
    
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
    '// 값 체크
    For i = 0 To oMat02.RowCount - 1
        JakName = Trim(oMat02.Columns("JakName").Cells(i + 1).Specific.VALUE)
        SubNo1 = Trim(oMat02.Columns("SubNo1").Cells(i + 1).Specific.VALUE)
        SubNo2 = Trim(oMat02.Columns("SubNo2").Cells(i + 1).Specific.VALUE)
        
        ReWkRn = Trim(oMat02.Columns("ReWkRn").Cells(i + 1).Specific.VALUE)
        ReWork = oForm01.Items("ReWork").Specific.VALUE
        
        sQry = "Select COUNT(*) From [@PS_PP020H] Where U_JakName = '" & JakName & "' And U_SubNo1 = '" & SubNo1 & "' And U_SubNo2 = '" & SubNo2 & "'"
        oRecordSet01.DoQuery sQry
        
        If oRecordSet01.Fields(0).VALUE > 0 Then
            ErrNum = 1
            GoTo Add_SubJakNum_Error
        Else
        
            If ReWork = "20" And ReWkRn = "" Then
                ErrNum = 2
                GoTo Add_SubJakNum_Error
            End If
            
        End If
    Next i
    
    For i = 0 To oMat02.RowCount - 1
    
        WorkGbn = Trim(oMat02.Columns("WorkGbn").Cells(i + 1).Specific.VALUE)
        JakName = Trim(oMat02.Columns("JakName").Cells(i + 1).Specific.VALUE)
        SubNo1 = Trim(oMat02.Columns("SubNo1").Cells(i + 1).Specific.VALUE)
        SubNo2 = Trim(oMat02.Columns("SubNo2").Cells(i + 1).Specific.VALUE)
        JakMyung = Trim(oMat02.Columns("JakMyung").Cells(i + 1).Specific.VALUE)
        JakSize = Trim(oMat02.Columns("JakSize").Cells(i + 1).Specific.VALUE)
        JakUnit = Trim(oMat02.Columns("JakUnit").Cells(i + 1).Specific.VALUE)
        RegNum = Trim(oMat02.Columns("RegNum").Cells(i + 1).Specific.VALUE)
        ItemCode = Trim(oMat02.Columns("ItemCode").Cells(i + 1).Specific.VALUE)
        ItemName = Make_ItemName(Trim(oMat02.Columns("ItemName").Cells(i + 1).Specific.VALUE))
        Material = Trim(oMat02.Columns("Material").Cells(i + 1).Specific.VALUE)
        Unit = Trim(oMat02.Columns("Unit").Cells(i + 1).Specific.VALUE)
        Size = Trim(oMat02.Columns("Size").Cells(i + 1).Specific.VALUE)
        ItmBsort = Trim(oMat02.Columns("ItmBSort").Cells(i + 1).Specific.VALUE)
        SjDocNum = Trim(oMat02.Columns("SjDocNum").Cells(i + 1).Specific.VALUE)
        SjLinNum = Trim(oMat02.Columns("SjLinNum").Cells(i + 1).Specific.VALUE)
        SjWeight = Trim(oMat02.Columns("SjWeight").Cells(i + 1).Specific.VALUE)
        WrWeight = Trim(oMat02.Columns("WrWeight").Cells(i + 1).Specific.VALUE)
        SjDcDate = Trim(oMat02.Columns("SjDcDate").Cells(i + 1).Specific.VALUE)
        SjDuDate = Trim(oMat02.Columns("SjDuDate").Cells(i + 1).Specific.VALUE)
        SlePrice = Trim(oMat02.Columns("SlePrice").Cells(i + 1).Specific.VALUE)
        CardCode = Trim(oMat02.Columns("CardCode").Cells(i + 1).Specific.VALUE)
        CardName = Trim(oMat02.Columns("CardName").Cells(i + 1).Specific.VALUE)
        ShipCode = Trim(oMat02.Columns("ShipCode").Cells(i + 1).Specific.VALUE)
        ShipName = Trim(oMat02.Columns("ShipName").Cells(i + 1).Specific.VALUE)
        Comments = Trim(oMat02.Columns("Comments").Cells(i + 1).Specific.VALUE)
        InOutGbn = Trim(oMat02.Columns("InOutGbn").Cells(i + 1).Specific.VALUE)
        JakDate = Trim(oMat02.Columns("JakDate").Cells(i + 1).Specific.VALUE)
        ProDate = Trim(oMat02.Columns("ProDate").Cells(i + 1).Specific.VALUE)
        ReDate = Trim(oMat02.Columns("ReDate").Cells(i + 1).Specific.VALUE)
        ReWkRn = Trim(oMat02.Columns("ReWkRn").Cells(i + 1).Specific.VALUE)
        OrderAmt = Trim(oMat02.Columns("OrderAmt").Cells(i + 1).Specific.VALUE)
        NegoAmt = Trim(oMat02.Columns("NegoAmt").Cells(i + 1).Specific.VALUE)
        TrgtAmt = Trim(oMat02.Columns("TrgtAmt").Cells(i + 1).Specific.VALUE)
        Status = "O"
        
        '// DocEntry
'        sQry = "Select IsNULL(Max(DocEntry), 0) From [@PS_PP020H] "
'        oRecordSet01.DoQuery sQry
'        If Trim(oRecordSet01.Fields(0).VALUE) = 0 Then
'            DocEntry = 1
'        Else
'            DocEntry = Trim(oRecordSet01.Fields(0).VALUE) + 1
'        End If
        
'        sQry = "INSERT INTO [@PS_PP020H]"
'        sQry = sQry & " ("
'        sQry = sQry & " DocEntry,"
'        sQry = sQry & " DocNum,"
'        sQry = sQry & " U_BPLId,"
'        sQry = sQry & " U_JakName,"
'        sQry = sQry & " U_SubNo1,"
'        sQry = sQry & " U_SubNo2,"
'        sQry = sQry & " U_JakMyung,"
'        sQry = sQry & " U_JakSize,"
'        sQry = sQry & " U_JakUnit,"
'        sQry = sQry & " U_RegNum,"
'        sQry = sQry & " U_ItemCode,"
'        sQry = sQry & " U_ItemName,"
'        sQry = sQry & " U_Material,"
'        sQry = sQry & " U_Unit,"
'        sQry = sQry & " U_Size,"
'        sQry = sQry & " U_ItmBSort,"
''        sQry = sQry & " U_CpName,"
'        sQry = sQry & " U_SjDocNum,"
'        sQry = sQry & " U_SjLinNum,"
''        sQry = sQry & " U_SjQty,"
'        sQry = sQry & " U_SjWeight,"
'        sQry = sQry & " U_WrWeight,"
'        sQry = sQry & " U_SjDcDate,"
'        sQry = sQry & " U_SjDuDate,"
'        sQry = sQry & " U_SlePrice,"
'        sQry = sQry & " U_WorkGbn,"
'        sQry = sQry & " U_CardCode,"
'        sQry = sQry & " U_CardName,"
'        sQry = sQry & " U_ShipCode,"
'        sQry = sQry & " U_ShipName,"
'        sQry = sQry & " U_InOutGbn,"
'        sQry = sQry & " U_JakDate,"
'        sQry = sQry & " U_ProDate,"
'        sQry = sQry & " U_ReDate,"
'        sQry = sQry & " U_Comments,"
'        sQry = sQry & " U_OrderAmt," '수주금액
'        sQry = sQry & " U_NegoAmt," 'Nego금액
'        sQry = sQry & " U_TrgtAmt," '목표금액
'        sQry = sQry & " U_Status,"
'        sQry = sQry & " U_ReWkRn,"
'        sQry = sQry & " CreateDate"
'        sQry = sQry & " ) "
'        sQry = sQry & "VALUES("
'        sQry = sQry & DocEntry & ","
'        sQry = sQry & DocEntry & ","
'        sQry = sQry & "'" & BPLID & "',"
'        sQry = sQry & "'" & JakName & "',"
'        sQry = sQry & "'" & SubNo1 & "',"
'        sQry = sQry & "'" & SubNo2 & "',"
'        sQry = sQry & "'" & JakMyung & "',"
'        sQry = sQry & "'" & JakSize & "',"
'        sQry = sQry & "'" & JakUnit & "',"
'        sQry = sQry & "'" & RegNum & "',"
'        sQry = sQry & "'" & ItemCode & "',"
'        sQry = sQry & "'" & ItemName & "',"
'        sQry = sQry & "'" & Material & "',"
'        sQry = sQry & "'" & Unit & "',"
'        sQry = sQry & "'" & Size & "',"
'        sQry = sQry & "'" & ItmBsort & "',"
''        sQry = sQry & "'" & CpName & "',"
'        sQry = sQry & "'" & SjDocNum & "',"
'        sQry = sQry & "'" & SjLinNum & "',"
''        sQry = sQry & "'" & SjQty & "',"
'        sQry = sQry & "'" & SjWeight & "',"
'        sQry = sQry & "'" & WrWeight & "',"
'        sQry = sQry & "'" & SjDcDate & "',"
'        sQry = sQry & "'" & SjDuDate & "',"
'        sQry = sQry & "'" & SlePrice & "',"
'        sQry = sQry & "'" & WorkGbn & "',"
'        sQry = sQry & "'" & CardCode & "',"
'        sQry = sQry & "'" & CardName & "',"
'        sQry = sQry & "'" & ShipCode & "',"
'        sQry = sQry & "'" & ShipName & "',"
'        sQry = sQry & "'" & InOutGbn & "',"
'        sQry = sQry & "'" & JakDate & "',"
'        sQry = sQry & "'" & ProDate & "',"
'        sQry = sQry & "'" & ReDate & "',"
'        sQry = sQry & "'" & Comments & "',"
'        sQry = sQry & "'" & OrderAmt & "'," '수주금액
'        sQry = sQry & "'" & NegoAmt & "'," 'Nego금액
'        sQry = sQry & "'" & TrgtAmt & "'," '목표금액
'        sQry = sQry & "'O',"
'        sQry = sQry & "'" & ReWkRn & "',"
'        sQry = sQry & " GETDATE()"
'        sQry = sQry & ")"

        sQry = "      EXEC PS_PP020_04 '"
        sQry = sQry & DocEntry & "','"
        sQry = sQry & BPLID & "','"
        sQry = sQry & JakName & "','"
        sQry = sQry & SubNo1 & "','"
        sQry = sQry & SubNo2 & "','"
        sQry = sQry & JakMyung & "','"
        sQry = sQry & JakSize & "','"
        sQry = sQry & JakUnit & "','"
        sQry = sQry & RegNum & "','"
        sQry = sQry & ItemCode & "','"
        sQry = sQry & ItemName & "','"
        sQry = sQry & Material & "','"
        sQry = sQry & Unit & "','"
        sQry = sQry & Size & "','"
        sQry = sQry & ItmBsort & "','"
        sQry = sQry & SjDocNum & "','"
        sQry = sQry & SjLinNum & "','"
        sQry = sQry & SjWeight & "','"
        sQry = sQry & WrWeight & "','"
        sQry = sQry & SjDcDate & "','"
        sQry = sQry & SjDuDate & "','"
        sQry = sQry & SlePrice & "','"
        sQry = sQry & WorkGbn & "','"
        sQry = sQry & CardCode & "','"
        sQry = sQry & CardName & "','"
        sQry = sQry & ShipCode & "','"
        sQry = sQry & ShipName & "','"
        sQry = sQry & InOutGbn & "','"
        sQry = sQry & JakDate & "','"
        sQry = sQry & ProDate & "','"
        sQry = sQry & ReDate & "','"
        sQry = sQry & Comments & "','"
        sQry = sQry & OrderAmt & "','"
        sQry = sQry & NegoAmt & "','"
        sQry = sQry & TrgtAmt & "','"
        sQry = sQry & Status & "','"
        sQry = sQry & ReWkRn & "'"

        Call oRecordSet01.DoQuery(sQry) '등록

    Next
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    MDC_Com.MDC_GF_Message "Sub작번등록 완료!", "S"
    
    oDS_PS_TEMPTABLE.Clear
    oMat02.Clear
    
    Set oRecordSet01 = Nothing
    Add_SubJakNum = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_SubJakNum_Error:
    Add_SubJakNum = False
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번째 라인의 작번이 이미 등록되어 있습니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "재작업인 경우는 재작업 사유를 필수로 입력해야합니다.  [" & i + 1 & "행]", "E"
    Else
        MDC_Com.MDC_GF_Message "Add_SubJakNum_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Function
