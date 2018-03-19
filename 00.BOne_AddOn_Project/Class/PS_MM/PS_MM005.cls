VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_MM005"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_MM005.cls
'//  Module         : MM
'//  Description    : 구매요청
'//  FormType       : PS_MM005
'//  Create Date    : 2010.09.19
'//  Modified Date  : 2015.12.03 NGY 저장부분을 PROCEDURE로 바꿈
'//  Creator        : Ryu Yung Jo
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01         As SAPbouiCOM.Form
Public oMat01          As SAPbouiCOM.Matrix
Private oDS_PS_MM005H  As SAPbouiCOM.DBDataSource    '등록헤더

Private oLast_Item_UID As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID  As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row  As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&
Private oForm01_Mode&
Private oLast_RightClick_CgNum&

Private oMenuUID$

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm(Optional MenuUID$)
On Error GoTo LoadForm_Error

    Dim i           As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01   As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_MM005.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    
    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_MM005_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm01.SupportedModes = -1
    oMenuUID = MenuUID
    
    If oMenuUID = "PS_MM005_1" Then
        oForm01.Mode = fm_FIND_MODE
        oForm01_Mode = fm_FIND_MODE
        oForm01.EnableMenu ("1281"), False        '// 찾기
        oForm01.EnableMenu ("1282"), False        '// 추가
        oForm01.EnableMenu ("1293"), False         '// 행삭제
        oForm01.EnableMenu ("1299"), False '행닫기
        oForm01.Title = "구매 요청 승인 [PS_MM005_1]"
    Else
        oForm01.Mode = fm_ADD_MODE
        oForm01_Mode = fm_ADD_MODE
        oForm01.EnableMenu ("1281"), True        '// 찾기
        oForm01.EnableMenu ("1282"), True        '// 추가
        oForm01.EnableMenu ("1293"), True         '// 행삭제
        oForm01.EnableMenu ("1299"), True '행닫기
    End If
    
    oMenuUID = MenuUID
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting(oMenuUID)
    Call Initialization
    Call Add_MatrixRow(0, True)
    Call LoadCaption
    Call FormItemEnabled(oMenuUID)
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1285"), False        '// 복원
    oForm01.EnableMenu ("1284"), False         '// 취소
            
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
    Dim ErrNum&
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Dim SumQty        As Long
    Dim SumWeight     As Currency
    Dim ItemCode$, ItemName$, Size$, Qty&, Weight As Currency, Unit$, RequestDate$, DueDate$, ItemType$, RequestNo$
    Dim RFC_Sender$
    Dim Seq&
    Dim ChildForm01   As Variant
    Set ChildForm01 = New PS_SM010
    Dim ChildForm02   As Variant
    Set ChildForm02 = New PS_SM030
    Dim sQry          As String
    Dim A          As String
    Dim oRecordSet01  As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "Btn01" Then
                    If oForm01_Mode = fm_ADD_MODE Then
                        If HeaderSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        If MatrixSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        If Add_PurchaseDemand(pval) = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        oMat01.Clear
                        oMat01.FlushToDataSource
                        oMat01.LoadFromDataSource
                        Add_MatrixRow 0, True
                    ElseIf oForm01_Mode = fm_UPDATE_MODE Then
                        If MatrixSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        If Update_PurchaseDemand(pval) = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        oForm01_Mode = fm_OK_MODE
                        Call LoadCaption
                    ElseIf oForm01_Mode = fm_OK_MODE Then
                        oForm01.Close
                    End If
                ElseIf pval.ItemUID = "Btn02" Then
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    oForm01_Mode = fm_OK_MODE
                    Call LoadCaption
                    Call LoadData
                    Call FormItemEnabled
               ElseIf pval.ItemUID = "Btn03" Then
               
                    If updateforuse_date = False Then ' 일괄업데이트 버튼 누를경우 실행.
                        BubbleEvent = False
                        Exit Sub
                    End If
                    
                End If
'et_KEY_DOWN ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "CntcCode" Then
                        If oForm01.Items("CntcCode").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "ItemCode" Then
                        If oForm01.Items("ItemCode").Specific.VALUE = "" Then
                            Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row)
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "ItemCode" Then
                            If oForm01.Items("OrdType").Specific.VALUE = "60" Then
                                If oMat01.Columns("ItemCode").Cells(pval.Row).Specific.VALUE = "" Then
                                    Sbo_Application.ActivateMenuItem ("7425")
                                    BubbleEvent = False
                                End If
                            Else
                                If oMat01.Columns("ItemCode").Cells(pval.Row).Specific.VALUE = "" Then
                                    If oForm01.Items("BOM_CHECK").Specific.Checked = True Then
                                        '//포장사업팀 제품BOM사용시
                                        Call ChildForm02.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row)
                                        BubbleEvent = False
                                    Else
                                        Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row)
                                        BubbleEvent = False
                                    End If
                                End If
                            End If
                        ElseIf pval.ColUID = "PP030DL" Then
                            If oMat01.Columns("PP030DL").Cells(pval.Row).Specific.VALUE = "" Then
                                Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                            End If
                        ElseIf pval.ColUID = "CntcCode" Then
                            If oMat01.Columns("CntcCode").Cells(pval.Row).Specific.VALUE = "" Then
                                Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                            End If
                        ElseIf pval.ColUID = "MachCode" Then
                            If oMat01.Columns("MachCode").Cells(pval.Row).Specific.VALUE = "" Then
                                Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                            End If
                        ElseIf pval.ColUID = "ProcCode" Then
                            If oMat01.Columns("ProcCode").Cells(pval.Row).Specific.VALUE = "" Then
                                Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                            End If
                        End If
                    End If
                End If
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            
            If pval.ItemUID = "Mat01" Then
        
                If pval.Row = 0 Then
                
                    oMat01.Columns(pval.ColUID).TitleObject.Sortable = True
                    oMat01.FlushToDataSource
                
                End If
            
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
                If pval.ItemUID = "OrdType" Or pval.ItemUID = "BPLId" Then
                    oForm01.Freeze True
                    oMat01.Clear
                    oDS_PS_MM005H.Clear
                    If oForm01_Mode = fm_ADD_MODE Then
                        Call Add_MatrixRow(0, False)
                    End If
                    
                    If oMenuUID = "PS_MM005_1" Then
                        oMat01.Columns("PP030DL").Editable = False
                        oMat01.Columns("ItemCode").Editable = False
'                        oMat01.Columns("OutItmCd").Editable = False
                        oMat01.Columns("OutSize").Editable = False
                        oMat01.Columns("OutUnit").Editable = False
                    Else
                        If Trim(oForm01.Items("OrdType").Specific.VALUE) = "30" And Trim(oForm01.Items("BPLId").Specific.VALUE) = "2" Then
                            oMat01.Columns("PP030DL").Editable = True
                            oMat01.Columns("ItemCode").Editable = False
                            oMat01.Columns("ItemName").Editable = True
'                            oMat01.Columns("OutItmCd").Editable = True
                            oMat01.Columns("OutSize").Editable = True
                            oMat01.Columns("OutUnit").Editable = True
                        ElseIf Trim(oForm01.Items("OrdType").Specific.VALUE) = "40" And Trim(oForm01.Items("BPLId").Specific.VALUE) = "2" Then
                            oMat01.Columns("PP030DL").Editable = True
                            oMat01.Columns("ItemCode").Editable = False
                            oMat01.Columns("ItemName").Editable = True
'                            oMat01.Columns("OutItmCd").Editable = True
                            oMat01.Columns("OutSize").Editable = True
                            oMat01.Columns("OutUnit").Editable = True
                        ElseIf Trim(oForm01.Items("OrdType").Specific.VALUE) = "60" Then
                            oMat01.Columns("ItemCode").Editable = True
                            oMat01.Columns("ItemName").Editable = True
                            oMat01.Columns("OutSize").Editable = True
                            oMat01.Columns("OutUnit").Editable = True
                            oMat01.Columns("Qty").Editable = False
                        Else
                            oMat01.Columns("PP030DL").Editable = False
                            oMat01.Columns("ItemCode").Editable = True
                            oMat01.Columns("ItemName").Editable = False
'                            oMat01.Columns("OutItmCd").Editable = False
                            oMat01.Columns("OutSize").Editable = False
                            oMat01.Columns("OutUnit").Editable = False
                        End If
                    End If
                    oForm01.Freeze False
                ElseIf pval.ItemUID = "Mat01" Then
                    If pval.ColUID = "OKYN" Then
                        oMat01.FlushToDataSource
                        If Trim(oMat01.Columns("OKYN").Cells(pval.Row).Specific.VALUE) = "Y" Then
                            oMat01.Columns("OKDate").Cells(pval.Row).Specific.VALUE = Format(Now, "YYYYMMDD")
                        ElseIf Trim(oMat01.Columns("OKYN").Cells(pval.Row).Specific.VALUE) = "N" Then
                            oMat01.Columns("OKDate").Cells(pval.Row).Specific.VALUE = ""
                        End If
                    End If
                    If oForm01_Mode = fm_ADD_MODE Then
                    Else
                        oForm01_Mode = fm_UPDATE_MODE
                        Call LoadCaption
                    End If
                End If
            Case et_CLICK: '//6
'et_DOUBLE_CLICK ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_DOUBLE_CLICK: '//7
                If pval.ItemUID = "Mat01" And pval.Row = "0" And pval.ColUID = "OKYN" And oMenuUID = "PS_MM005_1" Then
                    Dim OkYN$
                    oForm01.Freeze True
                    oMat01.FlushToDataSource
                    If Trim(oDS_PS_MM005H.GetValue("U_OKYN", 0)) = "" Or Trim(oDS_PS_MM005H.GetValue("U_OKYN", 0)) = "N" Then
                        OkYN = "Y"
                    ElseIf Trim(oDS_PS_MM005H.GetValue("U_OKYN", 0)) = "Y" Then
                        OkYN = "N"
                    End If
                    For i = 0 To oMat01.VisualRowCount - 1
                        oDS_PS_MM005H.setValue "U_OKYN", i, OkYN
                        If OkYN = "Y" Then
                            oDS_PS_MM005H.setValue "U_OKDate", i, Format(Now, "YYYYMMDD")
                        ElseIf OkYN = "N" Then
                            oDS_PS_MM005H.setValue "U_OKDate", i, ""
                        End If
                    Next i
                    oMat01.LoadFromDataSource
                    oMat01.Columns("OKDate").Cells(1).Click ct_Regular
                    If oForm01_Mode = fm_ADD_MODE Then
                    Else
                        oForm01_Mode = fm_UPDATE_MODE
                        Call LoadCaption
                    End If
                    oForm01.Freeze False
                End If
            Case et_MATRIX_LINK_PRESSED '//8
'et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "CntcCode" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "ItemCode" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "ItemCode" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        ElseIf pval.ColUID = "ItemName" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        ElseIf pval.ColUID = "PP030DL" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        ElseIf pval.ColUID = "Qty" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        ElseIf pval.ColUID = "Weight" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
'                        ElseIf pval.ColUID = "Price" Then
'                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
'                        ElseIf pval.ColUID = "LinTotal" Then
'                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        ElseIf pval.ColUID = "CntcCode" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        ElseIf pval.ColUID = "ProcCode" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        End If
                        
                        If oForm01_Mode = fm_ADD_MODE Then
                        Else
                            oForm01_Mode = fm_UPDATE_MODE
                            Call LoadCaption
                        End If
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
                Set oDS_PS_MM005H = Nothing
        End Select
    End If
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
    oForm01.Freeze False
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
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
'[1293:행삭제] //////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1293": '행삭제
                If oForm01_Mode <> fm_ADD_MODE Then
                    
                    sQry = "Select U_OKYN From [@PS_MM005H] "
                    sQry = sQry & "where U_CgNum = '" & oLast_RightClick_CgNum & "'"
                    oRecordSet01.DoQuery sQry
                    If oRecordSet01.Fields(0).VALUE = "Y" Then
                        MDC_Com.MDC_GF_Message "결재가 된 구매요청은 삭제할 수 없습니다.", "E"
                        BubbleEvent = False
                        Exit Sub
                    End If
                    
                    sQry = "Select Count(a.U_CGNo) From [@PS_MM010L] a Inner Join [@PS_MM010H] b On a.DocEntry = b.DocEntry "
                    sQry = sQry & "where a.U_CGNo = '" & oLast_RightClick_CgNum & "' And b.Status = 'O'"
                    oRecordSet01.DoQuery sQry
                    
                    If oRecordSet01.Fields(0).VALUE = 0 Then
                        ReturnValue = Sbo_Application.MessageBox("해당 라인의 구매요청을 삭제합니다. 삭제 후 복원할 수 없습니다. 삭제하시겠습니까?", 1, "&확인", "&취소")
                        Select Case ReturnValue
                            Case 1:
                                sQry = "Delete [@PS_MM005H] Where DocEntry = '" & oLast_RightClick_CgNum & "'"
                                oRecordSet01.DoQuery sQry
    
                                oLast_RightClick_CgNum = 0
                            Case 2:
                                Sbo_Application.MessageBox ("실행이 취소되었습니다.")
                                BubbleEvent = False
                                Exit Sub
                        End Select
                    Else
                        MDC_Com.MDC_GF_Message "해당 구매요청건에 해당하는 구매견적이 있습니다. 삭제할 수 없습니다.", "E"
                        BubbleEvent = False
                        Exit Sub
                    End If
                End If
            Case "1299": '행닫기(취소)
            
                If oForm01_Mode <> fm_ADD_MODE Then
                    sQry = "Select U_OKYN From [@PS_MM005H] "
                    sQry = sQry & "where U_CgNum = '" & oLast_RightClick_CgNum & "'"
                    oRecordSet01.DoQuery sQry
                    
                    If oRecordSet01.Fields(0).VALUE = "Y" Then
                        MDC_Com.MDC_GF_Message "결재가 된 구매요청은 닫기(취소)할 수 없습니다.", "E"
                        BubbleEvent = False
                        Exit Sub
                    End If
                    
                    sQry = "Select Count(a.U_CGNo) From [@PS_MM010L] a Inner Join [@PS_MM010H] b On a.DocEntry = b.DocEntry "
                    sQry = sQry & "where a.U_CGNo = '" & oLast_RightClick_CgNum & "' And b.Status = 'O'"
                    oRecordSet01.DoQuery sQry
                    
                    If oRecordSet01.Fields(0).VALUE = 0 Then
                        ReturnValue = Sbo_Application.MessageBox("해당 라인의 구매요청을 닫기(취소)합니다. 진행하시겠습니까?", 1, "&확인", "&취소")
                        Select Case ReturnValue
                            Case 1:
                                sQry = "           UPDATE [@PS_MM005H] "
                                sQry = sQry & " SET       Status = 'C',"
                                sQry = sQry & "             Canceled = 'Y',"
                                sQry = sQry & "             UpdateDate = GETDATE(),"
                                sQry = sQry & "             UserSign = '" & Sbo_Company.UserSignature & "'"
                                sQry = sQry & " Where    DocEntry = '" & oLast_RightClick_CgNum & "'"
                                oRecordSet01.DoQuery sQry
    
                                oLast_RightClick_CgNum = 0
                                
                                Call MDC_Com.MDC_GF_Message("해당 구매요청건이 닫기(취소)되었습니다.", "W")
                            Case 2:
                                Sbo_Application.MessageBox ("실행이 취소되었습니다.")
                                BubbleEvent = False
                                Exit Sub
                        End Select
                    Else
                        MDC_Com.MDC_GF_Message "해당 구매요청건에 해당하는 구매견적이 있습니다. 닫기(취소)할 수 없습니다.", "E"
                        BubbleEvent = False
                        Exit Sub
                    End If
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
                    oDS_PS_MM005H.RemoveRecord oDS_PS_MM005H.Size - 1       '// Mat01에 마지막라인(빈라인) 삭제
                    oMat01.Clear
                    oMat01.LoadFromDataSource
                    oForm01.Freeze False
                End If
'[1281:찾기] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1281": '찾기
                oForm01_Mode = fm_FIND_MODE
                oForm01.Items("SumQty").Specific.VALUE = 0
                oForm01.Items("SumWeight").Specific.VALUE = 0
                oMat01.Clear
                oMat01.FlushToDataSource
                oMat01.LoadFromDataSource
                Add_MatrixRow 0, True
                FormItemEnabled
                
                'oForm01_Mode = fm_OK_MODE
                Call LoadCaption
'                oForm01.Items("OrdType").Click ct_Regular
'[1282:추가] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1282": '추가
                oForm01.Freeze True
                oForm01_Mode = fm_ADD_MODE
                oForm01.Items("SumQty").Specific.VALUE = 0
                oForm01.Items("SumWeight").Specific.VALUE = 0
                oMat01.Clear
                oMat01.FlushToDataSource
                oMat01.LoadFromDataSource
                Add_MatrixRow 0, True
                FormItemEnabled
                Call LoadCaption
                oForm01.Freeze False
'[1288~1291:네비게이션] /////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                FormItemEnabled
                If oMat01.VisualRowCount > 0 Then
                    If oMat01.Columns("CGNo").Cells(oMat01.VisualRowCount).Specific.VALUE <> "" Then
                        If oDS_PS_MM005H.GetValue("Status", 0) = "O" Then
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
        If eventInfo.Row > 0 Then
            If Trim(oMat01.Columns("CgNum").Cells(eventInfo.Row).Specific.VALUE) <> "" Then
                oLast_RightClick_CgNum = Trim(oMat01.Columns("CgNum").Cells(eventInfo.Row).Specific.VALUE)
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
    Set oDS_PS_MM005H = oForm01.DataSources.DBDataSources("@PS_MM005H")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    
    oMat01.Columns("PP030DL").Editable = False
    oMat01.Columns("ItemCode").Editable = True
    
    Call oForm01.DataSources.UserDataSources.Add("BOM_CHECK", dt_SHORT_TEXT, 10)
    Call oForm01.Items("BOM_CHECK").Specific.DataBind.SetBound(True, "", "BOM_CHECK")
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
CreateItems_Error:
    MDC_Com.MDC_GF_Message "CreateItems_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub ComboBox_Setting(Optional MenuUID$)
On Error GoTo ComboBox_Setting_Error
    '//콤보에 기본값설정
    Dim oCombo       As SAPbouiCOM.ComboBox
    Dim sQry         As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.DataSources.UserDataSources.Add("DocDateFr", dt_DATE, 8)
    oForm01.Items("DocDateFr").Specific.DataBind.SetBound True, "", "DocDateFr"
    'oForm01.Items("DocDateFr").Specific.VALUE = Format(Date - 1, "YYYYMMDD") 'Format(Now, "YYYY-MM") & "-01"
    If oMenuUID = "PS_MM005_1" Then
        oForm01.DataSources.UserDataSources.Item("DocDateFr").VALUE = Format(Date - 90, "YYYY-MM") & "-01"
    Else
        oForm01.DataSources.UserDataSources.Item("DocDateFr").VALUE = Format(Now, "YYYY-MM") & "-01"
    End If
    
    Call oForm01.DataSources.UserDataSources.Add("DocDateTo", dt_DATE, 8)
    oForm01.Items("DocDateTo").Specific.DataBind.SetBound True, "", "DocDateTo"
    'oForm01.Items("DocDateTo").Specific.VALUE = Format(Date, "YYYYMMDD") 'Date
    oForm01.DataSources.UserDataSources.Item("DocDateTo").VALUE = Date

    Call oForm01.DataSources.UserDataSources.Add("SumQty", dt_SUM)
    oForm01.Items("SumQty").Specific.DataBind.SetBound True, "", "SumQty"
    
    Call oForm01.DataSources.UserDataSources.Add("SumWeight", dt_QUANTITY)
    oForm01.Items("SumWeight").Specific.DataBind.SetBound True, "", "SumWeight"
    
    '// 품목구분
    Set oCombo = oForm01.Items("OrdType").Specific
    sQry = "SELECT Code, Name From [@PSH_ORDTYP] Order by Code"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oMat01.Columns("OrdType").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    oCombo.Select "0", psk_Index
    
    '// 사업장
    Set oCombo = oForm01.Items("BPLId").Specific
    sQry = "SELECT BPLId, BPLName From [OBPL]"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oMat01.Columns("BPLId").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    '//청구부서
    Set oCombo = oForm01.Items("DeptCode").Specific
    sQry = "SELECT Code, Name From [OUDP] Order by Code"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oMat01.Columns("DeptCode").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    '//사용처
    sQry = "Select PrcCode, PrcName From [OPRC] Where DimCode = '1' Order by PrcCode"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oMat01.Columns("UseDept").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    '//구매방식
'    sQry = "SELECT Code, Name From [@PSH_RETYPE] Order by Code"
'    oRecordSet01.DoQuery sQry
'    Do Until oRecordSet01.EOF
'        oMat01.Columns("ReType").ValidValues.Add Trim(oRecordSet01.Fields(0).Value), Trim(oRecordSet01.Fields(1).Value)
'        oRecordSet01.MoveNext
'    Loop
    
    '//대분류
    Set oCombo = oForm01.Items("ItmBSort").Specific
    sQry = "SELECT Code, Name From [@PSH_ITMBSORT] Order by Code"
    oRecordSet01.DoQuery sQry
    oCombo.ValidValues.Add "ALL", "ALL"
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    '//품목타입
    Set oCombo = oForm01.Items("ItemType").Specific
    sQry = "SELECT Code, Name From [@PSH_SHAPE] Order by Code"
    oRecordSet01.DoQuery sQry
    oCombo.ValidValues.Add "ALL", "ALL"
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    '//결재여부
    Set oCombo = oForm01.Items("OKYN").Specific
    oCombo.ValidValues.Add "Y", "결재"
    oCombo.ValidValues.Add "N", "미결재"
    If oMenuUID = "PS_MM005_1" Then
        oCombo.Select 1, psk_Index
    Else
        oCombo.Select 0, psk_Index
    End If
    
    oMat01.Columns("DivideYN").ValidValues.Add "N", "분할입고"
    oMat01.Columns("DivideYN").ValidValues.Add "Y", "일반입고"
    
    '//외주사유코드
    sQry = "         SELECT     T1.U_Minor AS [Code],"
    sQry = sQry & "             T1.U_CdName AS [Value]"
    sQry = sQry & "  FROM       [@PS_SY001H] AS T0"
    sQry = sQry & "             INNER JOIN"
    sQry = sQry & "             [@PS_SY001L] AS T1"
    sQry = sQry & "                 ON T0.Code = T1.Code"
    sQry = sQry & "  WHERE      T0.Code = 'P201'"
    sQry = sQry & "             AND T1.U_UseYN = 'Y'"
    sQry = sQry & "  ORDER BY   T1.U_Seq"
    
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oMat01.Columns("OutCode").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
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

    Dim oCombo As SAPbouiCOM.ComboBox
    
    '//아이디별 사업장 세팅
    Set oCombo = oForm01.Items("BPLId").Specific
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
    '//아이디별 사번 세팅
    If oMenuUID = "PS_MM005_1" Then
        oForm01.Items("CntcCode").Specific.VALUE = ""
    Else
        oForm01.Items("CntcCode").Specific.VALUE = MDC_PS_Common.User_MSTCOD
    End If
    
    
    '//아이디별 부서 세팅
    Set oCombo = oForm01.Items("DeptCode").Specific
    oCombo.Select MDC_PS_Common.User_DeptCode, psk_ByValue
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

Sub FormItemEnabled(Optional MenuUID$)
On Error GoTo FormItemEnabled_Error
    If MenuUID = "PS_MM005_1" Then '승인
        If oForm01_Mode = fm_FIND_MODE Then
            oForm01.Items("CntcCode").Click ct_Regular
            oForm01.Items("OrdType").Enabled = True
            oForm01.Items("BPLId").Enabled = True
            oForm01.Items("CntcCode").Enabled = True
            oForm01.Items("DeptCode").Enabled = True
            oForm01.Items("DocDateFr").Enabled = True
            oForm01.Items("DocDateTo").Enabled = True
            oForm01.Items("CgNumFr").Enabled = True
            oForm01.Items("CgNumTo").Enabled = True
            oForm01.Items("ItemCode").Enabled = True
            oForm01.Items("ItmBSort").Enabled = True
            oForm01.Items("ItmMSort").Enabled = True
'            oForm01.Items("ItemType").Enabled = True
            oForm01.Items("OKYN").Enabled = True
            oForm01.Items("Btn02").Enabled = True
            
            oMat01.Columns("PP030DL").Editable = False
            oMat01.Columns("ItemCode").Editable = False
            oMat01.Columns("ItemName").Editable = False
            oMat01.Columns("Qty").Editable = False
            oMat01.Columns("Weight").Editable = False
            oMat01.Columns("BPLId").Editable = False
            oMat01.Columns("CgNum").Editable = False
            oMat01.Columns("DocDate").Editable = False
            oMat01.Columns("DueDate").Editable = False
            oMat01.Columns("CntcCode").Editable = False
            oMat01.Columns("CntcName").Editable = False
            oMat01.Columns("DeptCode").Editable = False
            oMat01.Columns("UseDept").Editable = False
            oMat01.Columns("Auto").Editable = False
            oMat01.Columns("QCYN").Editable = False
'            oMat01.Columns("ReType").Editable = False
            oMat01.Columns("Note").Editable = False
            oMat01.Columns("Comments").Editable = False
            oMat01.Columns("IvQty").Editable = False
            oMat01.Columns("IvWeight").Editable = False
            oMat01.Columns("OrdType").Editable = False
            oMat01.Columns("OKYN").Editable = True
            oMat01.Columns("OKDate").Editable = True
        End If
    Else
        If oForm01_Mode = fm_ADD_MODE Then
            oForm01.Items("CntcCode").Click ct_Regular
            oForm01.Items("OrdType").Enabled = True
            oForm01.Items("BPLId").Enabled = True
            oForm01.Items("CntcCode").Enabled = True
            oForm01.Items("DeptCode").Enabled = True
            oForm01.Items("DocDateFr").Enabled = False
            oForm01.Items("DocDateTo").Enabled = False
            oForm01.Items("CgNumFr").Enabled = False
            oForm01.Items("CgNumTo").Enabled = False
            oForm01.Items("ItemCode").Enabled = False
            oForm01.Items("ItmBSort").Enabled = False
            oForm01.Items("ItmMSort").Enabled = False
'            oForm01.Items("ItemType").Enabled = False
            oForm01.Items("OKYN").Enabled = False
            oForm01.Items("Btn02").Enabled = False
            
            oMat01.Columns("PP030DL").Editable = False
            oMat01.Columns("ItemCode").Editable = True
            oMat01.Columns("ItemName").Editable = False
            oMat01.Columns("Qty").Editable = True
            oMat01.Columns("Weight").Editable = True
            oMat01.Columns("BPLId").Editable = False
            oMat01.Columns("CgNum").Editable = False
            oMat01.Columns("DocDate").Editable = True
            oMat01.Columns("DueDate").Editable = True
            oMat01.Columns("CntcCode").Editable = False
            oMat01.Columns("CntcName").Editable = False
            oMat01.Columns("DeptCode").Editable = False
            oMat01.Columns("UseDept").Editable = True
            oMat01.Columns("Auto").Editable = True
            oMat01.Columns("QCYN").Editable = True
'            oMat01.Columns("ReType").Editable = True
            oMat01.Columns("Note").Editable = True
            oMat01.Columns("Comments").Editable = True
            oMat01.Columns("IvQty").Editable = False
            oMat01.Columns("IvWeight").Editable = False
            oMat01.Columns("OrdType").Editable = False
            oMat01.Columns("OKYN").Editable = False
            oMat01.Columns("OKDate").Editable = False
        ElseIf oForm01_Mode = fm_FIND_MODE Then
            oForm01.Items("CntcCode").Click ct_Regular
            oForm01.Items("OrdType").Enabled = True
            oForm01.Items("BPLId").Enabled = True
            oForm01.Items("CntcCode").Enabled = True
            oForm01.Items("DeptCode").Enabled = True
            oForm01.Items("DocDateFr").Enabled = True
            oForm01.Items("DocDateTo").Enabled = True
            oForm01.Items("CgNumFr").Enabled = True
            oForm01.Items("CgNumTo").Enabled = True
            oForm01.Items("ItemCode").Enabled = True
            oForm01.Items("ItmBSort").Enabled = True
            oForm01.Items("ItmMSort").Enabled = True
'            oForm01.Items("ItemType").Enabled = True
            oForm01.Items("OKYN").Enabled = True
            oForm01.Items("Btn02").Enabled = True
            
            oMat01.Clear
            
            oMat01.Columns("PP030DL").Editable = False
            oMat01.Columns("ItemCode").Editable = True
            oMat01.Columns("ItemName").Editable = False
            oMat01.Columns("Qty").Editable = True
            oMat01.Columns("Weight").Editable = True
            oMat01.Columns("BPLId").Editable = False
            oMat01.Columns("CgNum").Editable = False
            oMat01.Columns("DocDate").Editable = True
            oMat01.Columns("DueDate").Editable = True
            oMat01.Columns("CntcCode").Editable = False
            oMat01.Columns("CntcName").Editable = False
            oMat01.Columns("DeptCode").Editable = False
            oMat01.Columns("UseDept").Editable = True
            oMat01.Columns("Auto").Editable = True
            oMat01.Columns("QCYN").Editable = True
'            oMat01.Columns("ReType").Editable = True
            oMat01.Columns("Note").Editable = True
            oMat01.Columns("Comments").Editable = True
            oMat01.Columns("IvQty").Editable = False
            oMat01.Columns("IvWeight").Editable = False
            oMat01.Columns("OrdType").Editable = False
            oMat01.Columns("OKYN").Editable = False
            oMat01.Columns("OKDate").Editable = False
        ElseIf (oForm01_Mode = fm_OK_MODE) Then
            oForm01.Items("CntcCode").Click ct_Regular
            oForm01.Items("OrdType").Enabled = True
            oForm01.Items("BPLId").Enabled = True
            oForm01.Items("CntcCode").Enabled = True
            oForm01.Items("DeptCode").Enabled = True
            oForm01.Items("DocDateFr").Enabled = False
            oForm01.Items("DocDateTo").Enabled = False
            oForm01.Items("CgNumFr").Enabled = False
            oForm01.Items("CgNumTo").Enabled = False
            oForm01.Items("ItemCode").Enabled = False
            oForm01.Items("ItmBSort").Enabled = False
            oForm01.Items("ItmMSort").Enabled = False
'            oForm01.Items("ItemType").Enabled = False
            oForm01.Items("OKYN").Enabled = False
            oForm01.Items("Btn02").Enabled = False
            
            If oForm01.Items("OKYN").Specific.VALUE = "Y" Then
                oMat01.Columns("PP030DL").Editable = False
                oMat01.Columns("ItemCode").Editable = False
                oMat01.Columns("ItemName").Editable = False
                oMat01.Columns("Qty").Editable = False
                oMat01.Columns("Weight").Editable = False
                oMat01.Columns("BPLId").Editable = False
                oMat01.Columns("CgNum").Editable = False
                oMat01.Columns("DocDate").Editable = False
                oMat01.Columns("DueDate").Editable = False
                oMat01.Columns("CntcCode").Editable = False
                oMat01.Columns("CntcName").Editable = False
                oMat01.Columns("DeptCode").Editable = False
                oMat01.Columns("UseDept").Editable = False
                oMat01.Columns("Auto").Editable = False
                oMat01.Columns("QCYN").Editable = False
                oMat01.Columns("Note").Editable = False
                oMat01.Columns("Comments").Editable = False
                oMat01.Columns("IvQty").Editable = False
                oMat01.Columns("IvWeight").Editable = False
                oMat01.Columns("OrdType").Editable = False
                oMat01.Columns("OKYN").Editable = False
                oMat01.Columns("OKDate").Editable = False
                oMat01.Columns("Status").Editable = False
            Else
                oMat01.Columns("PP030DL").Editable = False
                oMat01.Columns("ItemCode").Editable = True
                oMat01.Columns("ItemName").Editable = False
                oMat01.Columns("Qty").Editable = True
                oMat01.Columns("Weight").Editable = True
                oMat01.Columns("BPLId").Editable = False
                oMat01.Columns("CgNum").Editable = False
                oMat01.Columns("DocDate").Editable = True
                oMat01.Columns("DueDate").Editable = True
                oMat01.Columns("CntcCode").Editable = False
                oMat01.Columns("CntcName").Editable = False
                oMat01.Columns("DeptCode").Editable = False
                oMat01.Columns("UseDept").Editable = True
                oMat01.Columns("Auto").Editable = True
                oMat01.Columns("QCYN").Editable = True
    
                oMat01.Columns("Note").Editable = True
                oMat01.Columns("Comments").Editable = True
                oMat01.Columns("IvQty").Editable = False
                oMat01.Columns("IvWeight").Editable = False
                oMat01.Columns("OrdType").Editable = False
                oMat01.Columns("OKYN").Editable = False
                oMat01.Columns("OKDate").Editable = False
                oMat01.Columns("Status").Editable = True
            End If
        End If
        
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_MM005H.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_MM005H.Offset = oRow
    oDS_PS_MM005H.setValue "DocNum", oRow, oRow + 1
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
    Dim ItemCode$, Qty&, SumQty As Long, SumWeight As Currency
    Dim Calculate_Weight As Currency
    Dim Calculate_Qty As Currency
    Dim oRecordSet01    As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
        Case "CntcCode"
            sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" & Trim(oForm01.Items("CntcCode").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            oForm01.Items("CntcName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
        Case "ItemCode"
            sQry = "Select ItemName From OITM Where ItemCode = '" & Trim(oForm01.Items("ItemCode").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            oForm01.Items("ItemName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
        Case "Mat01"
            If oCol = "ItemCode" Then
                oForm01.Freeze True
                
                '//부자재인경우 11자리(통합부자재코드)가 아니면 안됨
'                If oForm01.Items("OrdType").Specific.VALUE = "20" And Len(Trim(oMat01.Columns("ItemCode").Cells(oRow).Specific.VALUE)) <> 11 Then
'                    Sbo_Application.MessageBox ("통합부자재 코드로만 청구가능합니다.")
'                    oForm01.Freeze False
'                Else
                
                    If oForm01_Mode = fm_ADD_MODE Then
                        If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 0) And Trim(oMat01.Columns("ItemCode").Cells(oRow).Specific.VALUE) <> "" Then
                            oMat01.FlushToDataSource
                            Add_MatrixRow oMat01.RowCount, False
                        End If
                    End If
                    
                    
                    
                    If oForm01.Items("OrdType").Specific.VALUE = "60" Then
                        sQry = "Select U_ItemName, U_OutUnit From [@PS_MM005H] Where U_ItemCode = '" & Trim(oMat01.Columns("ItemCode").Cells(oRow).Specific.VALUE) & "'"
                        oRecordSet01.DoQuery sQry
                        oMat01.Columns("ItemName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                        oMat01.Columns("OutUnit").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields(1).VALUE)
                    Else
                        sQry = "Select ItemName, FrgnName, BuyUnitMsr From OITM Where ItemCode = '" & Trim(oMat01.Columns("ItemCode").Cells(oRow).Specific.VALUE) & "'"
                        oRecordSet01.DoQuery sQry
                        oMat01.Columns("ItemName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                        oMat01.Columns("OutUnit").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields(2).VALUE)
        '                oMat01.Columns("FrgnName").Cells(oRow).Specific.Value = Trim(oRecordSet01.Fields(1).Value)
                        
                        sQry = "Select Sum(OnHand) From OITW Where ItemCode = '" & Trim(oMat01.Columns("ItemCode").Cells(oRow).Specific.VALUE) & "'"
                        oRecordSet01.DoQuery sQry
                        Calculate_Qty = MDC_PS_Common.Calculate_Qty(Trim(oMat01.Columns("ItemCode").Cells(oRow).Specific.VALUE), Trim(oRecordSet01.Fields(0).VALUE))
                        oMat01.Columns("IvQty").Cells(oRow).Specific.VALUE = Calculate_Qty
                        oMat01.Columns("IvWeight").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                    End If
                    
                    oMat01.Columns("DocDate").Cells(oRow).Specific.VALUE = Format(Date, "YYYYMMDD")
                    oMat01.Columns("DueDate").Cells(oRow).Specific.VALUE = Format(Date, "YYYYMMDD")
                    oMat01.Columns("Auto").Cells(oRow).Specific.Select "N"
                    oMat01.Columns("QCYN").Cells(oRow).Specific.Select "N"
    '                oMat01.Columns("ReType").Cells(oRow).Specific.Select "10"
                    oMat01.Columns("OKYN").Cells(oRow).Specific.Select "N"
                    oMat01.Columns("DivideYN").Cells(oRow).Specific.Select "N"
    '                For i = 0 To oMat01.VisualRowCount - 2
    '                    SumQty = SumQty + oMat01.Columns("Qty").Cells(i + 1).Specific.Value
    '                    SumWeight = SumWeight + oMat01.Columns("Weight").Cells(i + 1).Specific.Value
    '                Next i
    '                oForm01.Items("SumQty").Specific.Value = SumQty
    '                oForm01.Items("SumWeight").Specific.Value = SumWeight
    '
    
                    oMat01.Columns("ItemCode").Cells(oRow).Click ct_Regular
                    oForm01.Freeze False
'                End If
            ElseIf oCol = "ItemName" Then
                If oRow = oMat01.RowCount Then
                    oMat01.FlushToDataSource
                    Add_MatrixRow oMat01.RowCount, False
                End If
                oMat01.Columns("Auto").Cells(oRow).Specific.Select "N"
                oMat01.Columns("QCYN").Cells(oRow).Specific.Select "N"
                oMat01.Columns("OKYN").Cells(oRow).Specific.Select "N"
                oMat01.Columns("DivideYN").Cells(oRow).Specific.Select "N"
                oMat01.Columns("ItemName").Cells(oRow).Click ct_Regular
            ElseIf oCol = "PP030DL" Then
                oForm01.Freeze True
                If oForm01_Mode = fm_ADD_MODE Then
                    If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 0) And Trim(oMat01.Columns("PP030DL").Cells(oRow).Specific.VALUE) <> "" Then
                        oMat01.FlushToDataSource
                        Add_MatrixRow oMat01.RowCount, False
                    End If
                End If
                                
'                sQry = "Select a.U_ItemCode, a.U_ItemName, c.FrgnName, b.U_CpCode, b.U_CpName "
'                sQry = sQry & "From [@PS_PP030H] a Inner Join [@PS_PP030M] b On a.DocEntry = b.DocEntry Inner Join [OITM] c On a.U_ItemCode = c.ItemCode "
'                sQry = sQry & "Where a.U_BPLId = '2' And b.U_WorkGbn = '30' And a.Status = 'O' "
'                sQry = sQry & "And Convert(Nvarchar(10), a.DocEntry) + '-' + Convert(Nvarchar(10), b.LineId) = '" & Trim(oMat01.Columns("PP030DL").Cells(oRow).Specific.Value) & "'"
                sQry = "EXEC [PS_MM005_02] '" & Trim(oForm01.Items("BPLId").Specific.VALUE) & "', '" & Trim(oForm01.Items("OrdType").Specific.VALUE) & "', "
                sQry = sQry & "'" & Trim(oMat01.Columns("PP030DL").Cells(oRow).Specific.VALUE) & "', '2'"
                oRecordSet01.DoQuery sQry
                                
                oMat01.Columns("ItemCode").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_ItemCode").VALUE)
                oMat01.Columns("ItemName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("ItemName").VALUE) '2014.03.14 송명규 추가
                oMat01.Columns("OutSize").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("Size").VALUE) '2014.03.14 송명규 추가
'                oMat01.Columns("FrgnName").Cells(oRow).Specific.Value = Trim(oRecordSet01.Fields("FrgnName").Value)
                
                oMat01.Columns("DocDate").Cells(oRow).Specific.VALUE = Format(Date, "YYYYMMDD")
                oMat01.Columns("DueDate").Cells(oRow).Specific.VALUE = Format(Date, "YYYYMMDD")
                oMat01.Columns("Auto").Cells(oRow).Specific.Select "N"
                oMat01.Columns("QCYN").Cells(oRow).Specific.Select "N"
'                oMat01.Columns("ReType").Cells(oRow).Specific.Select "10"
                oMat01.Columns("OKYN").Cells(oRow).Specific.Select "N"
                oMat01.Columns("DivideYN").Cells(oRow).Specific.Select "N"
                
                oMat01.Columns("ProcCode").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_CpCode").VALUE)
                oMat01.Columns("ProcName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_CpName").VALUE)
                
'                For i = 0 To oMat01.VisualRowCount - 2
'                    SumQty = SumQty + oMat01.Columns("Qty").Cells(i + 1).Specific.Value
'                    SumWeight = SumWeight + oMat01.Columns("Weight").Cells(i + 1).Specific.Value
'                Next i
'                oForm01.Items("SumQty").Specific.Value = SumQty
'                oForm01.Items("SumWeight").Specific.Value = SumWeight
                
                Call oMat01.AutoResizeColumns '2014.03.14 송명규 추가
                
                oMat01.Columns("PP030DL").Cells(oRow).Click ct_Regular
                oForm01.Freeze False
            ElseIf oCol = "Qty" Then
                oForm01.Freeze True
                oMat01.FlushToDataSource
                ItemCode = Trim(oDS_PS_MM005H.GetValue("U_ItemCode", oRow - 1))
                Qty = oDS_PS_MM005H.GetValue("U_Qty", oRow - 1)

                Calculate_Weight = MDC_PS_Common.Calculate_Weight(ItemCode, Qty, Trim(oForm01.Items("BPLId").Specific.VALUE))
                
'                            oMat01.Columns("Weight").Cells(pval.Row).Specific.Value = Calculate_Weight
                oDS_PS_MM005H.setValue "U_Weight", oRow - 1, Calculate_Weight '이론중량
                
                oMat01.LoadFromDataSource
                
'                oMat01.FlushToDataSource
'                If Trim(oMat01.Columns("Price").Cells(oRow).Specific.Value) = 0 Then
'                    oDS_PS_MM005H.setValue "U_LinTotal", oRow - 1, 0
'                Else
'                    oDS_PS_MM005H.setValue "U_LinTotal", oRow - 1, Trim(oMat01.Columns("Weight").Cells(oRow).Specific.Value) * Trim(oMat01.Columns("Price").Cells(oRow).Specific.Value)
'                End If
'                oMat01.LoadFromDataSource
                
                For i = 0 To oMat01.VisualRowCount - 2
                    If oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE = "" Then
                        SumQty = SumQty
                    Else
                        SumQty = SumQty + oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE
                    End If
                    SumWeight = SumWeight + oMat01.Columns("Weight").Cells(i + 1).Specific.VALUE
                Next i
                oForm01.Items("SumQty").Specific.VALUE = SumQty
                oForm01.Items("SumWeight").Specific.VALUE = SumWeight
                
                oMat01.Columns("Qty").Cells(oRow).Click
                oForm01.Freeze False
            ElseIf oCol = "Weight" Then
                oForm01.Freeze True
                For i = 0 To oMat01.VisualRowCount - 2
                    If oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE = "" Then
                        SumQty = SumQty
                    Else
                        SumQty = SumQty + oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE
                    End If
                    SumWeight = SumWeight + oMat01.Columns("Weight").Cells(i + 1).Specific.VALUE
                Next i
                oForm01.Items("SumQty").Specific.VALUE = SumQty
                oForm01.Items("SumWeight").Specific.VALUE = SumWeight
                
                oMat01.Columns("Weight").Cells(oRow).Click
                oForm01.Freeze False
'            ElseIf oCol = "Price" Then
'                oMat01.FlushToDataSource
'                If Trim(oMat01.Columns("Price").Cells(oRow).Specific.Value) = 0 Then
'                    oDS_PS_MM005H.setValue "U_LinTotal", oRow - 1, 0
'                Else
'                    oDS_PS_MM005H.setValue "U_LinTotal", oRow - 1, Trim(oMat01.Columns("Weight").Cells(oRow).Specific.Value) * Trim(oMat01.Columns("Price").Cells(oRow).Specific.Value)
'                End If
'                oMat01.LoadFromDataSource
'                oMat01.Columns("Price").Cells(oRow).Click
'             ElseIf oCol = "LinTotal" Then
'                oMat01.FlushToDataSource
'                If Trim(oMat01.Columns("LinTotal").Cells(oRow).Specific.Value) = 0 Then
'                    oDS_PS_MM005H.setValue "U_Price", oRow - 1, 0
'                Else
'                    If Trim(oMat01.Columns("Weight").Cells(oRow).Specific.Value) = 0 Then
'                        oDS_PS_MM005H.setValue "U_Price", oRow - 1, 0
'                    Else
'                        oDS_PS_MM005H.setValue "U_Price", oRow - 1, Trim(oMat01.Columns("LinTotal").Cells(oRow).Specific.Value) / Trim(oMat01.Columns("Weight").Cells(oRow).Specific.Value)
'                    End If
'                End If
'                oMat01.LoadFromDataSource
'                oMat01.Columns("Price").Cells(oRow).Click
            ElseIf oCol = "CntcCode" Then
'                sQry = "Select lastName + firstName From OHEM Where EmpID = '" & Trim(oMat01.Columns("CntcCode").Cells(oRow).Specific.Value) & "'"
                sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" & Trim(oMat01.Columns("CntcCode").Cells(oRow).Specific.VALUE) & "'"
                oRecordSet01.DoQuery sQry
            
                oMat01.Columns("CntcName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            ElseIf oCol = "ProcCode" Then
                sQry = "select U_CpName From [@ps_pp001L] Where U_CpCode = '" & Trim(oMat01.Columns("ProcCode").Cells(oRow).Specific.VALUE) & "'"
                oRecordSet01.DoQuery sQry
            
                oMat01.Columns("ProcName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            End If
    End Select

'    Set oRecordset = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
FlushToItemValue_Error:
    oForm01.Freeze False
    MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Function updateforuse_date() As Boolean '일괄 업데이트를 통해서 사용처, 납품일자를 수정함.
On Error GoTo updateforuse_date_Error
    Dim i As Integer
    Dim oRecordSet01    As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    oMat01.FlushToDataSource
    
    For i = 0 To oMat01.VisualRowCount - 3

        oDS_PS_MM005H.setValue "U_UseDept", i + 1, Trim(oDS_PS_MM005H.GetValue("U_UseDept", 0))  '행만큼 사용처 수정
        oDS_PS_MM005H.setValue "U_DueDate", i + 1, Trim(oDS_PS_MM005H.GetValue("U_DueDate", 0))  '행만큼 납품일자 수정

    Next i
    
    oMat01.LoadFromDataSource
    
    updateforuse_date = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
updateforuse_date_Error:
    oForm01.Freeze False
    MDC_Com.MDC_GF_Message "updateforuse_date_Error:" & Err.Number & " - " & Err.Description, "E"

End Function
Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error

    Dim ErrNum As Integer
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    ErrNum = 0

    '// Check
    Select Case True
        Case Trim(oForm01.Items("OrdType").Specific.VALUE) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("BPLId").Specific.VALUE) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
    End Select
    
    If oForm01_Mode = fm_ADD_MODE Then
        If Trim(oForm01.Items("CntcCode").Specific.VALUE) = "" Then
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
        ElseIf Trim(oForm01.Items("DeptCode").Specific.VALUE) = "" Then
            ErrNum = 4
            GoTo HeaderSpaceLineDel_Error
        End If
    End If

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "품목구분은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "사업장은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "청구인은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "청구부서는 필수사항입니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function

Private Function MatrixSpaceLineDel() As Boolean
On Error GoTo MatrixSpaceLineDel_Error

    Dim i&, j&
    Dim ErrNum          As Integer
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim Today           As String
    Dim ItemCode        As String
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    ErrNum = 0
    
    oMat01.FlushToDataSource

    '// 라인
    If oMat01.VisualRowCount = 0 Then
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    End If
    
    If oMat01.VisualRowCount = 1 And Trim(oDS_PS_MM005H.GetValue("U_ItemCode", 0)) = "" And Trim(oDS_PS_MM005H.GetValue("U_ItemName", 0)) = "" Then
        j = 0
    Else
        j = oMat01.VisualRowCount
    End If
    
    '//시스템 일자 '2015/07/01 부터 부자재 통합코드사용을 위해
'    sQry = " Select convert(char(8),GetDate(),112) "
'    oRecordSet01.DoQuery sQry
'    Today = Trim(oRecordSet01.Fields(0).VALUE)
            
    For i = 0 To j - 2
        Select Case True
            Case oDS_PS_MM005H.GetValue("U_ItemCode", i) = "" And Trim(oForm01.Items("OrdType").Specific.VALUE) <> "60"
                ErrNum = 2
                GoTo MatrixSpaceLineDel_Error
            Case oDS_PS_MM005H.GetValue("U_Weight", i) = 0
                ErrNum = 3
                GoTo MatrixSpaceLineDel_Error
            Case oDS_PS_MM005H.GetValue("U_DocDate", i) = ""
                ErrNum = 4
                GoTo MatrixSpaceLineDel_Error
            Case oDS_PS_MM005H.GetValue("U_DueDate", i) = ""
                ErrNum = 5
                GoTo MatrixSpaceLineDel_Error
            '외주제작, 외주가공 품의인 경우 외주사유코드 필수 입력
            Case (Trim(oForm01.Items("OrdType").Specific.VALUE) = "30" Or Trim(oForm01.Items("OrdType").Specific.VALUE) = "40") And oDS_PS_MM005H.GetValue("U_OutCode", i) = ""
                ErrNum = 8
                GoTo MatrixSpaceLineDel_Error
            '외주제작, 외주가공 품의인 경우 외주사유내용 필수 입력
            Case (Trim(oForm01.Items("OrdType").Specific.VALUE) = "30" Or Trim(oForm01.Items("OrdType").Specific.VALUE) = "40") And oDS_PS_MM005H.GetValue("U_OutNote", i) = ""
                ErrNum = 9
                GoTo MatrixSpaceLineDel_Error
            Case oDS_PS_MM005H.GetValue("U_DocDate", i) >= oDS_PS_MM005H.GetValue("U_DueDate", i)
                ErrNum = 10
                GoTo MatrixSpaceLineDel_Error
        End Select
        
        '마감상태 체크_S(2017.11.22 송명규 추가)
        If MDC_PS_Common.Check_Finish_Status(Trim(oForm01.Items("BPLId").Specific.VALUE), oDS_PS_MM005H.GetValue("U_DocDate", i), oForm01.TypeEx) = False Then
            ErrNum = 12
            GoTo MatrixSpaceLineDel_Error
        End If
        '마감상태 체크_E(2017.11.22 송명규 추가)
        
        '//자체로 생성해야하는 품목(연료, 자체수리등)은 자체코드생성해서 사용해야함 '2015/07/07 이영진과장 요청
'        If Trim(oForm01.Items("OrdType").Specific.VALUE) = "20" Then '//부재료일 경우
'
'
'            ItemCode = Trim(oDS_PS_MM005H.GetValue("U_ItemCode", i))
'
'            'If Today >= "20150701" Then '//2015년 7월1일 이후로 부자재 11자리가 아니면 청구안됨.
'            If oDS_PS_MM005H.GetValue("U_DocDate", i) >= "20150701" Then
'                If Len(ItemCode) <> 11 Then
'                    ErrNum = 11
'                    GoTo MatrixSpaceLineDel_Error
'                End If
'            End If
'
'        End If
'            If Trim(oDS_PS_MM005H.GetValue("U_OutItmCd", i)) = "" Then
'                ErrNum = 7
'                GoTo MatrixSpaceLineDel_Error
'            End If
'        End If
        
        If oMenuUID = "PS_MM005_1" Then
            If Trim(oDS_PS_MM005H.GetValue("U_OKYN", i)) = "Y" Then
                If Trim(oDS_PS_MM005H.GetValue("U_OKDate", i)) = "" Then
                    ErrNum = 6
                    GoTo MatrixSpaceLineDel_Error
                End If
            End If
        End If
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
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 품목코드가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 수량/중량이 0보다 작거나 같을 수 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 청구일자가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 납기일자가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 결재일자가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 7 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 외주품명이 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 8 Then
        MDC_Com.MDC_GF_Message "외주제작품의, 가공비품의인경우는 외주사유가 필수입니다. " & i + 1 & "번 라인의 [외주사유코드]가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 9 Then
        MDC_Com.MDC_GF_Message "외주제작품의, 가공비품의인경우는 외주사유가 필수입니다. " & i + 1 & "번 라인의 [외주사유내용]이 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 10 Then
        MDC_Com.MDC_GF_Message "납기일자가 청구일자와 같거나 이전일 입니다. " & i + 1 & "번 라인의 납기일을 확인하세요.", "E"
    ElseIf ErrNum = 11 Then
        MDC_Com.MDC_GF_Message "부자재 통합코드로 청구가 가능합니다. " & i + 1 & "번 라인의 품목코드를 확인하세요.", "E"
    ElseIf ErrNum = 12 Then
        MDC_Com.MDC_GF_Message "마감상태가 잠금입니다. 해당 일자로 등록할 수 없습니다. " & i + 1 & "번 라인의 청구일자를 확인하고, 회계부서로 문의하세요.", "E"
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
    If oForm01_Mode = fm_ADD_MODE Then
        oForm01.Items("Btn01").Specific.Caption = "추가"
    ElseIf oForm01_Mode = fm_OK_MODE Then
        oForm01.Items("Btn01").Specific.Caption = "확인"
    ElseIf oForm01_Mode = fm_UPDATE_MODE Then
        oForm01.Items("Btn01").Specific.Caption = "갱신"
    End If

    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadCaption_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Function Add_PurchaseDemand(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
On Error GoTo Add_PurchaseDemand_Error

    Dim i       As Integer, j&, K&
    Dim sQry    As String
    Dim Query02 As String
    
    Dim ErrOrdNum As String '선행프로세스보다 일자가 빨라서 저장되지 않는 작번을 저장
    Dim ErrNum    As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim RecordSet02 As SAPbobsCOM.Recordset
    Set RecordSet02 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry$, ItemCode$, ItemName$, FrgnName$, Qty&, Weight As Currency, Price As Currency, LinTotal As Currency, BPLID$
    Dim DocDate$, DueDate$, CntcCode$, CntcName$, DeptCode$, UseDept$, Auto$, QCYN$, ReType$, Note$, PP030DL$
    Dim IvQty&, IvWeight As Currency, OkYN$, OkDate$, OrdType$, ProcCode$, ProcName$, Comments$, PP030HNo$, PP030LNo$
    Dim OutItmCd$, OutSize$, OutUnit$
    Dim MaxItemCode$
    
    Dim OutCode As String
    Dim OutNote As String
    
    Dim MachCode As String
    Dim MachName As String
    
    Dim BaseEntry   As String
    Dim BaseLine    As String
    Dim DocType     As String
    Dim CurDocDate  As String
    
    oMat01.FlushToDataSource
    
    OrdType = Trim(oForm01.Items("OrdType").Specific.VALUE)
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
    CntcCode = Trim(oForm01.Items("CntcCode").Specific.VALUE)
    CntcName = Trim(oForm01.Items("CntcName").Specific.VALUE)
    DeptCode = Trim(oForm01.Items("DeptCode").Specific.VALUE)
    
    If Trim(oForm01.Items("OrdType").Specific.VALUE) = "60" Then
        sQry = "Select right(rtrim('00000' + Convert(Nvarchar(6), Convert(decimal(7,0), Isnull(Max(right(U_ItemCode,6)),'0')) + 1)),6) "
        sQry = sQry + " From [@PS_MM005H] Where U_BPLId = '" & BPLID & "'"
        sQry = sQry + " and U_OrdType = '" & OrdType & "' and U_ItemCode <> '1' "
        
        RecordSet01.DoQuery sQry
        
        MaxItemCode = "90" & BPLID & Trim(RecordSet01.Fields(0).VALUE)
    End If
    
    For i = 0 To oMat01.RowCount - 2
        DocDate = Trim(oDS_PS_MM005H.GetValue("U_DocDate", i))
        DueDate = Trim(oDS_PS_MM005H.GetValue("U_DueDate", i))
    
'        sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_MM005H] where Left(Convert(char, U_DocDate, 112), 6) = Left('" & DocDate & "', 6)"

'        sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_MM005H] where Left(DocEntry, 6) = Left('" & DocDate & "', 6)"
'        RecordSet01.DoQuery sQry
'        If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
'            DocEntry = Left(DocDate, 6) + "0001"
'        Else
'            DocEntry = Trim(RecordSet01.Fields(0).VALUE) + 1
'        End If
        
        If Trim(oDS_PS_MM005H.GetValue("U_PP030DL", i)) = "" Or Trim(oDS_PS_MM005H.GetValue("U_PP030DL", i)) = Null Then
            PP030DL = ""
            PP030HNo = ""
            PP030LNo = ""
        Else
            PP030DL = Trim(oDS_PS_MM005H.GetValue("U_PP030DL", i))
            For j = 1 To Len(PP030DL)
                If Mid(PP030DL, j, 1) = "-" Then
                    Exit For
                Else
                    PP030HNo = PP030HNo & Mid(PP030DL, j, 1)
                End If
            Next
            For j = 1 To Len(PP030DL)
                If Mid(PP030DL, j, 1) = "-" Then
                    K = j
                    Exit For
                End If
            Next
            PP030LNo = Mid(PP030DL, K + 1, Len(PP030DL) - K)
        End If
        
        
        ItemCode = Trim(oDS_PS_MM005H.GetValue("U_ItemCode", i))
        
        If Trim(oForm01.Items("OrdType").Specific.VALUE) = "60" Then
            '고정자산품의시 품목코드
            If Trim(ItemCode) = "" Then
                ItemCode = "90" & BPLID & Right(RTrim("00000" & LTrim(Str(Val(Right(MaxItemCode, 6)) + i))), 6)
            End If
        End If
        
        ItemName = Make_ItemName(Trim(oDS_PS_MM005H.GetValue("U_ItemName", i)))
        
'        FrgnName = Trim(oDS_PS_MM005H.GetValue("U_FrgnName", i))
        If Trim(oDS_PS_MM005H.GetValue("U_Qty", i)) = "" Then
            Qty = 0
        Else
            Qty = Trim(oDS_PS_MM005H.GetValue("U_Qty", i))
        End If
        Weight = Trim(oDS_PS_MM005H.GetValue("U_Weight", i))
'        Price = Trim(oDS_PS_MM005H.GetValue("U_Price", i))
'        LinTotal = Trim(oDS_PS_MM005H.GetValue("U_LinTotal", i))
        UseDept = Trim(oDS_PS_MM005H.GetValue("U_UseDept", i))
        Auto = Trim(oDS_PS_MM005H.GetValue("U_Auto", i))
        QCYN = Trim(oDS_PS_MM005H.GetValue("U_QCYN", i))
'        ReType = Trim(oDS_PS_MM005H.GetValue("U_ReType", i))
        Note = Trim(oDS_PS_MM005H.GetValue("U_Note", i))
        OutCode = Trim(oDS_PS_MM005H.GetValue("U_OutCode", i))
        OutNote = Trim(oDS_PS_MM005H.GetValue("U_OutNote", i))
'        IvQty = Trim(oDS_PS_MM005H.GetValue("U_IvQty", i))
'        IvWeight = Trim(oDS_PS_MM005H.GetValue("U_IvWeight", i))
        OkYN = Trim(oDS_PS_MM005H.GetValue("U_OKYN", i))
        OkDate = Trim(oDS_PS_MM005H.GetValue("U_OKDate", i))
        ProcCode = Trim(oDS_PS_MM005H.GetValue("U_ProcCode", i))
        ProcName = Trim(oDS_PS_MM005H.GetValue("U_ProcName", i))
        Comments = Trim(oDS_PS_MM005H.GetValue("U_Comments", i))
'        OutItmCd = Trim(oDS_PS_MM005H.GetValue("U_OutItmCd", i))
        OutSize = Trim(oDS_PS_MM005H.GetValue("U_OutSize", i))
        OutUnit = Trim(oDS_PS_MM005H.GetValue("U_OutUnit", i))
        
        MachCode = Trim(oDS_PS_MM005H.GetValue("U_MachCode", i)) '설비코드
        MachName = Trim(oDS_PS_MM005H.GetValue("U_MachName", i)) '설비명
        
        If Trim(ItemCode) <> "" Then
        
            sQry = "            EXEC [PS_MM005_05] "
            sQry = sQry & "'" & PP030DL & "',"
            sQry = sQry & "'" & PP030HNo & "',"
            sQry = sQry & "'" & PP030LNo & "',"
            sQry = sQry & "'" & ItemCode & "',"
            sQry = sQry & "'" & ItemName & "',"
            sQry = sQry & "'" & Qty & "',"
            sQry = sQry & "'" & Weight & "',"
            sQry = sQry & "'" & BPLID & "',"
            sQry = sQry & "'" & DocDate & "',"
            sQry = sQry & "'" & DueDate & "',"
            sQry = sQry & "'" & CntcCode & "',"
            sQry = sQry & "'" & CntcName & "',"
            sQry = sQry & "'" & DeptCode & "',"
            sQry = sQry & "'" & UseDept & "',"
            sQry = sQry & "'" & Auto & "',"
            sQry = sQry & "'" & QCYN & "',"
            sQry = sQry & "'" & Note & "',"
            sQry = sQry & "'" & OutCode & "',"
            sQry = sQry & "'" & OutNote & "',"
            sQry = sQry & "'" & OkYN & "',"
            sQry = sQry & "'" & OkDate & "',"
            sQry = sQry & "'" & OrdType & "',"
            sQry = sQry & "'" & ProcCode & "',"
            sQry = sQry & "'" & ProcName & "',"
            sQry = sQry & "'" & Comments & "',"
            sQry = sQry & "'" & OutSize & "',"
            sQry = sQry & "'" & OutUnit & "',"
            sQry = sQry & "'" & MachCode & "',"
            sQry = sQry & "'" & MachName & "'"
            
            '선행프로세스 대비 일자체크_S
            'Entry = Split(SjDocLin, "-")
            BaseEntry = PP030HNo
            BaseLine = "0"
            DocType = "PS_MM005"
            CurDocDate = DocDate
         
            Query02 = "         EXEC PS_Z_CHECK_DATE '"
            Query02 = Query02 & BaseEntry & "','"
            Query02 = Query02 & BaseLine & "','"
            Query02 = Query02 & DocType & "','"
            Query02 = Query02 & CurDocDate & "'"
        
            Call RecordSet02.DoQuery(Query02)
            
            If RecordSet02.Fields("ReturnValue").VALUE = "True" Then
         
                Call RecordSet01.DoQuery(sQry) '등록
         
            Else
             
                ErrOrdNum = ErrOrdNum & " [" & ItemCode & "]"
             
            End If
            
            PP030HNo = ""
            PP030LNo = ""
        
            If ErrOrdNum <> "" Then '하나라도 선행프로세스 일자가 빠른 작번이 있으면
                ErrNum = "1"
                GoTo Add_PurchaseDemand_Error
            End If
            '선행프로세스 대비 일자체크_E
        
        End If
        
    Next
    
    MDC_Com.MDC_GF_Message "구매요청완료!", "S"
    
    Set RecordSet01 = Nothing
    Add_PurchaseDemand = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_PurchaseDemand_Error:

    Add_PurchaseDemand = False
    Set RecordSet01 = Nothing
    
    If ErrNum = "1" Then
    
        Call Sbo_Application.MessageBox("작번등록일은 생산의뢰접수일과 같거나 늦어야합니다. 확인하십시오." & Chr(13) & ErrOrdNum, 1)
        
    Else
    
        MDC_Com.MDC_GF_Message "Add_PurchaseDemand_Error:" & Err.Number & " - " & Err.Description, "E"

    End If
    
End Function

Public Function Update_PurchaseDemand(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
On Error GoTo Update_PurchaseDemand_Error

    Dim i As Integer, j&, K&
    Dim sQry As String
    Dim Query02 As String
    
    Dim ErrOrdNum As String '선행프로세스보다 일자가 빨라서 저장되지 않는 작번을 저장
    Dim ErrNum    As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim RecordSet02 As SAPbobsCOM.Recordset
    Set RecordSet02 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry$, ItemCode$, ItemName$, Qty&, Weight As Currency, Price As Currency, LinTotal As Currency
    Dim DocDate$, DueDate$, CntcCode$, CntcName$, DeptCode$, UseDept$, Auto$, QCYN$, ReType$, Note$, BPLID$
    Dim IvQty&, IvWeight As Currency, OkYN$, OkDate$, OrdType$, ProcCode$, CgNum$, Status$, Comments$, PP030DL$, PP030HNo$, PP030LNo$
    Dim OutItmCd$, OutSize$, OutUnit$
    
    Dim OutCode As String
    Dim OutNote As String
    
    Dim MachCode As String
    Dim MachName As String
    
    Dim BaseEntry   As String
    Dim BaseLine    As String
    Dim DocType     As String
    Dim CurDocDate  As String
    
    oMat01.FlushToDataSource
    
'    OrdType = Trim(oForm01.Items("OrdType").Specific.Value)
'    BPLId = Trim(oForm01.Items("BPLId").Specific.Value)
'    CntcCode = Trim(oForm01.Items("CntcCode").Specific.Value)
'    CntcName = Trim(oForm01.Items("CntcName").Specific.Value)
'    DeptCode = Trim(oForm01.Items("DeptCode").Specific.Value)
        
    For i = 0 To oMat01.RowCount - 1
        DocDate = Trim(oDS_PS_MM005H.GetValue("U_DocDate", i))
        DueDate = Trim(oDS_PS_MM005H.GetValue("U_DueDate", i))
        CgNum = Trim(oDS_PS_MM005H.GetValue("U_CgNum", i))
        ItemCode = Trim(oDS_PS_MM005H.GetValue("U_ItemCode", i))
        ItemName = Make_ItemName(Trim(oDS_PS_MM005H.GetValue("U_ItemName", i)))
        If Trim(oDS_PS_MM005H.GetValue("U_Qty", i)) = "" Then
            Qty = 0
        Else
            Qty = Trim(oDS_PS_MM005H.GetValue("U_Qty", i))
        End If
        Weight = Trim(oDS_PS_MM005H.GetValue("U_Weight", i))
'        Price = Trim(oDS_PS_MM005H.GetValue("U_Price", i))
'        LinTotal = Trim(oDS_PS_MM005H.GetValue("U_LinTotal", i))
        UseDept = Trim(oDS_PS_MM005H.GetValue("U_UseDept", i))
        Auto = Trim(oDS_PS_MM005H.GetValue("U_Auto", i))
        QCYN = Trim(oDS_PS_MM005H.GetValue("U_QCYN", i))
'        ReType = Trim(oDS_PS_MM005H.GetValue("U_ReType", i))
        Note = Trim(oDS_PS_MM005H.GetValue("U_Note", i))
        OutCode = Trim(oDS_PS_MM005H.GetValue("U_OutCode", i))
        OutNote = Trim(oDS_PS_MM005H.GetValue("U_OutNote", i))
'        IvQty = Trim(oDS_PS_MM005H.GetValue("U_IvQty", i))
'        IvWeight = Trim(oDS_PS_MM005H.GetValue("U_IvWeight", i))
        OkYN = Trim(oDS_PS_MM005H.GetValue("U_OKYN", i))
        OkDate = Trim(oDS_PS_MM005H.GetValue("U_OKDate", i))
        ProcCode = Trim(oDS_PS_MM005H.GetValue("U_ProcCode", i))
        Comments = Trim(oDS_PS_MM005H.GetValue("U_Comments", i))
        Status = Trim(oDS_PS_MM005H.GetValue("U_Status", i))
        
        OrdType = Trim(oDS_PS_MM005H.GetValue("U_OrdType", i))
        BPLID = Trim(oDS_PS_MM005H.GetValue("U_BPLId", i))
        CntcCode = Trim(oDS_PS_MM005H.GetValue("U_CntcCode", i))
        CntcName = Trim(oDS_PS_MM005H.GetValue("U_CntcName", i))
        DeptCode = Trim(oDS_PS_MM005H.GetValue("U_DeptCode", i))
        
'        OutItmCd = Trim(oDS_PS_MM005H.GetValue("U_OutItmCd", i))
        OutSize = Trim(oDS_PS_MM005H.GetValue("U_OutSize", i))
        OutUnit = Trim(oDS_PS_MM005H.GetValue("U_OutUnit", i))
        
        MachCode = Trim(oDS_PS_MM005H.GetValue("U_MachCode", i)) '설비코드
        MachName = Trim(oDS_PS_MM005H.GetValue("U_MachName", i)) '설비명
        
        If Trim(oDS_PS_MM005H.GetValue("U_PP030DL", i)) = "" Or Trim(oDS_PS_MM005H.GetValue("U_PP030DL", i)) = Null Then
            PP030DL = ""
            PP030HNo = ""
            PP030LNo = ""
        Else
            PP030DL = Trim(oDS_PS_MM005H.GetValue("U_PP030DL", i))
            PP030HNo = ""
            PP030LNo = ""
            For j = 1 To Len(PP030DL)
                If Mid(PP030DL, j, 1) = "-" Then
                    Exit For
                Else
                    PP030HNo = PP030HNo & Mid(PP030DL, j, 1)
                End If
            Next
            For j = 1 To Len(PP030DL)
                If Mid(PP030DL, j, 1) = "-" Then
                    K = j
                    Exit For
                End If
            Next
            PP030LNo = Mid(PP030DL, K + 1, Len(PP030DL) - K)
        End If
        
        sQry = "UPDATE [@PS_MM005H] "
        sQry = sQry & "SET "
        
        sQry = sQry & "U_PP030DL = '" & PP030DL & "', "
        sQry = sQry & "U_PP030HNo = '" & PP030HNo & "', "
        sQry = sQry & "U_PP030LNo = '" & PP030LNo & "', "
        sQry = sQry & "U_ItemCode = '" & ItemCode & "', "
        sQry = sQry & "U_ItemName = '" & ItemName & "', "
        sQry = sQry & "U_Qty = '" & Qty & "', "
        sQry = sQry & "U_Weight = '" & Weight & "', "
'        sQry = sQry & "U_Price = '" & Price & "', "
'        sQry = sQry & "U_LinTotal = '" & LinTotal & "', "
        sQry = sQry & "U_CgNum = '" & CgNum & "', "
        sQry = sQry & "U_DocDate = '" & DocDate & "', "
        sQry = sQry & "U_DueDate = '" & DueDate & "', "
        sQry = sQry & "U_CntcCode = '" & CntcCode & "', "
        sQry = sQry & "U_CntcName = '" & CntcName & "', "
        sQry = sQry & "U_DeptCode = '" & DeptCode & "', "
        sQry = sQry & "U_UseDept = '" & UseDept & "', "
        sQry = sQry & "U_Auto = '" & Auto & "', "
        sQry = sQry & "U_QCYN = '" & QCYN & "', "
'        sQry = sQry & "U_ReType = '" & ReType & "', "
        sQry = sQry & "U_Note = '" & Note & "', "
        sQry = sQry & "U_OutCode = '" & OutCode & "', "
        sQry = sQry & "U_OutNote = '" & OutNote & "', "
'        sQry = sQry & "U_IvQty = '" & IvQty & "', "
'        sQry = sQry & "U_IvWeight = '" & IvWeight & "', "
        sQry = sQry & "U_OKYN = '" & OkYN & "', "
        If OkDate = "" Then
            sQry = sQry & "U_OKDate = NULL, "
        Else
            sQry = sQry & "U_OKDate = '" & OkDate & "', "
        End If
        sQry = sQry & "U_OrdType = '" & OrdType & "', "
        sQry = sQry & "U_ProcCode = '" & ProcCode & "', "
        sQry = sQry & "U_Comments = '" & Comments & "', "
'        sQry = sQry & "U_OutItmCd = '" & OutItmCd & "', "
        sQry = sQry & "U_OutSize = '" & OutSize & "', "
        sQry = sQry & "U_OutUnit = '" & OutUnit & "', "
        sQry = sQry & "U_Status = '" & Status & "', "
        sQry = sQry & "U_MachCode = '" & MachCode & "', "
        sQry = sQry & "U_MachName = '" & MachName & "', "
        sQry = sQry & "UpdateDate = GETDATE() "
        sQry = sQry & "Where DocEntry = '" & CgNum & "' "
    
    
        '선행프로세스 대비 일자체크_S
        'Entry = Split(SjDocLin, "-")
        BaseEntry = PP030HNo
        BaseLine = "0"
        DocType = "PS_MM005"
        CurDocDate = DocDate
     
        Query02 = "         EXEC PS_Z_CHECK_DATE '"
        Query02 = Query02 & BaseEntry & "','"
        Query02 = Query02 & BaseLine & "','"
        Query02 = Query02 & DocType & "','"
        Query02 = Query02 & CurDocDate & "'"
    
        Call RecordSet02.DoQuery(Query02)
        '선행프로세스 대비 일자체크_E
        
        If RecordSet02.Fields("ReturnValue").VALUE = "True" Then
     
            Call RecordSet01.DoQuery(sQry) '등록
     
        Else
         
            ErrOrdNum = ErrOrdNum & " [" & ItemCode & "]"
         
        End If
    
        If ErrOrdNum <> "" Then '하나라도 선행프로세스 일자가 빠른 작번이 있으면
            ErrNum = "1"
            GoTo Update_PurchaseDemand_Error
        End If
        
    Next
'    sQry = sQry & " U_DueDate,"

    MDC_Com.MDC_GF_Message "구매요청 갱신 완료!", "S"
    
    Set RecordSet01 = Nothing
    Update_PurchaseDemand = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Update_PurchaseDemand_Error:

    Update_PurchaseDemand = False
    Set RecordSet01 = Nothing
    
    If ErrNum = "1" Then
    
        Call Sbo_Application.MessageBox("작번등록일은 생산의뢰접수일과 같거나 늦어야합니다. 확인하십시오." & Chr(13) & ErrOrdNum, 1)
        
    Else
    
        MDC_Com.MDC_GF_Message "Update_PurchaseDemand_Error:" & Err.Number & " - " & Err.Description, "E"
    
    End If
    
End Function

Public Sub LoadData()
On Error GoTo LoadData_Error
    Dim i As Integer
    Dim sQry As String, SumQty As Currency, SumWeight As Currency
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim oRecordset02 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Set oRecordset02 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim OrdType$, BPLID$, CntcCode$, DeptCode$, CgNumFr$, CgNumTo$, ItemCode$, ItmBsort$, ItmMsort$, ItemType$, OkYN$
    Dim Calculate_Qty As Long
    Dim DocDateFr$, DocDateTo$
    
    
    OrdType = Trim(oForm01.Items("OrdType").Specific.VALUE)
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
    CntcCode = Trim(oForm01.Items("CntcCode").Specific.VALUE)
    DeptCode = Trim(oForm01.Items("DeptCode").Specific.VALUE)
    DocDateFr = Trim(oForm01.Items("DocDateFr").Specific.VALUE)
    DocDateTo = Trim(oForm01.Items("DocDateTo").Specific.VALUE)
    CgNumFr = Trim(oForm01.Items("CgNumFr").Specific.VALUE)
    CgNumTo = Trim(oForm01.Items("CgNumTo").Specific.VALUE)
    ItemCode = Trim(oForm01.Items("ItemCode").Specific.VALUE)
    ItmBsort = Trim(oForm01.Items("ItmBSort").Specific.VALUE)
    ItmMsort = Trim(oForm01.Items("ItmMSort").Specific.VALUE)
    ItemType = Trim(oForm01.Items("ItemType").Specific.VALUE)
    OkYN = Trim(oForm01.Items("OKYN").Specific.VALUE)
    
    If OrdType = "" Then OrdType = "%"
    If BPLID = "" Then BPLID = "%"
    If CntcCode = "" Then CntcCode = "%"
    If DeptCode = "" Then DeptCode = "%"
    'If DocDateFr = "" Then DocDateFr = "19000101"
    If DocDateFr = "" Then DocDateFr = Format(Date - 90, "YYYY-MM") & "-01"
    'If DocDateTo = "" Then DocDateTo = "20991231"
    If DocDateTo = "" Then DocDateTo = Format(Now, "YYYY-MM-DD")
    If CgNumFr = "" Then CgNumFr = "0000000000"
    If CgNumTo = "" Then CgNumTo = "9999999999"
    If ItemCode = "" Then ItemCode = "%"
    If ItmBsort = "" Or ItmBsort = "ALL" Then ItmBsort = "%"
    If ItmMsort = "" Or ItmMsort = "ALL" Then ItmMsort = "%"
    If ItemType = "" Or ItemType = "ALL" Then ItemType = "%"
    If OkYN = "" Or OkYN = "ALL" Then OkYN = "%"
    
    sQry = "EXEC [PS_MM005_01] '" & OrdType & "','" & BPLID & "','" & CntcCode & "','" & DeptCode & "','" & DocDateFr & "',"
    sQry = sQry & "'" & DocDateTo & "','" & CgNumFr & "','" & CgNumTo & "','" & ItemCode & "','" & ItmBsort & "','" & ItmMsort & "','" & ItemType & "','" & OkYN & "'"
    oRecordSet01.DoQuery sQry
    
    oMat01.Clear
    oDS_PS_MM005H.Clear
    
    If (oRecordSet01.RecordCount = 0) Then
        MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
        Set oRecordSet01 = Nothing
        Exit Sub
    End If
    
    oForm01.Freeze True
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_MM005H.Size Then
            oDS_PS_MM005H.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_MM005H.Offset = i
                
        oDS_PS_MM005H.setValue "DocNum", i, i + 1
        oDS_PS_MM005H.setValue "U_PP030DL", i, Trim(oRecordSet01.Fields("U_PP030DL").VALUE)
        oDS_PS_MM005H.setValue "U_ItemCode", i, Trim(oRecordSet01.Fields("U_ItemCode").VALUE)
        oDS_PS_MM005H.setValue "U_ItemName", i, Trim(oRecordSet01.Fields("U_ItemName").VALUE)
'        oDS_PS_MM005H.setValue "U_FrgnName", i, Trim(oRecordSet01.Fields("U_FrgnName").Value)
        oDS_PS_MM005H.setValue "U_BPLId", i, Trim(oRecordSet01.Fields("U_BPLId").VALUE)
        oDS_PS_MM005H.setValue "U_Qty", i, Trim(oRecordSet01.Fields("U_Qty").VALUE)
        oDS_PS_MM005H.setValue "U_Weight", i, Trim(oRecordSet01.Fields("U_Weight").VALUE)
'        oDS_PS_MM005H.setValue "U_Price", i, Trim(oRecordSet01.Fields("U_Price").Value)
'        oDS_PS_MM005H.setValue "U_LinTotal", i, Trim(oRecordSet01.Fields("U_LinTotal").Value)
        oDS_PS_MM005H.setValue "U_CgNum", i, Trim(oRecordSet01.Fields("U_CgNum").VALUE)
        oDS_PS_MM005H.setValue "U_DocDate", i, Format(Trim(oRecordSet01.Fields("U_DocDate").VALUE), "YYYYMMDD")
        oDS_PS_MM005H.setValue "U_DueDate", i, Format(Trim(oRecordSet01.Fields("U_DueDate").VALUE), "YYYYMMDD")
        oDS_PS_MM005H.setValue "U_CntcCode", i, Trim(oRecordSet01.Fields("U_CntcCode").VALUE)
        oDS_PS_MM005H.setValue "U_CntcName", i, Trim(oRecordSet01.Fields("U_CntcName").VALUE)
        oDS_PS_MM005H.setValue "U_DeptCode", i, Trim(oRecordSet01.Fields("U_DeptCode").VALUE)
        oDS_PS_MM005H.setValue "U_UseDept", i, Trim(oRecordSet01.Fields("U_UseDept").VALUE)
        oDS_PS_MM005H.setValue "U_Auto", i, Trim(oRecordSet01.Fields("U_Auto").VALUE)
        oDS_PS_MM005H.setValue "U_QCYN", i, Trim(oRecordSet01.Fields("U_QCYN").VALUE)
'        oDS_PS_MM005H.setValue "U_ReType", i, Trim(oRecordSet01.Fields("U_ReType").Value)f
        oDS_PS_MM005H.setValue "U_Note", i, Trim(oRecordSet01.Fields("U_Note").VALUE)
        oDS_PS_MM005H.setValue "U_OutCode", i, Trim(oRecordSet01.Fields("U_OutCode").VALUE)
        oDS_PS_MM005H.setValue "U_OutNote", i, Trim(oRecordSet01.Fields("U_OutNote").VALUE)
'        oDS_PS_MM005H.setValue "U_OutItmCd", i, Trim(oRecordSet01.Fields("U_OutItmCd").Value)
        oDS_PS_MM005H.setValue "U_OutSize", i, Trim(oRecordSet01.Fields("U_OutSize").VALUE)
        oDS_PS_MM005H.setValue "U_OutUnit", i, Trim(oRecordSet01.Fields("U_OutUnit").VALUE)
        oDS_PS_MM005H.setValue "U_OrdNum", i, Trim(oRecordSet01.Fields("U_OrdNum").VALUE)
        
        sQry = "Select Sum(OnHand) From OITW Where ItemCode = '" & Trim(oRecordSet01.Fields("U_ItemCode").VALUE) & "'"
        oRecordset02.DoQuery sQry
        Calculate_Qty = MDC_PS_Common.Calculate_Qty(Trim(oRecordSet01.Fields("U_ItemCode").VALUE), Trim(oRecordset02.Fields(0).VALUE))
        
        oDS_PS_MM005H.setValue "U_IvQty", i, Calculate_Qty
        oDS_PS_MM005H.setValue "U_IvWeight", i, Trim(oRecordset02.Fields(0).VALUE)
        oDS_PS_MM005H.setValue "U_OKYN", i, Trim(oRecordSet01.Fields("U_OKYN").VALUE)
        
        If Format(Trim(oRecordSet01.Fields("U_OKDate").VALUE), "YYYYMMDD") = "18991230" Then
            oDS_PS_MM005H.setValue "U_OKDate", i, ""
        Else
            oDS_PS_MM005H.setValue "U_OKDate", i, Format(Trim(oRecordSet01.Fields("U_OKDate").VALUE), "YYYYMMDD")
        End If
        
        oDS_PS_MM005H.setValue "U_OrdType", i, Trim(oRecordSet01.Fields("U_OrdType").VALUE)
        oDS_PS_MM005H.setValue "U_ProcCode", i, Trim(oRecordSet01.Fields("U_ProcCode").VALUE)
        oDS_PS_MM005H.setValue "U_ProcName", i, Trim(oRecordSet01.Fields("U_ProcName").VALUE)
        oDS_PS_MM005H.setValue "U_Comments", i, Trim(oRecordSet01.Fields("U_Comments").VALUE)
        oDS_PS_MM005H.setValue "U_Status", i, Trim(oRecordSet01.Fields("U_Status").VALUE)
        
        SumQty = SumQty + Trim(oRecordSet01.Fields("U_Qty").VALUE)
        SumWeight = SumWeight + Trim(oRecordSet01.Fields("U_Weight").VALUE)
                
        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
    Next
    
    oForm01.Items("SumQty").Specific.VALUE = SumQty
    oForm01.Items("SumWeight").Specific.VALUE = SumWeight
                
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    ProgBar01.Stop
    oForm01.Freeze False
  
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadData_Error:
    oForm01.Freeze False
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "LoadData_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub
