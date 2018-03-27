VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP215"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_PP215.cls
'//  Module         : PP
'//  Description    : 원재료 월소요량 등록
'//  FormType       : PS_PP215
'//  Create Date    : 2013.11.12
'//  Modified Date  :
'//  Creator        : N.G.Y
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01             As SAPbouiCOM.Form
Public oMat01             As SAPbouiCOM.Matrix
Private oDS_PS_PP215H As SAPbouiCOM.DBDataSource    '등록헤더
Private oDS_PS_PP215L As SAPbouiCOM.DBDataSource    '등록라인

Private oLast_Item_UID      As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID       As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row       As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&

Private oYYYYMM$

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP215.srf")
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
    
    oFormUniqueID01 = "PS_PP215_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
  
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    '화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
    oForm01.DataBrowser.BrowseBy = "Code"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting
    Call Initialization
    Call FormItemEnabled
    
    
    
    oForm01.EnableMenu ("1283"), True         '// 삭제
    oForm01.EnableMenu ("1287"), True         '// 복제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1284"), False        '// 취소
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
    Dim i&
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    Dim ItemCode        As String
    Dim rate As Currency
    Dim StdWgt As Currency
    
    Dim S_StdWgt As Double
    Dim S_Weight1 As Double
    Dim S_Weight2 As Double
    Dim S_PWeight1 As Double
    Dim S_PWeight2 As Double
                            
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
                        
                        Dim Code$, BPLId$, ym$, ItmBsort$
                        BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
                        ym = Trim(Right(oForm01.Items("YM").Specific.VALUE, 4))
                        ItmBsort = Trim(oForm01.Items("ItmBsort").Specific.VALUE)
                        
                        Code = ym & BPLId & ItmBsort
                        Call oDS_PS_PP215H.setValue("Code", 0, Code)
                    End If
                End If
'et_KEY_DOWN ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "ItmBsort" Then
                        If oForm01.Items("ItmBsort").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "ItemCode" Then
                            If oMat01.Columns("ItemCode").Cells(pval.Row).Specific.VALUE = "" Then
                                Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                            End If
                        End If
                        If pval.ColUID = "PItemCod" Then
                            If oMat01.Columns("PItemCod").Cells(pval.Row).Specific.VALUE = "" Then
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
                End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_MATRIX_LINK_PRESSED '//8
'et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "ItmBsort" Then
                        oForm01.Freeze True
                        
                        sQry = "Select Name From [@PSH_ITMBSORT] Where Code = '" & Trim(oForm01.Items("ItmBsort").Specific.VALUE) & "'"
                        oRecordSet01.DoQuery sQry
                        oForm01.Items("ItmBname").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
'                        Call LoadData
                        oForm01.Freeze False
                    ElseIf pval.ItemUID = "Mat01" Then
                        oMat01.FlushToDataSource
                        If pval.ColUID = "StdWgt" Then
                            oForm01.Freeze True
                            rate = oForm01.Items("Rate").Specific.VALUE
                            If rate <> 0 Then
                                '기준수량변경시 초기화
                                Call oDS_PS_PP215L.setValue("U_" & pval.ColUID, pval.Row - 1, Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                                Call oDS_PS_PP215L.setValue("U_Weight1", pval.Row - 1, oMat01.Columns("StdWgt").Cells(pval.Row).Specific.VALUE) '창원생산량
                                Call oDS_PS_PP215L.setValue("U_Weight2", pval.Row - 1, 0) '부산생산량
                                Call oDS_PS_PP215L.setValue("U_PWeight1", pval.Row - 1, Round(oMat01.Columns("StdWgt").Cells(pval.Row).Specific.VALUE / rate, 0)) '창원소요량
                                Call oDS_PS_PP215L.setValue("U_PWeight2", pval.Row - 1, 0) '부산소요량
                                
    '                            oMat01.Columns("PWeight1").Cells(pval.Row).Specific.VALUE = Round(oMat01.Columns("Weight1").Cells(pval.Row).Specific.VALUE / rate, 0)
    '                            oMat01.Columns("PWeight2").Cells(pval.Row).Specific.VALUE = 0
                                
                            Else
                                Call oDS_PS_PP215L.setValue("U_" & pval.ColUID, pval.Row - 1, 0)
                                MDC_Com.MDC_GF_Message "기준수율을 입력해야합니다.", "E"
                            End If
                            
                            oMat01.LoadFromDataSource
                            
                            For i = 0 To oMat01.VisualRowCount - 1
                                S_StdWgt = S_StdWgt + oMat01.Columns("StdWgt").Cells(i + 1).Specific.VALUE
                                S_Weight1 = S_Weight1 + oMat01.Columns("Weight1").Cells(i + 1).Specific.VALUE
                                S_Weight2 = S_Weight2 + oMat01.Columns("Weight2").Cells(i + 1).Specific.VALUE
                                S_PWeight1 = S_PWeight1 + oMat01.Columns("PWeight1").Cells(i + 1).Specific.VALUE
                                S_PWeight2 = S_PWeight2 + oMat01.Columns("PWeight2").Cells(i + 1).Specific.VALUE
                            Next i
                            
                            oForm01.Items("S_StdWgt").Specific.VALUE = S_StdWgt
                            oForm01.Items("S_Weight1").Specific.VALUE = S_Weight1
                            oForm01.Items("S_Weight2").Specific.VALUE = S_Weight2
                            oForm01.Items("S_PWeight1").Specific.VALUE = S_PWeight1
                            oForm01.Items("S_PWeight2").Specific.VALUE = S_PWeight2
                            
                            oForm01.Freeze False
                        ElseIf pval.ColUID = "Weight1" Then
                            '창원생산중량 입력
                            
                            StdWgt = oMat01.Columns("StdWgt").Cells(pval.Row).Specific.VALUE
                            rate = oForm01.Items("Rate").Specific.VALUE
                            oForm01.Freeze True
                            If rate <> 0 Then
                                If StdWgt > 0 Then
                                    '창원원재료 소요량
                                    Call oDS_PS_PP215L.setValue("U_" & pval.ColUID, pval.Row - 1, Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                                    Call oDS_PS_PP215L.setValue("U_PWeight1", pval.Row - 1, Round(oMat01.Columns("Weight1").Cells(pval.Row).Specific.VALUE / rate, 0)) '창원소요량
                                    
                                    '기계중량은 기준중량 - 창원생산량
                                    Call oDS_PS_PP215L.setValue("U_Weight2", pval.Row - 1, StdWgt - oMat01.Columns("Weight1").Cells(pval.Row).Specific.VALUE) '부산생산량
                                    Call oDS_PS_PP215L.setValue("U_PWeight2", pval.Row - 1, Round((StdWgt - oMat01.Columns("Weight1").Cells(pval.Row).Specific.VALUE) / rate, 0)) '부산소요량
                                    
                                Else
                                    Call oDS_PS_PP215L.setValue("U_Weight2", pval.Row - 1, 0) '부산생산량
                                    Call oDS_PS_PP215L.setValue("U_PWeight1", pval.Row - 1, 0) '창원소요량
                                    Call oDS_PS_PP215L.setValue("U_PWeight2", pval.Row - 1, 0) '부산소요량
                                    
                                End If
                                oMat01.LoadFromDataSource
                            Else
                                Call oDS_PS_PP215L.setValue("U_" & pval.ColUID, pval.Row - 1, 0)
                                MDC_Com.MDC_GF_Message "기준수율을 입력해야합니다.", "E"
                            End If
                            
                            For i = 0 To oMat01.VisualRowCount - 1
                                S_StdWgt = S_StdWgt + oMat01.Columns("StdWgt").Cells(i + 1).Specific.VALUE
                                S_Weight1 = S_Weight1 + oMat01.Columns("Weight1").Cells(i + 1).Specific.VALUE
                                S_Weight2 = S_Weight2 + oMat01.Columns("Weight2").Cells(i + 1).Specific.VALUE
                                S_PWeight1 = S_PWeight1 + oMat01.Columns("PWeight1").Cells(i + 1).Specific.VALUE
                                S_PWeight2 = S_PWeight2 + oMat01.Columns("PWeight2").Cells(i + 1).Specific.VALUE
                            Next i
                            
                            oForm01.Items("S_StdWgt").Specific.VALUE = S_StdWgt
                            oForm01.Items("S_Weight1").Specific.VALUE = S_Weight1
                            oForm01.Items("S_Weight2").Specific.VALUE = S_Weight2
                            oForm01.Items("S_PWeight1").Specific.VALUE = S_PWeight1
                            oForm01.Items("S_PWeight2").Specific.VALUE = S_PWeight2
                            
                            oMat01.Columns(pval.ColUID).Cells(pval.Row).Click
                           
                            oForm01.Freeze False
                        ElseIf pval.ColUID = "Weight2" Then
                            '부산생산중량 입력
                            
                            StdWgt = oMat01.Columns("StdWgt").Cells(pval.Row).Specific.VALUE
                            rate = oForm01.Items("Rate").Specific.VALUE
                            oForm01.Freeze True
                            If rate <> 0 Then
                                If StdWgt > 0 Then
                                    
                                    Call oDS_PS_PP215L.setValue("U_PWeight2", pval.Row - 1, Round(oMat01.Columns("Weight2").Cells(pval.Row).Specific.VALUE / rate, 0)) '부산소요량
                                    
                                    '창원중량은 기준중량 - 부산생산량
                                    Call oDS_PS_PP215L.setValue("U_Weight1", pval.Row - 1, StdWgt - oMat01.Columns("Weight2").Cells(pval.Row).Specific.VALUE)
                                    Call oDS_PS_PP215L.setValue("U_PWeight1", pval.Row - 1, Round((StdWgt - oMat01.Columns("Weight2").Cells(pval.Row).Specific.VALUE) / rate, 0))
                                    
                                Else
                                    Call oDS_PS_PP215L.setValue("U_Weight1", pval.Row - 1, 0) '창원생산량
                                    Call oDS_PS_PP215L.setValue("U_PWeight1", pval.Row - 1, 0) '창원소요량
                                    Call oDS_PS_PP215L.setValue("U_PWeight2", pval.Row - 1, 0) '부산소요량
                                    
                                    
                                End If
                                oMat01.LoadFromDataSource
                            Else
                                Call oDS_PS_PP215L.setValue("U_" & pval.ColUID, pval.Row - 1, 0)
                                MDC_Com.MDC_GF_Message "기준수율을 입력해야합니다.", "E"
                            End If
                            oMat01.Columns(pval.ColUID).Cells(pval.Row).Click
                            
                            For i = 0 To oMat01.VisualRowCount - 1
                                S_StdWgt = S_StdWgt + oMat01.Columns("StdWgt").Cells(i + 1).Specific.VALUE
                                S_Weight1 = S_Weight1 + oMat01.Columns("Weight1").Cells(i + 1).Specific.VALUE
                                S_Weight2 = S_Weight2 + oMat01.Columns("Weight2").Cells(i + 1).Specific.VALUE
                                S_PWeight1 = S_PWeight1 + oMat01.Columns("PWeight1").Cells(i + 1).Specific.VALUE
                                S_PWeight2 = S_PWeight2 + oMat01.Columns("PWeight2").Cells(i + 1).Specific.VALUE
                            Next i
                            
                            oForm01.Items("S_StdWgt").Specific.VALUE = S_StdWgt
                            oForm01.Items("S_Weight1").Specific.VALUE = S_Weight1
                            oForm01.Items("S_Weight2").Specific.VALUE = S_Weight2
                            oForm01.Items("S_PWeight1").Specific.VALUE = S_PWeight1
                            oForm01.Items("S_PWeight2").Specific.VALUE = S_PWeight2
                            
                            oForm01.Freeze False
                        ElseIf pval.ColUID = "ItemCode" Then
                            If (pval.Row = oMat01.RowCount Or oMat01.VisualRowCount = 0) And Trim(oMat01.Columns("ItemCode").Cells(pval.Row).Specific.VALUE) <> "" Then
                                oMat01.FlushToDataSource
                                Call Add_MatrixRow(oMat01.RowCount, False)
                                oMat01.Columns("ItemCode").Cells(pval.Row).Click ct_Regular
                            End If
                            
                            ItemCode = Trim(oDS_PS_PP215L.GetValue("U_ItemCode", pval.Row - 1))
                            sQry = "Select a.U_ItemCod1, a.U_ItemNam1, b.U_OutSize, b.U_CallSize, a.U_ItemCod2, a.U_ItemNam2"
                            sQry = sQry & " From [@PS_PP005H] a Inner Join OITM b On a.U_ItemCod1 = b.ItemCode "
                            sQry = sQry & " Where a.U_ItemCod1 = '" & ItemCode & "'"
                            oRecordSet01.DoQuery sQry
                            
                                
                            oForm01.Freeze True
                            oMat01.FlushToDataSource
                            
                            If oRecordSet01.RecordCount = 0 Then
                            
                                '매트릭스에 데이터를 직접 바인딩하면 이벤트가 실행되기 때문에 DataSource로 바인딩하는 방식으로 수정(2011.11.22 송명규)
                                Call oDS_PS_PP215L.setValue("U_ItemCode", pval.Row - 1, "")
                                Call oDS_PS_PP215L.setValue("U_ItemName", pval.Row - 1, "")
                                Call oDS_PS_PP215L.setValue("U_OutSize", pval.Row - 1, "")
                                Call oDS_PS_PP215L.setValue("U_CallSize", pval.Row - 1, "")
                                Call oDS_PS_PP215L.setValue("U_PItemCod", pval.Row - 1, "")
                                Call oDS_PS_PP215L.setValue("U_PItemNam", pval.Row - 1, "")
                                Call oDS_PS_PP215L.setValue("U_StdWgt", pval.Row - 1, 0)
                                Call oDS_PS_PP215L.setValue("U_Weight1", pval.Row - 1, 0)
                                Call oDS_PS_PP215L.setValue("U_Weight2", pval.Row - 1, 0)
                                Call oDS_PS_PP215L.setValue("U_PWeight1", pval.Row - 1, 0)
                                Call oDS_PS_PP215L.setValue("U_PWeight2", pval.Row - 1, 0)
                                
                                
                                MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
                                
                                Set oRecordSet01 = Nothing
                                oMat01.LoadFromDataSource
                                oForm01.Freeze False
                                Exit Sub
                            
                                
                            End If
                            
                            
                            '매트릭스에 데이터를 직접 바인딩하면 이벤트가 실행되기 때문에 DataSource로 바인딩하는 방식으로 수정(2011.11.22 송명규)
                            Call oDS_PS_PP215L.setValue("U_ItemCode", pval.Row - 1, Trim(oRecordSet01.Fields("U_ItemCod1").VALUE)) '품목코드
                            Call oDS_PS_PP215L.setValue("U_ItemName", pval.Row - 1, Trim(oRecordSet01.Fields("U_ItemNam1").VALUE)) '품목이름
                            Call oDS_PS_PP215L.setValue("U_OutSize", pval.Row - 1, Trim(oRecordSet01.Fields("U_OutSize").VALUE)) '외부규격
                            Call oDS_PS_PP215L.setValue("U_CallSize", pval.Row - 1, Trim(oRecordSet01.Fields("U_CallSize").VALUE)) '호칭규격
                            Call oDS_PS_PP215L.setValue("U_PItemCod", pval.Row - 1, Trim(oRecordSet01.Fields("U_ItemCod2").VALUE)) '원재료코드
                            Call oDS_PS_PP215L.setValue("U_PItemNam", pval.Row - 1, Trim(oRecordSet01.Fields("U_ItemNam2").VALUE)) '원재료명
                            
                            oMat01.LoadFromDataSource
                            oForm01.Freeze False
                        ElseIf pval.ColUID = "PItemCod" Then
                            ItemCode = Trim(oDS_PS_PP215L.GetValue("U_PItemCod", pval.Row - 1))
                            sQry = "Select ItemName"
                            sQry = sQry & " From OITM "
                            sQry = sQry & " Where ItemCode = '" & ItemCode & "'"
                            oRecordSet01.DoQuery sQry
                            oForm01.Freeze True
                            oMat01.FlushToDataSource
                        
                            If oRecordSet01.RecordCount = 0 Then
                                Call oDS_PS_PP215L.setValue("U_PItemNam", pval.Row - 1, "")  '원재료명
                            Else
                                Call oDS_PS_PP215L.setValue("U_PItemNam", pval.Row - 1, Trim(oRecordSet01.Fields("ItemName").VALUE)) '원재료명
                            End If
                            
                            
                            oMat01.LoadFromDataSource
                            oForm01.Freeze False
                                
                        End If
                    End If
                End If
            Case et_MATRIX_LOAD: '//11
                For i = 0 To oMat01.VisualRowCount - 1
                    S_StdWgt = S_StdWgt + oMat01.Columns("StdWgt").Cells(i + 1).Specific.VALUE
                    S_Weight1 = S_Weight1 + oMat01.Columns("Weight1").Cells(i + 1).Specific.VALUE
                    S_Weight2 = S_Weight2 + oMat01.Columns("Weight2").Cells(i + 1).Specific.VALUE
                    S_PWeight1 = S_PWeight1 + oMat01.Columns("PWeight1").Cells(i + 1).Specific.VALUE
                    S_PWeight2 = S_PWeight2 + oMat01.Columns("PWeight2").Cells(i + 1).Specific.VALUE
                    
                Next i
                
                oForm01.Items("S_StdWgt").Specific.VALUE = S_StdWgt
                oForm01.Items("S_Weight1").Specific.VALUE = S_Weight1
                oForm01.Items("S_Weight2").Specific.VALUE = S_Weight2
                oForm01.Items("S_PWeight1").Specific.VALUE = S_PWeight1
                oForm01.Items("S_PWeight2").Specific.VALUE = S_PWeight2
                
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
                Set oDS_PS_PP215H = Nothing
                Set oDS_PS_PP215L = Nothing
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
            Case "1286": '닫기
            Case "1293": '행삭제
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
                    For i = 0 To oMat01.VisualRowCount - 1
                        oMat01.Columns("LineNum").Cells(i + 1).Specific.VALUE = i + 1
                    Next i
                    
                    oMat01.FlushToDataSource
                    oDS_PS_PP215L.RemoveRecord oDS_PS_PP215L.Size - 1       '// Mat01에 마지막라인(빈라인) 삭제
                    oMat01.Clear
                    oMat01.LoadFromDataSource
                End If
            Case "1281": '찾기
                oForm01.Freeze True
                Call FormItemEnabled
                Call Initialization
'                oForm01.Items("CycleCod").Click ct_Regular
                oForm01.Freeze False
            Case "1282": '추가
                oForm01.Freeze True
                Call FormItemEnabled
                Call Initialization
                oForm01.Freeze False
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                oForm01.Freeze True
                Call FormItemEnabled
'                If oMat01.VisualRowCount > 0 Then
'                    If oMat01.Columns("CycleCod").Cells(oMat01.VisualRowCount).Specific.Value <> "" Then
'                        Add_MatrixRow oMat01.RowCount, False
'                    End If
'                End If
                oForm01.Freeze False
            Case "1287": '// 복제
                oForm01.Freeze True
                oDS_PS_PP215H.setValue "Code", 0, ""
                
                For i = 0 To oMat01.VisualRowCount - 1
                    oMat01.FlushToDataSource
                    oDS_PS_PP215L.setValue "Code", i, ""
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
    Set oDS_PS_PP215H = oForm01.DataSources.DBDataSources("@PS_PP215H")
    Set oDS_PS_PP215L = oForm01.DataSources.DBDataSources("@PS_PP215L")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    
    Call oForm01.DataSources.UserDataSources.Add("S_Weight1", dt_SUM)
    oForm01.Items("S_Weight1").Specific.DataBind.SetBound True, "", "S_Weight1"
    
    Call oForm01.DataSources.UserDataSources.Add("S_Weight2", dt_SUM)
    oForm01.Items("S_Weight2").Specific.DataBind.SetBound True, "", "S_Weight2"
    
    Call oForm01.DataSources.UserDataSources.Add("S_PWeight1", dt_SUM)
    oForm01.Items("S_PWeight1").Specific.DataBind.SetBound True, "", "S_PWeight1"
    
    Call oForm01.DataSources.UserDataSources.Add("S_PWeight2", dt_SUM)
    oForm01.Items("S_PWeight2").Specific.DataBind.SetBound True, "", "S_PWeight2"
    
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
    
    '매트릭스 거래처
    sQry = "SELECT T1.CardCode, T1.CardName "
    sQry = sQry + " FROM OCRD T1"
    sQry = sQry + " WHERE T1.CardType = 'C' And T1.frozenFor <> 'Y'"
    sQry = sQry + " and exists ( Select * From [@PS_SY001L] a "
    sQry = sQry + " Where a.Code= 'S005' and a.U_UseYN = 'Y'"
    sQry = sQry + " and a.U_Minor = T1.CardCode)"
    sQry = sQry + " Order by T1.CardFName"
    
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oMat01.Columns("CardCode").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
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
    
    '//아이디별 사업장 세팅
    Set oCombo = oForm01.Items("BPLId").Specific
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
    oForm01.Items("YM").Specific.VALUE = Left(Format(Now, "YYYYMMDD"), 6)
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
    '//ChooseFromList 설정
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

'Sub Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
'On Error GoTo Add_MatrixRow_Error
'    If RowIserted = False Then '//행추가여부
'        oDS_PS_PP215L.InsertRecord (oRow)
'    End If
'    oMat01.AddRow
'    oDS_PS_PP215L.Offset = oRow
'    oDS_PS_PP215L.setValue "U_LineNum", oRow, oRow + 1
'    oMat01.LoadFromDataSource
'    Exit Sub
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Add_MatrixRow_Error:
'    MDC_Com.MDC_GF_Message "Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
'End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum          As Integer
    Dim DocNum          As String

    ErrNum = 0

    '// Check
    Select Case True
        Case oDS_PS_PP215H.GetValue("U_ItmBsort", 0) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_PP215H.GetValue("U_BPLId", 0) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_PP215H.GetValue("U_YM", 0) = ""
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
        Case Len(Trim(oForm01.Items("YM").Specific.VALUE)) <> 6
            ErrNum = 4
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_PP215H.GetValue("U_Rate", 0) = ""
            ErrNum = 5
            GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "대분류는 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "사업장은 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "년월은 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "년월은 6자리 YYYYMM 형식으로 입력해야합니다. 확인하세요.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "기준수율은 필수입력사항입니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function

Public Sub LoadData()
On Error GoTo LoadData_Error
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim i As Integer
    Dim ym As String
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sQry = "Select Max(U_YM) From [@PS_PP215H] Where U_YM < '" & Trim(oForm01.Items("YM").Specific.VALUE) & "'"
    sQry = sQry & " And U_ItmBsort = '" & Trim(oForm01.Items("ItmBsort").Specific.VALUE) & "'"
    oRecordSet01.DoQuery sQry
    
    ym = Trim(oRecordSet01.Fields(0).VALUE)
    
    oForm01.Freeze True
    sQry = "Select a.ItemCode, a.ItemName, a.U_ItmMsort, ItmMname = (select U_CodeName from [@PSH_ITMMSORT] Where U_Code = a.U_ItmMsort) "
    sQry = sQry & " From [OITM] a "
    sQry = sQry & "Where a.U_ItmBsort = '" & Trim(oForm01.Items("ItmBsort").Specific.VALUE) & "' Order By a.U_ItmMsort, a.ItemCode"

    oRecordSet01.DoQuery sQry
    
    
    oDS_PS_PP215L.Clear
    oMat01.Clear
    oMat01.FlushToDataSource
    
    i = 0
    Do Until oRecordSet01.EOF
        oDS_PS_PP215L.InsertRecord i
        oDS_PS_PP215L.setValue "U_LineNum", i, i + 1
        oDS_PS_PP215L.setValue "U_ItemCode", i, Trim(oRecordSet01.Fields(0).VALUE)
        oDS_PS_PP215L.setValue "U_ItemName", i, Trim(oRecordSet01.Fields(1).VALUE)
        oDS_PS_PP215L.setValue "U_ItmMsort", i, Trim(oRecordSet01.Fields(2).VALUE)
        oDS_PS_PP215L.setValue "U_ItmMName", i, Trim(oRecordSet01.Fields(3).VALUE)
        i = i + 1
        oRecordSet01.MoveNext
    Loop
    
    oMat01.LoadFromDataSource
    oForm01.Freeze False

    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadData_Error:
    oForm01.Freeze False
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "LoadData_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_PP215L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_PP215L.Offset = oRow
    oDS_PS_PP215L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub
