VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_SD090"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_SD090.cls
'//  Module         : SD
'//  Desc           : 이동처리등록
'//  FormType       : PS_SD090
'//  Create Date    : 2010.08.23
'//  Modified Date  :
'//  Creator        : Dong sub Kim
'//  Copyright  (c) Poongsan Holdings
'//********************************************************************************
Option Explicit

Public oFormUniqueID01      As String
Public oForm01                 As SAPbouiCOM.Form
Public oForm02                 As SAPbouiCOM.Form
Public oMat01                  As SAPbouiCOM.Matrix
Private oDS_PS_SD090H       As SAPbouiCOM.DBDataSource    '등록헤더
Private oDS_PS_SD090L       As SAPbouiCOM.DBDataSource    '등록라인

Private oLast_Item_UID   As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID    As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row    As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

' 입고 DI를 위한 정보를 가지는 구조체
Private Type StockInfos
    CardCode As String            '고객코드
    ItemCode As String            '품목코드
    FromWarehouseCode As String   '창고코드
    ToWarehouseCode As String     '창고코드
    Weight  As Double             '중량
    UnWeight As Double
    BatchNum As String            '배치번호
    BatchWeight As Double         '배치중량
    Qty   As Long                 '수량
    TransNo As String             '재고이전문서번호
    Chk As String
    MatrixRow As Long
    StockTransDocEntry  As String '재고이전문서번호
    StockTransLineNum   As String '재고이전라인번호
    Indate  As String             '전기일
End Type

Private StockInfo() As StockInfos

Private oSeq&

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_SD090.srf")
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
    
    oFormUniqueID01 = "PS_SD090_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
        
    oForm01.Freeze True
    
    Call oForm01.EnableMenu("1293", True)
    Call CreateItems
    
    
    oForm01.EnableMenu ("1293"), True  '행삭제
    oForm01.EnableMenu ("1283"), False '제거
    oForm01.EnableMenu ("1284"), False '취소
    oForm01.EnableMenu ("1285"), False '복원
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    '화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
    oForm01.DataBrowser.BrowseBy = "DocNum"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////D////////////////////////////////////////////////
    
    Call ComboBox_Setting
    Call CF_ChooseFromList
'    Call Initial_Setting
    Call FormItemEnabled       '//폼의 아이템 활성화 컨트롤
  
    Call FormClear             '//UDO방식일때
    Call AddMatrixRow(0, oMat01.RowCount, True) '//UDO방식일때 - 매트릭스 Call Function
    
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    oForm01.Update
    oForm01.Freeze False
    oForm01.Visible = True
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
Dim sQry                As String
Dim sQry02              As String
Dim oRecordSet01        As SAPbobsCOM.Recordset
Dim oRecordset02        As SAPbobsCOM.Recordset
Dim TempForm01 As Variant
Dim ErrNum     As Integer

Dim SumQty As Long
Dim SumWeight As Currency

Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)   '// 객체 정의 및 데이터 할당
Set oRecordset02 = Sbo_Company.GetBusinessObject(BoRecordset)
   
                        
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
        
            Case et_ITEM_PRESSED: '//1
            If pval.ItemUID = "1" Then
                If oForm01.Mode = fm_ADD_MODE Or oForm01.Mode = fm_UPDATE_MODE Then
                   If HeaderSpaceLineDel = False Then
                        BubbleEvent = False             'BubbleEvent = True 이면, 사용자에게 제어권을 넘겨준다. BeforeAction = True일 경우만 쓴다.
                        Exit Sub
                    End If
                    
                    If MatrixSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    
                    If (oForm01.Mode = fm_ADD_MODE) Then    '// 재고 이동 DI API
                        If StockTrans() = True Then
                            Call UpdateUserField
                        Else
                            AddMatrixRow 1, oMat01.RowCount, True
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
                End If
            
            ElseIf pval.ItemUID = "ChulPrin" Then
                Call PS_SD090_Print_Report01
'            ElseIf pval.ItemUID = "GuraPrin" Then
            
            Else
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "Mat01" And pval.ColUID = "ItemCode" Then
                        FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                    End If
                End If
            End If
               
            Case et_KEY_DOWN:   '//2
            
            ' 거래처코드
                If oForm01.Items("CardCode").Specific.VALUE = "" Then
                    If pval.ItemUID = "CardCode" And pval.CharPressed = 9 Then
                    '//CharPressed: The character that was pressed to trigger this event.
                        oForm01.Items("CardCode").Click ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
                
            ' 아이템코드
                If pval.ItemUID = "Mat01" And pval.ColUID = "SD091HNo" And pval.CharPressed = 9 Then
                    If oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.String = "" Then
                        oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
                               
            ' 담당자
                If oForm01.Items("RepName").Specific.VALUE = "" Then
                    If pval.ItemUID = "RepName" And pval.CharPressed = 9 Then
                        oForm01.Items("RepName").Click ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
                
            ' 납품처
                If oForm01.Items("ShipTo").Specific.VALUE = "" Then
                    If pval.ItemUID = "ShipTo" And pval.CharPressed = 9 Then
                        oForm01.Items("ShipTo").Click ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
                
            ' 운송업체
            
            ' 차량번호
            
            ' 도착장소
            
            ' 질별
            
            ' 출고창고
                If pval.ItemUID = "OutWhCd" And pval.CharPressed = 9 Then
                    If oForm01.Items(pval.ColUID).Cells(pval.Row).Specific.String = "" Then
                        oForm01.Items(pval.ColUID).Cells(pval.Row).Click ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                     End If
                End If
                                
            ' 입고창고
                If pval.ItemUID = "InWhCd" And pval.CharPressed = 9 Then
                    If oForm01.Items(pval.ColUID).Cells(pval.Row).Specific.String = "" Then
                        oForm01.Items(pval.ColUID).Cells(pval.Row).Click ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
                
            Case et_COMBO_SELECT: '//5

            Case et_CLICK: '//6
            
                
            Case et_DOUBLE_CLICK: '//7
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "Mat01" And pval.ColUID = "ItemCode" Then
                        FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                    End If
                End If
            
            Case et_MATRIX_LINK_PRESSED '//8
            
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                
                   ' 거래처 이름 Query
                   If pval.ItemUID = "CardCode" Then
                       sQry = "Select CardName From [OCRD] Where CardCode = '" & Trim(oForm01.Items("CardCode").Specific.VALUE) & "'"
                       oRecordSet01.DoQuery sQry
                       oForm01.Items("CardName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                   End If
                   
                   ' 사원 이름 Query
                   If pval.ItemUID = "RepName" Then
                       sQry = "SELECT U_FULLNAME, U_MSTCOD FROM [OHEM] WHERE U_MSTCOD = '" & Trim(oForm01.Items("RepName").Specific.VALUE) & "'"
                       oRecordSet01.DoQuery sQry
                       oForm01.Items("RepNm1").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                   End If
                   
                   ' 납품처 이름 Query
                   If pval.ItemUID = "ShipTo" Then
                       sQry = "Select CardName From [OCRD] Where CardCode = '" & Trim(oForm01.Items("ShipTo").Specific.VALUE) & "'"
                       oRecordSet01.DoQuery sQry
                       oForm01.Items("ShipNm").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                   End If
                                  
                   ' 아이템 코드
                   If pval.ItemUID = "Mat01" And pval.ColUID = "ItemCode" Then
                       FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                   End If
                   
                  ' 출고 창고
                  If pval.ItemUID = "OutWhCd" Then
                      sQry = "Select WhsName From [OWHS] Where WhsCode = '" & Trim(oForm01.Items("OutWhCd").Specific.VALUE) & "'"
                      oRecordSet01.DoQuery sQry
                      oForm01.Items("OutWhName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                  End If
                   
                  ' 입고 창고
                  If pval.ItemUID = "InWhCd" Then
                      sQry = "Select WhsName From [OWHS] Where WhsCode = '" & Trim(oForm01.Items("InWhCd").Specific.VALUE) & "'"
                      oRecordSet01.DoQuery sQry
                      oForm01.Items("InWhCd").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                  End If
               
                  ' 작업요청
                   If pval.ItemUID = "Mat01" And pval.ColUID = "SD091HNo" Then
                       Dim i As Integer
                       Dim j As Integer
                       Dim DocEntry As Integer
                       Dim LineId As Integer
                       j = InStr(Trim(oMat01.Columns("SD091HNo").Cells(pval.Row).Specific.VALUE), "-")
                       DocEntry = Left(Trim(oMat01.Columns("SD091HNo").Cells(pval.Row).Specific.VALUE), j - 1)
                       LineId = Mid(Trim(oMat01.Columns("SD091HNo").Cells(pval.Row).Specific.VALUE), j + 1)
                       sQry = "Select U_ItemCode, U_ItemName, U_ItemGu, U_Qty, "
                       sQry = sQry & "U_Unweight, U_Weight, U_Comments, U_ItmBsort, U_ItmMsort, U_Unit1, U_Size, U_ItemType, "
                       sQry = sQry & "U_Quality, U_Mark, U_CallSize, U_SbasUnit "
                       sQry = sQry & "From [@PS_SD091L] Where DocEntry = '" & DocEntry & "' And LineId = '" & LineId & "'"
                       oRecordSet01.DoQuery sQry
                       oMat01.FlushToDataSource
                       Call oDS_PS_SD090L.setValue("U_ItemCode", pval.Row - 1, Trim(oRecordSet01.Fields(0).VALUE))
                       Call oDS_PS_SD090L.setValue("U_ItemName", pval.Row - 1, Trim(oRecordSet01.Fields(1).VALUE))
                       Call oDS_PS_SD090L.setValue("U_ItemGu", pval.Row - 1, Trim(oRecordSet01.Fields(2).VALUE))
                       Call oDS_PS_SD090L.setValue("U_Qty", pval.Row - 1, Trim(oRecordSet01.Fields(3).VALUE))
                       Call oDS_PS_SD090L.setValue("U_Unweight", pval.Row - 1, Trim(oRecordSet01.Fields(4).VALUE))
                       Call oDS_PS_SD090L.setValue("U_Weight", pval.Row - 1, Trim(oRecordSet01.Fields(5).VALUE))
                       Call oDS_PS_SD090L.setValue("U_Comments", pval.Row - 1, Trim(oRecordSet01.Fields(6).VALUE))
                       Call oDS_PS_SD090L.setValue("U_ItmBsort", pval.Row - 1, Trim(oRecordSet01.Fields(7).VALUE))
                       Call oDS_PS_SD090L.setValue("U_ItmMsort", pval.Row - 1, Trim(oRecordSet01.Fields(8).VALUE))
                       Call oDS_PS_SD090L.setValue("U_Unit1", pval.Row - 1, Trim(oRecordSet01.Fields(9).VALUE))
                       Call oDS_PS_SD090L.setValue("U_Size", pval.Row - 1, Trim(oRecordSet01.Fields(10).VALUE))
                       Call oDS_PS_SD090L.setValue("U_ItemType", pval.Row - 1, Trim(oRecordSet01.Fields(11).VALUE))
                       Call oDS_PS_SD090L.setValue("U_Quality", pval.Row - 1, Trim(oRecordSet01.Fields(12).VALUE))
                       Call oDS_PS_SD090L.setValue("U_Mark", pval.Row - 1, Trim(oRecordSet01.Fields(13).VALUE))
                       Call oDS_PS_SD090L.setValue("U_CallSize", pval.Row - 1, Trim(oRecordSet01.Fields(14).VALUE))
                       Call oDS_PS_SD090L.setValue("U_SbasUnit", pval.Row - 1, Trim(oRecordSet01.Fields(15).VALUE))
                       Set oRecordSet01 = Nothing

                       AddMatrixRow 1, oMat01.VisualRowCount, False
                       For i = 1 To oMat01.VisualRowCount
                           If oDS_PS_SD090L.GetValue("U_SD091HNo", i - 1) = "" Then
                               oDS_PS_SD090L.RemoveRecord i - 1
                               oMat01.LoadFromDataSource
                           End If
                       Next
                       
                       Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
                       sQry = "Select U_OutWhCd, U_InWhCd From [@PS_SD091H] Where DocEntry = '" & DocEntry & "'"
                       oRecordSet01.DoQuery sQry
                       Call oDS_PS_SD090H.setValue("U_OutWhCd", pval.Row - 1, Trim(oRecordSet01.Fields(0).VALUE))
                       Call oDS_PS_SD090H.setValue("U_InWhCd", pval.Row - 1, Trim(oRecordSet01.Fields(1).VALUE))
                       Set oRecordSet01 = Nothing
                       
                       
                       For i = 0 To oMat01.VisualRowCount - 1
                           If oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE = "" Then
                               SumQty = SumQty
                           Else
                               SumQty = SumQty + oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE
                           End If
                           SumWeight = SumWeight + oMat01.Columns("Weight").Cells(i + 1).Specific.VALUE
                    
                        Next i
                        oForm01.Items("SumQty").Specific.VALUE = SumQty
                        oForm01.Items("SumWeight").Specific.VALUE = SumWeight
                
                       
                       AddMatrixRow 1, oMat01.VisualRowCount, False
                   End If
                                
                   Set oRecordSet01 = Nothing
                End If
            Case et_MATRIX_LOAD: '//11
'                AddMatrixRow 1, oMat01.VisualRowCount, False
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
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
             '저장 후 추가 가능처리
              If pval.ItemUID = "1" Then
                If oForm01.Mode = fm_ADD_MODE And pval.Action_Success = True Then
                    oForm01.Mode = fm_OK_MODE
                    Call Sbo_Application.ActivateMenuItem("1282")
                ElseIf oForm01.Mode = fm_ADD_MODE And pval.Action_Success = False Then
                    FormItemEnabled
                    AddMatrixRow 1, oMat01.RowCount, True
                End If
              End If
            Case et_KEY_DOWN: '//2
                If pval.Action_Success = True Then
                    oSeq = 1
                End If
            Case et_COMBO_SELECT: '//5
                If pval.ItemChanged = True Then
                    Call oForm01.Freeze(True)
                        If (pval.ItemUID = "BPLId") Then
                            If oForm01.Items("BPLId").Specific.VALUE = "1" Then
                                Call oDS_PS_SD090H.setValue("U_OutWhCd", 0, "101")
                                Call oDS_PS_SD090H.setValue("U_InWhCd", 0, "104")
                            ElseIf oForm01.Items("BPLId").Specific.VALUE = "4" Then
                                Call oDS_PS_SD090H.setValue("U_OutWhCd", 0, "104")
                                Call oDS_PS_SD090H.setValue("U_InWhCd", 0, "101")
                            End If
                        End If
                    oForm01.Update
                    Call oForm01.Freeze(False)
                End If

            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
                 ' 이동요청 문서 Query
    
       
'            ' 품목 이름 Query
'                If (pval.ItemUID = "Mat01" And (pval.ColUID = "ItemCode")) Then
'                    sQry = "Select ItemName, U_ItmBsort, U_ItmMsort, U_Unit1, U_Size, U_ItemType, U_Quality, U_Mark, U_CallSize, U_SbasUnit From [OITM] Where "
'                    sQry = sQry & "ItemCode = '" & Trim(oMat01.Columns("ItemCode").Cells(pval.Row).Specific.Value) & "'"
'                    oRecordSet01.DoQuery sQry
'                    oMat01.Columns("ItemName").Cells(pval.Row).Specific.Value = Trim(oRecordSet01.Fields(0).Value)
'            ' 품목 대분류
'                    sQry02 = "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE Code = '" & Trim(oRecordSet01.Fields(1).Value) & "'"
'                    oRecordSet02.DoQuery sQry02
'                    Call oMat01.Columns("ItmBsort").Cells(pval.Row).Specific.Select(oRecordSet02.Fields(0).Value, psk_ByValue)
'            ' 품목 중분류
'                    sQry02 = "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE Code = '" & Trim(oRecordSet01.Fields(2).Value) & "'"
'                    oRecordSet02.DoQuery sQry02
'                    Call oMat01.Columns("ItmBsort").Cells(pval.Row).Specific.Select(oRecordSet02.Fields(0).Value, psk_ByValue)
'            ' 형태타입
'                    sQry02 = "SELECT Code, Name FROM [@PSH_SHAPE] WHERE Code = '" & Trim(oRecordSet01.Fields(5).Value) & "'"
'                    oRecordSet02.DoQuery sQry02
'                    Call oMat01.Columns("ItemType").Cells(pval.Row).Specific.Select(oRecordSet02.Fields(0).Value, psk_ByValue)
'            ' 질별
'                    sQry02 = "SELECT Code, Name FROM [@PSH_QUALITY] WHERE Code = '" & Trim(oRecordSet01.Fields(6).Value) & "'"
'                    oRecordSet02.DoQuery sQry02
'                    Call oMat01.Columns("Quality").Cells(pval.Row).Specific.Select(oRecordSet02.Fields(0).Value, psk_ByValue)
'            ' 인증기호
'                    sQry02 = "SELECT Code, Name FROM [@PSH_MARK] WHERE Code = '" & Trim(oRecordSet01.Fields(7).Value) & "'"
'                    oRecordSet02.DoQuery sQry02
'                    Call oMat01.Columns("Mark").Cells(pval.Row).Specific.Select(oRecordSet02.Fields(0).Value, psk_ByValue)
'            ' 판매기준단위
'                    sQry02 = "SELECT Code, Name FROM [@PSH_UOMORG] WHERE Code = '" & Trim(oRecordSet01.Fields(9).Value) & "'"
'                    oRecordSet02.DoQuery sQry02
'                    Call oMat01.Columns("SbasUnit").Cells(pval.Row).Specific.Select(oRecordSet02.Fields(0).Value, psk_ByValue)
'                    oMat01.Columns("Unit1").Cells(pval.Row).Specific.Value = Trim(oRecordSet01.Fields(3).Value)
'                    oMat01.Columns("Size").Cells(pval.Row).Specific.Value = Trim(oRecordSet01.Fields(4).Value)
'                End If
            Case et_MATRIX_LOAD: '//11
                
                For i = 0 To oMat01.VisualRowCount - 1
                    If oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE = "" Then
                        SumQty = SumQty
                    Else
                        SumQty = SumQty + oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE
                    End If
                    SumWeight = SumWeight + oMat01.Columns("Weight").Cells(i + 1).Specific.VALUE
                    
                Next i
                oForm01.Items("SumQty").Specific.VALUE = SumQty
                oForm01.Items("SumWeight").Specific.VALUE = SumWeight
        
                AddMatrixRow 1, oMat01.VisualRowCount, True
            Case et_FORM_ACTIVATE: '//18
            If oSeq = 1 Then
                oSeq = 0
            End If
            Case et_FORM_DEACTIVATE: '//19
            Case et_FORM_RESIZE '//20
            Case et_CHOOSE_FROM_LIST '//27
                
            Case et_GOT_FOCUS: '//3
                oLast_Item_UID = pval.ItemUID
'                If oLast_Item_UID = "매트릭스" Then
'                    If pval.Row > 0 Then
'                        oLast_Item_UID = pval.ItemUID
'                        oLast_Col_UID = pval.ColUID
'                        oLast_Col_Row = pval.Row
'                    End If
'                Else
'                    oLast_Item_UID = pval.ItemUID
'                    oLast_Col_UID = ""
'                    oLast_Col_Row = 0
'                End If
            Case et_LOST_FOCUS: '//4
            Case et_FORM_UNLOAD: '//17
                RemoveForms oFormUniqueID01
                Set oForm01 = Nothing
                Set oMat01 = Nothing
        End Select
      End If
    
    Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Raise_ItemEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_ItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub



Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_MenuEvent_Error
  Dim i&
  Dim SumQty As Long
  Dim SumWeight As Currency
  
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행닫기
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼

        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1281": '찾기
                FormItemEnabled
'                oForm01.Items("ItemCode").Click ct_Regular
            Case "1282": '추가
                FormItemEnabled
                FormClear
                AddMatrixRow 0, 0, True
                oForm01.Items("CardCode").Click ct_Collapsed
                
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                FormItemEnabled
                If oMat01.VisualRowCount > 0 Then
                    If oMat01.Columns("SD091HNo").Cells(oMat01.VisualRowCount).Specific.VALUE <> "" Then
                        AddMatrixRow 1, oMat01.RowCount, True
                    End If
                End If
            Case "1293": '행닫기
                If oMat01.RowCount <> oMat01.VisualRowCount Then
                    For i = 1 To oMat01.VisualRowCount
                        oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
                    Next i
                    oMat01.FlushToDataSource    ' DBDataSource에 레코드가 한줄 더 생긴다.
                    Call oDS_PS_SD090L.RemoveRecord(oDS_PS_SD090L.Size - 1)     ' 레코드 한 줄을 지운다.
                    oMat01.LoadFromDataSource   ' DBDataSource를 매트릭스에 올리고
                    If oMat01.RowCount = 0 Then
                        Call AddMatrixRow(1, 0, True)
                    Else
                        If Trim(oDS_PS_SD090L.GetValue("U_ItemCode", oMat01.RowCount - 1)) <> "" Then
                            Call AddMatrixRow(1, oMat01.RowCount, True)

                        End If
                    End If
                    
                    
                    For i = 0 To oMat01.VisualRowCount - 1
                        If oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE = "" Then
                            SumQty = SumQty
                        Else
                            SumQty = SumQty + oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE
                        End If
                        SumWeight = SumWeight + oMat01.Columns("Weight").Cells(i + 1).Specific.VALUE
                    
                    Next i
                    oForm01.Items("SumQty").Specific.VALUE = SumQty
                    oForm01.Items("SumWeight").Specific.VALUE = SumWeight
                End If
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
            Case et_FORM_DATA_ADD:  '//34 - 추가
            Case et_FORM_DATA_UPDATE: '//35 - 업데이트
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
        '//작업
    ElseIf (eventInfo.BeforeAction = False) Then
        '//작업
    End If
    Exit Sub
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function CreateItems() As Boolean
On Error GoTo CreateItems_Error
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_SD090H = oForm01.DataSources.DBDataSources("@PS_SD090H")
    Set oDS_PS_SD090L = oForm01.DataSources.DBDataSources("@PS_SD090L")
    Set oMat01 = oForm01.Items("Mat01").Specific                        '// 매트릭스 데이터 셋
    
    Call oForm01.DataSources.UserDataSources.Add("SumQty", dt_SUM)
    Call oForm01.DataSources.UserDataSources.Add("SumWeight", dt_QUANTITY)
    
    Call oForm01.Items("SumQty").Specific.DataBind.SetBound(True, "", "SumQty")
    Call oForm01.Items("SumWeight").Specific.DataBind.SetBound(True, "", "SumWeight")
    
    oDS_PS_SD090H.setValue "U_DocDate", 0, Format(Now, "yyyymmdd")
     
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL  Where BPLId = '1' Or BPLId = '4' order by BPLId", "1", False, False)
    
    Set oRecordSet01 = Nothing
    Exit Function
CreateItems_Error:
    Set oRecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub ComboBox_Setting()
On Error GoTo ComboBox_Setting_Error
    '//콤보에 기본값설정
' 반출등록
    Call MDC_SetMod.Set_ComboList(oForm01.Items("OutWhCd").Specific, "SELECT WhsCode, WhsName FROM [OWHS] order by WhsCode", "104", False, True)
    Call MDC_SetMod.Set_ComboList(oForm01.Items("InWhCd").Specific, "SELECT WhsCode, WhsName FROM [OWHS] order by WhsCode", "101", False, True)
    
' 품목대분류
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ItmBsort"), "SELECT Code, Name FROM [@PSH_ITMBSORT] ORDER BY Code")
' 품목중분류
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ItmMsort"), "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT] ORDER BY U_Code")
' 형태타입
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ItemType"), "SELECT Code, Name FROM [@PSH_SHAPE] ORDER BY Code")
' 질별
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("Quality"), "SELECT Code, Name FROM [@PSH_QUALITY] ORDER BY Code")
' 인증기호
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("Mark"), "SELECT Code, Name FROM [@PSH_MARK] ORDER BY Code")
' 매입기준단위
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("SbasUnit"), "SELECT Code, Name FROM [@PSH_UOMORG] ORDER BY Code")
    Exit Sub
ComboBox_Setting_Error:
    Sbo_Application.SetStatusBarMessage "ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub CF_ChooseFromList()
On Error GoTo CF_ChooseFromList_Error
    '//ChooseFromList 설정
    Exit Sub
CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub Initial_Setting()
On Error GoTo Initial_Setting_Error
 ' 사업장
    Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
 ' 인수자
    oForm01.Items("RepName").Specific.VALUE = MDC_PS_Common.User_MSTCOD()
    Exit Sub
Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub FormItemEnabled()
On Error GoTo FormItemEnabled_Error

    Dim oCombo          As SAPbouiCOM.ComboBox '콤보박스
    Dim oCombo1         As SAPbouiCOM.ComboBox '콤보박스
    Dim oColumn         As SAPbouiCOM.Column
    Dim lQuery As String
    Dim lRecordSet As SAPbobsCOM.Recordset
    
    Set lRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
        oForm01.Items("DocNum").Enabled = False
        oForm01.Items("CardCode").Enabled = True
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("RepName").Enabled = True
        oForm01.Items("DocDate").Enabled = True
        oForm01.Items("ShipTo").Enabled = True
        oForm01.Items("CarCo").Enabled = True
        oForm01.Items("CarNo").Enabled = True
        oForm01.Items("ArrSite").Enabled = True
        oForm01.Items("Fare").Enabled = True
        oForm01.Items("Specific").Enabled = True
        oForm01.Items("ChulPrin").Enabled = False
'        oForm01.Items("OutWhCd").Enabled = True
'        oForm01.Items("InWhCd").Enabled = True
        oMat01.Columns("ItemCode").Editable = True
        oMat01.Columns("ItemGu").Editable = True
        '////////////////////////////////////////////////폼상태변경//////////////////////////////////
        oForm01.Items("DocDate").Specific.String = Format(Now, "yyyymmdd")
        oForm01.Items("RepName").Click ct_Regular
        
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
        oForm01.Items("DocNum").Enabled = False
        oForm01.Items("CardCode").Enabled = True
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("RepName").Enabled = True
        oForm01.Items("DocDate").Enabled = True
        oForm01.Items("ShipTo").Enabled = True
        oForm01.Items("CarCo").Enabled = True
        oForm01.Items("CarNo").Enabled = True
        oForm01.Items("ArrSite").Enabled = True
        oForm01.Items("Fare").Enabled = False
        oForm01.Items("Specific").Enabled = False
        oForm01.Items("ChulPrin").Enabled = True 'False
        
'        oForm01.Items("OutWhCd").Enabled = False
'        oForm01.Items("InWhCd").Enabled = False
        oMat01.Columns("ItemCode").Editable = False
        oMat01.Columns("ItemGu").Editable = False

    ElseIf (oForm01.Mode = fm_OK_MODE) Then
   '////////////////////////////////////////////////폼상태변경//////////////////////////////////
'        lQuery = "Select Status From [@PS_SD090H] Where DocNum = '"
'        lQuery = lQuery + oForm01.Items("DocNum").Specific.Value
'        lQuery = lQuery + "'"
'
'        lRecordSet.DoQuery lQuery
'        If (lRecordSet.Fields(0).Value = "O") Then
'            oForm01.Items("DocNum").Enabled = False
'            oForm01.Items("BPLId").Enabled = False
'            oForm01.Items("RepName").Enabled = False
'            oMat01.Columns("ItemCode").Editable = False
'            oMat01.Columns("OutWhCd").Editable = False
'            oMat01.Columns("InWhCd").Editable = False
'            oMat01.Columns("BatchNum").Editable = False
'        ElseIf (lRecordSet.Fields(0).Value = "C") Then
'            oForm01.Items("DocNum").Enabled = False
'            oForm01.Items("BPLId").Enabled = False
'            oForm01.Items("RepName").Enabled = False
'            oMat01.Columns("ItemCode").Editable = False
'            oMat01.Columns("OutWhCd").Editable = False
'            oMat01.Columns("InWhCd").Editable = False
'            oMat01.Columns("BatchNum").Editable = False
'        End If
        '//Status 설정
        oForm01.Items("DocNum").Enabled = False
        oForm01.Items("CardCode").Enabled = False
        oForm01.Items("BPLId").Enabled = False
        oForm01.Items("RepName").Enabled = False
        oForm01.Items("DocDate").Enabled = False
        oForm01.Items("ShipTo").Enabled = False
        oForm01.Items("CarCo").Enabled = False
        oForm01.Items("CarNo").Enabled = False
        oForm01.Items("ArrSite").Enabled = False
        oForm01.Items("Fare").Enabled = False
        oForm01.Items("Specific").Enabled = False
        oForm01.Items("ChulPrin").Enabled = True
'        oForm01.Items("OutWhCd").Enabled = False
'        oForm01.Items("InWhCd").Enabled = False
        oMat01.Columns("ItemCode").Editable = False
        oMat01.Columns("ItemGu").Editable = False
        
        lQuery = "SELECT Status,Canceled FROM [@PS_SD090H] WHERE DocEntry = '" & oForm01.Items("DocNum").Specific.VALUE & "'"
        lRecordSet.DoQuery lQuery
        If (lRecordSet.Fields(0).VALUE = "O") Then
            oForm01.Items("Status").Specific.VALUE = "미결"
        ElseIf (lRecordSet.Fields(0).VALUE = "C") Then
            If (lRecordSet.Fields(1).VALUE = "Y") Then
                oForm01.Items("Status").Specific.VALUE = "취소"
            Else
                oForm01.Items("Status").Specific.VALUE = "종료"
            End If
        End If
        '////////////////////////////////////////////////폼상태변경//////////////////////////////////
    End If
    Call oMat01.AutoResizeColumns
    
    Exit Sub

FormItemEnabled_Error:
    Sbo_Application.SetStatusBarMessage "FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub AddMatrixRow(ByVal oSeq As Integer, ByVal oRow As Long, Optional RowIserted As Boolean)
'On Error GoTo AddMatrixRow_Error

   Select Case oSeq
        Case 0:
            oMat01.AddRow           ' 매트릭스에 새로운 로를 추가한다.
            oDS_PS_SD090L.setValue "U_LIneNum", oRow, oRow + 1
            oMat01.LoadFromDataSource
        Case 1:
            oDS_PS_SD090L.InsertRecord oRow
            oDS_PS_SD090L.setValue "U_LIneNum", oRow, oRow + 1
            oMat01.LoadFromDataSource
    End Select
'AddMatrixRow_Error:
'    Sbo_Application.SetStatusBarMessage "AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub FormClear()
On Error GoTo FormClear_Error
    Dim DocNum As String
    DocNum = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_SD090'", "")
    If DocNum = 0 Then
'        oForm01.Items("DocEntry").Specific.String = 1
        oDS_PS_SD090H.setValue "DocNum", 0, "1"
    Else
'        oForm01.Items("DocEntry").Specific.String = DocNum
        oDS_PS_SD090H.setValue "DocNum", 0, DocNum    ' 화면에 적용이 안되기 때문
    End If
    Exit Sub
FormClear_Error:
    Sbo_Application.SetStatusBarMessage "FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function DataValidCheck() As Boolean
On Error GoTo DataValidCheck_Error
    '//유효성검사
    Exit Function
DataValidCheck_Error:
    Sbo_Application.SetStatusBarMessage "DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
    '//데이터변화에 따른처리
    Dim i As Long
    Dim oRecordSet01        As SAPbobsCOM.Recordset
    Dim sQry                As String
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
    
        Case "Mat01"    'rowcount: 로 카운트를 반환, VisualRowCount: 삭제된 로를 제외하고 현재 보이는 로 카운트를 반환
'            If oCol = "ItemCode" Then
'                sQry = "SELECT U_ItmBsort, U_ItmMsort, U_Unit1, U_Size, U_ItemType, U_Quality, U_Mark, U_CallSize, U_SbasUnit"
'                sQry = sQry & "FROM [OITM] WHERE ItemCode ='" & Trim(oMat01.Columns("ItemCode").Cells(oRow).Specific) & "'"
'                oRecordSet01.DoQuery (sQry)
'
'                oRecordSet01.MoveFirst
'
'                For i = 0 To oRecordSet01.EOF
'                oMat01.Columns("ItmBsort").Cells(oRecordSet01.RecordCount).Specific.Value = Trim(oRecordSet01.Fields(0).Value)
'                oMat01.Columns("ItmMsort").Cells(oRecordSet01.RecordCount).Specific.Value = Trim(oRecordSet01.Fields(1).Value)
'                oMat01.Columns("Unit1").Cells(oRecordSet01.RecordCount).Specific.Value = Trim(oRecordSet01.Fields(2).Value)
'                oMat01.Columns("Size").Cells(oRecordSet01.RecordCount).Specific.Value = Trim(oRecordSet01.Fields(3).Value)
'                oMat01.Columns("ItemType").Cells(oRecordSet01.RecordCount).Specific.Value = Trim(oRecordSet01.Fields(4).Value)
'                oMat01.Columns("Quality").Cells(oRecordSet01.RecordCount).Specific.Value = Trim(oRecordSet01.Fields(5).Value)
'                oMat01.Columns("Mark").Cells(oRecordSet01.RecordCount).Specific.Value = Trim(oRecordSet01.Fields(6).Value)
'                oMat01.Columns("CallSize").Cells(oRecordSet01.RecordCount).Specific.Value = Trim(oRecordSet01.Fields(7).Value)
'                oMat01.Columns("SbasUnit").Cells(oRecordSet01.RecordCount).Specific.Value = Trim(oRecordSet01.Fields(8).Value)
'                oRecordSet01.MoveNext
'                Next i
'            End If
            
            If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 0) And _
            Trim(oMat01.Columns("ItemCode").Cells(oRow).Specific.VALUE) <> "" Then
                oMat01.FlushToDataSource      '데이터 소스를 지우고 매트릭스로부터 데이터 소스 레코드로 각 로를 복사한다.
                AddMatrixRow 1, oMat01.RowCount, True
                oMat01.Columns("ItemCode").Cells(oRow).Click ct_Regular 'Column: 칼럼 오브젝트의 collection을 반환한다.
            End If
    End Select
    Set oRecordSet01 = Nothing
End Sub

Private Sub MTX01()
On Error GoTo MTX01_Error
    '//메트릭스에 데이터 로드
    Exit Sub
MTX01_Error:
    Sbo_Application.SetStatusBarMessage "MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum      As Integer
    Dim DocNum      As String
    
    ErrNum = 0
    
    '// Check
    Select Case True
        Case oDS_PS_SD090H.GetValue("U_DocDate", 0) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_SD090H.GetValue("U_BPLId", 0) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_SD090H.GetValue("U_OutWhCd", 0) = ""
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_SD090H.GetValue("U_InWhCd", 0) = ""
            ErrNum = 4
            GoTo HeaderSpaceLineDel_Error
    End Select
    
    HeaderSpaceLineDel = True
Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "거래일자는 필수입력 사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "사업장은 필수입력 사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "출고창고는 필수입력 사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "입고창고는 필수입력 사항입니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function
        
Private Function MatrixSpaceLineDel() As Boolean
'------------------------------------------------------------------------------
' 저장할 데이터의 유효성을 점검한다
'------------------------------------------------------------------------------
On Error GoTo MatrixSpaceLineDel_Error
    Dim i           As Long
    Dim K           As Long
    Dim ErrNum      As Integer
    Dim Chk_Data    As String
    Dim oRow        As Integer
    Dim oRecordSet01  As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
            
    ErrNum = 0
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '// 화면상의 메트릭스에 입력된 내용을 모두 디비데이터소스로 넘긴다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    ' Flushes current data from the user interface to the bounded data source, as follows:
    ' 1. Cleans the data source.
    ' 2. Copies each row from the matrix to the corresponding data source record.
    oMat01.FlushToDataSource

    '// 라인
    If oMat01.VisualRowCount <= 1 Then
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    End If
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '// 맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
    '// 이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    If oMat01.VisualRowCount > 0 Then
    
        For i = 0 To oMat01.VisualRowCount - 2
            oDS_PS_SD090L.Offset = i
            Select Case True
                Case oDS_PS_SD090L.GetValue("U_ItemCode", i) = ""
                    ErrNum = 2
                    GoTo MatrixSpaceLineDel_Error
                    
                Case oDS_PS_SD090L.GetValue("U_Qty", i) = ""
                    ErrNum = 5
                    GoTo MatrixSpaceLineDel_Error
                
                Case oDS_PS_SD090L.GetValue("U_SD091HNo", i) = ""
                    ErrNum = 6
                    GoTo MatrixSpaceLineDel_Error
            End Select
        Next i
       
        If oDS_PS_SD090L.GetValue("U_SD091HNo", oMat01.VisualRowCount - 1) = "" Then
            oDS_PS_SD090L.RemoveRecord oMat01.VisualRowCount - 1
        End If
    End If
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '행을 삭제하였으니 DB데이터 소스를 다시 가져온다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    oMat01.LoadFromDataSource
    
    Set oRecordSet01 = Nothing
    MatrixSpaceLineDel = True
Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////
MatrixSpaceLineDel_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "라인 데이터가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "아이템 데이터는 필수입니다. 확인하세요.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "수량은 필수입니다. 확인하세요.", "E"
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "이동요청 문서는 필수입니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function

Private Function Exist_YN(DocNum$) As String

    Dim oRecordSet01   As SAPbobsCOM.Recordset
    Dim sQry           As String
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sQry = "SELECT Top 1 T1.DocNum FROM [@PS_SD090H] T1 "
    sQry = sQry & " WHERE T1.DocNum  = '" & DocNum & "'"
    oRecordSet01.DoQuery sQry
    
    Do Until oRecordSet01.EOF
        Exist_YN = oRecordSet01(0).VALUE
        oRecordSet01.MoveNext
    Loop
    
    If Trim(Exist_YN) = "" Then
        Exist_YN = ""
        Exit Function
    End If
    
    Set oRecordSet01 = Nothing
End Function

Private Sub Raise_EVENT_ROW_DELETE(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_ROW_DELETE_Error
    Dim i As Long
    If pval.BeforeAction = True Then
        If Sbo_Application.MessageBox("정말 삭제 하시겠습니까?", 1, "OK", "NO") <> 1 Then
            BubbleEvent = False
        End If
        '//행삭제전 행삭제가능여부검사
    ElseIf pval.BeforeAction = False Then
        For i = 1 To oMat01.VisualRowCount
            oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
        Next i
'        oMat01.Clear
        oMat01.FlushToDataSource
        Call oDS_PS_SD090L.RemoveRecord(oDS_PS_SD090L.Size - 1)
        oMat01.LoadFromDataSource
        If oMat01.RowCount = 0 Then
            Call AddMatrixRow(0, oMat01.RowCount, True)
        Else
            If Trim(oDS_PS_SD090L.GetValue("U_ItemCode", oMat01.RowCount - 1)) <> "" Then
                Call AddMatrixRow(1, oMat01.RowCount, True)
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function StockTrans() As Boolean
On Error GoTo StockTrans_Error
    Dim RetVal As Long
    Dim ErrCode As Long
    Dim ErrMsg As String
    Dim lQuery As String
    Dim lRecordSet As SAPbobsCOM.Recordset
    Dim oRecordSet As SAPbobsCOM.Recordset
    Dim lMaxBatchNum As Long    '해당 품목의 최대 배치번호
    Dim lBatchWeight As Double  '배치별 중량
    Dim lTypeCount As Integer   '전체 StockInfo 구조체배열의 RowCount
    Dim i, j, K, Q, z, r  As Long
    Dim DocCnt        As Long
    Dim Chk1_Val As String
    Dim sCur_ItemCode As String
    Dim sNxt_ItemCode As String
    Dim sCur_TrCardCode As String
    Dim sCur_TrOutWhs As String
    Dim sNxt_TrOutWhs As String
    Dim sCur_TrInWhs As String
    Dim sNxt_TrInWhs As String
    Dim RtnDocNum     As String
    Dim oStockTrans   As SAPbobsCOM.StockTransfer
    Dim oPrgBar       As SAPbouiCOM.ProgressBar
    Dim StockTransLineCounter As Long
    
    Set lRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BatchNum As String
    
        StockTrans = True
    
    For i = 0 To oMat01.RowCount - 1
        ReDim Preserve StockInfo(lTypeCount)
        'DI API
        StockInfo(lTypeCount).CardCode = Trim(oDS_PS_SD090H.GetValue("U_CardCode", 0))
        StockInfo(lTypeCount).ItemCode = Trim(oDS_PS_SD090L.GetValue("U_ItemCode", i))
        StockInfo(lTypeCount).FromWarehouseCode = Trim(oDS_PS_SD090H.GetValue("U_OutWhCd", 0))
        StockInfo(lTypeCount).ToWarehouseCode = Trim(oDS_PS_SD090H.GetValue("U_InWhCd", 0))
        StockInfo(lTypeCount).BatchNum = Trim(oDS_PS_SD090L.GetValue("U_BatchNum", i))
       
        StockInfo(lTypeCount).Weight = Round(Trim(oDS_PS_SD090L.GetValue("U_Weight", i)), 2)
        StockInfo(lTypeCount).UnWeight = Round(Trim(oDS_PS_SD090L.GetValue("U_Unweight", i)), 2)
        StockInfo(lTypeCount).BatchWeight = Round(Trim(oDS_PS_SD090L.GetValue("U_Qty", i)), 2)
        StockInfo(lTypeCount).Qty = Val(Trim(oDS_PS_SD090L.GetValue("U_Qty", i)))
   
        StockInfo(lTypeCount).TransNo = oForm01.Items("DocNum").Specific.VALUE & (i + 1)
        StockInfo(lTypeCount).Chk = "N"
        StockInfo(lTypeCount).MatrixRow = (i + 1)
        StockInfo(lTypeCount).Indate = oForm01.Items("DocDate").Specific.VALUE
        lTypeCount = lTypeCount + 1
    Next
    
    For i = 0 To (UBound(StockInfo))
        StockInfo(i).StockTransDocEntry = ""
    Next
    
    Sbo_Company.StartTransaction
    For i = 0 To (UBound(StockInfo))
    
        Chk1_Val = StockInfo(i).Chk
        
        If Chk1_Val <> "N" Then GoTo Continue_First
        
        sCur_TrOutWhs = StockInfo(i).FromWarehouseCode
        
        Set oStockTrans = Sbo_Company.GetBusinessObject(oStockTransfer)
        oStockTrans.CardCode = StockInfo(i).CardCode
        oStockTrans.DocDate = Format(StockInfo(i).Indate, "&&&&-&&-&&")
        oStockTrans.FromWarehouse = sCur_TrOutWhs
        oStockTrans.Comments = "재고이전" & oForm01.Items("DocNum").Specific.VALUE & "."
        
        StockTransLineCounter = -1
        For K = i To (UBound(StockInfo))
            Chk1_Val = StockInfo(K).Chk
            
            If Chk1_Val <> "N" Then GoTo Continue_Second
            sCur_TrCardCode = StockInfo(K).CardCode
            sNxt_TrOutWhs = StockInfo(K).FromWarehouseCode
            sCur_ItemCode = StockInfo(K).ItemCode
            sCur_TrInWhs = StockInfo(K).ToWarehouseCode
            
            If (sCur_TrOutWhs <> sNxt_TrOutWhs) Then
                GoTo Continue_Second
            End If

            If (i <> K) Then
                oStockTrans.Lines.Add
            End If
            StockTransLineCounter = StockTransLineCounter + 1
            '---------------------------------------------------------------------------< Line >----------
            With oStockTrans.Lines
                
                .ItemCode = StockInfo(K).ItemCode
                .UserFields("U_Qty").VALUE = Trim(StockInfo(K).Qty)                 '// 수량
                .UserFields("U_UnWeight").VALUE = Trim(StockInfo(K).UnWeight)       '// 단중
                .Quantity = Round(StockInfo(K).Qty, 2)                              '// 수량
                .WarehouseCode = StockInfo(K).ToWarehouseCode
                '//ManBatchNum = 'Y' 이면 배치번호를 입력하지 않는다.
                .UserFields("U_BatchNum").VALUE = StockInfo(K).BatchNum
                .BatchNumbers.BatchNumber = StockInfo(K).BatchNum
                .BatchNumbers.Quantity = Round(StockInfo(K).BatchWeight, 2)

                .BatchNumbers.Notes = "재고이전(Addon)"
                StockInfo(K).Chk = "Y"   '/ 적용한 라인에 대한 표시
                StockInfo(K).StockTransDocEntry = "Checked"
                StockInfo(K).StockTransLineNum = StockTransLineCounter
                
                For Q = K + 1 To (UBound(StockInfo))
                    Chk1_Val = StockInfo(Q).Chk
                    
                    If Chk1_Val <> "N" Then GoTo Continue_Sixth   '/ 체크2 에 않된건 Skip
                    
                    sNxt_TrOutWhs = StockInfo(Q).FromWarehouseCode
                    sNxt_ItemCode = StockInfo(Q).ItemCode
                    sNxt_TrInWhs = StockInfo(Q).ToWarehouseCode
    
                    If sNxt_TrOutWhs = sCur_TrOutWhs And sCur_ItemCode = sNxt_ItemCode And sCur_TrInWhs = sNxt_TrInWhs Then
                        '//ManBatchNum = 'Y' 이면 배치번호를 입력하지 않는다.
                        If MDC_PS_Common.GetValue("SELECT ManBatchNum FROM OITM WHERE ITEMCODE = ''", 0, 1) = "Y" Then
                            .BatchNumbers.Add
                            .BatchNumbers.BatchNumber = StockInfo(Q).BatchNum
                            .BatchNumbers.Quantity = Round(StockInfo(Q).BatchWeight, 2)
                            .UserFields("Quantity").VALUE = .UserFields("Quantity").VALUE + Trim(StockInfo(Q).Qty)           '//수량
                            .Quantity = .Quantity + Round(StockInfo(Q).Weight, 2)                                            '//중량을 합함
                            .BatchNumbers.Notes = "재고이전(Addon)"
                            StockInfo(Q).Chk = "Y"                                                                           '// 적용한 라인에 대한 표시
                            StockInfo(Q).StockTransDocEntry = "Checked"
                            StockInfo(Q).StockTransLineNum = StockTransLineCounter
                        End If
                    End If
Continue_Sixth:
                Next
            End With
Continue_Second:
        Next
        '---------------------------------------------------------------------------------------------

        RetVal = oStockTrans.Add
        If RetVal = 0 Then
            DocCnt = DocCnt + 1
            Call Sbo_Company.GetNewObjectCode(RtnDocNum) '//재고이전문서번호
            For r = 0 To UBound(StockInfo)
                If (StockInfo(r).StockTransDocEntry = "Checked") Then
                    StockInfo(r).StockTransDocEntry = RtnDocNum
                End If
            Next
            '// 데이터 업데이트
         Else
            GoTo StockTrans_Error
        End If
Continue_First:
    Next   '-----------------------------------------------------------------------------------------------< First For End
    
    If (Sbo_Company.InTransaction) Then
        Sbo_Company.EndTransaction (wf_Commit)
    End If
    Call Sbo_Application.SetStatusBarMessage(DocCnt & " 개의 재고이전 문서가 발행되었습니다 !", bmt_Short, False)
    Set oStockTrans = Nothing
    Set lRecordSet = Nothing
    Set oRecordSet = Nothing
    Exit Function
'************Error Process************
StockTrans_Error:
    If (Sbo_Company.InTransaction) Then
        Sbo_Company.EndTransaction (wf_RollBack)
    End If
    Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
    Call Sbo_Application.SetStatusBarMessage(ErrCode & " : " & ErrMsg, bmt_Short, True)
    StockTrans = False
    Set oStockTrans = Nothing
    Set lRecordSet = Nothing
    Set oRecordSet = Nothing
'************Error Process************

End Function

Private Function UpdateUserField() As Boolean
On Error GoTo UpdateUserField_Error
    Dim i As Long
    Dim lQuery As String
    Dim lRecordSet As SAPbobsCOM.Recordset
    Set lRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Call oDS_PS_SD090H.setValue("U_StoTrDoc", 0, (StockInfo(i).StockTransDocEntry))
'    oForm01.Items("StoTrDoc").Specific.Value = StockInfo(i).StockTransDocEntry

    UpdateUserField = True
    Exit Function
UpdateUserField_Error:
    UpdateUserField = False
End Function

Private Sub PS_SD090_Print_Report01()
On Error GoTo PS_SD090_Print_Report01_Error
    Dim DocNum As String
    Dim WinTitle                    As String
    Dim ReportName                  As String
    Dim sQry01                      As String
    
    Call ConnectODBC
    DocNum = oForm01.Items("DocNum").Specific.VALUE
    WinTitle = "[PS_SD090] 출고원부/반출증"
    ReportName = "PS_SD090_10.rpt"
    sQry01 = "EXEC PS_SD090_10 '" & DocNum & "'"
     ReDim gRpt_Formula(1)
        ReDim gRpt_Formula_Value(1)
        ReDim gRpt_SRptSqry(1)
        ReDim gRpt_SRptName(1)
        ReDim gRpt_SFormula(1, 1)
        ReDim gRpt_SFormula_Value(1, 1)

    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry01, "1", "Y", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    Exit Sub
PS_SD090_Print_Report01_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD090_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
       
       
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
