VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_QM020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_QM020.cls
'//  Module         : QM
'//  Description    : 검사성적서등록
'//  FormType       : PS_QM020
'//  Create Date    : 2010.11.03
'//  Modified Date  :
'//  Creator        : Youn Je Hyung
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01      As String
Public oForm01              As SAPbouiCOM.Form
Public oMat01               As SAPbouiCOM.Matrix
Private oDS_PS_QM020H       As SAPbouiCOM.DBDataSource    '등록헤더
Private oDS_PS_QM020L       As SAPbouiCOM.DBDataSource    '등록라인

Private oLast_Item_UID      As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID       As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row       As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&
Private Last_CardCode$, Last_CardName$, Last_InspPrsn$
Private oCode$

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm(Optional sCode As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_QM020.srf")
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
    
    oFormUniqueID01 = "PS_QM020_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
    
    If sCode <> "" Then
        oForm01.Mode = fm_FIND_MODE
    Else
        oForm01.Mode = fm_ADD_MODE
    End If
    
    
    
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    '화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
    oForm01.DataBrowser.BrowseBy = "Code"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting
    Call FormClear
'    Call Matrix_AddRow(1, 0, True) 'oMat01
'    FormItemEnabled
    
    oForm01.EnableMenu ("1283"), True         '// 삭제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1287"), True         '// 복제
    oForm01.EnableMenu ("1284"), False        '// 취소
    oForm01.EnableMenu ("1293"), False        '// 행삭제
        
    oForm01.Update
    oForm01.Freeze False
    
    oCode = sCode
    If oCode <> "" Then
    
        Call FormItemEnabled
    
        oForm01.Items("Code").Specific.VALUE = sCode
        Call oForm01.Items("1").Click
        
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

Private Sub CreateItems()
On Error GoTo CreateItems_Error
    '//디비데이터 소스 개체 할당
    Set oDS_PS_QM020H = oForm01.DataSources.DBDataSources("@PS_QM020H")
'    Set oDS_PS_QM020L = oForm01.DataSources.DBDataSources("@PS_QM020L")
    
    '// 메트릭스 개체 할당
'    Set oMat01 = oForm01.Items("Mat01").Specific

    oDS_PS_QM020H.setValue "U_InspDate", 0, Format(Now, "YYYYMMDD")
   
    oForm01.Items("Code").Visible = True 'False
   
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
    Dim oRecordSet      As SAPbobsCOM.Recordset
        
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    '// 검사자
    Set oCombo = oForm01.Items("InspPrsn").Specific
    sQry = "SELECT DISTINCT U_MSTCOD, lastName+firstName From [OHEM] " '중복제거를 위한 DISTINCT문 추가(2012.06.27 송명규 수정)
'    sQry = sQry & "Where dept='8'"
    oRecordSet.DoQuery sQry
    Do Until oRecordSet.EOF
        oCombo.ValidValues.Add Trim(oRecordSet.Fields(0).VALUE), Trim(oRecordSet.Fields(1).VALUE)
        oRecordSet.MoveNext
    Loop
    
    Call oForm01.Items("TS_Gbn").Specific.ValidValues.Add("10", "Kgf/m2")
    Call oForm01.Items("TS_Gbn").Specific.ValidValues.Add("20", "N/mm2")
    
    '외관검사
    Set oCombo = oForm01.Items("M_Visual").Specific
    oCombo.ValidValues.Add "Y", "합격"
    oCombo.ValidValues.Add "N", "불합격"
    oCombo.Select 0, psk_Index
    
    '포장상태
    Set oCombo = oForm01.Items("PACKYN").Specific
    oCombo.ValidValues.Add "Y", "양호"
    oCombo.ValidValues.Add "N", "부적합"
    oCombo.Select 0, psk_Index
    
    '합부판정
    Set oCombo = oForm01.Items("PASSYN").Specific
    oCombo.ValidValues.Add "Y", "합격"
    oCombo.ValidValues.Add "N", "불합격"
    oCombo.Select 0, psk_Index
    
    Set oCombo = Nothing
    
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
ComboBox_Setting_Error:
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    MDC_Com.MDC_GF_Message "ComboBox_Setting_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'****************************************************************************************************************
'// ItemEventHander
'****************************************************************************************************************
Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_ItemEvent_Error
    Dim i&
    Dim ErrNum          As Integer
    Dim SPEC_MIN        As Currency
    Dim SPEC_MAX        As Currency
    Dim VALUE           As Currency
    Dim VALUE_MIN       As Currency
    Dim VALUE_MAX       As Currency
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE Or oForm01.Mode = fm_UPDATE_MODE Then
                        If HeaderSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
'                        If MatrixSpaceLineDel = False Then
'                            BubbleEvent = False
'                            Exit Sub
'                        End If
                        If oForm01.Mode = fm_ADD_MODE Then
                            If Check_Exist = False Then
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    '//헤더
                    If pval.ItemUID = "OrdNum" Then
                        If oForm01.Items("OrdNum").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    End If
                    If pval.ItemUID = "CardCode" Then
                        If oForm01.Items("CardCode").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    End If
                    '//라인
'                    If pval.ItemUID = "Mat01" Then
'                        If pval.ColUID = "PP070No" Then
'                            If oMat01.Columns("PP070No").Cells(pval.Row).Specific.Value = "" Then
'                                Sbo_Application.ActivateMenuItem ("7425")
'                                BubbleEvent = False
'                            End If
'                        End If
'                    End If
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
        
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE Then
                        oForm01.Mode = fm_OK_MODE
                        Call Sbo_Application.ActivateMenuItem("1282")
                    ElseIf oForm01.Mode = fm_OK_MODE Then
                        FormItemEnabled
                        FlushToItemValue "ItemCode"
'                        Call Matrix_AddRow(1, oMat01.RowCount, False) 'oMat01
                    End If
                End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    '//헤더
                    If pval.ItemUID = "OrdNum" Or pval.ItemUID = "ItemCode" Or pval.ItemUID = "CardCode" Then
                        FlushToItemValue pval.ItemUID
                        FlushToItemValue "ItemCode"
                    End If
                    '//측정치 CHECK
                    If pval.ItemUID <> "OrdNum" And pval.ItemUID <> "InspDate" Then
                        '//인강강도 Min
                        If Val(oDS_PS_QM020H.GetValue("U_M_TS_min", 0)) <> 0 Then
                            VALUE_MIN = Val(oDS_PS_QM020H.GetValue("U_M_TS_min", 0))
                            SPEC_MIN = Val(oForm01.Items("TS_min").Specific.VALUE)
                            SPEC_MAX = Val(oForm01.Items("TS_max").Specific.VALUE)
                            If VALUE_MIN < SPEC_MIN Or VALUE_MIN > SPEC_MAX Then
                                Sbo_Application.MessageBox ("tensile strength 인장강도(kgf/㎡)의 Min 측정치가 SPEC 규격 범위를 벗어났습니다.")
                                GoTo Raise_ItemEvent_Error
                            End If
                        End If
                        '//인강강도 Max
                        If Val(oDS_PS_QM020H.GetValue("U_M_TS_max", 0)) <> 0 Then
                            VALUE_MAX = Val(oDS_PS_QM020H.GetValue("U_M_TS_max", 0))
                            SPEC_MIN = Val(oForm01.Items("TS_min").Specific.VALUE)
                            SPEC_MAX = Val(oForm01.Items("TS_max").Specific.VALUE)
                            If VALUE_MAX < SPEC_MIN Or VALUE_MAX > SPEC_MAX Then
                                Sbo_Application.MessageBox ("tensile strength 인장강도(kgf/㎡)의 Max 측정치가 SPEC 규격 범위를 벗어났습니다.")
                                GoTo Raise_ItemEvent_Error
                            End If
                        End If
                        '//연신율 Min
                        If Val(oDS_PS_QM020H.GetValue("U_M_El_min", 0)) <> 0 Then
                            VALUE_MIN = Val(oDS_PS_QM020H.GetValue("U_M_El_min", 0))
                            SPEC_MIN = Val(oForm01.Items("Elong").Specific.VALUE)
                            If VALUE_MIN < SPEC_MIN Then
                                Sbo_Application.MessageBox ("elongation 연신율(%)의 Min 측정치가 SPEC 규격 범위를 벗어났습니다.")
                                GoTo Raise_ItemEvent_Error
                            End If
                        End If
                        '//연신율 Max
                        If Val(oDS_PS_QM020H.GetValue("U_M_El_max", 0)) <> 0 Then
                            VALUE_MAX = Val(oDS_PS_QM020H.GetValue("U_M_El_max", 0))
                            SPEC_MIN = Val(oForm01.Items("Elong").Specific.VALUE)
                            If VALUE_MAX < SPEC_MIN Then
                                Sbo_Application.MessageBox ("elongation 연신율(%)의 Max 측정치가 SPEC 규격 범위를 벗어났습니다.")
                                GoTo Raise_ItemEvent_Error
                            End If
                        End If
                        '//경도 Min
                        If Val(oDS_PS_QM020H.GetValue("U_M_Hd_min", 0)) <> 0 Then
                            VALUE_MIN = Val(oDS_PS_QM020H.GetValue("U_M_Hd_min", 0))
                            SPEC_MIN = Val(oForm01.Items("Hard_min").Specific.VALUE)
                            SPEC_MAX = Val(oForm01.Items("Hard_max").Specific.VALUE)
                            If VALUE_MIN < SPEC_MIN Or VALUE_MIN > SPEC_MAX Then
                                Sbo_Application.MessageBox ("Hardness 경도(Hv)의 Min 측정치가 SPEC 규격 범위를 벗어났습니다.")
                                GoTo Raise_ItemEvent_Error
                            End If
                        End If
                        '//경도 Max
                        If Val(oDS_PS_QM020H.GetValue("U_M_Hd_max", 0)) <> 0 Then
                            VALUE_MAX = Val(oDS_PS_QM020H.GetValue("U_M_Hd_max", 0))
                            SPEC_MIN = Val(oForm01.Items("Hard_min").Specific.VALUE)
                            SPEC_MAX = Val(oForm01.Items("Hard_max").Specific.VALUE)
                            If VALUE_MAX < SPEC_MIN Or VALUE_MAX > SPEC_MAX Then
                                Sbo_Application.MessageBox ("Hardness 경도(Hv)의 Max 측정치가 SPEC 규격 범위를 벗어났습니다.")
                                GoTo Raise_ItemEvent_Error
                            End If
                        End If
                        '//Fe
                        If Val(oDS_PS_QM020H.GetValue("U_M_Ch_Fe", 0)) <> 0 Then
                            VALUE = Val(oDS_PS_QM020H.GetValue("U_M_Ch_Fe", 0))
                            SPEC_MIN = Val(oForm01.Items("C_Fe_min").Specific.VALUE)
                            SPEC_MAX = Val(oForm01.Items("C_Fe_max").Specific.VALUE)
                            If VALUE < SPEC_MIN Or VALUE > SPEC_MAX Then
                                Sbo_Application.MessageBox ("Fe 측정치가 SPEC 규격 범위를 벗어났습니다.")
                                GoTo Raise_ItemEvent_Error
                            End If
                        End If
                        '//P
                        If Val(oDS_PS_QM020H.GetValue("U_M_Ch_P", 0)) <> 0 Then
                            VALUE = Val(oDS_PS_QM020H.GetValue("U_M_Ch_P", 0))
                            SPEC_MIN = Val(oForm01.Items("C_P_min").Specific.VALUE)
                            SPEC_MAX = Val(oForm01.Items("C_P_max").Specific.VALUE)
                            If VALUE < SPEC_MIN Or VALUE > SPEC_MAX Then
                                Sbo_Application.MessageBox ("P 측정치가 SPEC 규격 범위를 벗어났습니다.")
                                GoTo Raise_ItemEvent_Error
                            End If
                        End If
                        '//측정항목A Min
                        If Val(oDS_PS_QM020H.GetValue("U_M_A_min", 0)) <> 0 Then
                            VALUE_MIN = Val(oDS_PS_QM020H.GetValue("U_M_A_min", 0))
                            SPEC_MIN = Val(oForm01.Items("A_Val").Specific.VALUE) - Val(oForm01.Items("A_Mi").Specific.VALUE)
                            SPEC_MAX = Val(oForm01.Items("A_Val").Specific.VALUE) + Val(oForm01.Items("A_Pl").Specific.VALUE)
                            If VALUE_MIN < SPEC_MIN Or VALUE_MIN > SPEC_MAX Then
                                Sbo_Application.MessageBox ("A 의 Min 측정치가 SPEC 규격 범위를 벗어났습니다.")
                                GoTo Raise_ItemEvent_Error
                            End If
                        End If
                        '//측정항목A Max
                        If Val(oDS_PS_QM020H.GetValue("U_M_A_max", 0)) <> 0 Then
                            VALUE_MAX = Val(oDS_PS_QM020H.GetValue("U_M_A_max", 0))
                            SPEC_MIN = Val(oForm01.Items("A_Val").Specific.VALUE) - Val(oForm01.Items("A_Mi").Specific.VALUE)
                            SPEC_MAX = Val(oForm01.Items("A_Val").Specific.VALUE) + Val(oForm01.Items("A_Pl").Specific.VALUE)
                            If VALUE_MAX < SPEC_MIN Or VALUE_MAX > SPEC_MAX Then
                                Sbo_Application.MessageBox ("A 의 Max 측정치가 SPEC 규격 범위를 벗어났습니다.")
                                GoTo Raise_ItemEvent_Error
                            End If
                        End If
                        '//측정항목B Min
                        If Val(oDS_PS_QM020H.GetValue("U_M_B_min", 0)) <> 0 Then
                            VALUE_MIN = Val(oDS_PS_QM020H.GetValue("U_M_B_min", 0))
                            SPEC_MIN = Val(oForm01.Items("B_Val").Specific.VALUE) - Val(oForm01.Items("B_Mi").Specific.VALUE)
                            SPEC_MAX = Val(oForm01.Items("B_Val").Specific.VALUE) + Val(oForm01.Items("B_Pl").Specific.VALUE)
                            If VALUE_MIN < SPEC_MIN Or VALUE_MIN > SPEC_MAX Then
                                Sbo_Application.MessageBox ("B 의 Min 측정치가 SPEC 규격 범위를 벗어났습니다.")
                                GoTo Raise_ItemEvent_Error
                            End If
                        End If
                        '//측정항목B Max
                        If Val(oDS_PS_QM020H.GetValue("U_M_B_max", 0)) <> 0 Then
                            VALUE_MAX = Val(oDS_PS_QM020H.GetValue("U_M_B_max", 0))
                            SPEC_MIN = Val(oForm01.Items("B_Val").Specific.VALUE) - Val(oForm01.Items("B_Mi").Specific.VALUE)
                            SPEC_MAX = Val(oForm01.Items("B_Val").Specific.VALUE) + Val(oForm01.Items("B_Pl").Specific.VALUE)
                            If VALUE_MAX < SPEC_MIN Or VALUE_MAX > SPEC_MAX Then
                                Sbo_Application.MessageBox ("B 의 Max 측정치가 SPEC 규격 범위를 벗어났습니다.")
                                GoTo Raise_ItemEvent_Error
                            End If
                        End If
                        '//측정항목B1 Min
                        If Val(oDS_PS_QM020H.GetValue("U_M_B1_min", 0)) <> 0 Then
                            VALUE_MIN = Val(oDS_PS_QM020H.GetValue("U_M_B1_min", 0))
                            SPEC_MIN = Val(oForm01.Items("B1_Val").Specific.VALUE) - Val(oForm01.Items("B1_Mi").Specific.VALUE)
                            SPEC_MAX = Val(oForm01.Items("B1_Val").Specific.VALUE) + Val(oForm01.Items("B1_Pl").Specific.VALUE)
                            If VALUE_MIN < SPEC_MIN Or VALUE_MIN > SPEC_MAX Then
                                Sbo_Application.MessageBox ("B1 의 Min 측정치가 SPEC 규격 범위를 벗어났습니다.")
                                GoTo Raise_ItemEvent_Error
                            End If
                        End If
                        '//측정항목B1 Max
                        If Val(oDS_PS_QM020H.GetValue("U_M_B1_max", 0)) <> 0 Then
                            VALUE_MAX = Val(oDS_PS_QM020H.GetValue("U_M_B1_max", 0))
                            SPEC_MIN = Val(oForm01.Items("B1_Val").Specific.VALUE) - Val(oForm01.Items("B1_Mi").Specific.VALUE)
                            SPEC_MAX = Val(oForm01.Items("B1_Val").Specific.VALUE) + Val(oForm01.Items("B1_Pl").Specific.VALUE)
                            If VALUE_MAX < SPEC_MIN Or VALUE_MAX > SPEC_MAX Then
                                Sbo_Application.MessageBox ("B1 의 Max 측정치가 SPEC 규격 범위를 벗어났습니다.")
                                GoTo Raise_ItemEvent_Error
                            End If
                        End If
                        '//측정항목C
                        If Val(oDS_PS_QM020H.GetValue("U_M_C", 0)) <> 0 Then
                            VALUE = Val(oDS_PS_QM020H.GetValue("U_M_C", 0))
                            SPEC_MIN = Val(oForm01.Items("C_Val").Specific.VALUE) - Val(oForm01.Items("C_Mi").Specific.VALUE)
                            SPEC_MAX = Val(oForm01.Items("C_Val").Specific.VALUE) + Val(oForm01.Items("C_Pl").Specific.VALUE)
                            If VALUE < SPEC_MIN Or VALUE > SPEC_MAX Then
                                Sbo_Application.MessageBox ("C 의 측정치가 SPEC 규격 범위를 벗어났습니다.")
                                GoTo Raise_ItemEvent_Error
                            End If
                        End If
                        '//측정항목D
                        If Val(oDS_PS_QM020H.GetValue("U_M_D", 0)) <> 0 Then
                            VALUE = Val(oDS_PS_QM020H.GetValue("U_M_D", 0))
                            SPEC_MIN = Val(oForm01.Items("D_Val").Specific.VALUE) - Val(oForm01.Items("D_Mi").Specific.VALUE)
                            SPEC_MAX = Val(oForm01.Items("D_Val").Specific.VALUE) + Val(oForm01.Items("D_Pl").Specific.VALUE)
                            If VALUE < SPEC_MIN Or VALUE > SPEC_MAX Then
                                Sbo_Application.MessageBox ("D 의 측정치가 SPEC 규격 범위를 벗어났습니다.")
                                GoTo Raise_ItemEvent_Error
                            End If
                        End If
    
                        '//측정항목E-E1
                        If Val(oDS_PS_QM020H.GetValue("U_M_EE1", 0)) <> 0 Then
                            VALUE = Val(oDS_PS_QM020H.GetValue("U_M_EE1", 0))
                            SPEC_MAX = Val(oForm01.Items("EE1_Val").Specific.VALUE)
                            If VALUE > SPEC_MAX Then
                                Sbo_Application.MessageBox ("E-E1 의 측정치가 SPEC 규격 범위를 벗어났습니다.")
                                GoTo Raise_ItemEvent_Error
                            End If
                        End If
                        '//측정항목F
                        If Val(oDS_PS_QM020H.GetValue("U_M_F", 0)) <> 0 Then
                            VALUE = Val(oDS_PS_QM020H.GetValue("U_M_F", 0))
                            If Val(oForm01.Items("F_Mi").Specific.VALUE) = 0 And Val(oForm01.Items("F_Pl").Specific.VALUE) = 0 Then
                                '//MAX 일때는 0과 MAX값 사이에 입력가능
                                If VALUE < 0 Or VALUE > Val(oForm01.Items("F_Val").Specific.VALUE) Then
                                    Sbo_Application.MessageBox ("F 의 측정치가 SPEC 규격 범위를 벗어났습니다.")
                                    GoTo Raise_ItemEvent_Error
                                End If
                            Else
                                SPEC_MIN = Val(oForm01.Items("F_Val").Specific.VALUE) - Val(oForm01.Items("F_Mi").Specific.VALUE)
                                SPEC_MAX = Val(oForm01.Items("F_Val").Specific.VALUE) + Val(oForm01.Items("F_Pl").Specific.VALUE)
    
                                If VALUE < SPEC_MIN Or VALUE > SPEC_MAX Then
                                    Sbo_Application.MessageBox ("F 의 측정치가 SPEC 규격 범위를 벗어났습니다.")
                                    GoTo Raise_ItemEvent_Error
                                End If
                            End If
                        End If
                        '//측정항목F
                        If Val(oDS_PS_QM020H.GetValue("U_M_F1", 0)) <> 0 Then
                            VALUE = Val(oDS_PS_QM020H.GetValue("U_M_F1", 0))
                            If Val(oForm01.Items("F_Mi").Specific.VALUE) = 0 And Val(oForm01.Items("F_Pl").Specific.VALUE) = 0 Then
                                If VALUE < 0 Or VALUE > Val(oForm01.Items("F1_Val").Specific.VALUE) Then
                                    Sbo_Application.MessageBox ("F1 의 측정치가 SPEC 규격 범위를 벗어났습니다.")
                                GoTo Raise_ItemEvent_Error
                                End If
                            Else
                                SPEC_MIN = Val(oForm01.Items("F1_Val").Specific.VALUE) - Val(oForm01.Items("F1_Mi").Specific.VALUE)
                                SPEC_MAX = Val(oForm01.Items("F1_Val").Specific.VALUE) + Val(oForm01.Items("F1_Pl").Specific.VALUE)
                                If VALUE < SPEC_MIN Or VALUE > SPEC_MAX Then
                                    Sbo_Application.MessageBox ("F1 의 측정치가 SPEC 규격 범위를 벗어났습니다.")
                                    GoTo Raise_ItemEvent_Error
                                End If
                            End If
                        End If
                        '//측정항목G
                        If Val(oDS_PS_QM020H.GetValue("U_M_G", 0)) <> 0 Then
                            VALUE = Val(oDS_PS_QM020H.GetValue("U_M_G", 0))
                            SPEC_MIN = Val(oForm01.Items("G_Val").Specific.VALUE) - Val(oForm01.Items("G_Mi").Specific.VALUE)
                            SPEC_MAX = Val(oForm01.Items("G_Val").Specific.VALUE) + Val(oForm01.Items("G_Pl").Specific.VALUE)
                            If VALUE < SPEC_MIN Or VALUE > SPEC_MAX Then
                                Sbo_Application.MessageBox ("G 의 측정치가 SPEC 규격 범위를 벗어났습니다.")
                                GoTo Raise_ItemEvent_Error
                            End If
                        End If
                        '//측정항목G1
                        If Val(oDS_PS_QM020H.GetValue("U_M_G1", 0)) <> 0 Then
                            VALUE = Val(oDS_PS_QM020H.GetValue("U_M_G1", 0))
                            SPEC_MIN = Val(oForm01.Items("G1_Val").Specific.VALUE) - Val(oForm01.Items("G1_Mi").Specific.VALUE)
                            SPEC_MAX = Val(oForm01.Items("G1_Val").Specific.VALUE) + Val(oForm01.Items("G1_Pl").Specific.VALUE)
                            If VALUE < SPEC_MIN Or VALUE > SPEC_MAX Then
                                Sbo_Application.MessageBox ("G1 의 측정치가 SPEC 규격 범위를 벗어났습니다.")
                                GoTo Raise_ItemEvent_Error
                            End If
                        End If
                        '//측정항목A-G1
                        If Val(oDS_PS_QM020H.GetValue("U_M_AG1", 0)) <> 0 Then
                            VALUE = Val(oDS_PS_QM020H.GetValue("U_M_AG1", 0))
                            SPEC_MIN = Val(oForm01.Items("AG1_Val").Specific.VALUE) - Val(oForm01.Items("AG1_Mi").Specific.VALUE)
                            SPEC_MAX = Val(oForm01.Items("AG1_Val").Specific.VALUE) + Val(oForm01.Items("AG1_Pl").Specific.VALUE)
                            If VALUE < SPEC_MIN Or VALUE > SPEC_MAX Then
                                Sbo_Application.MessageBox ("AG1 의 측정치가 SPEC 규격 범위를 벗어났습니다.")
                                GoTo Raise_ItemEvent_Error
                            End If
                        End If
                        '//측정항목H-H1
                        If Val(oDS_PS_QM020H.GetValue("U_M_HH1", 0)) <> 0 Then
                            VALUE = Val(oDS_PS_QM020H.GetValue("U_M_HH1", 0))
                            SPEC_MAX = Val(oForm01.Items("HH1_Val").Specific.VALUE)
                            If VALUE > SPEC_MAX Then
                                Sbo_Application.MessageBox ("H-H1 의 측정치가 SPEC 규격 범위를 벗어났습니다.")
                                GoTo Raise_ItemEvent_Error
                            End If
                        End If
                        '//측정항목I
                        If Val(oDS_PS_QM020H.GetValue("U_M_I", 0)) <> 0 Then
                            VALUE = Val(oDS_PS_QM020H.GetValue("U_M_I", 0))
                            SPEC_MIN = Val(oForm01.Items("I_Val").Specific.VALUE) - Val(oForm01.Items("I_Mi").Specific.VALUE)
                            SPEC_MAX = Val(oForm01.Items("I_Val").Specific.VALUE) + Val(oForm01.Items("I_Pl").Specific.VALUE)
                            If VALUE < SPEC_MIN Or VALUE > SPEC_MAX Then
                                Sbo_Application.MessageBox ("I 의 측정치가 SPEC 규격 범위를 벗어났습니다.")
                                GoTo Raise_ItemEvent_Error
                            End If
                        End If
                        '//측정항목I1
                        If Val(oDS_PS_QM020H.GetValue("U_M_I1", 0)) <> 0 Then
                            VALUE = Val(oDS_PS_QM020H.GetValue("U_M_I1", 0))
                            SPEC_MIN = Val(oForm01.Items("I1_Val").Specific.VALUE) - Val(oForm01.Items("I1_Mi").Specific.VALUE)
                            SPEC_MAX = Val(oForm01.Items("I1_Val").Specific.VALUE) + Val(oForm01.Items("I1_Pl").Specific.VALUE)
                            If VALUE < SPEC_MIN Or VALUE > SPEC_MAX Then
                                Sbo_Application.MessageBox ("I1 의 측정치가 SPEC 규격 범위를 벗어났습니다.")
                                GoTo Raise_ItemEvent_Error
                            End If
                        End If
                        '//측정항목J-J1
                        If Val(oDS_PS_QM020H.GetValue("U_M_JJ1", 0)) <> 0 Then
                            VALUE = Val(oDS_PS_QM020H.GetValue("U_M_JJ1", 0))
                            SPEC_MAX = Val(oForm01.Items("JJ1_Val").Specific.VALUE)
                            If VALUE > SPEC_MAX Then
                                Sbo_Application.MessageBox ("J-J1 의 측정치가 SPEC 규격 범위를 벗어났습니다.")
                                GoTo Raise_ItemEvent_Error
                            End If
                        End If
                        '//측정항목Edge Burr
                        If Val(oDS_PS_QM020H.GetValue("U_M_EdgeBu", 0)) <> 0 Then
                            VALUE = Val(oDS_PS_QM020H.GetValue("U_M_EdgeBu", 0))
                            SPEC_MAX = Val(oForm01.Items("EdgeBurr").Specific.VALUE)
                            If VALUE > SPEC_MAX Then
                                Sbo_Application.MessageBox ("Edge Burr 의 측정치가 SPEC 규격 범위를 벗어났습니다")
                                GoTo Raise_ItemEvent_Error
                            End If
                        End If
                        '//측정항목Camber
                        If Val(oDS_PS_QM020H.GetValue("U_M_Camber", 0)) <> 0 Then
                            VALUE = Val(oDS_PS_QM020H.GetValue("U_M_Camber", 0))
                            If VALUE > 0.5 Then
                                Sbo_Application.MessageBox ("Camber 의 측정치가 SPEC 규격 범위를 벗어났습니다")
                                GoTo Raise_ItemEvent_Error
                            End If
                        End If
                        '//측정항목Surface Roughness Min
                        If Val(oDS_PS_QM020H.GetValue("U_M_SR_min", 0)) <> 0 Then
                            VALUE_MIN = Val(oDS_PS_QM020H.GetValue("U_M_SR_min", 0))
                            SPEC_MAX = Val(oForm01.Items("Surface").Specific.VALUE)
                            If VALUE_MIN > SPEC_MAX Then
                                Sbo_Application.MessageBox ("surface roughness의 Min 측정치가 SPEC 규격 범위를 벗어났습니다")
                                GoTo Raise_ItemEvent_Error
                            End If
                        End If
                        '//측정항목Surface Roughness Max
                        If Val(oDS_PS_QM020H.GetValue("U_M_SR_max", 0)) <> 0 Then
                            VALUE_MAX = Val(oDS_PS_QM020H.GetValue("U_M_SR_max", 0))
                            SPEC_MAX = Val(oForm01.Items("Surface").Specific.VALUE)
                            If VALUE_MAX > SPEC_MAX Then
                                Sbo_Application.MessageBox ("surface roughness의 Max 측정치가 SPEC 규격 범위를 벗어났습니다")
                                GoTo Raise_ItemEvent_Error
                            End If
                        End If
                        '//측정항목Cross Bow
                        If Val(oDS_PS_QM020H.GetValue("U_M_CrossB", 0)) <> 0 Then
                            VALUE = Val(oDS_PS_QM020H.GetValue("U_M_CrossB", 0))
                            SPEC_MAX = Val(oForm01.Items("CrossBow").Specific.VALUE)
                            If VALUE > SPEC_MAX Then
                                Sbo_Application.MessageBox ("Cross Bow 의 측정치가 SPEC 규격 범위를 벗어났습니다")
                                GoTo Raise_ItemEvent_Error
                            End If
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
            Case et_FORM_UNLOAD: '//17
                RemoveForms oFormUniqueID01
                Set oForm01 = Nothing
                Set oMat01 = Nothing
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
'    Set ProgressBar01 = Nothing
    If ErrNum = 101 Then
        ErrNum = 0
        MDC_Com.MDC_GF_Message "Raise_ItemEvent_Error:" & Err.Number & " - " & Err.Description, "E"
        BubbleEvent = False
    End If
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
            Case "1285": '복원
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
        
    '-----------------------------------------------------------------------------------------------------------
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1285": '복원
            Case "1287": '복제
                oForm01.Freeze True
                oDS_PS_QM020H.setValue "U_OrdNum", 0, ""
                oDS_PS_QM020H.setValue "U_ItemCode", 0, ""
                oDS_PS_QM020H.setValue "U_ItemName", 0, ""
                oDS_PS_QM020H.setValue "U_LotNo", 0, ""
                oDS_PS_QM020H.setValue "U_InspDate", 0, Format(Now, "YYYYMMDD")
                oDS_PS_QM020H.setValue "U_InspPrsn", 0, ""
                oDS_PS_QM020H.setValue "U_CardCode", 0, ""
                oDS_PS_QM020H.setValue "U_CardName", 0, ""
                FormItemEnabled
                Call FormClear
                oForm01.Items("OrdNum").Click ct_Regular
                oForm01.Freeze False
            Case "1293": '행삭제
'                If oMat01.RowCount <> oMat01.VisualRowCount Then
'                    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
'                    '//맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
'                    '//이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
'                    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
'                    For i = 0 To oMat01.VisualRowCount - 1
'                        oMat01.Columns("LineId").Cells(i + 1).Specific.Value = i + 1
'                    Next i
'
'                    oMat01.FlushToDataSource
'                    oDS_PS_QM020L.RemoveRecord oDS_PS_QM020L.Size - 1   '// Mat1에 마지막라인(빈라인) 삭제
'                    oMat01.Clear
'                    oMat01.LoadFromDataSource
'                End If
            
            Case "1281": '찾기
                FormItemEnabled
                oForm01.Items("OrdNum").Click ct_Regular
                
            Case "1282": '추가
                Call FormItemEnabled
                Call FormClear
                oDS_PS_QM020H.setValue "U_InspDate", 0, Format(Now, "YYYYMMDD")
                oForm01.Items("OrdNum").Click ct_Regular
                
                oDS_PS_QM020H.setValue "U_CardCode", 0, Last_CardCode
                oDS_PS_QM020H.setValue "U_CardName", 0, Last_CardName
                oDS_PS_QM020H.setValue "U_InspPrsn", 0, Last_InspPrsn
                
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                FormItemEnabled
                FlushToItemValue "ItemCode"
'                If oMat01.VisualRowCount > 0 Then
'                    If oMat01.Columns("PP070No").Cells(oMat01.VisualRowCount).Specific.Value <> "" Then
'                        If oDS_PS_QM020H.GetValue("Status", 0) = "O" Then
'                            Matrix_AddRow 1, oMat01.RowCount, False
'                        End If
'                    End If
'                End If
  
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_MenuEvent_Error:
    MDC_Com.MDC_GF_Message "Raise_MenuEvent_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub Raise_RightClickEvent(ByRef FormUID As String, ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_RightClickEvent_Error
    If (eventInfo.BeforeAction = True) Then
        
    ElseIf (eventInfo.BeforeAction = False) Then
        '//작업
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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

Private Sub FormItemEnabled()

Dim oCombo          As SAPbouiCOM.ComboBox

    If oForm01.Mode = fm_FIND_MODE Then
        oForm01.Items("Code").Enabled = True
        oForm01.Items("CardCode").Enabled = True
        oForm01.Items("OrdNum").Enabled = True
        oForm01.Items("InspDate").Specific.VALUE = ""
    ElseIf oForm01.Mode = fm_ADD_MODE Then
        oForm01.Items("Code").Enabled = False
        oForm01.Items("CardCode").Enabled = True
        oForm01.Items("OrdNum").Enabled = True
        
        '외관검사
        Set oCombo = oForm01.Items("M_Visual").Specific
        oCombo.Select 0, psk_Index
        
        '포장상태
        Set oCombo = oForm01.Items("PACKYN").Specific
        oCombo.Select 0, psk_Index
        
        '합부판정
        Set oCombo = oForm01.Items("PASSYN").Specific
        oCombo.Select 0, psk_Index
        
    ElseIf oForm01.Mode = fm_OK_MODE Then
        oForm01.Items("CardCode").Enabled = True
        oForm01.Items("OrdNum").Enabled = False
    End If
    
    If Trim(oForm01.Items("TS_Gbn").Specific.VALUE) = "10" Then
        oForm01.Items("s6").Specific.Caption = "인장강도(kgf/㎡)"
    ElseIf Trim(oForm01.Items("TS_Gbn").Specific.VALUE) = "20" Then
        oForm01.Items("s6").Specific.Caption = "인장강도(N/㎟)"
    End If
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub FormClear()
On Error GoTo FormClear_Error
    Dim DocNum As String
    DocNum = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_QM020'", "")
    If DocNum = 0 Then
'        oForm01.Items("Code").Specific.Value = 1
        oDS_PS_QM020H.setValue "Code", 0, 1
    Else
'        oForm01.Items("Code").Specific.Value = DocNum
        oDS_PS_QM020H.setValue "Code", 0, DocNum
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormClear_Error:
    MDC_Com.MDC_GF_Message "FormClear_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub
'*******************************************************************
'// oPaneLevel ==> 0:All / 1:oForm.PaneLevel=1 / 2:oForm.PaneLevel=2
'*******************************************************************
Private Sub Matrix_AddRow(ByVal oMat As Integer, ByVal oRow As Long, Optional Insert_YN As Boolean)
On Error GoTo Matrix_AddRow_Error
        
    Select Case oMat
        Case 1: 'oMat01
'            If Insert_YN = False Then
'                oRow = oMat01.RowCount
'                oDS_PS_QM020L.InsertRecord (oRow)
'            End If
'            '수입내역
'            oDS_PS_QM020L.Offset = oRow
'            oDS_PS_QM020L.setValue "LineId", oRow, oRow + 1
'            oDS_PS_QM020L.setValue "U_PP070No", oRow, ""
'            oDS_PS_QM020L.setValue "U_ItemCode", oRow, ""
'            oDS_PS_QM020L.setValue "U_ItemName", oRow, ""
'            oDS_PS_QM020L.setValue "U_Size", oRow, ""
'            oDS_PS_QM020L.setValue "U_Mark", oRow, ""
'            oDS_PS_QM020L.setValue "U_Qty", oRow, ""
'            oDS_PS_QM020L.setValue "U_Weight", oRow, ""
'            oDS_PS_QM020L.setValue "U_DocDate", oRow, ""
'            oMat01.LoadFromDataSource
        
    End Select
Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Matrix_AddRow_Error:
    MDC_Com.MDC_GF_Message "Matrix_AddRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error
    Dim i&, DocNum$, LineId$
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim oRecordSet    As SAPbobsCOM.Recordset
        
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    '--------------------------------------------------------------
    'Header--------------------------------------------------------
    Select Case oUID
        Case "CardCode", "ItemCode"
            '납품처명
            sQry = "select cardname from ocrd where cardtype='C' and cardcode = '" & Trim(oDS_PS_QM020H.GetValue("U_CardCode", 0)) & "'"
            oRecordSet.DoQuery sQry
            oDS_PS_QM020H.setValue "U_CardName", 0, Trim(oRecordSet.Fields(0).VALUE)
            
            '검사사양서내역
            sQry = "select  U_TS_min,U_TS_max,U_Elong,U_Hard_min,U_Hard_max,U_C_Fe_min,U_C_Fe_max,U_C_P_min,U_C_P_max, U_Chem_Zn,U_Chem_Pb,U_Chem_Ni,"
            sQry = sQry & "U_A_Pl,U_A_Mi,U_A_Val,U_B_Pl,U_B_Mi,U_B_Val,U_B1_Pl,U_B1_Mi,U_B1_Val,U_C_Pl,U_C_Mi,U_C_Val,U_D_Pl,U_D_Mi,U_D_Val,U_EE1_Val,"
            sQry = sQry & "U_F_Pl,U_F_Mi,U_F_Val,U_F1_Pl,U_F1_Mi,U_F1_Val,U_G_Pl,U_G_Mi,U_G_Val,U_G1_Pl,U_G1_Mi,U_G1_Val,U_AG1_Pl,U_AG1_Mi,U_AG1_Val,U_HH1_Val,"
            sQry = sQry & "U_I_Pl,U_I_Mi,U_I_Val,U_I1_Pl,U_I1_Mi,U_I1_Val,U_JJ1_Val,U_EdgeBurr,U_Camber_C,U_Camber_M,U_Surface,U_CrossBow,U_CR_Val,U_CR_Pl,U_CR_Mi, U_BT_Val, U_TS_Gbn,U_Chem_Cu "
            sQry = sQry & "From [@PS_QM010H] "
            sQry = sQry & "where U_CardCode = '" & Trim(oDS_PS_QM020H.GetValue("U_CardCode", 0)) & "' "
            sQry = sQry & "and   U_ItemCode = '" & Trim(oDS_PS_QM020H.GetValue("U_ItemCode", 0)) & "'"
            oRecordSet.DoQuery sQry
            
            If Trim(oRecordSet.Fields("U_TS_Gbn").VALUE) = "10" Then
                oForm01.Items("s6").Specific.Caption = "인장강도(kgf/㎡)"
            Else
                oForm01.Items("s6").Specific.Caption = "인장강도(N/㎟)"
            End If
            
            oDS_PS_QM020H.setValue "U_TS_Gbn", 0, Trim(oRecordSet.Fields("U_TS_Gbn").VALUE)
            
            
            
            oForm01.Items("TS_min").Specific.VALUE = Trim(oRecordSet.Fields(0).VALUE)
            oForm01.Items("TS_max").Specific.VALUE = Trim(oRecordSet.Fields(1).VALUE)
            oForm01.Items("Elong").Specific.VALUE = Trim(oRecordSet.Fields(2).VALUE)
            oForm01.Items("Hard_min").Specific.VALUE = Trim(oRecordSet.Fields(3).VALUE)
            oForm01.Items("Hard_max").Specific.VALUE = Trim(oRecordSet.Fields(4).VALUE)
            oForm01.Items("C_Fe_min").Specific.VALUE = Trim(oRecordSet.Fields(5).VALUE)
            oForm01.Items("C_Fe_max").Specific.VALUE = Trim(oRecordSet.Fields(6).VALUE)
            oForm01.Items("C_P_min").Specific.VALUE = Trim(oRecordSet.Fields(7).VALUE)
            oForm01.Items("C_P_max").Specific.VALUE = Trim(oRecordSet.Fields(8).VALUE)
            oForm01.Items("Chem_Zn").Specific.VALUE = Trim(oRecordSet.Fields(9).VALUE)
            oForm01.Items("Chem_Pb").Specific.VALUE = Trim(oRecordSet.Fields(10).VALUE)
            oForm01.Items("Chem_Ni").Specific.VALUE = Trim(oRecordSet.Fields(11).VALUE)
            oForm01.Items("A_Pl").Specific.VALUE = Trim(oRecordSet.Fields(12).VALUE)
            oForm01.Items("A_Mi").Specific.VALUE = Trim(oRecordSet.Fields(13).VALUE)
            oForm01.Items("A_Val").Specific.VALUE = Trim(oRecordSet.Fields(14).VALUE)
            oForm01.Items("B_Pl").Specific.VALUE = Trim(oRecordSet.Fields(15).VALUE)
            oForm01.Items("B_Mi").Specific.VALUE = Trim(oRecordSet.Fields(16).VALUE)
            oForm01.Items("B_Val").Specific.VALUE = Trim(oRecordSet.Fields(17).VALUE)
            oForm01.Items("B1_Pl").Specific.VALUE = Trim(oRecordSet.Fields(18).VALUE)
            oForm01.Items("B1_Mi").Specific.VALUE = Trim(oRecordSet.Fields(19).VALUE)
            oForm01.Items("B1_Val").Specific.VALUE = Trim(oRecordSet.Fields(20).VALUE)
            oForm01.Items("C_Pl").Specific.VALUE = Trim(oRecordSet.Fields(21).VALUE)
            oForm01.Items("C_Mi").Specific.VALUE = Trim(oRecordSet.Fields(22).VALUE)
            oForm01.Items("C_Val").Specific.VALUE = Trim(oRecordSet.Fields(23).VALUE)
            oForm01.Items("D_Pl").Specific.VALUE = Trim(oRecordSet.Fields(24).VALUE)
            oForm01.Items("D_Mi").Specific.VALUE = Trim(oRecordSet.Fields(25).VALUE)
            oForm01.Items("D_Val").Specific.VALUE = Trim(oRecordSet.Fields(26).VALUE)
            oForm01.Items("EE1_Val").Specific.VALUE = Trim(oRecordSet.Fields(27).VALUE)
            oForm01.Items("F_Pl").Specific.VALUE = Trim(oRecordSet.Fields(28).VALUE)
            oForm01.Items("F_Mi").Specific.VALUE = Trim(oRecordSet.Fields(29).VALUE)
            oForm01.Items("F_Val").Specific.VALUE = Trim(oRecordSet.Fields(30).VALUE)
            oForm01.Items("F1_Pl").Specific.VALUE = Trim(oRecordSet.Fields(31).VALUE)
            oForm01.Items("F1_Mi").Specific.VALUE = Trim(oRecordSet.Fields(32).VALUE)
            oForm01.Items("F1_Val").Specific.VALUE = Trim(oRecordSet.Fields(33).VALUE)
            oForm01.Items("G_Pl").Specific.VALUE = Trim(oRecordSet.Fields(34).VALUE)
            oForm01.Items("G_Mi").Specific.VALUE = Trim(oRecordSet.Fields(35).VALUE)
            oForm01.Items("G_Val").Specific.VALUE = Trim(oRecordSet.Fields(36).VALUE)
            oForm01.Items("G1_Pl").Specific.VALUE = Trim(oRecordSet.Fields(37).VALUE)
            oForm01.Items("G1_Mi").Specific.VALUE = Trim(oRecordSet.Fields(38).VALUE)
            oForm01.Items("G1_Val").Specific.VALUE = Trim(oRecordSet.Fields(39).VALUE)
            oForm01.Items("AG1_Pl").Specific.VALUE = Trim(oRecordSet.Fields(40).VALUE)
            oForm01.Items("AG1_Mi").Specific.VALUE = Trim(oRecordSet.Fields(41).VALUE)
            oForm01.Items("AG1_Val").Specific.VALUE = Trim(oRecordSet.Fields(42).VALUE)
            oForm01.Items("HH1_Val").Specific.VALUE = Trim(oRecordSet.Fields(43).VALUE)
            oForm01.Items("I_Pl").Specific.VALUE = Trim(oRecordSet.Fields(44).VALUE)
            oForm01.Items("I_Mi").Specific.VALUE = Trim(oRecordSet.Fields(45).VALUE)
            oForm01.Items("I_Val").Specific.VALUE = Trim(oRecordSet.Fields(46).VALUE)
            oForm01.Items("I1_Pl").Specific.VALUE = Trim(oRecordSet.Fields(47).VALUE)
            oForm01.Items("I1_Mi").Specific.VALUE = Trim(oRecordSet.Fields(48).VALUE)
            oForm01.Items("I1_Val").Specific.VALUE = Trim(oRecordSet.Fields(49).VALUE)
            oForm01.Items("JJ1_Val").Specific.VALUE = Trim(oRecordSet.Fields(50).VALUE)
            oForm01.Items("EdgeBurr").Specific.VALUE = Trim(oRecordSet.Fields(51).VALUE)
            oForm01.Items("Camber_C").Specific.VALUE = Trim(oRecordSet.Fields(52).VALUE)
            oForm01.Items("Camber_M").Specific.VALUE = Trim(oRecordSet.Fields(53).VALUE)
            oForm01.Items("Surface").Specific.VALUE = Trim(oRecordSet.Fields(54).VALUE)
            oForm01.Items("CrossBow").Specific.VALUE = Trim(oRecordSet.Fields(55).VALUE)
            oForm01.Items("CR_Val").Specific.VALUE = Trim(oRecordSet.Fields(56).VALUE)
            oForm01.Items("CR_Pl").Specific.VALUE = Trim(oRecordSet.Fields(57).VALUE)
            oForm01.Items("CR_Mi").Specific.VALUE = Trim(oRecordSet.Fields(58).VALUE)
            oForm01.Items("BT_Val").Specific.VALUE = Trim(oRecordSet.Fields(59).VALUE) 'Bake Test 측정치 추가(2012.11.26 송명규)
            
            oForm01.Items("Chem_Cu").Specific.VALUE = Trim(oRecordSet.Fields("U_Chem_Cu").VALUE) '화학성분 Cu 추가(2014.2.21 노근용)
            
        Case "OrdNum"
            sQry = "select a.U_ItemCode,a.U_ItemName,b.U_BatchNum from [@PS_PP030H] a inner join  [@PS_PP030L] b on a.docentry=b.docentry "
            sQry = sQry & "where U_OrdNum = '" & Trim(oDS_PS_QM020H.GetValue("U_OrdNum", 0)) & "'"
            oRecordSet.DoQuery sQry
            oDS_PS_QM020H.setValue "U_ItemCode", 0, Trim(oRecordSet.Fields(0).VALUE)
            oDS_PS_QM020H.setValue "U_ItemName", 0, Trim(oRecordSet.Fields(1).VALUE)
            oDS_PS_QM020H.setValue "U_LotNo", 0, Left(Trim(oRecordSet.Fields(2).VALUE), 8)
            
            '기존저장된 Data select
            sQry = "select TOP 1 U_M_TS_min,U_M_TS_max,U_M_El_min,U_M_El_max,U_M_Hd_min,U_M_Hd_max,U_M_Ch_Fe,U_M_Ch_P, U_M_Ch_Cu From [@PS_QM020H] "
            sQry = sQry & "where U_ItemCode = '" & Trim(oDS_PS_QM020H.GetValue("U_ItemCode", 0)) & "' "
            sQry = sQry & "and U_LotNo = '" & Trim(oDS_PS_QM020H.GetValue("U_LotNo", 0)) & "'"
            oRecordSet.DoQuery sQry
            oDS_PS_QM020H.setValue "U_M_TS_min", 0, Trim(oRecordSet.Fields(0).VALUE)
            oDS_PS_QM020H.setValue "U_M_TS_max", 0, Trim(oRecordSet.Fields(1).VALUE)
            oDS_PS_QM020H.setValue "U_M_El_min", 0, Trim(oRecordSet.Fields(2).VALUE)
            oDS_PS_QM020H.setValue "U_M_El_max", 0, Trim(oRecordSet.Fields(3).VALUE)
            oDS_PS_QM020H.setValue "U_M_Hd_min", 0, Trim(oRecordSet.Fields(4).VALUE)
            oDS_PS_QM020H.setValue "U_M_Hd_max", 0, Trim(oRecordSet.Fields(5).VALUE)
            oDS_PS_QM020H.setValue "U_M_Ch_Fe", 0, Trim(oRecordSet.Fields(6).VALUE)
            oDS_PS_QM020H.setValue "U_M_Ch_P", 0, Trim(oRecordSet.Fields(7).VALUE)
            oDS_PS_QM020H.setValue "U_M_Ch_Cu", 0, Trim(oRecordSet.Fields(8).VALUE)
            
            '원소재성분이 입력[PS_QM030]된 Data Select
            sQry = "select Top 1 U_ChemC_Fe, U_ChemC_P, U_ChemC_Cu from [OBTN] "
            sQry = sQry & "where DistNumber like '" & Trim(oDS_PS_QM020H.GetValue("U_LotNo", 0)) & "' + '%'"
            oRecordSet.DoQuery sQry
            If oDS_PS_QM020H.GetValue("U_M_Ch_Fe", 0) = 0 Then
                oDS_PS_QM020H.setValue "U_M_Ch_Fe", 0, Trim(oRecordSet.Fields(0).VALUE)
            End If
            If oDS_PS_QM020H.GetValue("U_M_Ch_P", 0) = 0 Then
                oDS_PS_QM020H.setValue "U_M_Ch_P", 0, Trim(oRecordSet.Fields(1).VALUE)
            End If
            If oDS_PS_QM020H.GetValue("U_M_Ch_Cu", 0) = 0 Then
                oDS_PS_QM020H.setValue "U_M_Ch_Cu", 0, Trim(oRecordSet.Fields(2).VALUE)
            End If
            
    End Select
        
    '--------------------------------------------------------------
    'Line----------------------------------------------------------
    If oUID = "Mat01" Then
        Select Case oCol
        
        End Select
    End If
     
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FlushToItemValue_Error:
    MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum          As Integer
    Dim SPEC_MIN        As Currency
    Dim SPEC_MAX        As Currency
    Dim VALUE           As Currency
    Dim VALUE_MIN       As Currency
    Dim VALUE_MAX       As Currency
    
    ErrNum = 0

    '// Check
    Select Case True
        Case Trim(oDS_PS_QM020H.GetValue("U_OrdNum", 0)) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oDS_PS_QM020H.GetValue("U_ItemCode", 0)) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
'        Case Trim(oDS_PS_QM020H.GetValue("U_LotNo", 0)) = ""
'            ErrNum = 3
'            GoTo HeaderSpaceLineDel_Error
        Case Trim(oDS_PS_QM020H.GetValue("U_InspDate", 0)) = ""
            ErrNum = 4
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oDS_PS_QM020H.GetValue("U_InspPrsn", 0)) = ""
            ErrNum = 5
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oDS_PS_QM020H.GetValue("U_CardCode", 0)) = ""
            ErrNum = 6
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oDS_PS_QM020H.GetValue("U_CardName", 0)) = ""
            ErrNum = 7
            GoTo HeaderSpaceLineDel_Error
    End Select
    
    If oForm01.Items("PASSYN").Specific.VALUE = "Y" Then
    '--------------------------------------------------------------------------------------------
        '//측청치의 유효값 검사
        '//인강강도 Min
        VALUE_MIN = Val(oDS_PS_QM020H.GetValue("U_M_TS_min", 0))
        SPEC_MIN = Val(oForm01.Items("TS_min").Specific.VALUE)
        SPEC_MAX = Val(oForm01.Items("TS_max").Specific.VALUE)
        If VALUE_MIN < SPEC_MIN Or VALUE_MIN > SPEC_MAX Then
            ErrNum = 10
            oForm01.Items("M_TS_min").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
    
        '//인강강도 Max
        VALUE_MAX = Val(oDS_PS_QM020H.GetValue("U_M_TS_max", 0))
        SPEC_MIN = Val(oForm01.Items("TS_min").Specific.VALUE)
        SPEC_MAX = Val(oForm01.Items("TS_max").Specific.VALUE)
        If VALUE_MAX < SPEC_MIN Or VALUE_MAX > SPEC_MAX Then
            ErrNum = 11
            oForm01.Items("M_TS_max").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
        '//연신율 Min
        VALUE_MIN = Val(oDS_PS_QM020H.GetValue("U_M_El_min", 0))
        SPEC_MIN = Val(oForm01.Items("Elong").Specific.VALUE)
        If VALUE_MIN < SPEC_MIN Then
            ErrNum = 12
            oForm01.Items("M_El_min").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
        '//연신율 Max
        VALUE_MAX = Val(oDS_PS_QM020H.GetValue("U_M_El_max", 0))
        SPEC_MIN = Val(oForm01.Items("Elong").Specific.VALUE)
        If VALUE_MAX < SPEC_MIN Then
            ErrNum = 13
            oForm01.Items("M_El_max").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
        '//경도 Min
        VALUE_MIN = Val(oDS_PS_QM020H.GetValue("U_M_Hd_min", 0))
        SPEC_MIN = Val(oForm01.Items("Hard_min").Specific.VALUE)
        SPEC_MAX = Val(oForm01.Items("Hard_max").Specific.VALUE)
        If VALUE_MIN < SPEC_MIN Or VALUE_MIN > SPEC_MAX Then
            ErrNum = 14
            oForm01.Items("M_Hd_min").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
        '//경도 Max
        VALUE_MAX = Val(oDS_PS_QM020H.GetValue("U_M_Hd_max", 0))
        SPEC_MIN = Val(oForm01.Items("Hard_min").Specific.VALUE)
        SPEC_MAX = Val(oForm01.Items("Hard_max").Specific.VALUE)
        If VALUE_MAX < SPEC_MIN Or VALUE_MAX > SPEC_MAX Then
            ErrNum = 15
            oForm01.Items("M_Hd_max").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
        '//Fe
        VALUE = Val(oDS_PS_QM020H.GetValue("U_M_Ch_Fe", 0))
        SPEC_MIN = Val(oForm01.Items("C_Fe_min").Specific.VALUE)
        SPEC_MAX = Val(oForm01.Items("C_Fe_max").Specific.VALUE)
        If VALUE < SPEC_MIN Or VALUE > SPEC_MAX Then
            ErrNum = 16
            oForm01.Items("M_Ch_Fe").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
        '//P
        VALUE = Val(oDS_PS_QM020H.GetValue("U_M_Ch_P", 0))
        SPEC_MIN = Val(oForm01.Items("C_P_min").Specific.VALUE)
        SPEC_MAX = Val(oForm01.Items("C_P_max").Specific.VALUE)
        If VALUE < SPEC_MIN Or VALUE > SPEC_MAX Then
            ErrNum = 17
            oForm01.Items("M_Ch_P").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
        '//측정항목A Min
        VALUE_MIN = Val(oDS_PS_QM020H.GetValue("U_M_A_min", 0))
        SPEC_MIN = Val(oForm01.Items("A_Val").Specific.VALUE) - Val(oForm01.Items("A_Mi").Specific.VALUE)
        SPEC_MAX = Val(oForm01.Items("A_Val").Specific.VALUE) + Val(oForm01.Items("A_Pl").Specific.VALUE)
        If VALUE_MIN < SPEC_MIN Or VALUE_MIN > SPEC_MAX Then
            ErrNum = 18
            oForm01.Items("M_A_min").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
        '//측정항목A Max
        VALUE_MAX = Val(oDS_PS_QM020H.GetValue("U_M_A_max", 0))
        SPEC_MIN = Val(oForm01.Items("A_Val").Specific.VALUE) - Val(oForm01.Items("A_Mi").Specific.VALUE)
        SPEC_MAX = Val(oForm01.Items("A_Val").Specific.VALUE) + Val(oForm01.Items("A_Pl").Specific.VALUE)
        If VALUE_MAX < SPEC_MIN Or VALUE_MAX > SPEC_MAX Then
            ErrNum = 19
            oForm01.Items("M_A_max").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
        '//측정항목B Min
        VALUE_MIN = Val(oDS_PS_QM020H.GetValue("U_M_B_min", 0))
        SPEC_MIN = Val(oForm01.Items("B_Val").Specific.VALUE) - Val(oForm01.Items("B_Mi").Specific.VALUE)
        SPEC_MAX = Val(oForm01.Items("B_Val").Specific.VALUE) + Val(oForm01.Items("B_Pl").Specific.VALUE)
        If VALUE_MIN < SPEC_MIN Or VALUE_MIN > SPEC_MAX Then
            ErrNum = 20
            oForm01.Items("M_B_min").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
        '//측정항목B Max
        VALUE_MAX = Val(oDS_PS_QM020H.GetValue("U_M_B_max", 0))
        SPEC_MIN = Val(oForm01.Items("B_Val").Specific.VALUE) - Val(oForm01.Items("B_Mi").Specific.VALUE)
        SPEC_MAX = Val(oForm01.Items("B_Val").Specific.VALUE) + Val(oForm01.Items("B_Pl").Specific.VALUE)
        If VALUE_MAX < SPEC_MIN Or VALUE_MAX > SPEC_MAX Then
            ErrNum = 21
            oForm01.Items("M_B_max").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
        '//측정항목B1 Min
        VALUE_MIN = Val(oDS_PS_QM020H.GetValue("U_M_B1_min", 0))
        SPEC_MIN = Val(oForm01.Items("B1_Val").Specific.VALUE) - Val(oForm01.Items("B1_Mi").Specific.VALUE)
        SPEC_MAX = Val(oForm01.Items("B1_Val").Specific.VALUE) + Val(oForm01.Items("B1_Pl").Specific.VALUE)
        If VALUE_MIN < SPEC_MIN Or VALUE_MIN > SPEC_MAX Then
            ErrNum = 22
            oForm01.Items("M_B1_min").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
        '//측정항목B1 Max
        VALUE_MAX = Val(oDS_PS_QM020H.GetValue("U_M_B1_max", 0))
        SPEC_MIN = Val(oForm01.Items("B1_Val").Specific.VALUE) - Val(oForm01.Items("B1_Mi").Specific.VALUE)
        SPEC_MAX = Val(oForm01.Items("B1_Val").Specific.VALUE) + Val(oForm01.Items("B1_Pl").Specific.VALUE)
        If VALUE_MAX < SPEC_MIN Or VALUE_MAX > SPEC_MAX Then
            ErrNum = 23
            oForm01.Items("M_B1_max").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
        '//측정항목C
        VALUE = Val(oDS_PS_QM020H.GetValue("U_M_C", 0))
        SPEC_MIN = Val(oForm01.Items("C_Val").Specific.VALUE) - Val(oForm01.Items("C_Mi").Specific.VALUE)
        SPEC_MAX = Val(oForm01.Items("C_Val").Specific.VALUE) + Val(oForm01.Items("C_Pl").Specific.VALUE)
        If VALUE < SPEC_MIN Or VALUE > SPEC_MAX Then
            ErrNum = 24
            oForm01.Items("M_C").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
        '//측정항목D
        VALUE = Val(oDS_PS_QM020H.GetValue("U_M_D", 0))
        SPEC_MIN = Val(oForm01.Items("D_Val").Specific.VALUE) - Val(oForm01.Items("D_Mi").Specific.VALUE)
        SPEC_MAX = Val(oForm01.Items("D_Val").Specific.VALUE) + Val(oForm01.Items("D_Pl").Specific.VALUE)
        If VALUE < SPEC_MIN Or VALUE > SPEC_MAX Then
            ErrNum = 25
            oForm01.Items("M_D").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
    
        '//측정항목E-E1
        VALUE = Val(oDS_PS_QM020H.GetValue("U_M_EE1", 0))
        SPEC_MAX = Val(oForm01.Items("EE1_Val").Specific.VALUE)
        If VALUE > SPEC_MAX Then
            ErrNum = 26
            oForm01.Items("M_EE1").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
        '//측정항목F
        'If Val(oDS_PS_QM020H.GetValue("U_M_F", 0)) <> 0 Then
        VALUE = Val(oDS_PS_QM020H.GetValue("U_M_F", 0))
        If Val(oForm01.Items("F_Mi").Specific.VALUE) = 0 And Val(oForm01.Items("F_Pl").Specific.VALUE) = 0 Then
            '//MAX 일때는 0과 MAX값 사이에 입력가능
            If VALUE < 0 Or VALUE > Val(oForm01.Items("F_Val").Specific.VALUE) Then
                ErrNum = 27
                oForm01.Items("M_F").Click ct_Regular
                GoTo HeaderSpaceLineDel_Error
            End If
        Else
            SPEC_MIN = Val(oForm01.Items("F_Val").Specific.VALUE) - Val(oForm01.Items("F_Mi").Specific.VALUE)
            SPEC_MAX = Val(oForm01.Items("F_Val").Specific.VALUE) + Val(oForm01.Items("F_Pl").Specific.VALUE)
            
            If VALUE < SPEC_MIN Or VALUE > SPEC_MAX Then
                ErrNum = 27
                oForm01.Items("M_F").Click ct_Regular
                GoTo HeaderSpaceLineDel_Error
            End If
        End If
        'End If
     '//측정항목F
        VALUE = Val(oDS_PS_QM020H.GetValue("U_M_F1", 0))
        If Val(oForm01.Items("F_Mi").Specific.VALUE) = 0 And Val(oForm01.Items("F_Pl").Specific.VALUE) = 0 Then
            If VALUE < 0 Or VALUE > Val(oForm01.Items("F1_Val").Specific.VALUE) Then
                ErrNum = 28
                oForm01.Items("M_F1").Click ct_Regular
                GoTo HeaderSpaceLineDel_Error
            End If
        Else
            SPEC_MIN = Val(oForm01.Items("F1_Val").Specific.VALUE) - Val(oForm01.Items("F1_Mi").Specific.VALUE)
            SPEC_MAX = Val(oForm01.Items("F1_Val").Specific.VALUE) + Val(oForm01.Items("F1_Pl").Specific.VALUE)
            If VALUE < SPEC_MIN Or VALUE > SPEC_MAX Then
                ErrNum = 28
                oForm01.Items("M_F1").Click ct_Regular
                GoTo HeaderSpaceLineDel_Error
            End If
        End If
        '//측정항목G
        VALUE = Val(oDS_PS_QM020H.GetValue("U_M_G", 0))
        SPEC_MIN = Val(oForm01.Items("G_Val").Specific.VALUE) - Val(oForm01.Items("G_Mi").Specific.VALUE)
        SPEC_MAX = Val(oForm01.Items("G_Val").Specific.VALUE) + Val(oForm01.Items("G_Pl").Specific.VALUE)
        If VALUE < SPEC_MIN Or VALUE > SPEC_MAX Then
            ErrNum = 29
            oForm01.Items("M_G").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
       '//측정항목G1
        VALUE = Val(oDS_PS_QM020H.GetValue("U_M_G1", 0))
        SPEC_MIN = Val(oForm01.Items("G1_Val").Specific.VALUE) - Val(oForm01.Items("G1_Mi").Specific.VALUE)
        SPEC_MAX = Val(oForm01.Items("G1_Val").Specific.VALUE) + Val(oForm01.Items("G1_Pl").Specific.VALUE)
        If VALUE < SPEC_MIN Or VALUE > SPEC_MAX Then
            ErrNum = 30
            oForm01.Items("M_G1").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
      '//측정항목A-G1
        VALUE = Val(oDS_PS_QM020H.GetValue("U_M_AG1", 0))
        SPEC_MIN = Val(oForm01.Items("AG1_Val").Specific.VALUE) - Val(oForm01.Items("AG1_Mi").Specific.VALUE)
        SPEC_MAX = Val(oForm01.Items("AG1_Val").Specific.VALUE) + Val(oForm01.Items("AG1_Pl").Specific.VALUE)
        If VALUE < SPEC_MIN Or VALUE > SPEC_MAX Then
            ErrNum = 31
            oForm01.Items("M_AG1").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
      '//측정항목H-H1
        VALUE = Val(oDS_PS_QM020H.GetValue("U_M_HH1", 0))
        SPEC_MAX = Val(oForm01.Items("HH1_Val").Specific.VALUE)
        If VALUE > SPEC_MAX Then
            ErrNum = 32
            oForm01.Items("M_HH1").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
     '//측정항목I
        VALUE = Val(oDS_PS_QM020H.GetValue("U_M_I", 0))
        SPEC_MIN = Val(oForm01.Items("I_Val").Specific.VALUE) - Val(oForm01.Items("I_Mi").Specific.VALUE)
        SPEC_MAX = Val(oForm01.Items("I_Val").Specific.VALUE) + Val(oForm01.Items("I_Pl").Specific.VALUE)
        If VALUE < SPEC_MIN Or VALUE > SPEC_MAX Then
            ErrNum = 33
            oForm01.Items("M_I").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
     '//측정항목I1
        VALUE = Val(oDS_PS_QM020H.GetValue("U_M_I1", 0))
        SPEC_MIN = Val(oForm01.Items("I1_Val").Specific.VALUE) - Val(oForm01.Items("I1_Mi").Specific.VALUE)
        SPEC_MAX = Val(oForm01.Items("I1_Val").Specific.VALUE) + Val(oForm01.Items("I1_Pl").Specific.VALUE)
        If VALUE < SPEC_MIN Or VALUE > SPEC_MAX Then
            ErrNum = 34
            oForm01.Items("M_I1").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
      '//측정항목J-J1
        VALUE = Val(oDS_PS_QM020H.GetValue("U_M_JJ1", 0))
        SPEC_MAX = Val(oForm01.Items("JJ1_Val").Specific.VALUE)
        If VALUE > SPEC_MAX Then
            ErrNum = 35
            oForm01.Items("M_JJ1").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
        '//측정항목Edge Burr
        VALUE = Val(oDS_PS_QM020H.GetValue("U_M_EdgeBu", 0))
        SPEC_MAX = Val(oForm01.Items("EdgeBurr").Specific.VALUE)
        If VALUE > SPEC_MAX Then
            ErrNum = 36
            oForm01.Items("M_EdgeBu").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
      '//측정항목Camber
        VALUE = Val(oDS_PS_QM020H.GetValue("U_M_Camber", 0))
        If VALUE > 0.5 Then
            ErrNum = 37
            oForm01.Items("M_Camber").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
       '//측정항목Surface Roughness Min
        If Val(oDS_PS_QM020H.GetValue("U_M_SR_min", 0)) <> 0 Then
            VALUE_MIN = Val(oDS_PS_QM020H.GetValue("U_M_SR_min", 0))
            SPEC_MAX = Val(oForm01.Items("Surface").Specific.VALUE)
            If VALUE_MIN > SPEC_MAX Then
                ErrNum = 38
                oForm01.Items("M_SR_min").Click ct_Regular
                GoTo HeaderSpaceLineDel_Error
            End If
        End If
        '//측정항목Surface Roughness Max
        If Val(oDS_PS_QM020H.GetValue("U_M_SR_max", 0)) <> 0 Then
            VALUE_MAX = Val(oDS_PS_QM020H.GetValue("U_M_SR_max", 0))
            SPEC_MAX = Val(oForm01.Items("Surface").Specific.VALUE)
            If VALUE_MAX > SPEC_MAX Then
                ErrNum = 39
                oForm01.Items("M_SR_max").Click ct_Regular
                GoTo HeaderSpaceLineDel_Error
            End If
        End If
        '//측정항목Cross Bow
        If Val(oDS_PS_QM020H.GetValue("U_M_CrossB", 0)) <> 0 Then
            VALUE = Val(oDS_PS_QM020H.GetValue("U_M_CrossB", 0))
            SPEC_MAX = Val(oForm01.Items("CrossBow").Specific.VALUE)
            If VALUE > SPEC_MAX Then
                ErrNum = 40
                oForm01.Items("M_CrossB").Click ct_Regular
                GoTo HeaderSpaceLineDel_Error
            End If
        End If
    
        '//Corner Radius -U_CR_min
        VALUE = Val(oDS_PS_QM020H.GetValue("U_M_CR_min", 0))
        If Val(oForm01.Items("CR_Pl").Specific.VALUE) = 0 And Val(oForm01.Items("CR_Mi").Specific.VALUE) = 0 Then
            If VALUE < 0 Or VALUE > Val(oForm01.Items("CR_Val").Specific.VALUE) Then
                ErrNum = 41
                oForm01.Items("M_CR_min").Click ct_Regular
                GoTo HeaderSpaceLineDel_Error
            End If
        Else
            SPEC_MIN = Val(oForm01.Items("CR_Val").Specific.VALUE) - Val(oForm01.Items("CR_Mi").Specific.VALUE)
            SPEC_MAX = Val(oForm01.Items("CR_Val").Specific.VALUE) + Val(oForm01.Items("CR_Pl").Specific.VALUE)
            If VALUE < SPEC_MIN Or VALUE > SPEC_MAX Then
                ErrNum = 41
                oForm01.Items("M_CR_min").Click ct_Regular
                GoTo HeaderSpaceLineDel_Error
            End If
        End If
    
        '//Corner Radius -U_CR_max
        VALUE = Val(oDS_PS_QM020H.GetValue("U_M_CR_max", 0))
        If Val(oForm01.Items("CR_Pl").Specific.VALUE) = 0 And Val(oForm01.Items("CR_Mi").Specific.VALUE) = 0 Then
            If VALUE < 0 Or VALUE > Val(oForm01.Items("CR_Val").Specific.VALUE) Then
                ErrNum = 41
                oForm01.Items("M_CR_max").Click ct_Regular
                GoTo HeaderSpaceLineDel_Error
            End If
        Else
            SPEC_MIN = Val(oForm01.Items("CR_Val").Specific.VALUE) - Val(oForm01.Items("CR_Mi").Specific.VALUE)
            SPEC_MAX = Val(oForm01.Items("CR_Val").Specific.VALUE) + Val(oForm01.Items("CR_Pl").Specific.VALUE)
            If VALUE < SPEC_MIN Or VALUE > SPEC_MAX Then
                ErrNum = 41
                oForm01.Items("M_CR_max").Click ct_Regular
                GoTo HeaderSpaceLineDel_Error
            End If
        End If
    '    If Val(oDS_PS_QM020H.GetValue("U_M_CR_min", 0)) <> 0 Then
    '        VALUE_MIN = Val(oDS_PS_QM020H.GetValue("U_M_CR_min", 0))
    '        SPEC_MIN = Val(oForm01.Items("CR_Val").Specific.VALUE) - Val(oForm01.Items("CR_Mi").Specific.VALUE)
    '        SPEC_MAX = Val(oForm01.Items("CR_Val").Specific.VALUE) + Val(oForm01.Items("CR_Pl").Specific.VALUE)
    '        If VALUE_MIN < SPEC_MIN Or VALUE_MIN > SPEC_MAX Then
    '            ErrNum = 41
    '            oForm01.Items("M_CR_min").Click ct_Regular
    '            GoTo HeaderSpaceLineDel_Error
    '        End If
    '    End If
    '    If Val(oDS_PS_QM020H.GetValue("U_M_CR_max", 0)) <> 0 Then
    '        VALUE_MAX = Val(oDS_PS_QM020H.GetValue("U_M_CR_max", 0))
    '        SPEC_MIN = Val(oForm01.Items("CR_Val").Specific.VALUE) - Val(oForm01.Items("CR_Mi").Specific.VALUE)
    '        SPEC_MAX = Val(oForm01.Items("CR_Val").Specific.VALUE) + Val(oForm01.Items("CR_Pl").Specific.VALUE)
    '        If VALUE_MAX < SPEC_MIN Or VALUE_MAX > SPEC_MAX Then
    '            ErrNum = 42
    '            oForm01.Items("M_CR_max").Click ct_Regular
    '            GoTo HeaderSpaceLineDel_Error
    '        End If
    '    End If
    '    '//Cu  '임시로 막음 확인후 주석풀어야함. 2014/02/21
    '    If Val(oDS_PS_QM020H.GetValue("U_M_Ch_Cu", 0)) <> 0 Then
    '        VALUE = Val(oDS_PS_QM020H.GetValue("U_M_Ch_Cu", 0))
    '        SPEC_MAX = Val(oForm01.Items("Chem_Cu").Specific.VALUE)
    '        If VALUE > SPEC_MAX Then
    '            ErrNum = 43
    '            oForm01.Items("M_Ch_Cu").Click ct_Regular
    '            GoTo HeaderSpaceLineDel_Error
    '        End If
    '    End If
    End If
    
    '--------------------------------------------------------------------------------------------
    '//추가완료 후 다시 자동으로 CardCode,CardName,InspPrsn을 보여주기 위해 미리 저장
    Last_CardCode = Trim(oDS_PS_QM020H.GetValue("U_CardCode", 0))
    Last_CardName = Trim(oDS_PS_QM020H.GetValue("U_CardName", 0))
    Last_InspPrsn = Trim(oDS_PS_QM020H.GetValue("U_InspPrsn", 0))

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "작지번호는 필수사항입니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "제품코드는 필수사항입니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "LOT-NO는 필수사항입니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "검사일은 필수사항입니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "검사자는 필수사항입니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "납품처은 필수사항입니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 7 Then
        MDC_Com.MDC_GF_Message "납품처명이 없습니다. 납품처코드를 확인하여 주십시오.", "E"
        
    ElseIf ErrNum = 10 Then
        MDC_Com.MDC_GF_Message "tensile strength 인장강도(kgf/㎡)의 Min 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 11 Then
        MDC_Com.MDC_GF_Message "tensile strength 인장강도(kgf/㎡)의 Max 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 12 Then
        MDC_Com.MDC_GF_Message "elongation 연신율(%)의 Min 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 13 Then
        MDC_Com.MDC_GF_Message "elongation 연신율(%)의 Max 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 14 Then
        MDC_Com.MDC_GF_Message "Hardness 경도(Hv)의 Min 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 15 Then
        MDC_Com.MDC_GF_Message "Hardness 경도(Hv)의 Max 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 16 Then
        MDC_Com.MDC_GF_Message "Fe 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 17 Then
        MDC_Com.MDC_GF_Message "P 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 18 Then
        MDC_Com.MDC_GF_Message "A 의 Min 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 19 Then
        MDC_Com.MDC_GF_Message "A 의 Max 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 20 Then
        MDC_Com.MDC_GF_Message "B 의 Min 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 21 Then
        MDC_Com.MDC_GF_Message "B 의 Max 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 22 Then
        MDC_Com.MDC_GF_Message "B1 의 Min 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 23 Then
        MDC_Com.MDC_GF_Message "B1 의 Max 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 24 Then
        MDC_Com.MDC_GF_Message "C 의 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 25 Then
        MDC_Com.MDC_GF_Message "D 의 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 26 Then
        MDC_Com.MDC_GF_Message "E-E1 의 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 27 Then
        MDC_Com.MDC_GF_Message "F 의 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 28 Then
        MDC_Com.MDC_GF_Message "F1 의 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 29 Then
        MDC_Com.MDC_GF_Message "G 의 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 30 Then
        MDC_Com.MDC_GF_Message "G1 의 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 31 Then
        MDC_Com.MDC_GF_Message "AG1 의 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 32 Then
        MDC_Com.MDC_GF_Message "H-H1 의 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 33 Then
        MDC_Com.MDC_GF_Message "I 의 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 34 Then
        MDC_Com.MDC_GF_Message "I1 의 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 35 Then
        MDC_Com.MDC_GF_Message "J-J1 의 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 36 Then
        MDC_Com.MDC_GF_Message "Edge Burr 의 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 37 Then
        MDC_Com.MDC_GF_Message "Camber 의 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 38 Then
        MDC_Com.MDC_GF_Message "surface roughness의 Min 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 39 Then
        MDC_Com.MDC_GF_Message "surface roughness의 Max 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 40 Then
        MDC_Com.MDC_GF_Message "Cross Bow 의 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 41 Then
        MDC_Com.MDC_GF_Message "corner radius 의 Min 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 42 Then
        MDC_Com.MDC_GF_Message "corner radius 의 Max 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 43 Then
        MDC_Com.MDC_GF_Message "화학성분Cu 의 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
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
    If oMat01.VisualRowCount = 1 Then '// MAT01에 값이 있는지 확인 (ErrorNumber : 1)
'        ErrNum = 1
'        GoTo MatrixSpaceLineDel_Error
    End If
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '//마지막 행 하나를 빼고 i=0부터 시작하므로 하나를 빼므로
    '//oMat01.RowCount - 2가 된다..반드시 들어 가야 하는 필수값을 확인한다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    If oMat01.VisualRowCount > 0 Then
'        For i = 0 To oMat01.VisualRowCount - 2      '// Mat1에 입력값이 올바르게 들어갔는지 확인 (ErrorNumber : 2)
'            oDS_PS_QM020L.Offset = i
'            If Trim(oDS_PS_QM020L.GetValue("U_PP070No", i)) = "" Then
'                ErrNum = 2
'                oMat01.Columns("PP070No").Cells(i + 1).Click ct_Regular
'                GoTo MatrixSpaceLineDel_Error
'            End If
'        Next i
    End If
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '//맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
    '//이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    If oMat01.VisualRowCount > 0 Then
'        oDS_PS_QM020L.RemoveRecord oDS_PS_QM020L.Size - 1   '// Mat1에 마지막라인(빈라인) 삭제
    End If
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '행을 삭제하였으니 DB데이터 소스를 다시 가져온다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    oMat01.LoadFromDataSource

    Set oRecordSet = Nothing
    MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
MatrixSpaceLineDel_Error:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "라인 데이터가 없습니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "벌크포장문서 번호는 필수입니다. 확인하여 주십시오.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function

Private Function Check_Exist() As Boolean
On Error GoTo Check_Exist_Error
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim ErrNum          As Integer
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sQry = "select * from [@PS_QM020H] where U_OrdNum = '" & Trim(oDS_PS_QM020H.GetValue("U_OrdNum", 0)) & "'"
    oRecordSet.DoQuery sQry

    If oRecordSet.RecordCount <> 0 Then
        ErrNum = 1
        GoTo Check_Exist_Error
    End If
    
    Set oRecordSet = Nothing
    Check_Exist = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Check_Exist_Error:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "이미 동일한 작지번호가 존재합니다. 확인하여 주십시오.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    Check_Exist = False
End Function


