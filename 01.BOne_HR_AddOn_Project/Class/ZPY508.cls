VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZPY508"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//  SAP MANAGE UI API 2004 SDK Sample
'//****************************************************************************
'//  File           : ZPY508.cls
'//  Module         : 檣餌婦葬>薑骯婦葬
'//  Desc           : 翱旎盪蹴 模菜奢薯 貲撮 蛔煙
'//  FormType       : 2000060508
'//  Create Date    : 2011.01.03
'//  Modified Date  :
'//  Creator        : Choi Dong Kwon
'//  Modifier       :
'//  Copyright  (c) Morning Data
'//****************************************************************************

Option Explicit

Public oFormUniqueID    As String
Public oForm            As SAPbouiCOM.Form
Private oDS_ZPY508H     As SAPbouiCOM.DBDataSource    '衛蝶蠱囀萄 ④渦
Private oDS_ZPY508L     As SAPbouiCOM.DBDataSource    '衛蝶蠱囀萄 塭檣
Private oMat1           As SAPbouiCOM.Matrix
Private Last_Item       As String                     '贗楚蝶縑憮 摹鷗и 葆雖虞 嬴檜蠱 Uid高
Private Col_Last_Uid    As String                     '葆雖虞嬴檜蠱檜 詭お葛蝶橾唳辦縑 葆雖虞 摹鷗脹 Col曖 Uid高
Private Col_Last_Row    As Long                       '葆雖虞嬴檜蠱檜 詭お葛蝶橾唳辦縑 葆雖虞 摹鷗脹 Row高
Private oOLDCHK         As String

'*******************************************************************
' .srf だ橾煎睡攪 イ擊 煎萄и棻
'*******************************************************************
Public Sub LoadForm(Optional JSNYER$, Optional MSTCOD$, Optional CLTCOD$)
On Error GoTo LoadForm_Error
    Dim i           As Long
    Dim oXmlDoc             As New MSXML2.DOMDocument

    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\ZPY508.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = _
        oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    
    '天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天
    '//罹楝偃曖 詭お葛蝶陛 ぎ唳辦縑 類啗賅曄籀歲 煎萄 腎紫煙 虜萇 賅曄
    '天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetTotalFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetTotalFormsCount * 10)

    Sbo_Application.LoadBatchActions oXmlDoc.xml
    
    oFormUniqueID = "ZPY508_" & GetTotalFormsCount
    
    'イ й渡
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)
    
    '天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天
    '鏽滓暮縑 イ擊 氬朝棻   **鏽滓暮檜塢 偃羹蒂 氬嬴 場朝 寡翮煎憮 罹晦憮朝 �側瘓音Ц轀� 氈朝 イ擊 氬堅 氈棻
    '天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天
    AddForms Me, oFormUniqueID, "ZPY508"
    oForm.SupportedModes = -1
    oForm.Mode = fm_FIND_MODE
    
    '//////////////////////////////////////////////////////////////////////////////
    '***************************************************************
    '�飛橝偽�(�飛橦□� 嶸橾酈高擊 氬堅 氈朝 嬴檜蠱曖 Uid高)
    oForm.DataBrowser.BrowseBy = "DocNum"
    '***************************************************************
    '//////////////////////////////////////////////////////////////////////////////
    oForm.Freeze True
    CreateItems
    
    oForm.EnableMenu ("1293"), True  '/ ч餉薯
    oForm.EnableMenu ("1283"), True  '/ 薯剪
    oForm.EnableMenu ("1284"), False  '/ 鏃模
    
    
    If JSNYER <> "" Then
        Call ShowSource(JSNYER, MSTCOD, CLTCOD)
    End If
    
    oForm.Freeze False
    oForm.Update
    'oForm.Visible = True

    Set oXmlDoc = Nothing
    Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
LoadForm_Error:
    Set oXmlDoc = Nothing
    Sbo_Application.StatusBar.SetText "Form_Load Error:" & Err.Description, bmt_Short, smt_Error
    If (oForm Is Nothing) = False Then
        oForm.Freeze False
        Set oForm = Nothing
    End If
End Sub
'*******************************************************************
'// ItemEventHander
'*******************************************************************
Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)

    Dim sQry        As String
    Dim i           As Long
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns     As SAPbouiCOM.Columns
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo Raise_FormItemEvent_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case pval.EventType
'et_ITEM_PRESSED''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_ITEM_PRESSED
            If pval.BeforeAction Then
                If pval.ItemUID = "1" Then
                        '天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天
                        '//蹺陛塽 機等檜衛縑
                        '天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天
                        If oForm.Mode = fm_ADD_MODE Or oForm.Mode = fm_UPDATE_MODE Then
                            If oForm.Mode = fm_ADD_MODE Then
                                If MDC_SetMod.Value_ChkYn("[@ZPY508H]", "U_JSNYER", "'" & oForm.Items("JSNYER").Specific.String & "'", " AND U_MSTCOD = '" & oForm.Items("MSTCOD").Specific.String & "'") = False Then
                                    Sbo_Application.StatusBar.SetText "檜嘐 盪濰腎橫螳 氈朝 ④渦曖 頂辨婁 橾纂м棲棻.", bmt_Short, smt_Error
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            End If
                        If Trim$(oDS_ZPY508H.GetValue("U_ENDCHK", 0)) = "Y" And Trim$(oOLDCHK) = "Y" Then
                            Sbo_Application.StatusBar.SetText "濡旎 濠猿殮棲棻. �挫恉狤� 輿褊衛螃.", bmt_Short, smt_Error
                             BubbleEvent = False
                             Exit Sub
                        ElseIf MatrixSpaceLineDel = False Then
                                BubbleEvent = False
                            End If
                        End If
                ElseIf pval.ItemUID = "CBtn1" And oForm.Items("MSTCOD").Enabled = True Then '/ ChooseBtn餌錳葬蝶お
                    oForm.Items("MSTCOD").CLICK ct_Regular
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                End If
            Else
                If pval.ItemUID = "1" And pval.ActionSuccess = True And oForm.Mode = fm_ADD_MODE Then
                    oForm.Mode = fm_OK_MODE
                    Call Sbo_Application.ActivateMenuItem("1282")
                End If
            End If
'et_CLICK''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_CLICK
            If pval.BeforeAction = True And pval.ItemUID <> "1000001" And pval.ItemUID <> "2" And oForm.Mode <> fm_FIND_MODE Then
                If Last_Item = "MSTCOD" Then
                    If MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" & oForm.Items(Last_Item).Specific.String & "'", "") = True _
                    And oForm.Items(Last_Item).Specific.String <> "" And Last_Item <> pval.ItemUID Then
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
            End If
            If pval.FormUID = oForm.uniqueID And pval.BeforeAction = True And Last_Item = "Mat1" And Col_Last_Uid = "Col1" And Col_Last_Row > 0 _
              And (Col_Last_Uid <> pval.ColUID Or Col_Last_Row <> pval.Row) And pval.ItemUID <> "1000001" And pval.ItemUID <> "2" Then
                If Col_Last_Row > oMat1.VisualRowCount Then
                    Exit Sub
                End If
            End If
'et_VALIDATE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_VALIDATE
            If pval.BeforeAction = False And pval.ItemChanged = True Then
                If (pval.ItemUID = "MSTCOD" Or pval.ItemUID = "JSNYER") Then
                    FlushToItemValue pval.ItemUID
                ElseIf pval.ItemUID = "Mat1" And (pval.ColUID = "Col5") Then
                    FlushToItemValue pval.ColUID, pval.Row
                End If
            End If
        
'et_KEY_DOWN''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_KEY_DOWN
            '天天天天天天天天天天天天天天天天天天天天天天天天天天
            '//蹺陛賅萄縑憮 囀萄檜漸お陛 囀萄縑憮 橾橫 陬擊陽
            '天天天天天天天天天天天天天天天天天天天天天天天天天天
             If pval.BeforeAction = True And pval.ItemUID = "MSTCOD" And pval.CharPressed = 9 And pval.FormMode <> fm_FIND_MODE Then
                If MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" & oForm.Items(pval.ItemUID).Specific.String & "'", "") = True Then
                    oForm.Items("MSTCOD").CLICK ct_Regular
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                Else
                    If oMat1.RowCount > 0 Then
                        oMat1.Columns("Col5").Cells(oMat1.VisualRowCount).CLICK ct_Regular
                        BubbleEvent = False
                    End If
                End If
             End If
'et_GOT_FOCUS''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_GOT_FOCUS
            If Last_Item = "Mat1" Then
                If pval.Row > 0 Then
                    Last_Item = pval.ItemUID
                    Col_Last_Row = pval.Row
                    Col_Last_Uid = pval.ColUID
                End If
            Else
                Last_Item = pval.ItemUID
                Col_Last_Row = 0
                Col_Last_Uid = ""
            End If
'et_FORM_UNLOAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_FORM_UNLOAD
            '天天天天天天天天天天天天天天天天天天天天天天天天天天
            '鏽滓暮縑憮 餉薯塽 賅萇 詭賅葬 薯剪
            '天天天天天天天天天天天天天天天天天天天天天天天天天天
            If pval.BeforeAction = False Then
               RemoveForms oFormUniqueID
                Set oForm = Nothing
                Set oDS_ZPY508H = Nothing
                Set oDS_ZPY508L = Nothing
                Set oMat1 = Nothing
            End If
'et_MATRIX_LOAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_MATRIX_LOAD:
            If pval.BeforeAction = False Then
                FormItemEnabled
                Call Matrix_AddRow(oMat1.VisualRowCount)
            End If
            
    End Select
    
    Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Raise_FormItemEvent_Error:
    Sbo_Application.StatusBar.SetText "Raise_FormItemEvent_Error:", bmt_Short, smt_Error
End Sub
'*******************************************************************
'// MenuEventHander
'*******************************************************************
Public Sub Raise_FormMenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
    Dim i&
    
    If pval.BeforeAction = True Then
        Select Case pval.MenuUID
            Case "1283" '/ 薯剪
                If Trim$(oDS_ZPY508H.GetValue("U_ENDCHK", 0)) = "Y" Then
                    Sbo_Application.StatusBar.SetText "濡旎 濠猿殮棲棻. �挫恉狤� 輿褊衛螃.", bmt_Short, smt_Error
                    BubbleEvent = False
                    Exit Sub
                Else
                    If Sbo_Application.MessageBox("⑷營 �飛曈遛蹀�羹蒂 薯剪 ж衛啊蝗棲梱? 犒掘й 熱 橈蝗棲棻.", 2, "Yes", "No") = 2 Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                End If
            Case "1288", "1289", "1290", "1291"
                Call AuthorityCheck(oForm, "CLTCOD", "@ZPY508H", "DocNum")      '//蕾樓濠 掏и縑 評艇 餌機濰 爾晦
                
            
            Case Else
                Exit Sub
        End Select
    Else
    
        Select Case pval.MenuUID
            Case "1287" '/ 犒薯
            Case "1283" '/ 薯剪
                 FormItemEnabled
            Case "1281", "1282":
                FormItemEnabled
                If pval.MenuUID = "1282" Then
                    FormClear
                    Call Matrix_AddRow(0, True)
                    oForm.Items("JSNYER").CLICK ct_Regular
                End If
            Case "1288" To "1291":
            Case "1293"
                If oMat1.RowCount <> oMat1.VisualRowCount Then
                    '天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天
                    '//裔葆雖虞縑 等檜攪蒂 餉薯ж朝 檜嶸朝 ч擊 蹺陛 й唳辦縑 蛤綠等檜攪模蝶縑
                    '//檜嘐 等檜攪陛 菟橫陛 氈晦 陽僥縑 盪濰衛縑朝 葆雖虞 ч(DB等檜攪 模蝶縑)擊 餉薯и棻
                    '天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天
                    For i = 0 To oMat1.VisualRowCount - 1
                        oMat1.Columns("Col0").Cells(i + 1).Specific.Value = i + 1
                    Next i
                    
                    oMat1.FlushToDataSource
                    oDS_ZPY508L.RemoveRecord oDS_ZPY508L.Size - 1   '// Mat1縑 葆雖虞塭檣(綴塭檣) 餉薯
                    oMat1.Clear
                    oMat1.LoadFromDataSource
                    
                End If
                Call FlushToItemValue("Col3", 1)
        End Select
    End If
End Sub


Public Sub Raise_FormDataEvent(ByRef FormUID As String, ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
    Dim i       As Long
    Dim sQry    As String
    Dim oCombo  As SAPbouiCOM.ComboBox
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    
    On Error GoTo Raise_FormDataEvent_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If (BusinessObjectInfo.BeforeAction = False) Then
        Select Case BusinessObjectInfo.EventType
            Case et_FORM_DATA_LOAD:     '//33
            Case et_FORM_DATA_ADD:      '//34
            Case et_FORM_DATA_UPDATE:   '//35
            Case et_FORM_DATA_DELETE:   '//36
        End Select

    End If
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Exit Sub
    
Raise_FormDataEvent_Error:
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "Raise_FormDataEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True

End Sub
'*******************************************************************
'
'*******************************************************************
Private Sub CreateItems()
On Error GoTo Error_Message
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim oColumn         As SAPbouiCOM.Column
    Dim oCheck          As SAPbouiCOM.CheckBox
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '//蛤綠等檜攪 模蝶 偃羹 й渡
    Set oDS_ZPY508H = oForm.DataSources.DBDataSources("@ZPY508H")
    Set oDS_ZPY508L = oForm.DataSources.DBDataSources("@ZPY508L")
    
    Set oMat1 = oForm.Items("Mat1").Specific
    
    '//餌機濰
    Set oCombo = oForm.Items("CLTCOD").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oCombo.ValidValues.Add "%", "瞪羹"
    oForm.Items("CLTCOD").DisplayDesc = True
    oCombo.Select "%", psk_ByValue
    
    '// 婦啗囀萄
    Set oColumn = oMat1.Columns("Col1")
    oColumn.ValidValues.Add "11", "黴霜翱旎-斬煎濠黴霜晝罹爾濰徹"
    oColumn.ValidValues.Add "12", "黴霜翱旎-婁з晦獎檣奢薯��"
    oColumn.ValidValues.Add "21", "翱旎盪蹴-偃檣翱旎盪蹴"
    oColumn.ValidValues.Add "22", "翱旎盪蹴-翱旎盪蹴"
    oColumn.ValidValues.Add "31", "輿鷗葆溼-羶擒盪蹴"
    oColumn.ValidValues.Add "32", "輿鷗葆溼-輿鷗羶擒謙м盪蹴"
    oColumn.ValidValues.Add "33", "輿鷗葆溼-濰晦輿鷗葆溼盪蹴"
    oColumn.ValidValues.Add "34", "輿鷗葆溼-斬煎濠輿鷗葆溼盪蹴"
    oColumn.ValidValues.Add "41", "濰晦輿衝⑽盪蹴"
 
    '// 旎懦晦婦
    Set oColumn = oMat1.Columns("Col2")
    sQry = "SELECT BankCode, BankName FROM [ODSC]"
    oRecordSet.DoQuery sQry
    Do Until oRecordSet.EOF
        oColumn.ValidValues.Add oRecordSet.Fields(0).Value, oRecordSet.Fields(1).Value
        oRecordSet.MoveNext
    Loop
 
    '// 陶殮翱離
    Set oColumn = oMat1.Columns("Col4")
    oColumn.ValidValues.Add "00", "п渡橈擠"
    oColumn.ValidValues.Add "01", "1喇離"
    oColumn.ValidValues.Add "02", "2喇離"
    oColumn.ValidValues.Add "03", "3喇離"
    
    '// 謙(瞪) 罹睡
    Set oColumn = oMat1.Columns("Col7")
    oColumn.ValOff = "N":   oColumn.ValOn = "Y"
 
   '/ Check 幗が
    Set oCheck = oForm.Items("ENDCHK").Specific
    oCheck.ValOff = "N": oCheck.ValOn = "Y"

    Set oCheck = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oRecordSet = Nothing
   Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    Set oCheck = Nothing
    Set oColumn = Nothing
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Sbo_Application.StatusBar.SetText "CreateItems Error:" & Err.Description, bmt_Short, smt_Error
End Sub

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long)
    Dim iRow        As Long
    Dim oMast       As ZPAY_g_EmpID
    Dim TOTAMT      As Double
    
    Select Case oUID
      Case "JSNYER"
            If Trim$(oForm.Items(oUID).Specific.String) <> "" Then
                ZPAY_GBL_JSNYER = oForm.Items(oUID).Specific.String
            Else
                oDS_ZPY508H.setValue "U_JSNYER", 0, ZPAY_GBL_JSNYER
            End If
            oForm.Items(oUID).Update
      Case "MSTCOD"
            If oForm.Items(oUID).Specific.String = "" Then
                oDS_ZPY508H.setValue "U_MSTCOD", 0, ""
                oDS_ZPY508H.setValue "U_MSTNAM", 0, ""
                oDS_ZPY508H.setValue "U_EmpID", 0, ""
                oDS_ZPY508H.setValue "U_CLTCOD", 0, ""
            Else
                oDS_ZPY508H.setValue "U_MSTCOD", 0, UCase$(oForm.Items(oUID).Specific.String)
                oMast = MDC_SetMod.Get_EmpID_InFo(oForm.Items(oUID).Specific.String)
                oDS_ZPY508H.setValue "U_MSTNAM", 0, oMast.MSTNAM
                oDS_ZPY508H.setValue "U_EmpID", 0, oMast.EmpID
                oDS_ZPY508H.setValue "U_CLTCOD", 0, oMast.CLTCOD
            End If
            oForm.Items("MSTNAM").Update
            oForm.Items("EmpID").Update
            oForm.Items("CLTCOD").Update
            oForm.Items(oUID).Update
    End Select
    '天天天天天天天天天天天天天天天天天天天天天天天天天天天
    Select Case oUID
        Case "Col5"
            oMat1.FlushToDataSource

            For iRow = 1 To oMat1.VisualRowCount
                TOTAMT = TOTAMT + Val(oDS_ZPY508L.GetValue("U_SAVAMT", iRow - 1))
            Next iRow
            oDS_ZPY508H.setValue "U_TOTAMT", 0, TOTAMT
            oForm.Items("TOTAMT").Update
    End Select
    If Left$(oUID, 3) = "Col" Then
        oMat1.FlushToDataSource
        If oRow = oMat1.RowCount And Val(oDS_ZPY508L.GetValue("U_SAVAMT", oRow - 1)) <> 0 Then
            Matrix_AddRow oRow
            oMat1.Columns("Col5").Cells(oRow).CLICK ct_Regular
        End If
    End If
End Sub
Private Sub FormClear()
    Dim DocNum&
    
    DocNum = MDC_SetMod.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'ZPY508'", "")
    
    If DocNum = 0 Then
        oForm.Items("DocNum").Specific.String = 1
    Else
        oForm.Items("DocNum").Specific.String = DocNum
    End If
    Call FlushToItemValue("JSNYER")
    
End Sub

Private Sub FormItemEnabled()
    If oForm.Mode = fm_FIND_MODE Then
        oForm.Items("JSNYER").Enabled = True
        oForm.Items("MSTCOD").Enabled = True
        oForm.Items("MSTNAM").Enabled = True
        oForm.Items("DocNum").Enabled = True
        oForm.Items("ENDCHK").Enabled = True
    ElseIf oForm.Mode = fm_ADD_MODE Then
        oForm.Items("JSNYER").Enabled = True
        oForm.Items("MSTCOD").Enabled = True
        oForm.Items("MSTNAM").Enabled = False
        oForm.Items("DocNum").Enabled = False
        oForm.Items("ENDCHK").Enabled = True
    ElseIf oForm.Mode = fm_OK_MODE Then
        oForm.Items("JSNYER").Enabled = False
        oForm.Items("MSTCOD").Enabled = False
        oForm.Items("MSTNAM").Enabled = False
        oForm.Items("DocNum").Enabled = False
        '// 喇葆馬脹匙擎 綠�側瘓�
        oOLDCHK = oDS_ZPY508H.GetValue("U_ENDCHK", 0)
        If MDC_SetMod.Get_ReData("U_ENDCHK", "U_JOBYER", "[@ZPY509L]", "'" & oDS_ZPY508H.GetValue("U_JSNYER", 0) & "'", " AND Code = '" & oDS_ZPY508H.GetValue("U_CLTCOD", 0) & "'") = "Y" Then
            oForm.Items("ENDCHK").Enabled = False
        Else
            oForm.Items("ENDCHK").Enabled = True
        End If
        
    End If
End Sub

Private Function MatrixSpaceLineDel() As Boolean
'天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天
'盪濰й 等檜攪曖 嶸�蕉瑰� 薄匐и棻
'天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天
On Error GoTo Error_Message
    Dim iRow        As Long
    Dim kRow        As Long
    Dim ErrNum      As Integer
    Dim Chk_Data    As String
    
    ErrNum = 0
    '/ ④渦睡碟 羹觼
    Select Case True
      Case Len(Trim$(oDS_ZPY508H.GetValue("U_JSNYER", 0))) <> 4
           ErrNum = 2
           GoTo Error_Message
      Case oDS_ZPY508H.GetValue("U_MSTCOD", 0) = ""
           ErrNum = 3
           GoTo Error_Message
      Case oDS_ZPY508H.GetValue("U_CLTCOD", 0) = ""
           ErrNum = 4
           GoTo Error_Message
           
    End Select
    
    '天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天
    '�飛暺騔� 詭お葛蝶縑 殮溘脹 頂辨擊 賅舒 蛤綠等檜攪模蝶煎 剩曹棻
    '天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天
    oMat1.FlushToDataSource
        
    If oMat1.RowCount = 1 Then          '// Mat1縑 高檜 氈朝雖 �挫� (ErrorNumber : 1)
        ErrNum = 1
        GoTo Error_Message
    End If
    
    '天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天
    '//葆雖虞 ч ж釭蒂 貍堅 i=0睡攪 衛濛ж嘎煎 ж釭蒂 貍嘎煎
    '//oMat1.RowCount - 2陛 脹棻..奩萄衛 菟橫 陛撿 ж朝 в熱高擊 �挫恉挬�
    '天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天
    For iRow = 0 To oMat1.VisualRowCount - 2      '// Mat1縑 殮溘高檜 螢夥腦啪 菟橫鬲朝雖 �挫� (ErrorNumber : 3)
        oDS_ZPY508L.Offset = iRow
        If Trim$(oDS_ZPY508L.GetValue("U_SAVGBN", iRow)) = "" Then
            ErrNum = 6
            GoTo Error_Message
        ElseIf Trim$(oDS_ZPY508L.GetValue("U_SAVCOD", iRow)) = "" Then
            ErrNum = 7
            GoTo Error_Message
        ElseIf Trim$(oDS_ZPY508L.GetValue("U_SAVNUM", iRow)) = "" Then
            ErrNum = 8
            oMat1.Columns("Col3").Cells(iRow + 1).CLICK ct_Regular
            GoTo Error_Message
        ElseIf Trim$(oDS_ZPY508L.GetValue("U_SAVGBN", iRow)) = "41" And Trim$(oDS_ZPY508L.GetValue("U_STYEAR", iRow)) = "00" Then
            ErrNum = 9
            oMat1.Columns("Col3").Cells(iRow + 1).CLICK ct_Regular
            GoTo Error_Message
        Else
            '天天天天天天天天天天天天天天天天天天天天天天天天
            '醞犒羹觼濛機
            '天天天天天天天天天天天天天天天天天天天天天天天天
            Chk_Data = Trim$(oDS_ZPY508L.GetValue("U_SAVGBN", iRow)) & Trim$(oDS_ZPY508L.GetValue("U_SAVCOD", iRow)) & _
                       Trim$(oDS_ZPY508L.GetValue("U_SAVNUM", iRow)) & Trim$(oDS_ZPY508L.GetValue("U_STYEAR", iRow))
            For kRow = iRow + 1 To oMat1.VisualRowCount - 2
                oDS_ZPY508L.Offset = kRow
                If Trim$(Chk_Data) = Trim$(oDS_ZPY508L.GetValue("U_SAVGBN", kRow)) & Trim$(oDS_ZPY508L.GetValue("U_SAVCOD", kRow)) & _
                                    Trim$(oDS_ZPY508L.GetValue("U_SAVNUM", kRow)) & Trim$(oDS_ZPY508L.GetValue("U_STYEAR", kRow)) Then
                    ErrNum = 5
                    oMat1.Columns("Col3").Cells(iRow + 1).CLICK ct_Regular
                    GoTo Error_Message
                End If
            Next kRow
        End If
        
        If Trim$(oDS_ZPY508L.GetValue("U_SAVGBN", iRow)) <> "41" And Trim$(oDS_ZPY508L.GetValue("U_STYEAR", iRow)) <> "00" Then
            oDS_ZPY508L.Offset = iRow
            oDS_ZPY508L.setValue "U_STYEAR", iRow, "00"
            oMat1.SetLineData (iRow + 1)
        End If
    Next iRow
    
    '天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天
    '//裔葆雖虞縑 等檜攪蒂 餉薯ж朝 檜嶸朝 ч擊 蹺陛 й唳辦縑 蛤綠等檜攪模蝶縑
    '//檜嘐 等檜攪陛 菟橫陛 氈晦 陽僥縑 盪濰衛縑朝 葆雖虞 ч(DB等檜攪 模蝶縑)擊 餉薯и棻
    '天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天
    oDS_ZPY508L.RemoveRecord oDS_ZPY508L.Size - 1   '// Mat1縑 葆雖虞塭檣(綴塭檣) 餉薯
    
    '天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天
    'ч擊 餉薯ж艘戲棲 DB等檜攪 模蝶蒂 棻衛 陛螳螞棻
    '天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天天
    oMat1.LoadFromDataSource
    
    MatrixSpaceLineDel = True
    Exit Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    Select Case ErrNum
    Case 1:     Sbo_Application.StatusBar.SetText "殮溘й 等檜攪陛 橈蝗棲棻. 殮溘ж罹 輿褊衛螃.", bmt_Short, smt_Error
    Case 2:     Sbo_Application.StatusBar.SetText "敝樓喇紫蒂 �挫恉狤� 輿褊衛螃.", bmt_Short, smt_Error
    Case 3:     Sbo_Application.StatusBar.SetText "曖猿綠 雖晝旎擋檜 0殮棲棻. 殮溘ж罹 輿褊衛螃.", bmt_Short, smt_Error
    Case 4:     Sbo_Application.StatusBar.SetText "濠餌囀萄朝 в熱殮棲棻. 殮溘ж罹 輿褊衛螃.", bmt_Short, smt_Error
    Case 5:     Sbo_Application.StatusBar.SetText "醞犒殮溘腎歷蝗棲棻. 盪蹴掘碟/旎懦晦婦/啗謝廓��/陶殮翱離滌煎 餵啗ж罹 殮溘ж罹 輿褊衛螃.", bmt_Short, smt_Error
    Case 6:     Sbo_Application.StatusBar.SetText "盪蹴掘碟擎 в熱殮棲棻. 摹鷗ж罹 輿褊衛螃.", bmt_Short, smt_Error
    Case 7:     Sbo_Application.StatusBar.SetText "旎懦晦婦擎 в熱殮棲棻. 摹鷗ж罹 輿褊衛螃.", bmt_Short, smt_Error
    Case 8:     Sbo_Application.StatusBar.SetText "啗謝廓�ㄣ� в熱殮棲棻. 殮溘ж罹 輿褊衛螃.", bmt_Short, smt_Error
    Case 9:     Sbo_Application.StatusBar.SetText "濰晦輿衝⑽ 盪蹴檣 唳辦 陶殮翱離蒂 1喇離~3喇離煎 摹鷗ж罹撿 м棲棻.", bmt_Short, smt_Error
    Case Else
        Sbo_Application.StatusBar.SetText "MatrixSpaceLineDel Error:" & Err.Number & Space$(10) & Err.Description, bmt_Short, smt_Error
    End Select
    MatrixSpaceLineDel = False
End Function

Private Sub ShowSource(JSNYER$, MSTCOD$, CLTCOD$)
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim DocEntry$
    Dim oMast           As ZPAY_g_EmpID
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    sQry = "SELECT DocEntry FROM [@ZPY508H]"
    sQry = sQry & "   WHERE U_JSNYER = N'" & JSNYER & "'"
    sQry = sQry & "   AND   U_MSTCOD = N'" & MSTCOD & "'"
    sQry = sQry & "   AND   U_CLTCOD = N'" & CLTCOD & "'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        Do Until oRecordSet.EOF
            DocEntry = oRecordSet.Fields(0).Value
            oRecordSet.MoveNext
        Loop
        oForm.Items("JSNYER").Specific.Value = JSNYER
        oForm.Items("MSTCOD").Specific.String = MSTCOD
        
        Call oForm.Items("CLTCOD").Specific.Select(CLTCOD, psk_ByValue)
        oForm.Items("DocNum").Specific.Value = DocEntry
        
        oForm.Items("DocNum").Update
        oMat1.LoadFromDataSource
        oForm.Update
        oForm.Items("1").CLICK ct_Regular
        
    Else
        oForm.Mode = fm_OK_MODE
        Call Sbo_Application.ActivateMenuItem("1282")
        
        Call oDS_ZPY508H.setValue("U_JSNYER", 0, JSNYER)
        Call oDS_ZPY508H.setValue("U_MSTCOD", 0, MSTCOD)
        Call oDS_ZPY508H.setValue("U_CLTCOD", 0, CLTCOD)
        oMast = MDC_SetMod.Get_EmpID_InFo(MSTCOD)
        Call oDS_ZPY508H.setValue("U_MSTNAM", 0, oMast.MSTNAM)
        Call oDS_ZPY508H.setValue("U_EmpID", 0, oMast.EmpID)
        
        oForm.Update
        
        Sbo_Application.SendKeys "{TAB}"
    End If
    
    Set oRecordSet = Nothing
    
End Sub

'*******************************************************************
'// oPaneLevel ==> 0:All / 1:oForm.PaneLevel=1 / 2:oForm.PaneLevel=2
'*******************************************************************
Private Sub Matrix_AddRow(ByVal oRow As Long, Optional Insert_YN As Boolean)
    If Insert_YN = False Then
        oDS_ZPY508L.InsertRecord (oRow)
    End If
    oDS_ZPY508L.Offset = oRow
    oDS_ZPY508L.setValue "U_LineNum", oRow, oRow + 1
    oDS_ZPY508L.setValue "U_SAVGBN", oRow, ""
    oDS_ZPY508L.setValue "U_SAVCOD", oRow, ""
    oDS_ZPY508L.setValue "U_JONGYN", oRow, "N"
    oDS_ZPY508L.setValue "U_SAVNAM", oRow, ""
    oDS_ZPY508L.setValue "U_SAVNUM", oRow, ""
    oDS_ZPY508L.setValue "U_STYEAR", oRow, "00"
    oDS_ZPY508L.setValue "U_SAVAMT", oRow, 0
    oDS_ZPY508L.setValue "U_SARAMT", oRow, 0
    oMat1.LoadFromDataSource
End Sub


