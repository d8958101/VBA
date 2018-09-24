Sub 整合()
    Dim sheetName As String
    sheetName = "BEFORE"
    Worksheets(sheetName).Activate
    '先新增欄位
    Dim InsertColumns As Variant
    InsertColumns = Array("Grade", "Customer", "Sales", "Pull In、Push Out(依Request Date)", "HUB", _
    "AIT P/N", "R", "Unit Price(NTD)", "Ordered Qty(K)", "Ordered Amt(K/NTD)", "Ordered Amt(K/USD)", _
    "本月已開發票QTY(K)", "月FCST", "月FCST", "月FCST", "月FCST", "月FCST", "月FCST", "BKG")
    For i = LBound(InsertColumns) To UBound(InsertColumns)
        'Sheet1.Columns("A:A").Insert Shift:=xlToRight
        Sheets(sheetName).Columns("A:A").Insert Shift:=xlToRight
        
        Sheets(sheetName).Cells(1, 1) = InsertColumns(i)
    Next
    
    
    '再來排序
    Dim ColumnOrder As Variant, ndx As Integer
    Dim Found As Range, counter As Integer
        ColumnOrder = Array("Grade", "Customer", "Sales", "Pull In、Push Out(依Request Date)", "HUB", _
        "Plan Ship Date", "Schedule Ship Date", "Request Date", "Ordered Date", "Territory", "Pre Sch Ship Date", _
        "Customer Name", "AIT P/N", "Product_no", "Package Type", "Currency", "Unit Price", "R", "Unit Price(NTD)", _
        "Ordered Qty", "Ordered Qty(K)", "Ordered Amt(K/NTD)", "Ordered Amt(K/USD)", "Fcst Nonship Qty", "本月已開發票QTY(K)", _
        "月FCST", "月FCST", "月FCST", "月FCST", "月FCST", "月FCST", "End Customer", "Customer PO", "Sample End Customer", _
        "Key Account", "Order Number", "Line ID", "Line No.", "Cust Line No.", "Pick Date", "Grouping Date", "Move Order_no", _
        "Split Flag", "Order Status", "Packing No", "Delivery No", "Subinventory", "Pick From", "LATEST_UPDATED_FLAG", _
        "Hold Reason", "Shipping Instructions", "Application Field", "Original Product", "Sec. Cust PO", "Product Substitution Date", _
        "Schedule Change Date", "Planner Remark", "PC Remark", "Sale Person", "Shipping Method", "SA Planner", "BKG")
    counter = 1
    
    '關掉畫面上的資料的更新：
    '執行巨集之前，先把畫面更新關掉，可以比較快速跑完巨集，不過資料量不大
    '的時候，也沒必要就是了，記得程式碼的最後要把他再打開
    Application.ScreenUpdating = False
       
    For ndx = LBound(ColumnOrder) To UBound(ColumnOrder)
        '從左上角Rows("1:1")開始尋找，找字串ColumnOrder(ndx)，找儲存格的數值符合的LookIn:=xlValues
        '一字不漏比對value相同LookAt:=xlWhole，一個column一個column的順序去找SearchOrder:=xlByColumns
        '找的方向是下一個SearchDirection:=xlNext，大小寫不用完全相符合MatchCase:=False
        Set Found = Sheets(sheetName).Rows("1:1").Find(ColumnOrder(ndx), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
        
        If Not Found Is Nothing Then
            If ColumnOrder(ndx) = "月FCST" Then
                Debug.Print ("月FCST")
                
            End If
            
            
            If Found.Column <> counter Then
                '整個column剪下之後
                Found.EntireColumn.Cut
                '剪下的整個column依序insert到第1個Column、第2個Column………的位置
                '被人家卡位排擠的，就自動往右移動Shift:=xlToRight
                Sheets(sheetName).Columns(counter).Insert Shift:=xlToRight
                '清空記憶體裡面的內容，以免效能越來越差
                Application.CutCopyMode = False
            End If
        counter = counter + 1
        End If
    Next ndx
    '開啟畫面上的資料的更新
    Application.ScreenUpdating = True
    
    '設定某某欄位的日期格式dd-mmm-yy
    
    Dim FoundDate As Range
    '找出某某欄位
    Set FoundDate = Sheets(sheetName).Rows("1:1").Find("Pull In、Push Out(依Request Date)", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    '把這個欄位的型態轉成date(看起來是date，其實不是，因此使用排序功能的話會有問題
    'Sheets(sheetName).Columns(Chr(FoundDate.Column + 64)).Select
    Sheets(sheetName).Columns(FoundDate.Column).Select
    Selection.Value = Selection.Value
    Sheets(sheetName).Columns(FoundDate.Column).NumberFormat = "dd-mmm-yy"
     
    '設定Ordered Amt(K/NTD)欄位的公式：
    'Ordered Amt(K/NTD) = Unit Price(NTD) * Ordered Qty(K)
    '找出欄位Ordered Amt(K/NTD)
    Dim FoundOrderAmtKNTD As Range
    Set FoundOrderAmtKNTD = Sheets(sheetName).Rows("1:1").Find("Ordered Amt(K/NTD)", LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    'MsgBox (FoundOrderAmtKNTD.Column)
    
    '找出欄位:Unit Price(NTD)
    Dim FoundUnitPriceNTD As Range
    Set FoundUnitPriceNTD = Sheets(sheetName).Rows("1:1").Find("Unit Price(NTD)", LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    'MsgBox (FoundUnitPriceNTD.Column)
    
    '找出某某欄位:Ordered Qty(K)
            
    Dim FoundOrderedQtyK As Range
    Set FoundOrderedQtyK = Sheets(sheetName).Rows("1:1").Find("Ordered Qty(K)", LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    'MsgBox (FoundOrderedQtyK.Column)
        
    '設定公式$A2*$B2
    'lastrow最好用欄位Schedule Ship Date去找會比較好，因為其他欄位可能沒資料
    Dim FoundScheduleShipDate As Range
    Set FoundScheduleShipDate = Sheets(sheetName).Rows("1:1").Find("Schedule Ship Date", LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    lastrow = Sheets(sheetName).Cells(Rows.Count, FoundScheduleShipDate.Column).End(xlUp).Row
    Sheets(sheetName).Range(Chr(FoundOrderAmtKNTD.Column + 64) & "2:" & Chr(FoundOrderAmtKNTD.Column + 64) & lastrow).Formula = _
    "=$" & Chr(FoundUnitPriceNTD.Column + 64) & "2*$" & Chr(FoundOrderedQtyK.Column + 64) & "2"
    
    '順便設定欄位 Pull In、Push Out(依Request Date) 的粗體以及紅色文字
    Sheets(sheetName).Range(Chr(FoundDate.Column + 64) & "2:" & Chr(FoundDate.Column + 64) & lastrow).Font.Bold = True
    Sheets(sheetName).Range(Chr(FoundDate.Column + 64) & "2:" & Chr(FoundDate.Column + 64) & lastrow).Font.Color = vbRed
   
    
    
    '排序Ordered Date, Schedule Ship Date, Request Date, Product_no
    '這4個欄位的range先找出來
    Dim FoundOrderedDate As Range
    Set FoundOrderedDate = Sheets(sheetName).Rows("1:1").Find("Ordered Date", LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    Dim FoundRequestDate As Range
    Set FoundRequestDate = Sheets(sheetName).Rows("1:1").Find("Request Date", LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    Dim FoundProductNo As Range
    Set FoundProductNo = Sheets(sheetName).Rows("1:1").Find("Product_no", LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
        
    '把欄位Ordered Date, Schedule Ship Date, Request Date重新給值一次，不然使用排序功能的話會有問題
    Sheets(sheetName).Columns(FoundOrderedDate.Column).Select
    Selection.Value = Selection.Value
    Sheets(sheetName).Columns(FoundScheduleShipDate.Column).Select
    Selection.Value = Selection.Value
    Sheets(sheetName).Columns(FoundRequestDate.Column).Select
    Selection.Value = Selection.Value
    'Product_no也必須重新給值
    Sheets(sheetName).Columns(FoundProductNo.Column).Select
    Selection.Value = Selection.Value
    '把Product_no的格式轉成文字，不然使用排序功能會有問題(儲存格若是通用格式無法排序)
    Sheets(sheetName).Columns(FoundProductNo.Column).NumberFormat = "@"
    
                
    With ActiveSheet.Sort
        '要排序的第一個欄位, 要排序A欄位的話，就寫A1
        'Order：xlAscending表示排序遞增，xlDecending表示排序遞減
         .SortFields.Add Key:=Sheets(sheetName).Range(Chr(FoundOrderedDate.Column + 64) & "1"), Order:=xlAscending
         '要排序的第2個欄位
         .SortFields.Add Key:=Sheets(sheetName).Range(Chr(FoundScheduleShipDate.Column + 64) & "1"), Order:=xlAscending
         '要排序的第3個欄位
         .SortFields.Add Key:=Sheets(sheetName).Range(Chr(FoundRequestDate.Column + 64) & "1"), Order:=xlAscending
         '要排序的第4個欄位
         .SortFields.Add Key:=Sheets(sheetName).Range(Chr(FoundProductNo.Column + 64) & "1"), Order:=xlAscending
         '下面這個是寫死Range的寫法，不太建議這樣
         '.SetRange Range("A1:C13")
         '下面這個是活用的寫法，Excel的資料筆數通常是不固定，這樣寫最建議！
         '.SetRange Sheets(sheetName).Range(Chr(FoundScheduleShipDate.Column + 64) & "1:" & Chr(FoundProductNo.Column + 64) & lastrow)
         .SetRange Sheets(sheetName).Range(Chr(FoundScheduleShipDate.Column + 64) & "1:" & Chr(FoundProductNo.Column + 64) & lastrow)
         '資料是否包含標頭
         .Header = xlYes
         .Apply
    End With
    
    
    'vlook對照：if 處理中的sheet.Line ID == PD 102.Supplier So Shipment No
    'PD 102.Cust Name+Sales Name copy回 處理中的sheet.Customer+Sales
    
    '找出處理中的sheet的Line ID欄位
    Dim FoundBeforeLineID As Range
    Set FoundBeforeLineID = Sheets("BEFORE").Rows("1:1").Find("Line ID", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    
    '找出處理中的sheet的Customer欄位
    Dim FoundBeforeCustomer As Range
    '找出某某欄位
    Set FoundBeforeCustomer = Sheets("BEFORE").Rows("1:1").Find("Customer", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    
    '找出處理中的sheet的Sales欄位
    Dim FoundBeforeSales As Range
    Set FoundBeforeSales = Sheets("BEFORE").Rows("1:1").Find("Sales", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    
    
    
    '找出PD 102的Supplier So Shipment No
    Dim FoundPD102Supplier As Range
    Set FoundPD102Supplier = Sheets("PD 102").Rows("1:4").Find("Supplier So Shipment No", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    
    '找出PD 102的Cust Name
    Dim FoundPD102CustName As Range
    Set FoundPD102CustName = Sheets("PD 102").Rows("1:4").Find("Cust Name", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
   
    '找出PD 102的Sales Name
    Dim FoundPD102SalesName As Range
    Set FoundPD102SalesName = Sheets("PD 102").Rows("1:4").Find("Sales Name", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
   
    '開始逐筆檢查Line ID
    ''找出Line ID的最後一筆
    lastrow = Sheets("BEFORE").Cells(Rows.Count, FoundBeforeLineID.Column).End(xlUp).Row
    For i = 2 To lastrow
        
        Dim lineID As String
        lineID = Sheets("BEFORE").Cells(i, FoundBeforeLineID.Column)
        '到PD 102工作表跟Supplier So Shipment No欄位比對
        lastrowPD102 = Sheets("PD 102").Cells(Rows.Count, FoundPD102Supplier.Column).End(xlUp).Row
        Dim custNamePD102 As String
        custNamePD102 = "N/A"
        Dim salesNamePD102 As String
        salesNamePD102 = "N/A"
        
        For ii = 5 To lastrowPD102
            Dim compareValue As String
            compareValue = Sheets("PD 102").Cells(ii, FoundPD102Supplier.Column)
            If lineID = compareValue Then
                '順利比對到key值的時候, 就要複製回去處理中的工作表
                custNamePD102 = Sheets("PD 102").Cells(ii, FoundPD102CustName.Column)
                salesNamePD102 = Sheets("PD 102").Cells(ii, FoundPD102SalesName.Column)
                                
                Exit For
            End If
            
        Next
                
        Sheets("BEFORE").Cells(i, FoundBeforeCustomer.Column) = custNamePD102
        Sheets("BEFORE").Cells(i, FoundBeforeSales.Column) = salesNamePD102
        
        
    Next
    
    
    'vlook對照：if 處理中的sheet.Product_no == AIT PN處理後.Product_no (處理前)
    'AIT PN處理後.AIT P/N (處理後) copy回 處理中的sheet.AIT P/N
    Dim FoundBeforeProductNo As Range
    Set FoundBeforeProductNo = Sheets("BEFORE").Rows("1:1").Find("Product_no", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)

    Dim FoundBeforeAITPin As Range
    Set FoundBeforeAITPin = Sheets("BEFORE").Rows("1:1").Find("AIT P/N", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)


    Dim FoundMappingProductNo As Range
    Set FoundMappingProductNo = Sheets("AIT PN處理後").Rows("1:1").Find("Product_no (處理前)", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)

    Dim FoundMappingAITPin As Range
    Set FoundMappingAITPin = Sheets("AIT PN處理後").Rows("1:1").Find("AIT P/N (處理後)", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)

    lastrow = Sheets("BEFORE").Cells(Rows.Count, FoundBeforeProductNo.Column).End(xlUp).Row
    For i = 2 To lastrow
        Dim productNo As String
        productNo = Sheets("BEFORE").Cells(i, FoundBeforeProductNo.Column)
        lastrowMapping = Sheets("AIT PN處理後").Cells(Rows.Count, FoundMappingProductNo.Column).End(xlUp).Row
        Dim aitPinMapping As String
        aitPinMapping = "N/A"
        For ii = 2 To lastrowMapping
            Dim compareProductNo As String
            compareProductNo = Sheets("AIT PN處理後").Cells(ii, FoundMappingProductNo.Column)
            If productNo = compareProductNo Then
                '順利比對到key值的時候, 就要複製回去處理中的工作表
                aitPinMapping = Sheets("AIT PN處理後").Cells(ii, FoundMappingAITPin.Column)
                Exit For
            End If
        Next
        '複製回去處理中的工作表
        Sheets("BEFORE").Cells(i, FoundBeforeAITPin.Column) = aitPinMapping
    
        
    Next
    
    
    
    'vlook對照：if WorkingSheet.AIT P/N == AC.AIT PN
    'AC.Oracle Attribute copy到 WorkingSheet.Grade
    Dim FoundBeforeGrade As Range
    Set FoundBeforeGrade = Sheets("BEFORE").Rows("1:1").Find("Grade", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
        
    Dim FoundACAITPin As Range
    Set FoundACAITPin = Sheets("AC").Rows("1:1").Find("AIT PN", LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    
    Dim FoundACOracle As Range
    Set FoundACOracle = Sheets("AC").Rows("1:1").Find("Oracle Attribute", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    
    '開始逐筆檢查WorkingSheet.AIT P/N
    ''找出AIT P/N的最後一筆
    lastrow = Sheets("BEFORE").Cells(Rows.Count, FoundBeforeAITPin.Column).End(xlUp).Row
    For i = 2 To lastrow
        Dim workingAITPin As String
        workingAITPin = Sheets("BEFORE").Cells(i, FoundBeforeAITPin.Column)
        lastrowAC = Sheets("AC").Cells(Rows.Count, FoundACAITPin.Column).End(xlUp).Row
        Dim oracleAC As String
        oracleAC = "N/A"
        For ii = 2 To lastrowAC
            Dim compareAITPin As String
            compareAITPin = Sheets("AC").Cells(ii, FoundACAITPin.Column)
            If workingAITPin = compareAITPin Then
                '順利比對到key值的時候, 就要複製回去處理中的工作表
                oracleAC = Sheets("AC").Cells(ii, FoundACOracle.Column)
                                            
                Exit For
            End If
            
        Next
                
        Sheets("BEFORE").Cells(i, FoundBeforeGrade.Column) = oracleAC
        
        
    Next
    
    '透過inputbox輸入匯率
    Dim FoundRate As Range
    Set FoundRate = Sheets(sheetName).Rows("1:1").Find("R", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    Dim rateRange As Range
    Set rateRange = Range(Chr(FoundRate.Column + 64) & "2:" & Chr(FoundRate.Column + 64) & lastrow)
    rateRange.Value = InputBox("請輸入匯率")
    
    '設定Unit Price(NTD)公式:(UNIT PRICE)*(RATE)
    Dim FoundUnitPrice As Range
    Set FoundUnitPrice = Sheets(sheetName).Rows("1:1").Find("Unit Price", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    lastrow = Sheets(sheetName).Cells(Rows.Count, FoundUnitPrice.Column).End(xlUp).Row
    Sheets(sheetName).Range(Chr(FoundUnitPriceNTD.Column + 64) & "2:" & Chr(FoundUnitPriceNTD.Column + 64) & lastrow).Formula = _
    "=$" & Chr(FoundUnitPrice.Column + 64) & "2*$" & Chr(FoundRate.Column + 64) & "2"
    
    
    '設定公式Ordered Qty(K) = (Ordered Qty) / 1000
    Dim FoundOrderQty As Range
    Set FoundOrderQty = Sheets(sheetName).Rows("1:1").Find("Ordered Qty", LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    lastrow = Sheets(sheetName).Cells(Rows.Count, FoundOrderQty.Column).End(xlUp).Row
    Sheets(sheetName).Range(Chr(FoundOrderedQtyK.Column + 64) & "2:" & Chr(FoundOrderedQtyK.Column + 64) & lastrow).Formula = _
    "=$" & Chr(FoundOrderQty.Column + 64) & "2/1000"
    
    '設定公式Ordered Amt (K/USD) = (Unit Price)*(Ordered Qty(K))
    Dim FoundOrderAmtUSD As Range
    Set FoundOrderAmtUSD = Sheets(sheetName).Rows("1:1").Find("Ordered Amt(K/USD)", LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    lastrow = Sheets(sheetName).Cells(Rows.Count, FoundUnitPrice.Column).End(xlUp).Row
    Sheets(sheetName).Range(Chr(FoundOrderAmtUSD.Column + 64) & "2:" & Chr(FoundOrderAmtUSD.Column + 64) & lastrow).Formula = _
    "=$" & Chr(FoundUnitPrice.Column + 64) & "2*$" & Chr(FoundOrderedQtyK.Column + 64) & "2"
    
End Sub

