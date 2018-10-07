Sub 整合()
    Dim sheetName As String
    sheetName = "ExportShipPlan"
    Worksheets(sheetName).Activate
    '先新增欄位
    Dim InsertColumns As Variant
    InsertColumns = Array("Grade", "Customer", "Sales", "Pull In、Push Out(依Request Date)", "HUB", _
    "AIT P/N", "R", "Unit Price(NTD)", "Ordered Qty(K)", "Ordered Amt(K/NTD)", "Ordered Amt(K/USD)", _
    "本月已開發票QTY(K)", "月FCST", "月FCST", "月FCST", "月FCST", "月FCST", "月FCST")
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
    
    Application.ScreenUpdating = False
    '設定某某欄位的日期格式dd-mmm-yy
    Dim FoundDate As Range
    '找出某某欄位
    Set FoundDate = Sheets(sheetName).Rows("1:1").Find("Pull In、Push Out(依Request Date)", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    '把這個欄位的型態轉成date(看起來是date，其實不是，因此使用排序功能的話會有問題
    'Sheets(sheetName).Columns(Chr(FoundDate.Column + 64)).Select
    Sheets(sheetName).Columns(FoundDate.Column).Select
    Selection.Value = Selection.Value
    Sheets(sheetName).Columns(FoundDate.Column).numberFormat = "dd-mmm-yy"
    Application.ScreenUpdating = True
     
    Application.ScreenUpdating = False
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
    Application.ScreenUpdating = True
    
    Application.ScreenUpdating = False
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
    '順便把Ordered Date格式也改成dd-mmm-yy
    Sheets(sheetName).Columns(FoundOrderedDate.Column).numberFormat = "dd-mmm-yy"
    
    Sheets(sheetName).Columns(FoundScheduleShipDate.Column).Select
    Selection.Value = Selection.Value
    '順便把Schedule Ship Date格式也改成dd-mmm-yy
    Sheets(sheetName).Columns(FoundScheduleShipDate.Column).numberFormat = "dd-mmm-yy"
    
    Sheets(sheetName).Columns(FoundRequestDate.Column).Select
    Selection.Value = Selection.Value
    '順便把Request Date格式也改成dd-mmm-yy
    Sheets(sheetName).Columns(FoundRequestDate.Column).numberFormat = "dd-mmm-yy"
    'Product_no也必須重新給值
    Sheets(sheetName).Columns(FoundProductNo.Column).Select
    Selection.Value = Selection.Value
    '把Product_no的格式轉成文字，不然使用排序功能會有問題(儲存格若是通用格式無法排序)
    Sheets(sheetName).Columns(FoundProductNo.Column).numberFormat = "@"
   
    
    '排序到最後一個column
    Dim lastColumn As Long
    lastColumn = Cells(1, Columns.Count).End(xlToLeft).Column
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
         '.SetRange Sheets(sheetName).Range(Chr(FoundScheduleShipDate.Column + 64) & "1:" & Chr(FoundProductNo.Column + 64) & lastrow)
         .SetRange Sheets(sheetName).Range("A1:" & Col_Letter(lastColumn) & lastrow)
         '資料是否包含標頭
         .Header = xlYes
         .Apply
    End With
    Application.ScreenUpdating = True
    
    
    Application.ScreenUpdating = False
    'vlook對照：if 處理中的sheet.Line ID == PD 102.Supplier So Shipment No
    'PD 102.Cust Name+Sales Name copy回 處理中的sheet.Customer+Sales
    '找出處理中的sheet的Line ID欄位
    Dim FoundBeforeLineID As Range
    Set FoundBeforeLineID = Sheets(sheetName).Rows("1:1").Find("Line ID", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    
    '找出處理中的sheet的Customer欄位
    Dim FoundBeforeCustomer As Range
    '找出某某欄位
    Set FoundBeforeCustomer = Sheets(sheetName).Rows("1:1").Find("Customer", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    
    '找出處理中的sheet的Sales欄位
    Dim FoundBeforeSales As Range
    Set FoundBeforeSales = Sheets(sheetName).Rows("1:1").Find("Sales", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    
    ''''''''''''''
    Dim wkbPD102 As Workbook

    Dim strPD102FileToOpen As String
    strPD102FileToOpen = ""
    '透過dialog視窗取得檔案名稱
    strPD102FileToOpen = Application.GetOpenFilename _
    (Title:="請選擇PD 102的檔案", _
    FileFilter:="Excel Files *.xls* (*.xls*),")
        
    If strPD102FileToOpen = "False" Then
        MsgBox "選取PD 102檔案失敗！.", vbExclamation, "Sorry!"
        Exit Sub
    Else
        Set wkbPD102 = Workbooks.Open(strPD102FileToOpen)
        '找出PD 102的Supplier So Shipment No
        Dim FoundPD102Supplier As Range
        Set FoundPD102Supplier = wkbPD102.Sheets("page").Rows("1:1").Find("Supplier So Shipment No", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
  
        '找出PD 102的Cust Name
        Dim FoundPD102CustName As Range
        Set FoundPD102CustName = wkbPD102.Sheets("page").Rows("1:1").Find("Cust Name", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
        
        '找出PD 102的Sales Name
        Dim FoundPD102SalesName As Range
        Set FoundPD102SalesName = wkbPD102.Sheets("page").Rows("1:1").Find("Sales Name", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
        
        '開始逐筆檢查Line ID
        ''找出Line ID的最後一筆
        ThisWorkbook.Activate
        Worksheets(sheetName).Activate
        lastrow = Sheets(sheetName).Cells(Rows.Count, FoundBeforeLineID.Column).End(xlUp).Row
        For i = 2 To lastrow
            
            Dim lineID As String
            lineID = Sheets(sheetName).Cells(i, FoundBeforeLineID.Column)
            '到PD 102工作表跟Supplier So Shipment No欄位比對
            wkbPD102.Activate
            Worksheets("page").Activate
            lastrowPD102 = wkbPD102.Sheets("page").Cells(Rows.Count, FoundPD102Supplier.Column).End(xlUp).Row
            Dim custNamePD102 As String
            custNamePD102 = "N/A"
            Dim salesNamePD102 As String
            salesNamePD102 = "N/A"
            
            For ii = 2 To lastrowPD102
                Dim compareValue As String
                compareValue = wkbPD102.Sheets("page").Cells(ii, FoundPD102Supplier.Column)
                If lineID = compareValue Then
                    '順利比對到key值的時候, 就要複製回去處理中的工作表
                    custNamePD102 = wkbPD102.Sheets("page").Cells(ii, FoundPD102CustName.Column)
                    salesNamePD102 = wkbPD102.Sheets("page").Cells(ii, FoundPD102SalesName.Column)
                                    
                    Exit For
                End If
                
            Next
            ThisWorkbook.Activate
            Worksheets(sheetName).Activate
            Sheets(sheetName).Cells(i, FoundBeforeCustomer.Column) = custNamePD102
            Sheets(sheetName).Cells(i, FoundBeforeSales.Column) = salesNamePD102
            
            
        Next
        
        '關閉檔案
        wkbPD102.Activate
        Worksheets("page").Activate
        wkbPD102.Close SaveChanges:=False
    End If
    
    '''''''''''''''
    ThisWorkbook.Activate
    Worksheets(sheetName).Activate
    Application.ScreenUpdating = True
      
       
       
    
    
    Application.ScreenUpdating = False
    'vlook對照：if 處理中的sheet.Product_no == AIT PN處理後.Product_no (處理前)
    'AIT PN處理後.AIT P/N (處理後) copy回 處理中的sheet.AIT P/N
    Dim FoundBeforeProductNo As Range
    Set FoundBeforeProductNo = Sheets(sheetName).Rows("1:1").Find("Product_no", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)

    Dim FoundBeforeAITPin As Range
    Set FoundBeforeAITPin = Sheets(sheetName).Rows("1:1").Find("AIT P/N", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)

    '開啟料號對照表
    Dim wkbMapping As Workbook
    Dim strMappingFileToOpen As String
    strMappingFileToOpen = ""
    '透過dialog視窗取得檔案名稱
    strMappingFileToOpen = Application.GetOpenFilename _
    (Title:="請選擇 料號對照表 的檔案", _
    FileFilter:="Excel Files *.xls* (*.xls*),")
        
    If strMappingFileToOpen = "False" Then
        MsgBox "選取 料號對照表 檔案失敗！.", vbExclamation, "Sorry!"
        Exit Sub
    Else
        Set wkbMapping = Workbooks.Open(strMappingFileToOpen)
        Dim FoundMappingProductNo As Range
        Set FoundMappingProductNo = Sheets(1).Rows("1:1").Find("Product_no", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    
        Dim FoundMappingAITPin As Range
        Set FoundMappingAITPin = Sheets(1).Rows("1:1").Find("AIT P/N", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
        
        ThisWorkbook.Activate
        Worksheets(sheetName).Activate
        lastrow = Sheets(sheetName).Cells(Rows.Count, FoundBeforeProductNo.Column).End(xlUp).Row
        For i = 2 To lastrow
            Dim productNo As String
            productNo = Sheets(sheetName).Cells(i, FoundBeforeProductNo.Column)
            wkbMapping.Activate
            Worksheets(1).Activate
            lastrowMapping = Sheets(1).Cells(Rows.Count, FoundMappingProductNo.Column).End(xlUp).Row
            Dim aitPinMapping As String
            aitPinMapping = "N/A"
            For ii = 2 To lastrowMapping
                Dim compareProductNo As String
                compareProductNo = Sheets(1).Cells(ii, FoundMappingProductNo.Column)
                If productNo = compareProductNo Then
                    '順利比對到key值的時候, 就要複製回去處理中的工作表
                    aitPinMapping = Sheets(1).Cells(ii, FoundMappingAITPin.Column)
                    Exit For
                End If
            Next
            '複製回去處理中的工作表
            ThisWorkbook.Activate
            Worksheets(sheetName).Activate
            Sheets(sheetName).Cells(i, FoundBeforeAITPin.Column) = aitPinMapping
        
            
        Next
                
        wkbMapping.Activate
        Worksheets(1).Activate
        wkbMapping.Close SaveChanges:=False
    End If
    '''''''''''''''''''''
    ThisWorkbook.Activate
    Worksheets(sheetName).Activate
    Application.ScreenUpdating = True
    
    
    
    
    Application.ScreenUpdating = False
    'vlook對照：if WorkingSheet.AIT P/N == AC.AIT PN
    'AC.Oracle Attribute copy到 WorkingSheet.Grade
    Dim FoundBeforeGrade As Range
    Set FoundBeforeGrade = Sheets(sheetName).Rows("1:1").Find("Grade", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
        
    '開啟AC檔案
    Dim wkbAC As Workbook

    Dim strACFileToOpen As String
    strACFileToOpen = ""
    '透過dialog視窗取得檔案名稱
    strACFileToOpen = Application.GetOpenFilename _
    (Title:="請選擇AC的檔案", _
    FileFilter:="Excel Files *.xls* (*.xls*),")
        
    If strACFileToOpen = "False" Then
        MsgBox "選取AC檔案失敗！.", vbExclamation, "Sorry!"
        Exit Sub
    Else
        Set wkbAC = Workbooks.Open(strACFileToOpen)
        Dim FoundACAITPin As Range
        Set FoundACAITPin = Sheets(1).Rows("1:1").Find("AIT PN", LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
          
        Dim FoundACOracle As Range
        Set FoundACOracle = Sheets(1).Rows("1:1").Find("Oracle Attribute", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
        
        '開始逐筆檢查WorkingSheet.AIT P/N
        ''找出AIT P/N的最後一筆
        ThisWorkbook.Activate
        Worksheets(sheetName).Activate
        lastrow = Sheets(sheetName).Cells(Rows.Count, FoundBeforeAITPin.Column).End(xlUp).Row
        For i = 2 To lastrow
            Dim workingAITPin As String
            workingAITPin = Sheets(sheetName).Cells(i, FoundBeforeAITPin.Column)
            wkbAC.Activate
            Worksheets(1).Activate
            lastrowAC = Sheets(1).Cells(Rows.Count, FoundACAITPin.Column).End(xlUp).Row
            Dim oracleAC As String
            oracleAC = "N/A"
            For ii = 2 To lastrowAC
                Dim compareAITPin As String
                compareAITPin = Sheets(1).Cells(ii, FoundACAITPin.Column)
                If workingAITPin = compareAITPin Then
                    '順利比對到key值的時候, 就要複製回去處理中的工作表
                    oracleAC = Sheets(1).Cells(ii, FoundACOracle.Column)
                                                
                    Exit For
                End If
                
            Next
            ThisWorkbook.Activate
            Worksheets(sheetName).Activate
            Sheets(sheetName).Cells(i, FoundBeforeGrade.Column) = oracleAC
            
            
        Next
                    
        wkbAC.Activate
        Worksheets(1).Activate
        wkbAC.Close SaveChanges:=False
    End If
    '''''''''''''''
    ThisWorkbook.Activate
    Worksheets(sheetName).Activate
    Application.ScreenUpdating = True
    
    
    
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
    
    Application.ScreenUpdating = False
    設定群組 sheetName, "Territory", "Customer Name"
    設定群組 sheetName, "Currency", "Currency"
    設定群組 sheetName, "R", "R"
    設定群組 sheetName, "Fcst Nonship Qty", "月FCST"
    設定群組 sheetName, "Sample End Customer", "Key Account"
    設定群組 sheetName, "Grouping Date", "Grouping Date"
    設定群組 sheetName, "Split Flag", "Order Status"
    設定群組 sheetName, "Subinventory", "PC Remark"
    設定群組 sheetName, "Shipping Method", "Shipping Method"
        
    
    Application.ScreenUpdating = True
    '這樣就會凍結第1個 Row以及ABCDEFGH column
    ActiveWindow.FreezePanes = False
    Range("I2").Select
    ActiveWindow.FreezePanes = True
    
    '設定欄位置左、置中、置右
    Dim ColumnAlignLeft As Variant, ColumnAlignCenter As Variant, ColumnAlignRight As Variant, index As Integer
    Dim FoundColumns As Range
        ColumnAlignLeft = Array("Grade", "Customer", "Schedule Ship Date", "Request Date", "Ordered Date", "Ordered Date", _
        "Territory", "Pre Sch Ship Date", "Customer Name", "AIT P/N", "Product_no", "Package Type", "Currency", _
        "Order Status", "LATEST_UPDATED_FLAG", "Hold Reason", "Application Field", "Schedule Change Date", _
        "Planner Remark", "Sale Person", "Shipping Method", "SA Planner")
        
        ColumnAlignCenter = Array("End Customer", "Customer PO", "Sample End Customer", "Order Number", "Line ID", "Line No.", _
        "Cust Line No.", "Split Flag", "BKG")
        
        ColumnAlignRight = Array("Unit Price", "R", "Unit Price(NTD)", "Ordered Qty", "Ordered Qty(K)", _
        "Ordered Amt(K/NTD)", "Ordered Amt(K/USD)")

    Application.ScreenUpdating = False
    '置左
    For index = LBound(ColumnAlignLeft) To UBound(ColumnAlignLeft)
        Set FoundColumns = Rows("1:1").Find(ColumnAlignLeft(index), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
        If Not FoundColumns Is Nothing Then
            Dim columnEng As String
            columnEng = Col_Letter(FoundColumns.Column)
            '該欄位所有資料(包括header)置左
            Columns(columnEng & ":" & columnEng).HorizontalAlignment = xlLeft
            '表頭另外設定(header一般都是置中)
            Range(columnEng & "1").HorizontalAlignment = xlLeft
        
        End If
    Next index
    
    '置中
    For index = LBound(ColumnAlignCenter) To UBound(ColumnAlignCenter)
        Set FoundColumns = Rows("1:1").Find(ColumnAlignCenter(index), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
        If Not FoundColumns Is Nothing Then
            columnEng = Col_Letter(FoundColumns.Column)
            Columns(columnEng & ":" & columnEng).HorizontalAlignment = xlCenter
            Range(columnEng & "1").HorizontalAlignment = xlLeft
        
        End If
    Next index
    
    '靠右
    For index = LBound(ColumnAlignRight) To UBound(ColumnAlignRight)
        Set FoundColumns = Rows("1:1").Find(ColumnAlignRight(index), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
        If Not FoundColumns Is Nothing Then
            columnEng = Col_Letter(FoundColumns.Column)
            '該欄位所有資料(包括header)置左
            Columns(columnEng & ":" & columnEng).HorizontalAlignment = xlRight
            '表頭另外設定(header一般都是置中)
            Range(columnEng & "1").HorizontalAlignment = xlLeft
        
        End If
    Next index
    
    Application.ScreenUpdating = True
    
    '收合折疊群組
    ActiveSheet.Outline.ShowLevels ColumnLevels:=1
    
    設定小數點幾位 sheetName, "Unit Price", 5
    設定小數點幾位 sheetName, "Unit Price(NTD)", 2
    設定小數點幾位 sheetName, "Ordered Qty(K)", 2
    設定小數點幾位 sheetName, "Ordered Amt(K/NTD)", 2
    設定小數點幾位 sheetName, "Ordered Amt(K/USD)", 2
    
End Sub

Sub 設定群組(sheetName As String, startHeader As String, endHeader As String)
    Dim FoundStart As Range
    '找出某某欄位
    Set FoundStart = Sheets(sheetName).Rows("1:1").Find(startHeader, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    Dim FoundEnd As Range
    '找出某某欄位
    Set FoundEnd = Sheets(sheetName).Rows("1:1").Find(endHeader, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    
    '設定群組
    Sheets(sheetName).Columns(Col_Letter(FoundStart.Column) & ":" & Col_Letter(FoundEnd.Column)).Columns.Group
    
    
End Sub


'轉換column index為英文letter
Function Col_Letter(lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function

Sub 設定小數點幾位(sheetName As String, columnName As String, numberOfDigits As Integer)
    '設定欄位格式小數點幾位
    Dim FoundFloatingNumber As Range
    'Unit Price小數點五位
    Set FoundFloatingNumber = Sheets(sheetName).Rows("1:1").Find(columnName, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    '把這個欄位的value重新給一次，確保萬一不會出錯
    Sheets(sheetName).Columns(FoundFloatingNumber.Column).Select
    Selection.Value = Selection.Value
    Dim numberFormat As String
    numberFormat = "0."
    numberFormat = numberFormat & Replace(Space(numberOfDigits), " ", "0")
    Sheets(sheetName).Columns(FoundFloatingNumber.Column).numberFormat = numberFormat

End Sub


