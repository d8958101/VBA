Sub ��X()
    Dim sheetName As String
    sheetName = "BEFORE"
    Worksheets(sheetName).Activate
    '���s�W���
    Dim InsertColumns As Variant
    InsertColumns = Array("Grade", "Customer", "Sales", "Pull In�BPush Out(��Request Date)", "HUB", _
    "AIT P/N", "R", "Unit Price(NTD)", "Ordered Qty(K)", "Ordered Amt(K/NTD)", "Ordered Amt(K/USD)", _
    "����w�}�o��QTY(K)", "��FCST", "��FCST", "��FCST", "��FCST", "��FCST", "��FCST", "BKG")
    For i = LBound(InsertColumns) To UBound(InsertColumns)
        'Sheet1.Columns("A:A").Insert Shift:=xlToRight
        Sheets(sheetName).Columns("A:A").Insert Shift:=xlToRight
        
        Sheets(sheetName).Cells(1, 1) = InsertColumns(i)
    Next
    
    
    '�A�ӱƧ�
    Dim ColumnOrder As Variant, ndx As Integer
    Dim Found As Range, counter As Integer
        ColumnOrder = Array("Grade", "Customer", "Sales", "Pull In�BPush Out(��Request Date)", "HUB", _
        "Plan Ship Date", "Schedule Ship Date", "Request Date", "Ordered Date", "Territory", "Pre Sch Ship Date", _
        "Customer Name", "AIT P/N", "Product_no", "Package Type", "Currency", "Unit Price", "R", "Unit Price(NTD)", _
        "Ordered Qty", "Ordered Qty(K)", "Ordered Amt(K/NTD)", "Ordered Amt(K/USD)", "Fcst Nonship Qty", "����w�}�o��QTY(K)", _
        "��FCST", "��FCST", "��FCST", "��FCST", "��FCST", "��FCST", "End Customer", "Customer PO", "Sample End Customer", _
        "Key Account", "Order Number", "Line ID", "Line No.", "Cust Line No.", "Pick Date", "Grouping Date", "Move Order_no", _
        "Split Flag", "Order Status", "Packing No", "Delivery No", "Subinventory", "Pick From", "LATEST_UPDATED_FLAG", _
        "Hold Reason", "Shipping Instructions", "Application Field", "Original Product", "Sec. Cust PO", "Product Substitution Date", _
        "Schedule Change Date", "Planner Remark", "PC Remark", "Sale Person", "Shipping Method", "SA Planner", "BKG")
    counter = 1
    
    '�����e���W����ƪ���s�G
    '���楨�����e�A����e����s�����A�i�H����ֳt�]�������A���L��ƶq���j
    '���ɭԡA�]�S���n�N�O�F�A�O�o�{���X���̫�n��L�A���}
    Application.ScreenUpdating = False
       
    For ndx = LBound(ColumnOrder) To UBound(ColumnOrder)
        '�q���W��Rows("1:1")�}�l�M��A��r��ColumnOrder(ndx)�A���x�s�檺�ƭȲŦX��LookIn:=xlValues
        '�@�r���|���value�ۦPLookAt:=xlWhole�A�@��column�@��column�����ǥh��SearchOrder:=xlByColumns
        '�䪺��V�O�U�@��SearchDirection:=xlNext�A�j�p�g���Χ����۲ŦXMatchCase:=False
        Set Found = Sheets(sheetName).Rows("1:1").Find(ColumnOrder(ndx), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
        
        If Not Found Is Nothing Then
            If ColumnOrder(ndx) = "��FCST" Then
                Debug.Print ("��FCST")
                
            End If
            
            
            If Found.Column <> counter Then
                '���column�ŤU����
                Found.EntireColumn.Cut
                '�ŤU�����column�̧�insert���1��Column�B��2��Column�K�K�K����m
                '�Q�H�a�d��������A�N�۰ʩ��k����Shift:=xlToRight
                Sheets(sheetName).Columns(counter).Insert Shift:=xlToRight
                '�M�ŰO����̭������e�A�H�K�į�V�ӶV�t
                Application.CutCopyMode = False
            End If
        counter = counter + 1
        End If
    Next ndx
    '�}�ҵe���W����ƪ���s
    Application.ScreenUpdating = True
    
    '�]�w�Y�Y��쪺����榡dd-mmm-yy
    
    Dim FoundDate As Range
    '��X�Y�Y���
    Set FoundDate = Sheets(sheetName).Rows("1:1").Find("Pull In�BPush Out(��Request Date)", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    '��o����쪺���A�নdate(�ݰ_�ӬOdate�A��ꤣ�O�A�]���ϥαƧǥ\�઺�ܷ|�����D
    'Sheets(sheetName).Columns(Chr(FoundDate.Column + 64)).Select
    Sheets(sheetName).Columns(FoundDate.Column).Select
    Selection.Value = Selection.Value
    Sheets(sheetName).Columns(FoundDate.Column).NumberFormat = "dd-mmm-yy"
     
    '�]�wOrdered Amt(K/NTD)��쪺�����G
    'Ordered Amt(K/NTD) = Unit Price(NTD) * Ordered Qty(K)
    '��X���Ordered Amt(K/NTD)
    Dim FoundOrderAmtKNTD As Range
    Set FoundOrderAmtKNTD = Sheets(sheetName).Rows("1:1").Find("Ordered Amt(K/NTD)", LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    'MsgBox (FoundOrderAmtKNTD.Column)
    
    '��X���:Unit Price(NTD)
    Dim FoundUnitPriceNTD As Range
    Set FoundUnitPriceNTD = Sheets(sheetName).Rows("1:1").Find("Unit Price(NTD)", LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    'MsgBox (FoundUnitPriceNTD.Column)
    
    '��X�Y�Y���:Ordered Qty(K)
            
    Dim FoundOrderedQtyK As Range
    Set FoundOrderedQtyK = Sheets(sheetName).Rows("1:1").Find("Ordered Qty(K)", LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    'MsgBox (FoundOrderedQtyK.Column)
        
    '�]�w����$A2*$B2
    'lastrow�̦n�����Schedule Ship Date�h��|����n�A�]����L���i��S���
    Dim FoundScheduleShipDate As Range
    Set FoundScheduleShipDate = Sheets(sheetName).Rows("1:1").Find("Schedule Ship Date", LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    lastrow = Sheets(sheetName).Cells(Rows.Count, FoundScheduleShipDate.Column).End(xlUp).Row
    Sheets(sheetName).Range(Chr(FoundOrderAmtKNTD.Column + 64) & "2:" & Chr(FoundOrderAmtKNTD.Column + 64) & lastrow).Formula = _
    "=$" & Chr(FoundUnitPriceNTD.Column + 64) & "2*$" & Chr(FoundOrderedQtyK.Column + 64) & "2"
    
    '���K�]�w��� Pull In�BPush Out(��Request Date) ������H�ά����r
    Sheets(sheetName).Range(Chr(FoundDate.Column + 64) & "2:" & Chr(FoundDate.Column + 64) & lastrow).Font.Bold = True
    Sheets(sheetName).Range(Chr(FoundDate.Column + 64) & "2:" & Chr(FoundDate.Column + 64) & lastrow).Font.Color = vbRed
   
    
    
    '�Ƨ�Ordered Date, Schedule Ship Date, Request Date, Product_no
    '�o4����쪺range����X��
    Dim FoundOrderedDate As Range
    Set FoundOrderedDate = Sheets(sheetName).Rows("1:1").Find("Ordered Date", LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    Dim FoundRequestDate As Range
    Set FoundRequestDate = Sheets(sheetName).Rows("1:1").Find("Request Date", LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    Dim FoundProductNo As Range
    Set FoundProductNo = Sheets(sheetName).Rows("1:1").Find("Product_no", LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
        
    '�����Ordered Date, Schedule Ship Date, Request Date���s���Ȥ@���A���M�ϥαƧǥ\�઺�ܷ|�����D
    Sheets(sheetName).Columns(FoundOrderedDate.Column).Select
    Selection.Value = Selection.Value
    Sheets(sheetName).Columns(FoundScheduleShipDate.Column).Select
    Selection.Value = Selection.Value
    Sheets(sheetName).Columns(FoundRequestDate.Column).Select
    Selection.Value = Selection.Value
    'Product_no�]�������s����
    Sheets(sheetName).Columns(FoundProductNo.Column).Select
    Selection.Value = Selection.Value
    '��Product_no���榡�ন��r�A���M�ϥαƧǥ\��|�����D(�x�s��Y�O�q�ή榡�L�k�Ƨ�)
    Sheets(sheetName).Columns(FoundProductNo.Column).NumberFormat = "@"
    
                
    With ActiveSheet.Sort
        '�n�ƧǪ��Ĥ@�����, �n�Ƨ�A��쪺�ܡA�N�gA1
        'Order�GxlAscending��ܱƧǻ��W�AxlDecending��ܱƧǻ���
         .SortFields.Add Key:=Sheets(sheetName).Range(Chr(FoundOrderedDate.Column + 64) & "1"), Order:=xlAscending
         '�n�ƧǪ���2�����
         .SortFields.Add Key:=Sheets(sheetName).Range(Chr(FoundScheduleShipDate.Column + 64) & "1"), Order:=xlAscending
         '�n�ƧǪ���3�����
         .SortFields.Add Key:=Sheets(sheetName).Range(Chr(FoundRequestDate.Column + 64) & "1"), Order:=xlAscending
         '�n�ƧǪ���4�����
         .SortFields.Add Key:=Sheets(sheetName).Range(Chr(FoundProductNo.Column + 64) & "1"), Order:=xlAscending
         '�U���o�ӬO�g��Range���g�k�A���ӫ�ĳ�o��
         '.SetRange Range("A1:C13")
         '�U���o�ӬO���Ϊ��g�k�AExcel����Ƶ��Ƴq�`�O���T�w�A�o�˼g�̫�ĳ�I
         '.SetRange Sheets(sheetName).Range(Chr(FoundScheduleShipDate.Column + 64) & "1:" & Chr(FoundProductNo.Column + 64) & lastrow)
         .SetRange Sheets(sheetName).Range(Chr(FoundScheduleShipDate.Column + 64) & "1:" & Chr(FoundProductNo.Column + 64) & lastrow)
         '��ƬO�_�]�t���Y
         .Header = xlYes
         .Apply
    End With
    
    
    'vlook��ӡGif �B�z����sheet.Line ID == PD 102.Supplier So Shipment No
    'PD 102.Cust Name+Sales Name copy�^ �B�z����sheet.Customer+Sales
    
    '��X�B�z����sheet��Line ID���
    Dim FoundBeforeLineID As Range
    Set FoundBeforeLineID = Sheets("BEFORE").Rows("1:1").Find("Line ID", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    
    '��X�B�z����sheet��Customer���
    Dim FoundBeforeCustomer As Range
    '��X�Y�Y���
    Set FoundBeforeCustomer = Sheets("BEFORE").Rows("1:1").Find("Customer", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    
    '��X�B�z����sheet��Sales���
    Dim FoundBeforeSales As Range
    Set FoundBeforeSales = Sheets("BEFORE").Rows("1:1").Find("Sales", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    
    
    
    '��XPD 102��Supplier So Shipment No
    Dim FoundPD102Supplier As Range
    Set FoundPD102Supplier = Sheets("PD 102").Rows("1:4").Find("Supplier So Shipment No", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    
    '��XPD 102��Cust Name
    Dim FoundPD102CustName As Range
    Set FoundPD102CustName = Sheets("PD 102").Rows("1:4").Find("Cust Name", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
   
    '��XPD 102��Sales Name
    Dim FoundPD102SalesName As Range
    Set FoundPD102SalesName = Sheets("PD 102").Rows("1:4").Find("Sales Name", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
   
    '�}�l�v���ˬdLine ID
    ''��XLine ID���̫�@��
    lastrow = Sheets("BEFORE").Cells(Rows.Count, FoundBeforeLineID.Column).End(xlUp).Row
    For i = 2 To lastrow
        
        Dim lineID As String
        lineID = Sheets("BEFORE").Cells(i, FoundBeforeLineID.Column)
        '��PD 102�u�@���Supplier So Shipment No�����
        lastrowPD102 = Sheets("PD 102").Cells(Rows.Count, FoundPD102Supplier.Column).End(xlUp).Row
        Dim custNamePD102 As String
        custNamePD102 = "N/A"
        Dim salesNamePD102 As String
        salesNamePD102 = "N/A"
        
        For ii = 5 To lastrowPD102
            Dim compareValue As String
            compareValue = Sheets("PD 102").Cells(ii, FoundPD102Supplier.Column)
            If lineID = compareValue Then
                '���Q����key�Ȫ��ɭ�, �N�n�ƻs�^�h�B�z�����u�@��
                custNamePD102 = Sheets("PD 102").Cells(ii, FoundPD102CustName.Column)
                salesNamePD102 = Sheets("PD 102").Cells(ii, FoundPD102SalesName.Column)
                                
                Exit For
            End If
            
        Next
                
        Sheets("BEFORE").Cells(i, FoundBeforeCustomer.Column) = custNamePD102
        Sheets("BEFORE").Cells(i, FoundBeforeSales.Column) = salesNamePD102
        
        
    Next
    
    
    'vlook��ӡGif �B�z����sheet.Product_no == AIT PN�B�z��.Product_no (�B�z�e)
    'AIT PN�B�z��.AIT P/N (�B�z��) copy�^ �B�z����sheet.AIT P/N
    Dim FoundBeforeProductNo As Range
    Set FoundBeforeProductNo = Sheets("BEFORE").Rows("1:1").Find("Product_no", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)

    Dim FoundBeforeAITPin As Range
    Set FoundBeforeAITPin = Sheets("BEFORE").Rows("1:1").Find("AIT P/N", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)


    Dim FoundMappingProductNo As Range
    Set FoundMappingProductNo = Sheets("AIT PN�B�z��").Rows("1:1").Find("Product_no (�B�z�e)", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)

    Dim FoundMappingAITPin As Range
    Set FoundMappingAITPin = Sheets("AIT PN�B�z��").Rows("1:1").Find("AIT P/N (�B�z��)", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)

    lastrow = Sheets("BEFORE").Cells(Rows.Count, FoundBeforeProductNo.Column).End(xlUp).Row
    For i = 2 To lastrow
        Dim productNo As String
        productNo = Sheets("BEFORE").Cells(i, FoundBeforeProductNo.Column)
        lastrowMapping = Sheets("AIT PN�B�z��").Cells(Rows.Count, FoundMappingProductNo.Column).End(xlUp).Row
        Dim aitPinMapping As String
        aitPinMapping = "N/A"
        For ii = 2 To lastrowMapping
            Dim compareProductNo As String
            compareProductNo = Sheets("AIT PN�B�z��").Cells(ii, FoundMappingProductNo.Column)
            If productNo = compareProductNo Then
                '���Q����key�Ȫ��ɭ�, �N�n�ƻs�^�h�B�z�����u�@��
                aitPinMapping = Sheets("AIT PN�B�z��").Cells(ii, FoundMappingAITPin.Column)
                Exit For
            End If
        Next
        '�ƻs�^�h�B�z�����u�@��
        Sheets("BEFORE").Cells(i, FoundBeforeAITPin.Column) = aitPinMapping
    
        
    Next
    
    
    
    'vlook��ӡGif WorkingSheet.AIT P/N == AC.AIT PN
    'AC.Oracle Attribute copy�� WorkingSheet.Grade
    Dim FoundBeforeGrade As Range
    Set FoundBeforeGrade = Sheets("BEFORE").Rows("1:1").Find("Grade", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
        
    Dim FoundACAITPin As Range
    Set FoundACAITPin = Sheets("AC").Rows("1:1").Find("AIT PN", LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    
    Dim FoundACOracle As Range
    Set FoundACOracle = Sheets("AC").Rows("1:1").Find("Oracle Attribute", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    
    '�}�l�v���ˬdWorkingSheet.AIT P/N
    ''��XAIT P/N���̫�@��
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
                '���Q����key�Ȫ��ɭ�, �N�n�ƻs�^�h�B�z�����u�@��
                oracleAC = Sheets("AC").Cells(ii, FoundACOracle.Column)
                                            
                Exit For
            End If
            
        Next
                
        Sheets("BEFORE").Cells(i, FoundBeforeGrade.Column) = oracleAC
        
        
    Next
    
    '�z�Linputbox��J�ײv
    Dim FoundRate As Range
    Set FoundRate = Sheets(sheetName).Rows("1:1").Find("R", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    Dim rateRange As Range
    Set rateRange = Range(Chr(FoundRate.Column + 64) & "2:" & Chr(FoundRate.Column + 64) & lastrow)
    rateRange.Value = InputBox("�п�J�ײv")
    
    '�]�wUnit Price(NTD)����:(UNIT PRICE)*(RATE)
    Dim FoundUnitPrice As Range
    Set FoundUnitPrice = Sheets(sheetName).Rows("1:1").Find("Unit Price", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    lastrow = Sheets(sheetName).Cells(Rows.Count, FoundUnitPrice.Column).End(xlUp).Row
    Sheets(sheetName).Range(Chr(FoundUnitPriceNTD.Column + 64) & "2:" & Chr(FoundUnitPriceNTD.Column + 64) & lastrow).Formula = _
    "=$" & Chr(FoundUnitPrice.Column + 64) & "2*$" & Chr(FoundRate.Column + 64) & "2"
    
    
    '�]�w����Ordered Qty(K) = (Ordered Qty) / 1000
    Dim FoundOrderQty As Range
    Set FoundOrderQty = Sheets(sheetName).Rows("1:1").Find("Ordered Qty", LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    lastrow = Sheets(sheetName).Cells(Rows.Count, FoundOrderQty.Column).End(xlUp).Row
    Sheets(sheetName).Range(Chr(FoundOrderedQtyK.Column + 64) & "2:" & Chr(FoundOrderedQtyK.Column + 64) & lastrow).Formula = _
    "=$" & Chr(FoundOrderQty.Column + 64) & "2/1000"
    
    '�]�w����Ordered Amt (K/USD) = (Unit Price)*(Ordered Qty(K))
    Dim FoundOrderAmtUSD As Range
    Set FoundOrderAmtUSD = Sheets(sheetName).Rows("1:1").Find("Ordered Amt(K/USD)", LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    lastrow = Sheets(sheetName).Cells(Rows.Count, FoundUnitPrice.Column).End(xlUp).Row
    Sheets(sheetName).Range(Chr(FoundOrderAmtUSD.Column + 64) & "2:" & Chr(FoundOrderAmtUSD.Column + 64) & lastrow).Formula = _
    "=$" & Chr(FoundUnitPrice.Column + 64) & "2*$" & Chr(FoundOrderedQtyK.Column + 64) & "2"
    
End Sub

