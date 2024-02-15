codeunit 50110 "Import Tracking Info"
{
    var
        TempExcelBuffer: Record "Excel Buffer" temporary;
        Text001: Label 'Analyzing Data...\\';


    procedure ImportExcelData(var TransferHeader: Record "Transfer Header")
    var
        InStream: InStream;
        FromFile: Text;
        SheetName: Text;
    begin
        //This procedure allows you to upload a file to Business Central
        UploadIntoStream('Select the Excel file to Import', '', '', FromFile, InStream);

        if FromFile = '' then
            Error('File not found');

        SheetName := TempExcelBuffer.SelectSheetsNameStream(InStream);

        //Subsequently, it loads the Excel Buffer data type with the information from the excel file.
        TempExcelBuffer.Reset();
        TempExcelBuffer.DeleteAll();
        TempExcelBuffer.OpenBookStream(InStream, SheetName);
        TempExcelBuffer.ReadSheet();

        //And finally, it invokes the procedure that will allow obtaining the information from each cell of the excel
        //file and inserting it into the Gen. Journal Lines.
        InsertExcelData(TransferHeader);
    end;

    local procedure InsertExcelData(var TransferHeader: Record "Transfer Header")
    var
        Window: Dialog;
        TransferLine: Record "Transfer Line";
        ReservationEntry: Record "Reservation Entry";
        RowNo, MaxRowNo : Integer;
        ItemNo: Code[20];
        VariantCode: Code[10];
        SerialNo, LotNo : Code[50];
        Confirmed: Boolean;
        LineNo: Integer;
        Question: TextConst ENU = 'Previous records were detected in the Item Tracking table. Do you want to delete them to avoid duplicates?';
    begin
        Window.OPEN(Text001 + '@1@@@@@@@@@@@@@@@@@@@@@@@@@\');
        Window.UPDATE(1, 0);

        RowNo := 0;
        MaxRowNo := 0;

        //To know how many lines we are going to iterate over, we calculate the value of the last line
        TempExcelBuffer.Reset();
        if TempExcelBuffer.FindLast() then begin
            MaxRowNo := TempExcelBuffer."Row No.";
        end;

        ReservationEntry.Reset();
        ReservationEntry.SetRange("Source ID", TransferHeader."No.");
        ReservationEntry.SetRange("Source Type", 39);
        ReservationEntry.SetRange("Source Subtype", 1);
        if ReservationEntry.FindSet() then begin
            Confirmed := Dialog.Confirm(Question, false);
            if Confirmed then begin
                repeat
                    ReservationEntry.Delete();
                until ReservationEntry.Next() = 0;
            end;
        end;

        //We iterate from line or row 2, since the first one is the Headers.
        for RowNo := 2 to MaxRowNo do begin
            Window.UPDATE(1, ROUND(RowNo / MaxRowNo * 10000, 1));

            Evaluate(LineNo, GetValueAtCell(RowNo, 1));
            ItemNo := GetValueAtCell(RowNo, 2);
            LotNo := GetValueAtCell(RowNo, 4);
            SerialNo := GetValueAtCell(RowNo, 5);

            TransferLine.Reset();
            TransferLine.SetRange("Document No.", TransferHeader."No.");
            TransferLine.SetRange("Line No.", LineNo);
            if TransferLine.FindLast() then begin
                CreateTrackingInfo(TransferLine, '', SerialNo, 0D);
            end;
        end;

        TempExcelBuffer.DELETEALL;
        Window.CLOSE;
    end;

    local procedure GetValueAtCell(RowNo: Integer; ColNo: Integer): Text
    begin
        TempExcelBuffer.Reset();
        If TempExcelBuffer.Get(RowNo, ColNo) then
            exit(TempExcelBuffer."Cell Value as Text")
        else
            exit('');
    end;


    /// <summary>
    /// Creates tracking information for items in specific transfer lines. This procedure is crucial
    /// for businesses requiring detailed tracking of inventory items through lot numbers,
    /// serial numbers, and expiration dates. By integrating this functionality,
    /// companies can enhance inventory management, ensure compliance, and improve product safety.
    /// </summary>
    /// <param name="TransLine">The transfer line record for which tracking information will be created.</param>
    /// <param name="LotNo">The lot number associated with the item in the transfer line.</param>
    /// <param name="SN">The serial number associated with the item in the transfer line.</param>
    /// <param name="ExperationDate">The expiration date of the item, important for perishable goods.</param>
    internal procedure CreateTrackingInfo(
        TransLine: Record "Transfer Line";
        LotNo: Code[50];
        SN: Code[50];
        ExperationDate: Date)
    var
        Item: Record Item; // The item record associated with the transfer line.
        ItemTrackingCode: Record "Item Tracking Code"; // Tracking code setup for the item.
        TempReservEntry: Record "Reservation Entry" temporary; // Temporary record for reservation entries.
        CreateReservEntry: Codeunit "Create Reserv. Entry"; // Codeunit to handle reservation entry creation.
        ItemTrackingMgt: Codeunit "Item Tracking Management"; // Codeunit for item tracking management.
        ExpirationDateTracking, LotWarehouseTracking, SNWarehouseTracking : Boolean; // Flags for tracking types.
        ReservStatus: Enum "Reservation Status"; // Reservation status for the entry.
        I: Integer; // Counter or identifier, initially set to 0.
        CurrentSourceRowID: Text[250]; // Source row ID for tracking synchronization.
        SecondSourceRowID: Text[250]; // Secondary source row ID for tracking synchronization.
    begin
        // Retrieve the item based on the item number in the transfer line.
        Item.Get(TransLine."Item No.");
        // Ensure the item has a tracking code assigned.
        if Item."Item Tracking Code" = '' then
            Error('Item Tracking code cannot be empty in Item No =%1', Item."No.");

        // Get the item tracking code setup.
        ItemTrackingCode.Get(Item."Item Tracking Code");
        // Determine the types of tracking required for the item.
        LotWarehouseTracking := ItemTrackingCode."Lot Specific Tracking";
        SNWarehouseTracking := ItemTrackingCode."SN Specific Tracking";
        ExpirationDateTracking := ItemTrackingCode."Man. Expir. Date Entry Reqd.";

        // Initialize the temporary reservation entry.
        TempReservEntry.DeleteAll();
        TempReservEntry.Init();
        // Set tracking information based on the item tracking setup.
        if LotWarehouseTracking then
            TempReservEntry."Lot No." := LotNo;
        if SNWarehouseTracking then
            TempReservEntry."Serial No." := SN;
        if ExperationDate <> 0D then
            TempReservEntry."Expiration Date" := ExperationDate;
        // Validate the necessity of expiration date tracking.
        if (ExpirationDateTracking) and (TempReservEntry."Expiration Date" = 0D) then
            Error('You must use a Expiration Date Tracking');

        // Insert the temporary reservation entry.
        TempReservEntry."Entry No." := I;
        TempReservEntry.Quantity := 1;
        TempReservEntry.Insert();

        // Process each temporary reservation entry to create actual reservation entries.
        if TempReservEntry.FindSet() then
            repeat
                // Create reservation entries for the item tracking.
                CreateReservEntry.SetDates(0D, TempReservEntry."Expiration Date");
                CreateReservEntry.CreateReservEntryFor(
                  Database::"Transfer Line", 0,
                  TransLine."Document No.", '', TransLine."Derived From Line No.", TransLine."Line No.", TransLine."Qty. per Unit of Measure",
                  TempReservEntry.Quantity, TempReservEntry.Quantity * TransLine."Qty. per Unit of Measure", TempReservEntry);
                CreateReservEntry.CreateEntry(
                  TransLine."Item No.", TransLine."Variant Code", TransLine."Transfer-from Code", '', TransLine."Receipt Date", 0D, 0, ReservStatus::Surplus);

                // Synchronize item tracking information.
                CurrentSourceRowID := ItemTrackingMgt.ComposeRowID(5741, 0, TransLine."Document No.", '', 0, TransLine."Line No.");
                SecondSourceRowID := ItemTrackingMgt.ComposeRowID(5741, 1, TransLine."Document No.", '', 0, TransLine."Line No.");
                ItemTrackingMgt.SynchronizeItemTracking(CurrentSourceRowID, SecondSourceRowID, '');
            until TempReservEntry.Next() = 0;
    end;


    //Exporting Template.
    procedure ExportTemplate(var TransferHeader: Record "Transfer Header")
    var
    begin
        ExcelHeaders();
        Process_TransferOrder(TransferHeader);
    end;

    local procedure ExcelHeaders()
    var
    begin
        TempExcelBuffer.NewRow();
        TempExcelBuffer.AddColumn('Line No.', false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn('Item No.', false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn('Variant Code', false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn('Serial Info', false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn('Lot Info', false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
    end;

    local procedure Process_TransferOrder(var TransferHeader: Record "Transfer Header"): Text
    var
        Window: Dialog;
        CalculatingLinesMsg: Label 'Exporting Template...\\';
        CurrentItemMsg: Label 'Item No. #1########  Variant Code #2############', Comment = '%1,%2 = counters';
        TransferLine: Record "Transfer Line";
        Item: Record Item;
        ControlVariable: Integer;
        StartNumber: Integer;
        EndNumber: Integer;
    begin
        Window.Open(CalculatingLinesMsg + CurrentItemMsg);

        TransferLine.Reset();
        TransferLine.SetRange("Document No.", TransferHeader."No.");
        if TransferLine.FindSet() then begin
            repeat
                if Item.Get(TransferLine."Item No.") then begin
                    if Item."Item Tracking Code" <> '' then begin
                        StartNumber := 1;
                        EndNumber := TransferLine.Quantity;
                        for ControlVariable := StartNumber to EndNumber do begin
                            Window.UPDATE(1, TransferLine."Item No.");
                            Window.UPDATE(2, TransferLine."Variant Code");
                            TempExcelBuffer.NewRow();
                            TempExcelBuffer.AddColumn(TransferLine."Line No.", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                            TempExcelBuffer.AddColumn(TransferLine."Item No.", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                            TempExcelBuffer.AddColumn(TransferLine."Variant Code", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                        end;
                    end;
                end;
            until TransferLine.Next() = 0;
        end;

        CloseExcel();

        Window.Close();
    end;

    local procedure CloseExcel()
    var
        NameLbl: Label 'SerialTemplate';
        ExcelFileName: Label 'SerialTemplate_%1_%2_%3';
    begin
        TempExcelBuffer.CreateNewBook(NameLbl);
        TempExcelBuffer.WriteSheet(NameLbl, CompanyName, UserId);
        TempExcelBuffer.CloseBook();
        TempExcelBuffer.SetFriendlyFilename(StrSubstNo(ExcelFileName, CurrentDateTime, UserId, CompanyName));
        TempExcelBuffer.OpenExcel();
    end;
}
