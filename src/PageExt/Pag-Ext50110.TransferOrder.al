pageextension 50110 "Tranfer Order MSCE" extends "Transfer Order"
{
    actions
    {
        addafter(GetReceiptLines)
        {
            group(SerialTemplate)
            {
                Caption = 'Serial Template';
                action(SerialImport)
                {
                    ApplicationArea = All;
                    Image = ImportExcel;
                    Caption = 'Import Tracking Info by Excel';
                    ToolTip = 'Executes the Import Excel action.';

                    trigger OnAction()
                    var
                        QuestionCreateManual: Label 'Do you want to import tracking info by Excel?';
                    begin
                        Confirmed := Dialog.Confirm(QuestionCreateManual, false);
                        if not Confirmed then
                            exit;

                        ImportExcelforTORD.ImportExcelData(Rec);
                    end;
                }
                action(SerialExport)
                {
                    ApplicationArea = All;
                    Image = ExportToExcel;
                    Caption = 'Export Tracking Info Template';
                    ToolTip = 'Executes the Export Tracking Info Template action.';

                    trigger OnAction()
                    begin
                        Confirmed := Dialog.Confirm(QuestionCreateManual, false);
                        if not Confirmed then
                            exit;

                        ImportExcelforTORD.ExportTemplate(Rec);
                    end;
                }
            }
        }
    }

    var
        ImportExcelforTORD: Codeunit "Import Tracking Info";
        QuestionCreateManual: Label 'Do you want to export tracking info Template?';
        Confirmed: Boolean;
}
