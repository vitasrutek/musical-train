unit Unit5;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Grids, Excel2010, Vcl.StdCtrls, ComObj;

type
  TForm5 = class(TForm)
    StringGrid1: TStringGrid;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
  public
  function Xls_To_StringGrid(AGrid: TStringGrid; AXLSFile: string): Boolean;
    { Public declarations }
  end;

var
  Form5: TForm5;

implementation

{$R *.dfm}

uses Unit1;

procedure TForm5.Button1Click(Sender: TObject);
begin
Xls_To_StringGrid(StringGrid1, mainform.PDsoubor)
end;

procedure TForm5.Button2Click(Sender: TObject);
var
   obsah, datum, cena, technik: string;  //bylo variant
   i: integer;
begin
for i := 1 to 300 do
begin
obsah := StringGrid1.Cells[0, i];
  if obsah = mainform.Edit1.text         //bylo obsah.value
    then
      begin
        mainform.Edit9.Text := StringGrid1.Cells[7, i];
        mainform.Edit8.Text := StringGrid1.Cells[4, i];
        datum := StringGrid1.Cells[3, i];
          if datum = '' then
            begin
             datum := StringGrid1.Cells[2, i];
              mainform.Edit6.Text := (datum);       //bylo datetostr (datum)
            end else
              if datum <> '' then
                begin
                  datum := StringGrid1.Cells[3, i];
                  mainform.Edit6.Text := (datum);
                end;

      end;
end;

end;

function TForm5.Xls_To_StringGrid(AGrid: TStringGrid; AXLSFile: string): Boolean;
const
  xlCellTypeLastCell = $0000000B;
var
  XLApp, Sheet: OLEVariant;
  RangeMatrix: Variant;
  x, y, k, r: Integer;
begin
  Result := False;
  // Create Excel-OLE Object
  XLApp := CreateOleObject('Excel.Application');
  try
    // Hide Excel
    XLApp.Visible := False;
    // Open the Workbook
    XLApp.Workbooks.Open(AXLSFile);
    // Sheet := XLApp.Workbooks[1].WorkSheets[1];
    Sheet := XLApp.Workbooks[ExtractFileName(AXLSFile)].WorkSheets[1];
    // In order to know the dimension of the WorkSheet, i.e the number of rows
    // and the number of columns, we activate the last non-empty cell of it
    Sheet.Cells.SpecialCells(xlCellTypeLastCell, EmptyParam).Activate;
    // Get the value of the last row
    x := XLApp.ActiveCell.Row;
    // Get the value of the last column
    y := XLApp.ActiveCell.Column;
    // Set Stringgrid's row &col dimensions.
    AGrid.RowCount := x;
    AGrid.ColCount := y;
    // Assign the Variant associated with the WorkSheet to the Delphi Variant
    RangeMatrix := XLApp.Range['A1', XLApp.Cells.Item[X, Y]].Value;
    //  Define the loop for filling in the TStringGrid
    k := 1;
    repeat
      for r := 1 to y do
        AGrid.Cells[(r - 1), (k - 1)] := RangeMatrix[K, R];
      Inc(k, 1);
      AGrid.RowCount := k + 1;
    until k > x;
    // Unassign the Delphi Variant Matrix
    RangeMatrix := Unassigned;
  finally
    // Quit Excel
    if not VarIsEmpty(XLApp) then
    begin
      XLApp.DisplayAlerts := False;
      XLApp.Quit;
      XLAPP := Unassigned;
      Sheet := Unassigned;
      Result := True;
    end;
  end;
end;

end.
