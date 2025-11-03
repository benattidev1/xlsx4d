unit Xlsx4D.Types;

interface

uses
  System.Generics.Collections,
  System.SysUtils;

type
  TCellType = (ctEmpty, ctString, ctNumber, ctBoolean, ctDate, ctFormula, ctError);

  TCell = record
    Row: Integer;           // Linha (1-based)
    Col: Integer;           // Coluna (1-based)
    Value: Variant;         // Valor da célula
    CellType: TCellType;    // Tipo da célula
    FormattedValue: string; // Valor formatado como string

    class function Empty(ARow, ACol: Integer): TCell; static;
    class function Create(ARow, ACol: Integer; const AValue: Variant; ACellType: TCellType = ctString): TCell; static;
    function IsEmpty: Boolean;
    function AsString: string;
    function AsInteger: Integer;
    function AsFloat: Double;
    function AsBoolean: Boolean;
    function AsDateTime: TDateTime;
  end;

  TRowCells = TList<TCell>;

  TWorksheet = class
  private
    FName: string;
    FRows: TObjectList<TRowCells>;
    FMaxRow: Integer;
    FMaxCol: Integer;

    function GetCell(ARow, ACol: Integer): TCell;
    procedure SetCell(ARow, ACol: Integer; const Value: TCell);
    function GetRowCount: Integer;
    function GetColCount: Integer;
  public
    constructor Create;
    destructor Destroy; override;

    procedure Clear;
    procedure AddCell(const ACell: TCell);
    function GetRow(ARow: Integer): TRowCells;

    property Name: string read FName write FName;
    property Cells[ARow, ACol: Integer]: TCell read GetCell write SetCell; default;
    property RowCount: Integer read GetRowCount;
    property ColCount: Integer read GetColCount;
  end;

  TWorksheets = class(TObjectList<TWorksheet>)
  public
    function FindByName(const AName: string): TWorksheet;
  end;

  EXlsx4DException = class(Exception);

implementation

{ TCell }

function TCell.AsBoolean: Boolean;
begin

end;

function TCell.AsDateTime: TDateTime;
begin

end;

function TCell.AsFloat: Double;
begin

end;

function TCell.AsInteger: Integer;
begin

end;

function TCell.AsString: string;
begin

end;

class function TCell.Create(ARow, ACol: Integer; const AValue: Variant;
  ACellType: TCellType): TCell;
begin

end;

class function TCell.Empty(ARow, ACol: Integer): TCell;
begin

end;

function TCell.IsEmpty: Boolean;
begin

end;

{ TWorksheet }

procedure TWorksheet.AddCell(const ACell: TCell);
begin

end;

procedure TWorksheet.Clear;
begin

end;

constructor TWorksheet.Create;
begin

end;

destructor TWorksheet.Destroy;
begin

  inherited;
end;

function TWorksheet.GetCell(ARow, ACol: Integer): TCell;
begin

end;

function TWorksheet.GetColCount: Integer;
begin

end;

function TWorksheet.GetRow(ARow: Integer): TRowCells;
begin

end;

function TWorksheet.GetRowCount: Integer;
begin

end;

procedure TWorksheet.SetCell(ARow, ACol: Integer; const Value: TCell);
begin

end;

{ TWorksheets }

function TWorksheets.FindByName(const AName: string): TWorksheet;
begin

end;

end.

