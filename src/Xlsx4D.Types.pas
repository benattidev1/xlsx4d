unit Xlsx4D.Types;

interface

uses
  System.Generics.Collections,
  System.SysUtils,
  System.Variants;

type
  TCellType = (ctEmpty, ctString, ctNumber, ctBoolean, ctDate, ctFormula, ctError);

  TCell = record
    Row: Integer;
    Col: Integer;
    Value: Variant;
    CellType: TCellType;
    FormattedValue: string;

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
  if IsEmpty then
    Result := False
  else
  try
    Result := VarAsType(Value, varBoolean)
  except
    Result := False;
  end;
end;

function TCell.AsDateTime: TDateTime;
begin
  if IsEmpty then
    Result := 0
  else
  try
    if VarIsNumeric(Value) then
      Result := Double(Value) + EncodeDate(1899, 12, 30)
    else
      Result := StrToDateTime(VarToStr(Value));
  except
    Result := 0;
  end;
end;

function TCell.AsFloat: Double;
begin
  if IsEmpty then
    Result := 0.00
  else
  try
    Result := VarAsType(Value, varDouble);
  except
    Result := 0.00;
  end;
end;

function TCell.AsInteger: Integer;
begin
  if IsEmpty then
    Result := 0
  else
    Result := StrToIntDef(VarToStrDef(Value, '0'), 0);
end;

function TCell.AsString: string;
begin
  if IsEmpty then
    Result := ''
  else
    Result := VarToStrDef(Value, '');
end;

class function TCell.Create(ARow, ACol: Integer; const AValue: Variant; ACellType: TCellType): TCell;
begin
  Result.Row := ARow;
  Result.Col := ACol;
  Result.Value := AValue;
  Result.CellType := ACellType;
  Result.FormattedValue := VarToStrDef(AValue, '');
end;

class function TCell.Empty(ARow, ACol: Integer): TCell;
begin
  Result := TCell.Create(ARow, ACol, Null, ctEmpty);
end;

function TCell.IsEmpty: Boolean;
begin
  Result := (CellType = ctEmpty) or VarIsNull(Value) or VarIsEmpty(Value);
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

