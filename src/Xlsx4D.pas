unit Xlsx4D;

interface

uses
  System.SysUtils,
  Xlsx4D.Engine.XLSX,
  Xlsx4D.Engine.XLS,
  XLSX4D.Types;

type
  TReader = class
  private
    FFileName: string;
    FWorksheets: TWorksheets;

    function GetWorksheetCount: Integer;
    function GetWorksheet(Index: Integer): TWorksheet;
    function IsXLSXFile(const AFileName: string): Boolean;
  public
    constructor Create;
    destructor Destroy; override;

    procedure LoadFromFile(const AFileName: string);
    procedure Clear;
    function FindWorksheet(const AName: string): TWorksheet;
    function GetCellValue(ASheetIndex, ARow, ACol: Integer): string; overload;
    function GetCellValue(const ASheetName: string; ARow, ACol: Integer): string; overload;

    property FileName: string read FFileName write FFileName;
    property WorksheetCount: Integer read GetWorksheetCount;
    property Worksheets[Index: Integer]: TWorksheet read GetWorksheet; default;
  end;

  TReaderHelper = class
  public
    class function QuickLoadFirstSheet(const AFileName: string): TWorksheet;
    class function QuickReadCell(const AFileName: string; ASheetIndex, ARow, ACol: Integer): string;
    class procedure CellRefToRowCol(const ACellRef: string; out ARow, ACol: Integer);
    class function RowColToCellRef(ARow, ACol: Integer): string;
  end;

implementation

uses
  System.IOUtils;

{ TReader }

procedure TReader.Clear;
begin
  FWorksheets.Clear;
  FFileName := '';
end;

constructor TReader.Create;
begin
  inherited Create;
  FWorksheets := TWorksheets.Create(True);
end;

destructor TReader.Destroy;
begin
  FWorksheets.Free;
  inherited;
end;

function TReader.FindWorksheet(const AName: string): TWorksheet;
begin
  Result := FWorksheets.FindByName(AName);
end;

function TReader.GetCellValue(ASheetIndex, ARow, ACol: Integer): string;
begin
  Result := '';
  if (ASheetIndex > 0) and (ASheetIndex < FWorksheets.Count) then
    Result := FWorksheets[ASheetIndex].Cells[ARow, ACol].AsString;
end;

function TReader.GetCellValue(const ASheetName: string; ARow,
  ACol: Integer): string;
var
  Sheet: TWorksheet;
begin
  Result := '';
  Sheet := FindWorksheet(ASheetName);
  if Assigned(Sheet) then
    Result := Sheet.Cells[ARow, ACol].AsString;
end;

function TReader.GetWorksheet(Index: Integer): TWorksheet;
begin
  if (Index >= 0) and (Index < FWorksheets.Count) then
    Result := FWorksheets[Index]
  else
    raise EXlsx4DException.CreateFmt('Índice de planilha inválido: %d', [Index]);
end;

function TReader.GetWorksheetCount: Integer;
begin
  Result := FWorksheets.Count;
end;

function TReader.IsXLSXFile(const AFileName: string): Boolean;
var
  Ext: string;
begin
  Ext := LowerCase(TPath.GetExtension(AFileName));
  Result := (Ext = '.xlsx') or (Ext = '.xlsm');
end;

procedure TReader.LoadFromFile(const AFileName: string);
var
  XlsxEngine: TXLSXEngine;
  XlsEngine: TXLSEngine;
  LoadedSheets: TWorksheets;
begin
  if not FileExists(AFileName) then
    raise EXlsx4DException.CreateFmt('Arquivo não encontrado: %s', [AFileName]);

  Clear;
  FFileName := AFileName;

  if IsXLSXFile(AFileName) then
  begin
    XlsxEngine := TXLSXEngine.Create;
    try
      LoadedSheets := XlsxEngine.LoadFromFile(AFileName);

      FWorksheets.Free;
      FWorksheets := LoadedSheets;
    finally
      XlsxEngine.Free;
    end;
  end
  else
  begin
    XlsEngine := TXLSEngine.Create;
    try
      LoadedSheets := XlsEngine.LoadFromFile(AFileName);

      FWorksheets.Free;
      FWorksheets := LoadedSheets;
    finally
      XlsEngine.Free;
    end;
  end;
end;


{ TReaderHelper }

class procedure TReaderHelper.CellRefToRowCol(const ACellRef: string; out ARow,
  ACol: Integer);
var
  I: Integer;
  ColPart, RowPart: string;
begin
  ColPart := '';
  RowPart := '';

  for I := 1 to Length(ACellRef) do
  begin
    if CharInSet(ACellRef[I], ['A'..'Z', 'a'..'z']) then
      ColPart := ColPart + UpCase(ACellRef[I])
    else
      RowPart := RowPart + ACellRef[I];
  end;

  ACol := 0;
  for I := 1 to Length(ColPart) do
    ACol := ACol * 26 + (Ord(ColPart[I]) - Ord('A') + 1);

  ARow := StrToIntDef(RowPart, 1);
end;

class function TReaderHelper.QuickLoadFirstSheet(
  const AFileName: string): TWorksheet;
var
  Reader: TReader;
begin
  Reader := TReader.Create;
  try
    Reader.LoadFromFile(AFileName);
    if Reader.WorksheetCount > 0 then
    begin
      Result := TWorksheet.Create(Reader.Worksheets[0].Name);
      Result := Reader.Worksheets[0];
    end
    else
      Result := nil;
  finally

  end;
end;

class function TReaderHelper.QuickReadCell(const AFileName: string; ASheetIndex,
  ARow, ACol: Integer): string;
var
  Reader: TReader;
begin
  Result := '';
  Reader := TReader.Create;
  try
    Reader.LoadFromFile(AFileName);
    if ASheetIndex < Reader.WorksheetCount then
      Result := Reader.GetCellValue(ASheetIndex, ARow, ACol);
  finally
    Reader.Free;
  end;
end;

class function TReaderHelper.RowColToCellRef(ARow, ACol: Integer): string;
var
  ColStr: string;
  Temp: Integer;
begin
  ColStr := '';
  Temp := ACol;
  while Temp > 0 do
  begin
    Dec(Temp);
    ColStr := Chr(Ord('A') + (Temp mod 26)) + ColStr;
    Temp := Temp div 26;
  end;

  Result := ColStr + IntToStr(ARow);
end;

end.

