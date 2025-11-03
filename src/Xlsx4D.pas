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
begin

end;

class function TReaderHelper.QuickLoadFirstSheet(
  const AFileName: string): TWorksheet;
begin

end;

class function TReaderHelper.QuickReadCell(const AFileName: string; ASheetIndex,
  ARow, ACol: Integer): string;
begin

end;

class function TReaderHelper.RowColToCellRef(ARow, ACol: Integer): string;
begin

end;

end.

