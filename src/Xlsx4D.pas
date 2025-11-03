unit Xlsx4D;

interface

uses
  XLSX4D.Types;

type
  TReader = class
  private
    FFileName: string;
    FWorksheetCount: Integer;

    function GetWorksheetCount: Integer;
    function GetWorksheet(Index: Integer): TWorksheet;
    function IsXLSXFile(const AFileName: string): Boolean;
  public
    constructor Create;
    destructor Destroy; override;

    procedure LoadFromFile(const AFileName: string);
    procedure Clear;
    function FindWorksheet(const AName: string): Integer;
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

{ TReader }

procedure TReader.Clear;
begin

end;

constructor TReader.Create;
begin

end;

destructor TReader.Destroy;
begin

  inherited;
end;

function TReader.FindWorksheet(const AName: string): Integer;
begin

end;

function TReader.GetCellValue(ASheetIndex, ARow, ACol: Integer): string;
begin

end;

function TReader.GetCellValue(const ASheetName: string; ARow,
  ACol: Integer): string;
begin

end;

function TReader.GetWorksheet(Index: Integer): TWorksheet;
begin

end;

function TReader.GetWorksheetCount: Integer;
begin

end;

function TReader.IsXLSXFile(const AFileName: string): Boolean;
begin

end;

procedure TReader.LoadFromFile(const AFileName: string);
begin

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

