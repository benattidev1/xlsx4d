unit Xlsx4D;

interface

uses
  Xlsx4D.Types, System.SysUtils;

type
  TXlsx4D = class
  private
    FWorkSheets: TWorksheets;
    FFileName: string;
  public
    constructor Create;
    destructor Destroy; override;

    function LoadFromFile(const AFileName: string): Boolean;
    function GetWorksheets: TWorksheets;
    function GetWorksheet(const AName: string): TWorksheet;
    function GetWorksheetByIndex(const AIndex: Integer): TWorksheet;
    function GetWorksheetCount: Integer;

    property FileName: string read FFileName;
    property Worksheets: TWorksheets read GetWorksheets;
  end;

implementation

uses 
  Xlsx4D.Engine.XLSX;

{ TXlsx4D }

constructor TXlsx4D.Create;
begin
  inherited Create;
  FWorkSheets := nil;
  FFileName := '';
end;

destructor TXlsx4D.Destroy;
begin
  if FWorkSheets <> nil then
    FWorkSheets.Free;
  inherited;
end;

function TXlsx4D.GetWorksheet(const AName: string): TWorksheet;
begin
  if FWorkSheets = nil then 
    Result := Nil
  else
    Result := FWorkSheets.FindByName(AName);
end;

function TXlsx4D.GetWorksheetByIndex(const AIndex: Integer): TWorksheet;
begin
  if (FWorkSheets = nil) or (AIndex < 0) or (AIndex >= FWorkSheets.Count) then 
    Result := Nil
  else
    Result := FWorkSheets.Items[AIndex];
end;

function TXlsx4D.GetWorksheetCount: Integer;
begin
  if FWorkSheets = nil then
    Result := 0
  else
    Result := FWorkSheets.Count;
end;

function TXlsx4D.GetWorksheets: TWorksheets;
begin
  Result := FWorkSheets;
end;

function TXlsx4D.LoadFromFile(const AFileName: string): Boolean;
var
  Engine: TXLSXEngine;
begin
  Result := False;

  if not FileExists(AFileName) then
    raise EXlsx4DException.CreateFmt('File not found: %s', [AFileName]);

  // free previous worksheets if any
  if FWorkSheets <> nil then
    FreeAndNil(FWorkSheets);

  Engine := TXLSXEngine.Create;
  try
    FWorkSheets := Engine.LoadFromFile(AFileName);
    FFileName := AFileName;

    Result := (Assigned(FWorkSheets));
  finally
    Engine.Free;    
  end;
end;

end.

