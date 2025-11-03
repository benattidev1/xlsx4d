unit Xlsx4D.Engine.XLSX;

interface

uses
  System.Classes,
  System.Zip,
  Xml.XMLIntf,
  Xlsx4D.Types;

type
  TXLSXEngine = class
  private
    FSharedStrings: TStringList;
    FWorksheets: TWorksheets;

    procedure LoadSharedStrings(const AZipFile: TZipFile);
    procedure LoadWorkbookInfo(const AZipFile: TZipFile);
    procedure LoadWorksheet(const AZipFile: TZipFile; const ASheetPath: string;
      AWorksheet: TWorksheet);
    procedure ParseCellRef(const ACellRef: string; out ARow, ACol: Integer);
    function ColLetterToNumber(const AColLetter: string): Integer;
    function GetCellValue(const ANode: IXMLNode; out ACellType: TCellType): Variant;
  public
    constructor Create;
    destructor Destroy; override;

    function LoadFromFile(const AFileName: string): TWorksheets;
  end;

implementation

{ TXLSXEngine }

function TXLSXEngine.ColLetterToNumber(const AColLetter: string): Integer;
begin

end;

constructor TXLSXEngine.Create;
begin

end;

destructor TXLSXEngine.Destroy;
begin

  inherited;
end;

function TXLSXEngine.GetCellValue(const ANode: IXMLNode;
  out ACellType: TCellType): Variant;
begin

end;

function TXLSXEngine.LoadFromFile(const AFileName: string): TWorksheets;
begin

end;

procedure TXLSXEngine.LoadSharedStrings(const AZipFile: TZipFile);
begin

end;

procedure TXLSXEngine.LoadWorkbookInfo(const AZipFile: TZipFile);
begin

end;

procedure TXLSXEngine.LoadWorksheet(const AZipFile: TZipFile;
  const ASheetPath: string; AWorksheet: TWorksheet);
begin

end;

procedure TXLSXEngine.ParseCellRef(const ACellRef: string; out ARow,
  ACol: Integer);
begin

end;

end.

