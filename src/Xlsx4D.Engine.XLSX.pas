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

uses
  System.SysUtils,
  System.Variants,
  Xml.XMLDoc;

{ TXLSXEngine }

function TXLSXEngine.ColLetterToNumber(const AColLetter: string): Integer;
var
  I: Integer;
begin
  Result := 0;
  for I := 1 to Length(AColLetter) do
  begin
    Result := Result * 26 + (Ord(AColLetter[I]) - Ord('A') + 1);
  end;
end;

constructor TXLSXEngine.Create;
begin
  inherited Create;
  FSharedStrings := TStringList.Create;
end;

destructor TXLSXEngine.Destroy;
begin
  FSharedStrings.Free;
  inherited;
end;

function TXLSXEngine.GetCellValue(const ANode: IXMLNode;
  out ACellType: TCellType): Variant;
var
  ValueNode: IXMLNode;
  CellTypeAttr: string;
  ValueStr: string;
  SharedStringIndex: Integer;
begin
  Result := Null;
  ACellType := ctEmpty;

  CellTypeAttr := ANode.Attributes['t'];

  ValueNode := ANode.ChildNodes.FindNode('v');
  if not Assigned(ValueNode) then
    Exit;

  ValueStr := ValueNode.Text;

  if CellTypeAttr = 's' then
  begin
    ACellType := ctString;
    SharedStringIndex := StrToIntDef(ValueStr, -1);
    if (SharedStringIndex >= 0) and (SharedStringIndex < FSharedStrings.Count) then
      Result := FSharedStrings[SharedStringIndex]
    else
      Result := '';
  end
  else if CellTypeAttr = 'b' then
  begin
    ACellType := ctBoolean;
    Result := (ValueStr = '1');
  end
  else if CellTypeAttr = 'e' then
  begin
    ACellType := ctError;
    Result := ValueStr;
  end
  else if CellTypeAttr = 'str' then
  begin
    ACellType := ctString;
    Result := ValueStr;
  end
  else
  begin
    ACellType := ctNumber;
    Result := StrToFloatDef(StringReplace(ValueStr, '.', ',', [rfReplaceAll]), 0);
  end;
end;

function TXLSXEngine.LoadFromFile(const AFileName: string): TWorksheets;
var
  ZipFile: TZipFile;
  I: Integer;
begin
  if not FileExists(AFileName) then
    raise EXlsx4DException.CreateFmt('Arquivo não encontrado: %s', [AFileName]);

  FWorksheets := TWorksheets.Create(True);
  Result := FWorksheets;

  ZipFile := TZipFile.Create;
  try
    try
      ZipFile.Open(AFileName, TZipMode.zmRead);

      LoadSharedStrings(ZipFile);

      LoadWorkbookInfo(ZipFile);

      for I := 0 to FWorksheets.Count - 1 do
      begin
        LoadWorksheet(ZipFile, Format('xl/worksheets/sheet%d.xml', [I + 1]),
          FWorksheets[I]);
      end;
    except
      on E: Exception do
        raise EXlsx4DException.CreateFmt('Erro ao ler arquivo XLSX: %s', [E.Message]);
    end;
  finally
    ZipFile.Free;
  end;
end;

procedure TXLSXEngine.LoadSharedStrings(const AZipFile: TZipFile);
var
  Stream: TMemoryStream;
  XMLDoc: TXMLDocument;
  Node, TextNode: IXMLNode;
  I: Integer;
  Bytes: TBytes;
begin
  FSharedStrings.Clear;

  if AZipFile.IndexOf('xl/sharedStrings.xml') < 0 then
    Exit;

  Stream := TMemoryStream.Create;
  try
    AZipFile.Read('xl/sharedStrings.xml', Bytes);

    if Length(Bytes) > 0 then
    begin
      Stream.WriteBuffer(Bytes[0], Length(Bytes));
      Stream.Position := 0;
    end;

    XMLDoc := TXMLDocument.Create(nil);
    XMLDoc.LoadFromStream(Stream);
    XMLDoc.Active := True;

    Node := XMLDoc.DocumentElement;
    if Assigned(Node) then
    begin
      for I := 0 to Node.ChildNodes.Count - 1 do
      begin
        if SameText(Node.ChildNodes[I].NodeName, 'si') then
        begin
          TextNode := Node.ChildNodes[I].ChildNodes.FindNode('t');
          if Assigned(TextNode) then
            FSharedStrings.Add(TextNode.Text)
          else
            FSharedStrings.Add('');
        end;
      end;
    end;
  finally
    Stream.Free;
  end;
end;

procedure TXLSXEngine.LoadWorkbookInfo(const AZipFile: TZipFile);
var
  Stream: TMemoryStream;
  XMLDoc: TXMLDocument;
  SheetsNode, SheetNode: IXMLNode;
  I: Integer;
  SheetName: string;
  Bytes: TBytes;
begin
  Stream := TMemoryStream.Create;
  try
    AZipFile.Read('xml/workbook.xml', Bytes);

    if Length(Bytes) > 0 then
    begin
      Stream.WriteBuffer(Bytes[0], Length(Bytes));
      Stream.Position := 0;
    end;

    XMLDoc := TXMLDocument.Create(nil);
    XMLDoc.LoadFromStream(Stream);
    XMLDoc.Active := True;

    SheetsNode := XMLDoc.DocumentElement.ChildNodes.FindNode('sheets');
    if Assigned(SheetsNode) then
    begin
      for I := 0 to SheetsNode.ChildNodes.Count - 1 do
      begin
        SheetNode := SheetsNode.ChildNodes[I];
        if SameText(SheetNode.NodeName, 'sheet') then
        begin
          SheetName := SheetNode.Attributes['name'];
          if SheetName = '' then
            SheetName := Format('Sheet%d', [I + 1]);

          FWorksheets.Add(TWorksheet.Create(SheetName));
        end;
      end;
    end;
  finally
    Stream.Free;
  end;
end;

procedure TXLSXEngine.LoadWorksheet(const AZipFile: TZipFile;
  const ASheetPath: string; AWorksheet: TWorksheet);
var
  Stream: TMemoryStream;
  XMLDoc: TXMLDocument;
  SheetDataNode, RowNode, CellNode: IXMLNode;
  I, J: Integer;
  CellRef: string;
  Row, Col: Integer;
  Cell: TCell;
  CellType: TCellType;
   Bytes: TBytes;
begin
  if AZipFile.IndexOf(ASheetPath) < 0 then
    Exit;

  Stream := TMemoryStream.Create;
  try
    AZipFile.Read(ASheetPath, Bytes);

    if Length(Bytes) > 0 then
    begin
      Stream.WriteBuffer(Bytes[0], Length(Bytes));
      Stream.Position := 0;
    end;

    XMLDoc := TXMLDocument.Create(nil);
    XMLDoc.LoadFromStream(Stream);
    XMLDoc.Active := True;

    SheetDataNode := XMLDoc.DocumentElement.ChildNodes.FindNode('sheetData');
    if not Assigned(SheetDataNode) then
      Exit;

    for I := 0 to SheetDataNode.ChildNodes.Count - 1 do
    begin
      RowNode := SheetDataNode.ChildNodes[I];
      if not SameText(RowNode.NodeName, 'row') then
        Continue;

      for J := 0 to RowNode.ChildNodes.Count - 1 do
      begin
        CellNode := RowNode.ChildNodes[J];
        if not SameText(CellNode.NodeName, 'c') then
          Continue;

        CellRef := CellNode.Attributes['r'];
        ParseCellRef(CellRef, Row, Col);

        Cell := TCell.Create(Row, Col, GetCellValue(CellNode, CellType), CellType);
        AWorksheet.AddCell(Cell);
      end;
    end;
  finally
    Stream.Free;
  end;
end;

procedure TXLSXEngine.ParseCellRef(const ACellRef: string; out ARow,
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

  ACol := ColLetterToNumber(ColPart);
  ARow := StrToIntDef(RowPart, 1);
end;

end.

