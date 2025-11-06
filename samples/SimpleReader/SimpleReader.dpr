program SimpleReader;

{$APPTYPE CONSOLE}

{$R *.res}

uses
  System.SysUtils,
  Xlsx4D in '..\..\src\Xlsx4D.pas',
  Xlsx4D.Types in '..\..\src\Xlsx4D.Types.pas',
  Xlsx4D.Engine.XLS in '..\..\src\Xlsx4D.Engine.XLS.pas',
  Xlsx4D.Engine.XLSX in '..\..\src\Xlsx4D.Engine.XLSX.pas';

procedure Op1_ListSheets;
var
  Reader: TReader;
  I: Integer;
begin
  Writeln('=== Exemplo 1 ====');
  Writeln;

  Reader := TReader.Create;
  try
    Writeln('Carregando arquivo...');
    Reader.LoadFromFile('example.xlsx');

    Writeln('Arquivo carregado: ', Reader.FileName);
    Writeln('Número de planilhas: ', Reader.WorksheetCount);
    Writeln;

//    for I := 0 to Reader.WorksheetCount-1 do

  finally
    Reader.Free;
  end;
end;

procedure Op2_FindSheetByName;
begin
  
end;

procedure Op3_ReadSpecificCell;
begin
  
end;

procedure Op4_QuickRead;
begin
  
end;

procedure Op5_IterateLines;
begin
  
end;

var
  Option: Integer;
begin
  try
    WriteLn('========================================');
    WriteLn('  XLSReader - Biblioteca de Leitura   ');
    WriteLn('  de Arquivos Excel para Delphi        ');
    WriteLn('========================================');
    WriteLn;

    WriteLn('IMPORTANTE: coloque um arquivo "exemplo.xlsx" no diretório do executável');
    WriteLn('para testar os exemplos');
    WriteLn;
    WriteLn('Escolha um exemplo: ');
    WriteLn('1 - Listar todas as planilhas');
    WriteLn('2 - Buscar planilha por nome');
    WriteLn('3 - Ler célula específica');
    WriteLn('4 - Leitura rápida');
    WriteLn('5 - Iterar por linhas');
    WriteLn('0 - Sair');
    WriteLn;
    Write('Opção: ');
    Readln(Option);
    WriteLn;

    case Option of
      1: Op1_ListSheets;
      2: Op2_FindSheetByName;
      3: Op3_ReadSpecificCell;
      4: Op4_QuickRead;
      5: Op5_IterateLines;
      0: begin end;
    else
      WriteLn('Opção não encontrada');
    end;

    WriteLn;
    WriteLn('Preccione ENTER para sair...');
    Readln;
  except
    on E: Exception do
    begin
      WriteLn('ERRO: ', E.Message);
      WriteLn;
      WriteLn('Preccione ENTER para sair...');
      Readln;
    end;
  end;
end.
