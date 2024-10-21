unit untPrincipal;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Imaging.pngimage, Vcl.ExtCtrls,
  Vcl.StdCtrls, Vcl.CheckLst, FireDAC.Stan.Intf, FireDAC.Stan.Option,
  FireDAC.Stan.Param, FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf,
  FireDAC.DApt.Intf, FireDAC.Stan.Async, FireDAC.DApt, Data.DB,
  FireDAC.Comp.DataSet, FireDAC.Comp.Client, System.Win.ComObj, GraphicEx,
  Vcl.ComCtrls;

type
  TfrmPrincipal = class(TForm)
    pgcPrincipal: TPageControl;
    tsConectar: TTabSheet;
    tsExportar: TTabSheet;
    pnlConexao: TPanel;
    lblDriverID: TLabel;
    lblDatabase: TLabel;
    lblUser_name: TLabel;
    lblPassword: TLabel;
    lblPort: TLabel;
    lblServer: TLabel;
    lblVendorLib: TLabel;
    edtDatabase: TEdit;
    cbxDriverID: TComboBox;
    edtUser_name: TEdit;
    edtPassword: TEdit;
    edtPort: TEdit;
    edtServer: TEdit;
    edtVendorLib: TEdit;
    pnlLocDatabase: TPanel;
    imgLocDatabase: TImage;
    pnlLocVendorLib: TPanel;
    imgLocVendorLib: TImage;
    pnlConectar: TPanel;
    imgConectar: TImage;
    Label1: TLabel;
    OpenDialog: TOpenDialog;
    qrTable: TFDQuery;
    qrRegistros: TFDQuery;
    save: TSaveDialog;
    pnlTable: TPanel;
    chkListaTabelas: TCheckListBox;
    pnlMarcar: TPanel;
    pnlDesmarcarTodos: TPanel;
    imgDesmarcarTodos: TImage;
    lblDesmarcarTodos: TLabel;
    pnlMarcarTodos: TPanel;
    imgMarcarTodos: TImage;
    lblMarcarTodos: TLabel;
    Panel2: TPanel;
    pnlExportXLS: TPanel;
    Image1: TImage;
    Label2: TLabel;
    pnlExportCSV: TPanel;
    Image2: TImage;
    Label3: TLabel;
    pnlMensagem: TPanel;
    procedure cbxDriverIDSelect(Sender: TObject);
    procedure imgLocDatabaseClick(Sender: TObject);
    procedure imgLocVendorLibClick(Sender: TObject);
    procedure pnlConectarClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure pnlMarcarTodosClick(Sender: TObject);
    procedure pnlDesmarcarTodosClick(Sender: TObject);
    procedure pnlExportCSVClick(Sender: TObject);
    procedure pnlExportXLSClick(Sender: TObject);
  private
    { Private declarations }
    function EliminaBrancos(sTexto:String):String;
    procedure PreencherCamposConexao;
    procedure CarregarListaTable;
    procedure LocalizarTableFB;
    procedure LocalizaRegistros(Table: string);
    procedure excel(const query: TFDQuery; const exibir_titulo: Boolean;const caminho_arquivo: string);
    procedure csv(const table: TFDQuery; const exibir_titulo: Boolean;const caminho_arquivo: string);
    procedure Exportar(tipo: string);
    procedure Menssagem(txt: string);
    function ContarQtdMarcada: Integer;
  public
    { Public declarations }
  end;

var
  frmPrincipal: TfrmPrincipal;

implementation

uses
  untDmConexao;

{$R *.dfm}

{ TForm1 }

procedure TfrmPrincipal.CarregarListaTable;
var
  i : Integer;
begin
  if cbxDriverID.Text = 'FB' then
    LocalizarTableFB;

  chkListaTabelas.Items.Clear();
  qrTable.First;
  while not qrTable.Eof do
  begin
    chkListaTabelas.Items.Add(EliminaBrancos(qrTable.Fields[0].AsString));
    qrTable.Next;
  end;

  for I := 0 to chkListaTabelas.Count -1 do
  begin
    chkListaTabelas.Checked[i] := True;
  end;
end;

procedure TfrmPrincipal.cbxDriverIDSelect(Sender: TObject);
begin
  PreencherCamposConexao;
end;

function TfrmPrincipal.ContarQtdMarcada: Integer;
var
  Qtd, i: Integer;
begin
  Qtd := 0;

  for i := 0 to chkListaTabelas.Count -1 do
  begin
    if chkListaTabelas.Checked[i] then
    begin
      Qtd := Qtd + 1;
    end;
  end;

  Result := Qtd;
end;

procedure TfrmPrincipal.csv(const table: TFDQuery; const exibir_titulo: Boolean; const caminho_arquivo: string);
var
  SLISTA: TStringList;
  NCAMPO: Integer;
  CLinha, patch: string;
begin
  try
    patch := caminho_arquivo;
    try
      SLISTA := TStringList.Create;
      SLISTA := TStringList.Create;
      CLinha := '';
      for NCAMPO := 0 to table.Fields.Count - 1 do
        if table.Fields[NCAMPO].visible = True then
          CLinha := CLinha + table.Fields[NCAMPO].DisplayLabel + ';';
      SLISTA.Add(CLinha);
      table.First;
      while not table.Eof do
      begin
        CLinha := '';
        for NCAMPO := 0 to table.Fields.Count - 1 do
          if table.Fields[NCAMPO].visible = True then
            CLinha := CLinha + table.Fields[NCAMPO].DisplayText + ';';
        SLISTA.Add(CLinha);
        table.Next;
      end;
      SLISTA.SaveToFile(patch);
    finally
      SLISTA.Free;
    end;
  except
    on E: Exception do
    begin
      Application.MessageBox
        (PChar('Não foi possivel gerar o relatório, restição = ' + E.ToString),
        'Atenção', MB_OK + MB_ICONWARNING);
    end;
  end;
end;

function TfrmPrincipal.EliminaBrancos(sTexto: String): String;
// Elimina todos os espaços em branco da string
//(inclusive os espaços entre as palavras)
var
  nPos : Integer;
begin
  nPos := 1;
  while Pos(' ',sTexto) > 0 do
  begin
    nPos := Pos(' ',sTexto);
    (*Text[nPos] := ''; *)
    Delete(sTexto,nPos,1);
  end;
  Result := sTexto;
end;

procedure TfrmPrincipal.excel(const query: TFDQuery; const exibir_titulo: Boolean; const caminho_arquivo: string);
var
  objExcel, Sheet: Variant;
  patch, cTitulo: string;
  linha, coluna: Integer;
begin
  try
    patch := caminho_arquivo;
    // Cria uma instancia para utilizar o Excel
    cTitulo := ExtractFileName(caminho_arquivo);
    objExcel := CreateOleObject('Excel.Application');
    objExcel.visible := True;
    objExcel.Caption := cTitulo;

    // Adiciona planilha(sheet)
    objExcel.Workbooks.Add;
    objExcel.Workbooks[1].Sheets.Add;
    objExcel.Workbooks[1].WorkSheets[1].Name := cTitulo;
    Sheet := objExcel.Workbooks[1].WorkSheets[cTitulo];
    query.First;
    // para inserir o titulo no relatorio
    if exibir_titulo then
    begin
      for coluna := 1 to query.FieldCount do
      begin
        if query.Fields[coluna - 1].visible = True then
        begin
          Sheet.cells[1, coluna] := query.Fields[coluna - 1].DisplayLabel;
        end;
      end;
      Sheet.columns.Autofit;
    end;

    for linha := 2 to query.RecordCount + 1 do
    begin
      for coluna := 1 to query.FieldCount do
      begin
        if query.Fields[coluna - 1].visible = True then
        begin
          case query.Fields[coluna - 1].DataType of
            ftDate, ftDateTime, ftTimeStamp:
              begin
                Sheet.cells[linha, coluna].NumberFormat := 'dd/mm/aaaa';
                Sheet.cells[linha, coluna] := query.Fields[coluna - 1]
                  .AsDateTime;
              end;
          else
            Sheet.cells[linha, coluna] := query.Fields[coluna - 1].AsString;
          end;
        end;
      end;
      query.Next;
    end;

    Sheet.SaveAs(patch);
  except
    on E: Exception do
    begin
      Application.MessageBox
        (PChar('Atenção verifique sua licença do Microsoft office, para geração deste relatório é necessária uma licença. Restrição retornada = '
        + E.ToString), 'Atenção', MB_OK + MB_ICONWARNING);
    end;
  end;
end;

procedure TfrmPrincipal.Exportar(tipo: string);
var
  tabela, patch: string;
  i, RecNo, RecordCount: Integer;
begin
  patch := '';

  if Tipo = 'CSV' then
  begin
    save.FileName := 'exportar.csv';
    save.Filter := '|*.csv';
  end
  else
  if tipo = 'XLS' then
  begin
    save.FileName := 'exportar.xls';
    save.Filter := '|*.xls';
  end;

  if not save.Execute then
    Exit;
  patch := save.FileName;

  RecNo := 0;
  RecordCount := ContarQtdMarcada;

  for i := 0 to chkListaTabelas.Count -1 do
  begin
    if chkListaTabelas.Checked[i] then
    begin
      RecNo := RecNo + 1;
      Menssagem('Exportando '+RecNo.ToString+'/'+RecordCount.ToString);

      tabela := chkListaTabelas.Items[i].Trim;

      LocalizaRegistros(tabela);
      if Tipo = 'CSV' then
        csv(qrRegistros, True, ExtractFileDir(save.FileName) + '/'+tabela+ '.csv')
      else
      if tipo = 'XLS' then
        excel(qrRegistros, True, ExtractFileDir(save.FileName) + '/'+tabela+ '.xls');
    end;
  end;
  Menssagem('Exportação concluída.');
end;

procedure TfrmPrincipal.FormShow(Sender: TObject);
begin
  Menssagem('Aguardando comando para exportação.');
  pgcPrincipal.ActivePage := tsConectar;
  tsExportar.TabVisible := False;
  PreencherCamposConexao;
end;

procedure TfrmPrincipal.imgLocDatabaseClick(Sender: TObject);
begin
  try
    OpenDialog.Execute;
    if Trim(OpenDialog.FileName) <> '' then
      edtDatabase.Text := OpenDialog.FileName;
  finally
    OpenDialog.Files.Clear;
  end;
end;

procedure TfrmPrincipal.imgLocVendorLibClick(Sender: TObject);
begin
  try
    OpenDialog.Execute;
    if Trim(OpenDialog.FileName) <> '' then
      edtVendorLib.Text := OpenDialog.FileName;
  finally
    OpenDialog.Files.Clear;
  end;
end;

procedure TfrmPrincipal.LocalizaRegistros(Table: string);
begin
  qrRegistros.Close;
  qrRegistros.SQL.Clear;
  qrRegistros.SQL.Add('SELECT * FROM ' + Table);
  qrRegistros.Open;
end;

procedure TfrmPrincipal.LocalizarTableFB;
begin
  qrTable.Close;
  qrTable.SQL.Clear;
  qrTable.SQL.Add('SELECT RDB$RELATION_NAME');
  qrTable.SQL.Add('FROM RDB$RELATIONS');
  qrTable.SQL.Add('WHERE RDB$SYSTEM_FLAG = 0'); // Ignora tabelas de sistema
  qrTable.SQL.Add('ORDER BY RDB$RELATION_NAME');
  qrTable.Open;
end;

procedure TfrmPrincipal.Menssagem(txt: string);
begin
  Application.ProcessMessages;
  pnlMensagem.Caption := txt;
  Application.ProcessMessages;
end;

procedure TfrmPrincipal.pnlConectarClick(Sender: TObject);
begin
  if frmDmConexao.Conectar(cbxDriverID.Text, edtDatabase.Text, edtUser_name.Text, edtPassword.Text, edtPort.Text, edtServer.Text, edtVendorLib.Text) then
  begin
    ShowMessage('Conetado com sucesso!');
    tsExportar.TabVisible := True;
    CarregarListaTable;
  end
  else
  begin
    ShowMessage('Falha ao conectar!');
    tsExportar.TabVisible := False;
  end;
end;

procedure TfrmPrincipal.pnlDesmarcarTodosClick(Sender: TObject);
var
  i : integer;
begin
  for I := 0 to chkListaTabelas.Count -1 do
  begin
    chkListaTabelas.Checked[i] := False;
  end;
end;

procedure TfrmPrincipal.pnlExportCSVClick(Sender: TObject);
begin
  Exportar('CSV');
end;

procedure TfrmPrincipal.pnlExportXLSClick(Sender: TObject);
begin
  Exportar('XLS');
end;

procedure TfrmPrincipal.pnlMarcarTodosClick(Sender: TObject);
var
  i : integer;
begin
  for I := 0 to chkListaTabelas.Count -1 do
  begin
    chkListaTabelas.Checked[i] := True;
  end;
end;

procedure TfrmPrincipal.PreencherCamposConexao;
begin
  if cbxDriverID.Text = 'FB' then
  begin
    edtDatabase.Text             := ExtractFileDir(Application.ExeName) + '\DADOS.FDB';
    edtUser_name.Text            := 'sysdba';
    edtPassword.Text             := 'masterkey';
    edtPort.Text                 := '3050';
    edtServer.Text               := 'localhost';
    edtVendorLib.Text            := ExtractFileDir(Application.ExeName) + '\fbclient.dll';
  end;
end;

end.
