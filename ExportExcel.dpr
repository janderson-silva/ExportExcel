program ExportExcel;

uses
  Vcl.Forms,
  untPrincipal in 'view\untPrincipal.pas' {frmPrincipal},
  untDmConexao in 'data.module\untDmConexao.pas' {frmDmConexao: TDataModule},
  Vcl.Themes,
  Vcl.Styles;

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  TStyleManager.TrySetStyle('Onyx Blue');
  Application.CreateForm(TfrmPrincipal, frmPrincipal);
  Application.CreateForm(TfrmDmConexao, frmDmConexao);
  Application.Run;
end.
