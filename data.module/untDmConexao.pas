{*******************************************************************************}
{ Projeto: Gerador de API                                                       }
{                                                                               }
{ O objetivo da aplicação é facilitar a criação de Interface, model e controller}
{ para Insert, Update, Delete e Select a partir de tabelas do banco de dados    }
{ (Postgres ou Firebird), respeitando a tipagem, PK e FK                        }
{*******************************************************************************}
{                                                                               }
{ Desenvolvido por JANDERSON APARECIDO DA SILVA                                 }
{ Email: janderson_rm@hotmail.com                                               }
{                                                                               }
{*******************************************************************************}

unit untDmConexao;

interface

uses
  Data.DB,
  FireDAC.Comp.Client,
  FireDAC.Comp.UI,
  FireDAC.Phys,
  FireDAC.Phys.FB,
  FireDAC.Phys.FBDef,
  FireDAC.Phys.IBBase,
  FireDAC.Phys.Intf,
  FireDAC.Phys.PG,
  FireDAC.Phys.PGDef,
  FireDAC.Stan.Async,
  FireDAC.Stan.Def,
  FireDAC.Stan.Error,
  FireDAC.Stan.Intf,
  FireDAC.Stan.Option,
  FireDAC.Stan.Pool,
  FireDAC.UI.Intf,
  FireDAC.VCLUI.Wait,
  System.Classes,
  System.SysUtils;

type
  TfrmDmConexao = class(TDataModule)
    FDConnection: TFDConnection;
    FDPhysPgDriverLink1: TFDPhysPgDriverLink;
    FDPhysFBDriverLink1: TFDPhysFBDriverLink;
    FDGUIxWaitCursor1: TFDGUIxWaitCursor;
  private
    { Private declarations }
  public
    { Public declarations }
    function Conectar(DriverID,
                      Database,
                      User_name,
                      Password,
                      Port,
                      Server,
                      VendorLib: string): Boolean;
  end;

var
  frmDmConexao: TfrmDmConexao;

implementation

{%CLASSGROUP 'Vcl.Controls.TControl'}

{$R *.dfm}

{ TfrmDmConexao }

function TfrmDmConexao.Conectar(DriverID,
                                Database,
                                User_name,
                                Password,
                                Port,
                                Server,
                                VendorLib: string): Boolean;
begin
  FDConnection.Connected := False;
  FDConnection.Params.Clear;
  FDConnection.Params.Values['DriverID']  := DriverID;
  FDConnection.Params.Values['Database']  := Database;
  FDConnection.Params.Values['User_name'] := User_name;
  FDConnection.Params.Values['Password']  := Password;
  FDConnection.Params.Values['Port']      := Port;
  FDConnection.Params.Values['Server']    := Server;

  if DriverID = 'FB' then
    FDPhysFBDriverLink1.VendorLib := VendorLib
  else
  if DriverID = 'PG' then
    FDPhysPgDriverLink1.VendorLib := VendorLib;

  try
    FDConnection.Connected := True;
    if FDConnection.Connected then
      Result := True
    else
      Result := False;
  except
    Result := False;
  end;
end;

end.
