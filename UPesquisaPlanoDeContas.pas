unit UPesquisaPlanoDeContas;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms, 
  Dialogs, StdCtrls, Grids, DBGrids;

type
  TFramePesquisaPlanoDeContas = class(TFrame)
    DBGrid1: TDBGrid;
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

implementation

uses Coneccoes, Ufuncoes, uRuntimeFields;
{$R *.dfm}


procedure TFramePesquisaPlanoDeContas.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
