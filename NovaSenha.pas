unit NovaSenha;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, JvExControls, JvButton, JvNavigationPane, ExtCtrls,
  JvExExtCtrls, StdCtrls;

type
  TFrm_NovaSenha = class(TForm)
    Label1: TLabel;
    JvNavPanelDivider1: TJvNavPanelDivider;
    Label2: TLabel;
    JNPBImediatamente: TJvNavPanelButton;
    JNPBEsperar: TJvNavPanelButton;
    Label3: TLabel;
    Label4: TLabel;
    procedure JNPBImediatamenteClick(Sender: TObject);
    procedure JNPBEsperarClick(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_NovaSenha: TFrm_NovaSenha;

implementation

uses tabelas, uRuntimeFields;

{$R *.dfm}

procedure TFrm_NovaSenha.JNPBImediatamenteClick(Sender: TObject);
begin
  DM_Tabelas.ZQLoc_Senha.Edit;
  DM_Tabelas.ZQLoc_Senha.FieldByName('pedesenha').AsString := 'S';
  DM_Tabelas.ZQLoc_Senha.Post;
  Close;
end;

procedure TFrm_NovaSenha.JNPBEsperarClick(Sender: TObject);
begin
  Close;
end;


procedure TFrm_NovaSenha.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
