unit calculo_mod_11;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TForm1 = class(TForm)
    Button1: TButton;
    xnosso: TEdit;
    nosso: TEdit;
    procedure Button1Click(Sender: TObject);
//    function inttostrZero(numero: Double; tamanho : Integer) : string;
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation
uses
  uRuntimeFields;

{$R *.dfm}
function Modulo11(Valor: String; Base: Integer = 9; Resto : boolean = false) : string;
var
   Soma : integer;
   Contador, Peso, Digito : integer;
begin
   Soma := 0;
   Peso := 2;
   for Contador := Length(Valor) downto 1 do
   begin
      Soma := Soma + (StrToInt(Valor[Contador]) * Peso);
      if Peso < Base then
         Peso := Peso + 1
      else
         Peso := 2;
   end;

   if Resto then
      Result := IntToStr(Soma mod 11)
   else
   begin
      Digito := 11 - (Soma mod 11);
      if (Digito > 9) then
         Digito := 0;
      Result := IntToStr(Digito);
   end
end;

function inttostrZero(numero: Double; tamanho : integer) : string;
Var
  Varvezes, Xtama : integer;
  Varzeros : String;
Begin
  varzeros := '';
  Xtama := length(floattostr(numero));
  for varvezes := Xtama+1 to tamanho do
    Varzeros := Varzeros + '0';
  result := Varzeros+floattostr(numero);
End;

procedure TForm1.Button1Click(Sender: TObject);
var
 varnossnum:string;
 vary,varx:integer;
begin
 //   varnossnum:='32765432765432';
//   varnossnum:='2765432';
   varnossnum:='234567892345';
   vary := 0;
   for varx := 1 to 12 do
       vary := vary + (strtoint(copy(varnossnum,varx,1)) * strtoint(copy(trim(xnosso.Text),13-varx,1)));
       vary := (vary mod 11);
//        vary := strtoint(Modulo11(xnosso.Text,9));
       if vary = 10 Then
           varnossnum := inttostrZero(strtofloat(trim(xnosso.text)),7)+'1'
       else if vary = 1 Then
          varnossnum := inttostrZero(strtofloat(trim(xnosso.text)),7)+'0'
       else if vary = 0 Then
          varnossnum := inttostrZero(strtofloat(trim(xnosso.text)),7)+'0'
       else
       begin
         vary := 11 - vary ;
          varnossnum := inttostrZero(strtofloat(trim(xnosso.text)),7)+inttostr(vary);
       end;
       nosso.Text := varnossnum;
end;


procedure TForm1.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
