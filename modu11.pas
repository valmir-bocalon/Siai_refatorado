unit modu11;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs;

type
  TForm2 = class(TForm)

  private
  function Modulo11(Valor: String; Base: Integer = 9; Resto : boolean = false) : string;
  Function Mod11( iNumero: Integer ) : Integer;
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Form2: TForm2;

implementation
uses
  uRuntimeFields;

{$R *.dfm}
Function Mod11( iNumero: Integer ) : Integer;
var   sCadeia  : String;
      iX       : Integer;
      iY       : Integer;
      iValor   : Integer;
      iDigito  : Integer;
      sPosicao : String;
      iPosicao : Integer;
 
Begin
   iValor := 0;
   sCadeia := StrZero( iNumero, 16 );
   For iY := 2 DownTo 1 do
   Begin
      For iX := 8 DownTo 1 do
      Begin
         sPosicao    := Copy( sCadeia, ( 17 - ( iX + ( 8 * ( iY - 1 ) ) ) ), 1 );
         iPosicao    := StrToInt( sPosicao );
         iValor      := iValor + ( iPosicao * ( iX + 1 ) )
      End;
   End;
 
   iDigito := ( ( iValor * 10 ) mod 11 );
 
   If iDigito >= 10 Then
      iDigito := 0;
 
   Result := StrToInt(StrZero( iDigito, 1 ));
 
End;

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


procedure TForm2.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
