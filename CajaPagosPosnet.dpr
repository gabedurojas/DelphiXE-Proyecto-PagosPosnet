program CajaPagosPosnet;

uses
  Forms,
  Principal in 'Principal.pas' {Principal_Menu},
  Modulo in 'Modulo.pas' {Modulo2: TDataModule},
  Unidad in 'Unidad.pas',
  Acceso1 in 'Acceso1.pas' {Acceso};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TModulo2, Modulo2);
  Application.CreateForm(TPrincipal_Menu, Principal_Menu);
  Application.Run;
end.
