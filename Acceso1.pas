unit Acceso1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Mask, DB, ADODB, Buttons, ExtCtrls, JvExButtons,
  JvBitBtn;

type
  TAcceso = class(TForm)
    Edit2: TMaskEdit;
    edit1: TEdit;
    GroupBox1: TGroupBox;
    Label4: TLabel;
    Label1: TLabel;
    Image2: TImage;
    Image1: TImage;
    Image3: TImage;
    GroupBox2: TGroupBox;
    Image4: TImage;
    Label2: TLabel;
    Label3: TLabel;
    Button2: TJvBitBtn;
    Boton1: TJvBitBtn;
    procedure FormActivate(Sender: TObject);
    procedure edit1Change(Sender: TObject);
    procedure edit1KeyPress(Sender: TObject; var Key: Char);
    procedure Edit2KeyPress(Sender: TObject; var Key: Char);
    procedure Edit2Exit(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    function GetAppVersion: string;
    procedure VerVersion(nro:byte);
    procedure Button2Click(Sender: TObject);
    procedure Boton1Click(Sender: TObject);
  private
    DataSet1,DataSet2:TADODataSet;
    Com1:TADOCommand;{ Private declarations }
  public
    control:boolean;
    cuenta:integer;
    usuario,dato:string;{ Public declarations }
  end;

var
  Acceso: TAcceso;

implementation

uses modulo;

{$R *.dfm}
procedure TAcceso.FormCreate(Sender: TObject);
begin
DataSet1:=TADODataSet.Create(nil);
DataSet1.ConnectionString:=modulo2.ConectDbMuni;
DataSet2:=TADODataSet.Create(nil);
end;

procedure TAcceso.FormActivate(Sender: TObject);
begin
boton1.Enabled:=false;
control:=false;
Edit2.SetFocus;
cuenta:=1;
label1.Caption:='Version - ' + GetAppVersion;
GroupBox2.Visible :=false;
VerVersion(1);
end;

procedure TAcceso.VerVersion(nro:byte);
 var   Cnx2:TADOConnection;
begin
 Cnx2:=TADOConnection.Create(nil);
 Cnx2.LoginPrompt:=false;
 Cnx2.ConnectionString:=modulo2.ConectDbMuni;
 if nro=1 then begin
  DataSet2:=TADODataSet.Create(nil);
  DataSet2.Close;
  DataSet2.Connection:=Cnx2;
  DataSet2.CommandText:='SELECT NroVersion FROM VersionSoft where (NroAplicacion=11) and (activo=1)';
  DataSet2.Open;
  if not Dataset2.IsEmpty  then dato:=DataSet2.Fields[0].AsString
    else dato:='0';
  DataSet2.Connection:=nil;
 end;
 Cnx2.Close;
 FreeAndNil(Cnx2);
end;

function TAcceso.GetAppVersion: string;
var
    Size, Size2: DWord;
    Pt, Pt2: Pointer;
begin
 Size := GetFileVersionInfoSize(PChar(ParamStr(0)), Size2);
 if Size > 0 then
 begin
 GetMem(Pt, Size);
  try
   GetFileVersionInfo(PChar(ParamStr(0)), 0, Size, Pt);
   VerQueryValue(Pt, '\', Pt2, Size2);
   with TVSFixedFileInfo(Pt2^) do
   begin
    Result := '' +
    IntToStr(HiWord(dwFileVersionMS)) + IntToStr(LoWord(dwFileVersionMS))
    + IntToStr(HiWord(dwFileVersionLS)) + IntToStr(LoWord(dwFileVersionLS));
    //IntToStr(HiWord(dwFileVersionMS)) + '.' +
    //IntToStr(LoWord(dwFileVersionMS)) + '.' +
    //IntToStr(HiWord(dwFileVersionLS));
    //+ ' Build ' + IntToStr(LoWord(dwFileVersionLS));
   end;
  finally
   FreeMem(Pt);
  end;
 end;
end;


procedure TAcceso.edit1Change(Sender: TObject);
begin
if trim(edit1.Text)='' then begin
 boton1.Enabled:=false;
 end
else begin
 boton1.Enabled:=true;
end;
end;

procedure TAcceso.edit1KeyPress(Sender: TObject; var Key: Char);
begin
if (key = chr(13))  then begin
 if boton1.Enabled = true then boton1.SetFocus;
end;
end;

procedure TAcceso.Edit2KeyPress(Sender: TObject; var Key: Char);
begin
if key = chr(13) then edit1.SetFocus;
end;

procedure TAcceso.Edit2Exit(Sender: TObject);
begin
 if button2.Focused then button2.SetFocus
 else begin
  try
   DataSet1.CommandText:='SELECT * FROM Usuarios2 WHERE (codusu =' + trim(edit2.Text) + ');';
   DataSet1.Open;
   if DataSet1.IsEmpty then begin
    dataset1.Close;
   end
   else begin
    if DataSet1.RecordCount = 0 then begin
     messagedlg('No exite el Nº de Operador Ingresado',mterror,[mbok],0);
     edit2.SetFocus;
    end
    else begin
     Label4.Caption:=trim(DataSet1.Fields[1].AsString);
     usuario:=trim(DataSet1.Fields[1].AsString);
     if DataSet1.Fields[3].AsInteger < 1 then cuenta:=1
     else cuenta:=DataSet1.Fields[3].AsInteger;
     modulo2.UsuNro:=DataSet1.Fields[0].AsString;
    end;
   end;
  except
    control:=true;
  end;
  DataSet1.Close;
 end;
end;

procedure TAcceso.FormClose(Sender: TObject; var Action: TCloseAction);
begin
//DataSet1.Free;
 DataSet1.Close;
end;

procedure TAcceso.Button2Click(Sender: TObject);
begin
close;
end;

procedure TAcceso.Boton1Click(Sender: TObject);
var clav,clav1:string;
begin
 if GetAppVersion <> dato then begin
  GroupBox2.Visible:=true;
  boton1.Enabled :=false;
  button2.SetFocus;
 end
 else begin
  GroupBox2.Visible:=false;
  if (cuenta < 4) then begin
   if not(cuenta > 3) then begin
    DataSet1.CommandText:='todos_loguinSinmod ' + trim(Edit2.Text) + ' , ' + chr(39) + trim(Edit1.Text) + chr(39) + ';';
    DataSet1.Open;
    if DataSet1.IsEmpty then begin
     DataSet1.Close;
    end
    else begin
     clav1:=DataSet1.Fields[0].AsString;
     DataSet1.Close;
     if clav1 <> 'OK' then begin
      messagedlg( clav1 + '!!!'+chr(13)+chr(13)+'   Acceso DENEGADO!!!',mterror,[mbok],0);
      modulo2.UsuNvl:=0;
      cuenta:=cuenta+1;
      edit1.SetFocus;
     end
     else begin
      control:=true;
      modulo2.UsuNro:=trim(Edit2.Text);
      modulo2.UsuNombre:=usuario;
      close;
     end;
    end;
   end
   else begin
    cuenta:=cuenta + 1;
    messagedlg('El Nº de Operador ha Caducado su Cuenta'+chr(13)+chr(13)
    +'   Acceso DENEGADO!!!'+chr(13)+'Consulte con el Administrador',mterror,[mbok],0);
    close;
   end;
  end
  else begin
   Com1:=TADOCommand.Create(nil);
   Com1.ConnectionString:=modulo2.ConectDbMuni;
   Com1.CommandText:='UPDATE Usuarios2 SET intentos =' + chr(39) + inttostr(cuenta+1) + chr(39) +
   'WHERE (codusu =' + modulo2.UsuNro + ')';
   Com1.Execute;
   Com1.Free;
   close;
  end;
 end;
end;

end.
