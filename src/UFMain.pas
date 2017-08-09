unit UFMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs;

type
  TFMain = class(TForm)
    procedure FormCanResize(Sender : TObject; var NewWidth, NewHeight : Integer; var Resize : Boolean);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FMain : TFMain;

implementation

{$R *.dfm}

procedure TFMain.FormCanResize(Sender : TObject; var NewWidth, NewHeight : Integer; var Resize : Boolean);
begin
  Resize := false;
end;

end.
