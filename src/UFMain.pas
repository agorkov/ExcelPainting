unit UFMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, jpeg, ExtCtrls;

type
  TFMain = class(TForm)
    Img : TImage;
    BCreateExcel : TButton;
    procedure FormCanResize(Sender : TObject; var NewWidth, NewHeight : Integer; var Resize : Boolean);
    procedure BCreateExcelClick(Sender : TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FMain : TFMain;

implementation

{$R *.dfm}

uses
  ComObj;

procedure TFMain.BCreateExcelClick(Sender : TObject);
var
  Excel : Variant;
  FileName : string;
begin
  FileName := ExtractFilePath(Application.ExeName) + 'test.xlsx';
  Excel := CreateOleObject('Excel.Application');
  Excel.Workbooks.Open[FileName];

  Excel.Range['a1'].ColumnWidth := 100;
  Excel.Range['a1'].RowHeight := 100;

  Excel.ActiveWorkbook.Save;
  Excel.ActiveWorkbook.Close;
  Excel.Application.Quit;
end;

procedure TFMain.FormCanResize(Sender : TObject; var NewWidth, NewHeight : Integer; var Resize : Boolean);
begin
  Resize := false;
end;

end.
