unit UFMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, jpeg, ExtCtrls, ComCtrls;

type
  TFMain = class(TForm)
    Img : TImage;
    BCreateExcel : TButton;
    ProgressBar1 : TProgressBar;
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
  i, j : Integer;
begin
  ProgressBar1.Position := 0;
  ProgressBar1.Step := 1;
  ProgressBar1.Max := Img.Picture.Height * Img.Picture.Width;

  FileName := ExtractFilePath(Application.ExeName) + 'test.xlsx';
  Excel := CreateOleObject('Excel.Application');
  Excel.Workbooks.Open[FileName];

  for i := 0 to Img.Picture.Height - 1 do
  begin
    Excel.Rows[i + 1].RowHeight := 3.75;
  end;

  for i := 0 to Img.Picture.Width - 1 do
  begin
    Excel.Columns[i + 1].ColumnWidth := 0.42;
  end;

  for i := 0 to Img.Picture.Height - 1 do
  begin
    for j := 0 to Img.Picture.Width - 1 do
    begin
      Excel.Cells.Item[i + 1, j + 1].Interior.Color := Img.Canvas.Pixels[j, i];
      ProgressBar1.StepIt;
      Application.ProcessMessages;
    end;
  end;

  Excel.ActiveWorkbook.Save;
  Excel.ActiveWorkbook.Close;
  Excel.Application.Quit;

  Application.Terminate;
end;

procedure TFMain.FormCanResize(Sender : TObject; var NewWidth, NewHeight : Integer; var Resize : Boolean);
begin
  Resize := false;
end;

end.
