unit UFMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, jpeg, ExtCtrls, ComCtrls, ExtDlgs;

type
  TFMain = class(TForm)
    Img : TImage;
    BCreateExcel : TButton;
    pbCurrentTask : TProgressBar;
    OPD : TOpenPictureDialog;
    pbAll : TProgressBar;
    procedure FormCanResize(Sender : TObject; var NewWidth, NewHeight : Integer; var Resize : Boolean);
    procedure BCreateExcelClick(Sender : TObject);
    procedure ImgDblClick(Sender : TObject);
    procedure FormActivate(Sender : TObject);
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
  ComObj, UBitmapFunctions, Math;

type
  TRGBArray = ARRAY [0 .. 32767] OF TRGBTriple;
  pRGBArray = ^TRGBArray;

procedure TFMain.BCreateExcelClick(Sender : TObject);
var
  Excel : Variant;
  FileName : string;
  i, j : Integer;
begin
  pbCurrentTask.Position := 0;
  pbCurrentTask.Step := 1;
  pbCurrentTask.Max := Img.Picture.Height * Img.Picture.Width;

  pbAll.Position := 0;
  pbAll.Step := 1;
  pbAll.Max := 3;

  FileName := ExtractFilePath(Application.ExeName) + 'test.xlsx';
  Excel := CreateOleObject('Excel.Application');
  Excel.Workbooks.Open[FileName];

  pbCurrentTask.Position := 0;
  pbCurrentTask.Step := 1;
  pbCurrentTask.Max := Img.Picture.Height;
  for i := 0 to Img.Picture.Height - 1 do
  begin
    Excel.Rows[i + 1].RowHeight := 3.75;
    pbCurrentTask.StepIt;
    Application.ProcessMessages;
  end;
  pbAll.StepIt;
  Application.ProcessMessages;

  pbCurrentTask.Position := 0;
  pbCurrentTask.Step := 1;
  pbCurrentTask.Max := Img.Picture.Width;
  for i := 0 to Img.Picture.Width - 1 do
  begin
    Excel.Columns[i + 1].ColumnWidth := 0.42;
    pbCurrentTask.StepIt;
    Application.ProcessMessages;
  end;
  pbAll.StepIt;
  Application.ProcessMessages;

  pbCurrentTask.Position := 0;
  pbCurrentTask.Step := 1;
  pbCurrentTask.Max := Img.Picture.Height * Img.Picture.Width;
  for i := 0 to Img.Picture.Height - 1 do
  begin
    for j := 0 to Img.Picture.Width - 1 do
    begin
      Excel.Cells.Item[i + 1, j + 1].Interior.Color := Img.Canvas.Pixels[j, i];
      pbCurrentTask.StepIt;
      Application.ProcessMessages;
    end;
  end;
  pbAll.StepIt;
  Application.ProcessMessages;

  Excel.ActiveWorkbook.Save;
  Excel.ActiveWorkbook.Close;
  Excel.Application.Quit;

  Application.Terminate;
end;

procedure TFMain.FormActivate(Sender : TObject);
begin
  ImgDblClick(nil);
end;

procedure TFMain.FormCanResize(Sender : TObject; var NewWidth, NewHeight : Integer; var Resize : Boolean);
begin
  Resize := false;
end;

procedure TFMain.ImgDblClick(Sender : TObject);
var
  BM : TBitmap;
  h, w : word;
  k : real;
begin
  if OPD.Execute then
  begin
    BM := UBitmapFunctions.LoadFromFile(OPD.FileName);
    h := BM.Height;
    w := BM.Width;
    k := Max(h, w) / 64;
    h := Round(h / k);
    w := Round(w / k);
    UBitmapFunctions.BMResize(BM, w, h);
    Img.Picture.Assign(BM);
    BM.Free;
  end;
end;

end.
