unit UPochta;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.ExtCtrls,
  Vcl.Menus;

type
  TFPochta = class(TForm)
    pnl1: TPanel;
    Button1: TButton;
    dtp1: TDateTimePicker;
    lst1: TListBox;
    redt1: TRichEdit;
    spl1: TSplitter;
    pm1: TPopupMenu;
    N2: TMenuItem;
    N1: TMenuItem;
    procedure Button1Click(Sender: TObject);
    procedure lst1DblClick(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure N1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FPochta: TFPochta;

implementation

uses
  Main, UnewBRAK, mainBRAK;

{$R *.dfm}

procedure TFPochta.Button1Click(Sender: TObject);
var
  SR: TSearchRec;
  FindRes: Integer;
 S,Dir:string;
H:HWND;
begin
    S := ExtractFileDir(ParamStr(0));
    Dir:=S+'\Почта\';
  lst1.Clear;

  FindRes := FindFirst(Dir+'\*.*', faAnyFile, SR);
  while FindRes = 0 do
  begin
    if ((SR.Attr and faDirectory) = faDirectory) and
      ((SR.Name = '.') or (SR.Name = '..')) then
    begin
      FindRes := FindNext(SR);
      Continue;
    end;
    // если у файла (каталога) дата создания меньше,
    // чем установлено в DateTimePicker1, то
    if FileDateToDateTime(SR.Time) < dtp1.Date then
    begin
      FindRes := FindNext(SR); // продолжить поиск
      Continue; // продолжить цикл
    end;

    lst1.Items.Add(SR.Name);
    FindRes := FindNext(SR);
  end;
  FindClose(SR);
end;

procedure TFPochta.FormShow(Sender: TObject);
begin
     dtp1.DateTime:=Now-1;
     Button1.Click;
end;

procedure TFPochta.lst1DblClick(Sender: TObject);
Var
  S,Dir,S1:string;
  H:HWND;
begin
    S1:=  lst1.Items[lst1.ItemIndex];
    S := ExtractFileDir(ParamStr(0));
    Dir:=S+'\Почта\';
    redt1.Lines.LoadFromFile(Dir+lst1.Items[lst1.ItemIndex]);
    FNewBRAK.Put:=Dir+lst1.Items[lst1.ItemIndex];
end;

procedure TFPochta.N1Click(Sender: TObject);
var S,S1,Idgp,Idko,id:string;
I,Res1:Integer;
begin
{
  Id:=StrToInt(Label13.Caption);
  Idgp:=StrToInt(Labgp.Caption);
  Idko:=StrToInt(Labko.Caption);}
          for I := 0 to redt1.Lines.Count-1 do
          Begin
             Id:= redt1.Lines[i];
             Res1:=Pos('ID: ',Id);
             if Res1<>0 then
             Begin
                Delete(Id,1,4);
                FNewBRAK.Label13.Caption:=id;
             End;
             //
             Idgp:= redt1.Lines[i];
             Res1:=Pos('Idgp:',Idgp);
             if Res1<>0 then
             Begin
                Delete(Idgp,1,6);
                FNewBRAK.Labgp.Caption:=Idgp;
             End;
             //
             Idko:= redt1.Lines[i];
             Res1:=Pos('Idko:',Idko);
             if Res1<>0 then
             Begin
                Delete(Idko,1,6);
                FNewBRAK.Labko.Caption:=Idko;
             End;
          End;
         // FNewBRAK.LabIndex.Caption:='1'; //Клапана
          FBRAK.IZD_G:=0;
          FNewBRAK.btn1.Enabled:=True;
          FNewBRAK.ShowModal;

end;

procedure TFPochta.N2Click(Sender: TObject);
Var I,Res1:Integer;
S,S1:string;
begin
        for I := 0 to redt1.Lines.Count-1 do
          Begin
             S1:= redt1.Lines[i];
             Res1:=Pos('ЗавНомер: ',S1);
             if Res1<>0 then
             Begin
                Delete(S1,1,10);
                Break;
             End;
          End;
          S:= Form1.ADOQuery1.ConnectionString;
        Form1.ADOQuery1.Close;
        Form1.ADOQuery1.SQL.Clear;
        Form1.ADOQuery1.SQL.Text :=
                'Select * from Запуск Where  НачНомер='+#39+S1+#39;
        Form1.ADOQuery1.Open;
        if Form1.ADOQuery1.RecordCount<>0 then
        Begin
              Form1.PageControl1.ActivePageIndex:=1;
              Form1.Edit31.Text:=S1;
              Form1.Button36.Click;
        End
        else
        Begin
            Form1.ADOQuery1.Close;
            Form1.ADOQuery1.SQL.Clear;
            Form1.ADOQuery1.SQL.Text :=
                'Select * from ЗапускВозд Where  НачНом='+#39+S1+#39;
            Form1.ADOQuery1.Open;
            if Form1.ADOQuery1.RecordCount<>0 then
            Begin
                Form1.PageControl1.ActivePageIndex:=3;
                Form1.Edit32.Text:=S1;
                Form1.Button46.Click;
            End
        End;


end;

end.
