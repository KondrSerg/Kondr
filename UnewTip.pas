unit UnewTip;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.Menus;

type
  TFNewTip = class(TForm)
    Edit1: TEdit;
    btn1: TButton;
    btn2: TButton;
    Memo1: TMemo;
    Label2: TLabel;
    rg1: TRadioGroup;
    Edit2: TEdit;
    Label1: TLabel;
    Label3: TLabel;
    rg2: TRadioGroup;
    Edit3: TEdit;
    Edit4: TEdit;
    Memo2: TMemo;
    btn3: TButton;
    Label4: TLabel;
    Label5: TLabel;
    pm1: TPopupMenu;
    N1: TMenuItem;
    btn4: TButton;
    procedure FormShow(Sender: TObject);
    procedure btn1Click(Sender: TObject);
    procedure btn2Click(Sender: TObject);
    procedure btn3Click(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure btn4Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FNewTip: TFNewTip;

implementation

uses
  Main, UIzv;

{$R *.dfm}

procedure TFNewTip.btn1Click(Sender: TObject);
begin
  if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s '+
  'WHERE (Обозначение='+#39+Edit1.Text+#39+')',
  ['Тип']) then
  exit;
  if Form1.ADOQuery1.RecordCount=0 then
  begin
    if not Form1.mkQueryInsert(Form1.ADOQuery2,
    'Insert Into %s ' +
    '([Обозначение],ID,Код) ' +
    'Values (%s,%s,%s)',
    ['Тип',#39 + Edit1.Text + #39,#39 +Label2.Caption+ #39,#39 + Edit2.Text + #39 ])
    then
    exit;
    //Close;
     btn4.Click;
  end
  else
  Begin
    MessageDlg('Такой клапан уже есть!',mtInformation,[mbOk],0)
  End;


end;

procedure TFNewTip.btn2Click(Sender: TObject);
begin
  Close;
end;

procedure TFNewTip.btn3Click(Sender: TObject);
begin
  if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s '+
  'WHERE (Признак='+#39+Edit3.Text+#39+')',
  ['Признаки']) then
  exit;
  if Form1.ADOQuery1.RecordCount=0 then
  begin
    if not Form1.mkQueryInsert(Form1.ADOQuery2,
    'Insert Into %s ' +
    '([Признак],ID,Описание) ' +
    'Values (%s,%s,%s)',
    ['Признаки',#39 + Edit3.Text + #39,#39 +Label5.Caption+ #39,#39 + Edit4.Text + #39 ])
    then
    exit;
    btn4.Click;
    //Close;
  end
  else
  Begin
    MessageDlg('Такой Признак уже есть!',mtInformation,[mbOk],0);
  End;
end;

procedure TFNewTip.btn4Click(Sender: TObject);
var Id,I:Integer;
S,id_S:string;
begin
  Memo1.Lines.Clear;
  if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s ',
  ['Тип']) then
  exit;
  for I := 0 to Form1.ADOQuery1.RecordCount - 1 do
  begin
    ID_S := Form1.ADOQuery1.FieldByName('Id').AsString;
    S    := Form1.ADOQuery1.FieldByName('Обозначение').AsString;
    Memo1.Lines.Add(ID_S+' '+S);
    Form1.ADOQuery1.Next;
  end;
    if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select MAX(ID) AS I from %s ',
  ['Тип']) then
  exit;
  ID := Form1.ADOQuery1.FieldByName('I').AsInteger;
  Label2.Caption:=IntToStr(ID+1);
  //______________________________________________________________________

  Memo2.Lines.Clear;
  if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s ',
  ['Признаки']) then
  exit;
  for I := 0 to Form1.ADOQuery1.RecordCount - 1 do
  begin
    ID_S := Form1.ADOQuery1.FieldByName('Id').AsString;
    S    := Form1.ADOQuery1.FieldByName('Признак').AsString;
    Memo2.Lines.Add(ID_S+' '+S);
    Form1.ADOQuery1.Next;
  end;
    if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select MAX(ID) AS I from %s ',
  ['Признаки']) then
  exit;
  ID := Form1.ADOQuery1.FieldByName('I').AsInteger;
  Label5.Caption:=IntToStr(ID+1);
end;

procedure TFNewTip.FormShow(Sender: TObject);

begin
   btn4.Click;
end;

procedure TFNewTip.N1Click(Sender: TObject);
var S:string;
LineNumber,ID,k:Integer;
begin
  LineNumber := SendMessage(Memo2.Handle, EM_LINEFROMCHAR, Memo2.SelStart,0);
  S:=Memo2.Lines.Strings[LineNumber];
  K:=Pos(' ',S);
  if K<>0 then
     Delete(S,K,200);
  ID:=StrToInt(S );
  if not Form1.mkQueryDelete(Form1.ADOQuery1, 'DELETE FROM %s Where (Id= ' + #39 +
  IntToStr(ID)+ #39 + ') ',
  ['Признаки']) then
      Exit;
       btn4.Click;
end;

end.
