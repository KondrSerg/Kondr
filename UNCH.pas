unit UNCH;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.StdCtrls;

type
  TFnch = class(TForm)
    pnl1: TPanel;
    Button1: TButton;
    Button2: TButton;
    ComboBox1: TComboBox;
    Label1: TLabel;
    ComboBox2: TComboBox;
    rg1: TRadioGroup;
    Edit1: TEdit;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Edit2: TEdit;
    Edit3: TEdit;
    Edit4: TEdit;
    Memo1: TMemo;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Edit5: TEdit;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Fnch: TFnch;

implementation

uses
  Main;

{$R *.dfm}

procedure TFnch.Button1Click(Sender: TObject);
var i:Integer;
S,ISP,per,Priv,Lop,Fl,SB,P_SB:string;
begin
      //--------------------------------------------------
         if ComboBox1.ItemIndex=-1 then
         MessageDlg('Выберите исполнение!', mtError,[mbOk], 0);
         if ComboBox1.ItemIndex=0 then
         Isp:='o';
         if ComboBox1.ItemIndex=1 then
         Isp:='d';
         //__________________________________________________
         Per:=ComboBox2.Text;
         Fl:= IntToStr(rg1.ItemIndex);
         Lop :=Edit1.Text;
         Priv:=Edit2.Text;

         SB :=StringReplace(Edit3.Text, ',', '.', [rfReplaceAll]);
         P_Sb:=StringReplace(Edit4.Text, ',', '.', [rfReplaceAll]);
         I:=StrToInt(Label5.Caption);
         if I>=0 then

         if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [сборка]=' +
         #39 + Sb + #39 +',[подсборка1]='+
         #39 + P_Sb + #39 +
         ' WHERE  (исполнение='+#39+Isp+#39+') and (КолФланцев='+#39+Fl+#39+')'+
         ' AND (колЛопаток='+#39+Lop+#39+')'+
         ' AND (периметр='+#39+Per+#39+') AND (колПривод='+#39+Priv+#39+')',
         ['КПУ']) then
         Exit;
         if I=-1 then
         begin
          if not Form1.mkQueryInsert(Form1.ADOQuery1, 'Insert Into %s ' +
          '(исполнение, КолФланцев, [колЛопаток], [периметр], [колПривод], [сборка], [подсборка1], [Прим]) ' +
          'Values (%s,%s,%s,%s,%s,%s,%s,%s)',
            ['КПУ', #39 + Isp + #39, #39 + Fl + #39, #39 + Lop + #39, #39 + Per + #39
            , #39 + Priv + #39, #39 + SB + #39, #39 + P_Sb + #39, #39 + Edit5.Text + #39]) then
            exit;
            MessageDlg('Ok!', mtError,[mbOk], 0);
         end;
end;

procedure TFnch.Button2Click(Sender: TObject);
var i:Integer;
S,ISP,per,Priv,Lop,Fl:string;
begin
         Memo1.Lines.Clear;
         //--------------------------------------------------
         if ComboBox1.ItemIndex=-1 then
         MessageDlg('Выберите исполнение!', mtError,[mbOk], 0);
         if ComboBox1.ItemIndex=0 then
         Isp:='o';
         if ComboBox1.ItemIndex=1 then
         Isp:='d';
         //__________________________________________________
         Per:=ComboBox2.Text;
         Fl:= IntToStr(rg1.ItemIndex);
         Lop :=Edit1.Text;
         Priv:=Edit2.Text;
         S:=' Where (исполнение='+#39+Isp+#39+') and (КолФланцев='+#39+Fl+#39+')'+
         ' AND (колЛопаток='+#39+Lop+#39+')'+
         ' AND (периметр='+#39+Per+#39+') AND (колПривод='+#39+Priv+#39+')';

         if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s ' +
         S, ['КПУ']) then
                exit;
         if Form1.ADOQuery1.RecordCount<>0 then
         for I := 0 to Form1.ADOQuery1.RecordCount-1 do
         Begin
             Memo1.Lines.Add(Form1.ADOQuery1.FieldByName('сборка').AsString+'   '
             +Form1.ADOQuery1.FieldByName('подсборка1').AsString);

             Edit3.Text:=Form1.ADOQuery1.FieldByName('сборка').AsString;
             Edit4.Text:=Form1.ADOQuery1.FieldByName('подсборка1').AsString;
             Label5.Caption:=Form1.ADOQuery1.FieldByName('Прим').AsString;
             Form1.ADOQuery1.Next;
         End;
         if Form1.ADOQuery1.RecordCount=0 then
         Begin
           Edit3.Text:='0';
           Edit4.Text:='0';
           Label5.Caption:='-1';
           if MessageDlg('Норм для таких условий не существует!#13#10Введите нормы, и нажмите - "Сохранить"',
           mtConfirmation, [ mbYes, mbCancel ], 0 ) = mrYes then
           Edit3.SelectAll
           Else
           exit
         End;
end;

procedure TFnch.FormShow(Sender: TObject);
var I:Integer;
begin
         ComboBox2.Clear;
         Memo1.Lines.Clear;
         if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select DISTINCT Периметр from %s '
         , ['КПУ']) then
         exit;
         for I := 0 to Form1.ADOQuery1.RecordCount-1   do
         Begin
           ComboBox2.Items.Add(Form1.ADOQuery1.FieldByName('Периметр').AsString);
           Form1.ADOQuery1.Next;
         End;
end;

end.
