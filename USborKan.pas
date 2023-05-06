unit USborKan;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls;

type
  TFSborKan = class(TForm)
    CBB1: TComboBox;
    Button1: TButton;
    Button2: TButton;
    Label1: TLabel;
    Label2: TLabel;
    procedure Button2Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FSborKan: TFSborKan;

implementation

uses
  Main;

{$R *.dfm}

procedure TFSborKan.Button1Click(Sender: TObject);
begin
  Close;
end;

procedure TFSborKan.Button2Click(Sender: TObject);
begin
            if not Form1.mkQueryUpdate(Form1.ADOQuery1,          //    ,[%s]=%s,[%s]=%s,[%s]=%s
                'UPDATE %s SET [%s]=%s WHERE ([IDГП]=%s) AND ([IDКО]=%s)',
                ['[Запуск750]', 'Сборщик', #39 + CBB1.Text
                + #39,
                #39 + Label1.Caption + #39,
                #39 + Label2.Caption + #39]) then
                Exit;
                Close;
end;

procedure TFSborKan.FormShow(Sender: TObject);
var I:Integer;
begin

               if not Form1.mkQuerySelect(Form1.ADOQuery2,
                'Select * from %s WHERE МаркаАвто='+#39+'Канал'+#39+' Order by Фамилия',
                ['Сборщик']) then
                exit;
        CBB1.Items.Clear;
        CBB1.Items.Add('');
        CBB1.DropDownCount := 30;

        for I := 0 to Form1.ADOQuery2.RecordCount - 1 do
        begin
                CBB1.Items.Add(Form1.ADOQuery2.FieldByName('Фамилия').AsString);
                Form1.ADOQuery2.Next;
        end
end;

end.
