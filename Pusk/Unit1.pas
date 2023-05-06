unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, ActnList;

type
  TFPusk = class(TForm)
    Image1: TImage;
    Timer1: TTimer;
    actlst1: TActionList;
    act1: TAction;
    procedure FormShow(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure act1Execute(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FPusk: TFPusk;

implementation

{$R *.dfm}

procedure TFPusk.FormShow(Sender: TObject);
begin
Width:=99;//���������� ������ ����
Height:=31;//���������� ������
Left:=-200;//������ ���� �� ����� ������� ������
end;

procedure TFPusk.Timer1Timer(Sender: TObject);
var
i:Integer;
h:THandle;
begin
Visible:=true; //������� ���� �������
//���������� ������� ������� ���� � ����� ������ ���� ������
Top:=Screen.Height-Height;
Left:=1;
//��������� ������ ��������� h, ������� ����� �������������� ��� ��������
h:=CreateEvent(nil, true,false, 'et' ) ;
//������ ����� ��������� ������
//�� 1 �� 80 ��������� �������� �� begin �� end
for i:=1 to 80 do
begin
//��������� �������� ������� ������� ���� .� �������
Top:=Screen.Height-Height-i*5;
Repaint; //������������ ����
WaitForSingleObject(h,15);//�������� � 5 �����������
End;
//������ ���� ��������� ������. �������� ��� ��,
//������ ���������� ���� � �������� �������
for i:=80 downto 1 do
begin
Top:=Screen. Height-Height-i*5;
Repaint;
WaitForSingleObject(h,15) ;
end;
Closehandle(h); //������������ ��������� h
Visible:=false; //�������� ����,
end;

procedure TFPusk.act1Execute(Sender: TObject);
var
i:Integer;
h:THandle;
begin
  h:=CreateEvent(nil, true,false, 'et' ) ;
Closehandle(h); //������������ ��������� h
Visible:=false; //�������� ����,
Timer1.Enabled:=False;
Close;
end;

end.
