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
Width:=99;//Установить ширину окна
Height:=31;//Установить высоту
Left:=-200;//Убрать окно за левую границу экрана
end;

procedure TFPusk.Timer1Timer(Sender: TObject);
var
i:Integer;
h:THandle;
begin
Visible:=true; //Сделать окно видимым
//Установить верхнюю позицию окна в левый нижний угол экрана
Top:=Screen.Height-Height;
Left:=1;
//Создается пустой указатель h, который будет использоваться для задержки
h:=CreateEvent(nil, true,false, 'et' ) ;
//Сейчас будем поднимать кнопку
//От 1 до 80 выполнять действия от begin до end
for i:=1 to 80 do
begin
//Увеличить значение верхней позиции окна .с кнопкой
Top:=Screen.Height-Height-i*5;
Repaint; //Перерисовать окно
WaitForSingleObject(h,15);//Задержка в 5 миллисекунд
End;
//Дальше идет опускание кнопки. Алгоритм тот же,
//просто выполнение идет в обратном порядке
for i:=80 downto 1 do
begin
Top:=Screen. Height-Height-i*5;
Repaint;
WaitForSingleObject(h,15) ;
end;
Closehandle(h); //Уничтожается указатель h
Visible:=false; //Прячется окно,
end;

procedure TFPusk.act1Execute(Sender: TObject);
var
i:Integer;
h:THandle;
begin
  h:=CreateEvent(nil, true,false, 'et' ) ;
Closehandle(h); //Уничтожается указатель h
Visible:=false; //Прячется окно,
Timer1.Enabled:=False;
Close;
end;

end.
