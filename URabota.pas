unit URabota;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Grids, AdvObj, BaseGrid, AdvGrid,Vcl.Clipbrd,
  Vcl.ExtCtrls, Vcl.StdCtrls, Vcl.Menus,ShellAPI, Vcl.Buttons;

type
  TFRabota = class(TForm)
    pnl1: TPanel;
    SG: TAdvStringGrid;
    Button1: TButton;
    pm1: TPopupMenu;
    Copy1: TMenuItem;
    btn1: TSpeedButton;
    Label1: TLabel;
    Edit1: TEdit;
    Label2: TLabel;
    Label3: TLabel;
    Memo1: TMemo;
    Memo2: TMemo;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    procedure Button1Click(Sender: TObject);
    procedure Summ(Sbor: string;Col:Integer);
    procedure SGGetCellColor(Sender: TObject; ARow, ACol: Integer;
      AState: TGridDrawState; ABrush: TBrush; AFont: TFont);
    procedure Copy1Click(Sender: TObject);
    procedure btn1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SGMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
  private
    { Private declarations }
  public
    { Public declarations }
    Vozd,YY:Integer;
  end;

var
  FRabota: TFRabota;

implementation

uses
  Main;

{$R *.dfm}
procedure TFRabota.Copy1Click(Sender: TObject);
begin
Clipboard.AsText := SG.Cells[SG.Col,SG.Row];
end;

procedure TFRabota.FormShow(Sender: TObject);
begin
       Button1.Click;
end;

procedure TFRabota.SGGetCellColor(Sender: TObject; ARow, ACol: Integer;
  AState: TGridDrawState; ABrush: TBrush; AFont: TFont);
  var S,Sbor,Dat:string;
  R,D,I,F:Integer;
  DD1,DD2,DTab:TDate;
  DD,DTS:String;
begin
  if (ARow>0) then
  Begin
    DD1:=StrToDate(Label2.Caption);
    DD2:=StrToDate(Label3.Caption);
    //
    if SG.Cells[6,ARow]='' then
     DTab:=0
     Else
    DTab:=StrToDate(SG.Cells[6,ARow]);
    F:=0;
    for I := trunc(DD1) to trunc(DD2) do
    Begin
      DD:=FormatDateTime('dd.mm.YYYY',I);
      DTS:=FormatDateTime('dd.mm.YYYY',DTab);
      R:=Pos(DD,DTS);
      if R<>0 then
      F:=1;
    End;
    S:=AnsiUpperCase(SG.Cells[7,ARow]);
    Sbor:=AnsiUpperCase(FRabota.Caption);
    R:=Pos(Sbor,S);

    if (R<>0) AND (F=1) AND ((ACol=6) or (ACol=7) or (ACol=8)) then
    begin
      ABrush.Color := RGB(0, 255, 0); //Св зеленый
    end;
        //=================================
    if SG.Cells[9,ARow]='' then
     DTab:=0
     Else
    DTab:=StrToDate(SG.Cells[9,ARow]);
    F:=0;
    for I := trunc(DD1) to trunc(DD2) do
    Begin
      DD:=FormatDateTime('dd.mm.YYYY',I);
      DTS:=FormatDateTime('dd.mm.YYYY',DTab);
      R:=Pos(DD,DTS);
      if R<>0 then
      F:=1;
    End;
    //
    S:=AnsiUpperCase(SG.Cells[10,ARow]);
    //Sbor:=FRabota.Caption;
    R:=Pos(Sbor,S);
    if (R<>0) AND (F=1) AND ((ACol=9) or (ACol=10) or (ACol=11)) then
    begin
      ABrush.Color := RGB(0, 255, 0); //Св зеленый
    end;
    //  ===============================
        //
    if SG.Cells[12,ARow]='' then
     DTab:=0
     Else
    DTab:=StrToDate(SG.Cells[12,ARow]);
    F:=0;
    for I := trunc(DD1) to trunc(DD2) do
    Begin
      DD:=FormatDateTime('dd.mm.YYYY',I);
      DTS:=FormatDateTime('dd.mm.YYYY',DTab);
      R:=Pos(DD,DTS);
      if R<>0 then
      F:=1;
    End;
    //
    S:=AnsiUpperCase(SG.Cells[13,ARow]);
    //Sbor:=FRabota.Caption;
    R:=Pos(Sbor,S);
    if (R<>0) AND (F=1) AND ((ACol=12) or (ACol=13) or (ACol=14)) then
    begin
      ABrush.Color := RGB(0, 255, 0); //Св зеленый
    end;
    //==========================================
    if SG.Cells[15,ARow]='' then
     DTab:=0
     Else
    DTab:=StrToDate(SG.Cells[15,ARow]);
    F:=0;
    for I := trunc(DD1) to trunc(DD2) do
    Begin
      DD:=FormatDateTime('dd.mm.YYYY',I);
      DTS:=FormatDateTime('dd.mm.YYYY',DTab);
      R:=Pos(DD,DTS);
      if R<>0 then
      F:=1;
    End;
    S:=AnsiUpperCase(SG.Cells[16,ARow]);
    R:=Pos(Sbor,S);
    if (R<>0) AND (F=1) AND ((ACol=15) or (ACol=16) or (ACol=17)) then
    begin
      ABrush.Color := RGB(0, 255, 0); //Св зеленый
    end;
    // ================================================
    S:=AnsiUpperCase(SG.Cells[18,ARow]);
    R:=Pos(Sbor,S);
    if (R<>0) AND ((ACol=17) or (ACol=18) or (ACol=19)) then
    begin
      ABrush.Color := RGB(0, 255, 0); //Св зеленый
    end;
    //

  End;

end;

procedure TFRabota.SGMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  I, J: Integer;
  myRect: TGridRect;
  Sum: Double;
begin
  Sum := 0;
  myRect := SG.Selection;

  if (SG.Col = 5) or (SG.Col = 8) or
  (SG.Col = 11) or (SG.Col = 14) or
  (SG.Col = 17) or (SG.Col = 19) or
  (SG.Col = 20) or
  (SG.Col = 21) then
  begin

    for I := myRect.Left to myRect.Right do
      for J := myRect.Top to myRect.Bottom do
      begin
        try
          Sum := Sum + StrToFloat(SG.Cells[I, J]);
        except
          Break;
        end;
      end;
      Edit1.Text := FloatToStr(Sum);
  end;
end;

procedure TFRabota.Summ(Sbor: string;Col:Integer);
var  S:string;
I,II,Y,R,R1:Integer;
N:Double;
C:TColor;
  D,F:Integer;
  DD1,DD2,DTab:TDate;
  DD,DTS:String;
begin
    //FRabota.SG.Cells[7,I+1]:= ADOQuery2.FieldByName('Подсборка1').AsString;  Col
    //FRabota.SG.Cells[8,I+1]:= ADOQuery2.FieldByName('Н\ч Сборка Клапана').AsString;  Col+1

      DD1:=StrToDate(Label2.Caption);
      DD2:=StrToDate(Label3.Caption);
      N:=0;
      Y:=0;
      for I := 0 to SG.RowCount-1 do
     begin
         if Col=18 then            //7 09.06.23
          S:=AnsiUpperCase(SG.Cells[10,I+1])
          Else
         S:=AnsiUpperCase(SG.Cells[Col,I+1]);
         //C:=SG.GetVisualProperties(Col,I+1).Abrush.Color;// (Rec Colors[Col,I+1];
         R:=Pos(AnsiUpperCase(Sbor),S);
             //==========================================
         //if (Col<>18)then
         //Begin
         if (Col=18)then
         begin       //6
          if SG.Cells[9,I+1]='' then
            DTab:=0
          Else                      //6
            DTab:=StrToDate(SG.Cells[9,I+1]);
          F:=0;
          for II := trunc(DD1) to trunc(DD2) do
          Begin
              DD:=FormatDateTime('dd.mm.YYYY',II);
              DTS:=FormatDateTime('dd.mm.YYYY',DTab);
              R1:=Pos(DD,DTS);
              if R1<>0 then
              F:=1;
          End;
         End
         else
         begin
          if SG.Cells[Col-1,I+1]='' then
            DTab:=0
          Else
            DTab:=StrToDate(SG.Cells[Col-1,I+1]);
          F:=0;
          for II := trunc(DD1) to trunc(DD2) do
          Begin
              DD:=FormatDateTime('dd.mm.YYYY',II);
              DTS:=FormatDateTime('dd.mm.YYYY',DTab);
              R1:=Pos(DD,DTS);
              if R1<>0 then
              F:=1;
          End;
         end;

          //============================
         if SG.Cells[1,I+1]<>'' then
         Inc(Y);
         if (R<>0) AND (F=1) then
         begin
            if SG.Cells[Col+1,I+1]='' then
            begin
              Continue;
            end
            else
            begin
                N:=N+StrToFloat(SG.Cells[Col+1,I+1]);
                //Inc(Y);
            end;
         end;
     end;
     SG.Cells[Col,Y+1]:='Summ';
     SG.Cells[Col+1,Y+1]:=FloatToStr(N);
     YY:=Y+1;
end;

procedure TFRabota.btn1Click(Sender: TObject);
begin
      Form1.ExportGridtoExcel2(SG);
end;

procedure TFRabota.Button1Click(Sender: TObject);
var I,R,R1:Integer;
S,Sbor:string;
N,N1,N2,N3:Double;
begin
    Sbor:=FRabota.Caption;
    Summ(Sbor,7);
    Summ(Sbor,10);
    Summ(Sbor,13);
    Summ(Sbor,16);
    Summ(Sbor,18);
    if Vozd=1 then
    Begin
      for I := 0 to Memo1.Lines.Count-1 do
      begin
        if Memo1.Lines.Strings[i]<>'' then //Расключение
        begin
            N:=N+StrToFloat(Memo1.Lines.Strings[i]);
        end;
      end;
//
      for I := 0 to Memo2.Lines.Count-1 do
      begin
        if Memo2.Lines.Strings[i]<>'' then //Обварка
        begin
            N1:=N1+StrToFloat(Memo2.Lines.Strings[i]);
        end;
      end;
      SG.Cells[7,YY+1]:=SG.Cells[8,YY]+'+'+ SG.Cells[17,YY]+'+'+FloatToStr(N1)+'+'+ SG.Cells[19,YY];
      SG.Cells[8,YY+1]:=  FloatToStr(N1+StrToFloat(SG.Cells[8,YY])+StrToFloat(SG.Cells[17,YY])+StrToFloat(SG.Cells[19,YY]));

      SG.Cells[10,YY+1]:=SG.Cells[11,YY]+'+'+ SG.Cells[14,YY];
      SG.Cells[11,YY+1]:=  FloatToStr(StrToFloat(SG.Cells[14,YY])+StrToFloat(SG.Cells[11,YY]));
      //SG.Cells[7+1,YY]:=FloatToStr(N);
    End;
        if Vozd=0 then
    Begin

      SG.Cells[7,YY+1]:=SG.Cells[8,YY]{+'+'+ SG.Cells[19,YY]};
      SG.Cells[8,YY+1]:=  FloatToStr(StrToFloat(SG.Cells[8,YY]){+StrToFloat(SG.Cells[19,YY])});
      //Мотвеенко 09.06.23
      SG.Cells[10,YY+1]:=SG.Cells[11,YY]+'+'+ SG.Cells[19,YY];
      SG.Cells[11,YY+1]:=  FloatToStr(StrToFloat(SG.Cells[11,YY])+StrToFloat(SG.Cells[19,YY]));
    End;

    S :=ExtractFileDir(ParamStr(0)) ;

end;

end.
