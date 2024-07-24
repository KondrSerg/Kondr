//------------------------------------------
// Barcod-to-Bitmap
// Тестовый проект модуля кодировщика EAN-13
// Author: Бадло Сергей Григорьевич
// H-page: http://raxp.radioliga.com
// Cистемные требования: win32
//------------------------------------------

unit UShtrih_EKV;

interface

uses
  Windows, Messages, SysUtils, Classes, System.Actions, Vcl.ActnList, Vcl.Menus,
  Vcl.StdCtrls, Vcl.Mask, Vcl.Imaging.pngimage, Vcl.Buttons, Vcl.ExtCtrls,Dialogs,ComCtrls,Math,AdvObj, BaseGrid, AdvGrid,
  Vcl.Controls, Vcl.Grids, Vcl.Graphics,ExcelXP, Forms,Variants,
  barcod,  shellapi,Clipbrd,ComObj
   ;

type
  TFShtrih_EKV = class(TForm)
    Image2: TImage;
    StringGrid1: TStringGrid;
    pnl1: TPanel;
    btn1: TSpeedButton;
    Button1: TButton;
    medt1: TMaskEdit;
    img1: TImage;

    Button2: TButton;Label1: TLabel;
    pm1: TPopupMenu;
    Copy1: TMenuItem;
    actlst1: TActionList;
    act1: TAction;
    Memo1: TMemo;
    img3: TImage;
    img2: TImage;
    img4: TImage;
    procedure btn1Click(Sender: TObject);
    procedure medt1Change(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure img1DblClick(Sender: TObject);
    procedure Image2Click(Sender: TObject);
    function chek(var n:string):string;
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Clear_StringGrid(StringGrid: TStringGrid);
    procedure StringGrid1MouseWheelDown(Sender: TObject; Shift: TShiftState;
      MousePos: TPoint; var Handled: Boolean);
    procedure StringGrid1MouseWheelUp(Sender: TObject; Shift: TShiftState;
      MousePos: TPoint; var Handled: Boolean);
    procedure Copy1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure StringGrid1KeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
  public
    { Public declarations }
    SSS,S1:string;
    EKO:Integer;//1
  end;

var
  FShtrih_EKV: TFShtrih_EKV;

implementation

uses
  Main, barcod_EKV;

{$R *.dfm}
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//Очистка Грида

procedure TFShtrih_EKV.Clear_StringGrid(StringGrid: TStringGrid);
var
  i, j: integer;
  Str:string;
  SG: TStringGrid;
  Zak,Nam,NK,KK,NN,KN,BZ,BRIK,Kol:String;
begin
  sg:= StringGrid;
  Str:=SG.Name;
  i:=0;
  J:=0;
  for i := 0 to sg.ColCount - 1 do
    sg.Cols[i].Clear;
end;

procedure TFShtrih_EKV.Copy1Click(Sender: TObject);
var S1:string;
L:Integer;
begin
S1:=StringGrid1.Cells[4,StringGrid1.Row];
      L:=Length(S1);
      Delete(S1,L,1);
Clipboard.AsText :=S1;
end;

procedure TFShtrih_EKV.Button1Click(Sender: TObject);
var
  XL2,XL1,Sheet1,Sheet2,Sheet3: Variant;
  i, j,J1, y, PosTrud, l,Kol_Zak,Kol_Klap,R,R1,R2,R3,R5,R4,R6,Res1: integer;
  Color1: TColor;
        //trud:Float;
  TrudStr,dir,God,mes,Dat_S1,Dat,Zakaz,izd,Nom: string;
  Zak,Nam,Nam1,NK,KK,NN,KN,BZ,BRIK,Kol,WZ,TU,Moch,VER,Front,Nom_1,fileName,ZI:String;
  Sh,Nom_I:Int64;
  Before:IDispatch;
  bmp1 : tbitmap;
  Kol_T,Kol_Kl,Moch_F:Double;
begin

  Dat_S1:=FormatDateTime('mm.dd.YYYY',Now);
  Dat:=FormatDateTime('dd.mm.YYYY',Now);
  God:= FormatDateTime('YYYY',Now);
  Mes:= FormatDateTime('mmmm',Now);
  bz:=FShtrih_EKV.Caption;
  ZI:='';

  Dir := 'V:\ОИТ\CKlapana\Шильды ЭКО\' + God;
  CreateDir(Dir);

  Dir := 'V:\ОИТ\CKlapana\Шильды ЭКО\'  + God + '\' + mes + '\';
  CreateDir(Dir);

  XL2 := CreateOleObject('Excel.Application');
  Memo1.Lines.Add('XL2 := CreateOleObject(');
  XL2.Application.EnableEvents := false;
  CopyFile(PWideChar('V:\ОИТ\CKlapana2\2013\ШтрихВЗЭКО.xlsx'), PWideChar(Dir +
  '\'  + bz +' Штрих.xlsx'), False);

  XL2.Workbooks.Open(Dir + '\'+bz +' Штрих.xlsx');
  Memo1.Lines.Add('XL2.Workbooks.Open(Dir + ');
  XL2.Visible := True;
  Kol_Zak:=StringGrid1.RowCount-1;
  J:=1;
  Sheet3 := XL2.WorkSheets.Item[1];
  Sheet3.Activate;
  eko:=0;
   for I := 0 to Kol_Zak-1 do
    Begin
        Nam:= StringGrid1.Cells[2,I+1];
        eko:=0;
        R1:=Pos('ЭКВ-К-',Nam);
        R2:=Pos('ЭКВ-K-',Nam);//K Латинская
        if (R1<>0) or (R2<>0) then
                eko:=1;
        if StringGrid1.Cells[0,I+1]='' then
        Break;
        Kol_Klap:= StrToInt(StringGrid1.Cells[0,I+1]);
      for Y := 0 to Kol_Klap-1 do
      Begin
          XL2.ActiveWorkBook.Worksheets[3].Select;
          XL2.ActiveWorkBook.Worksheets[3].Rows['1:7'].Select;
          XL2.Selection.Copy;
          XL2.ActiveWorkBook.Worksheets[1].Select;
          XL2.ActiveWorkBook.Worksheets[1].Range['A' + IntToStr(J)].Select;
          XL2.ActiveWorkBook.Worksheets[1].Paste;
          Before:=XL2.ActiveWorkBook.Worksheets[1].Range['A' + IntToStr(J+7)];
          XL2.ActiveWorkBook.Worksheets[1].HPageBreaks.Add(Before);
          J:=J+7;
      End;
    End;
    XL2.ActiveWorkBook.Worksheets[1].Range['A' + IntToStr(J)+':H' + IntToStr(J+7)].Delete;
   // XL2.Visible := True;
  J:=1;
  Sheet3 := XL2.WorkSheets.Item[2];
  Sheet3.Activate;
   for I := 0 to Kol_Zak-1 do
    Begin
        R1:=Pos('ЭКВ-К-',Nam);
        R2:=Pos('ЭКВ-K-',Nam);//K Латинская
        if (R1<>0) or (R2<>0) then
                eko:=1;
        if eko=0 then
        Break;
        if StringGrid1.Cells[0,I+1]='' then
        Break;
        Kol_Klap:= StrToInt(StringGrid1.Cells[0,I+1]);
      for Y := 0 to Kol_Klap-1 do
      Begin
        XL2.ActiveWorkBook.Worksheets[2].Rows['1:5'].Select;
        XL2.Selection.Copy;
        XL2.ActiveWorkBook.Worksheets[2].Range['A' + IntToStr(J+5)].Select;
        XL2.ActiveWorkBook.Worksheets[2].Paste;
        Before:=XL2.ActiveWorkBook.Worksheets[2].Range['A' + IntToStr(J+5)];
        XL2.ActiveWorkBook.Worksheets[2].HPageBreaks.Add(Before);
        J:=J+5;
      End;
    End;
    XL2.ActiveWorkBook.Worksheets[2].Range['A' + IntToStr(J)+':I' + IntToStr(J+5)].Delete;
    J:=1;
  Sheet3 := XL2.WorkSheets.Item[1];
  Sheet3.Activate;
  for I := 0 to Kol_Zak-1 do
    Begin
        if StringGrid1.Cells[0,I+1]='' then
        Break;

            //clipboard.Clear;
      Kol_Klap:= StrToInt(StringGrid1.Cells[0,I+1]);
      bz:= StringGrid1.Cells[1,I+1];
      Nam:= StringGrid1.Cells[2,I+1];
      R1:=Pos('_ZI',Nam);  //Воздухонагреватель Канал-ЭКВ-K-200-6,0_ZI
      if R1<>0 then
        ZI:='_ZI';
      Nam1:= StringGrid1.Cells[2,I+1];
      Izd:=  StringGrid1.Cells[2,I+1];
      Zak:= StringGrid1.Cells[8,I+1];
      Moch:= StringGrid1.Cells[9,I+1];
      SyStem.SysUtils.FormatSettings.DecimalSeparator :=(',');
     // Moch_F:=StrToFloat(StringReplace(Moch,'.',',',[rfReplaceAll]));
     // Kol_T:=StrToFloat(StringReplace(StringGrid1.Cells[10,I+1],'.',',',[rfReplaceAll]));
      //Взрывозащита
      WZ:='';

     {if EKO=1 then
     Begin
      //-------
      R1:=Pos('ВЕНЭ-600',Nam1);
      if R1<>0 then
      TU:='ТУ 3442-159-40149153-2011';
      //------- Электрокалорифер ВЕНЭВ-К2-632-465-645-15,6-IP66_RAL9005_191036493а-ТМН
      R1:=Pos('ВЕНЭВ',Nam1);
      if R1<>0 then
      begin
      TU:='ТУ 3442-159-40149153-2011, Центр "ПрофЭкс", RU C-RU.АЖ58.В.01110/20';
      //Взрывозащита
      WZ:='1 Ex e mb IIC T3 Gb X';
      end;
     End;}

      nn:= StringGrid1.Cells[6,I+1];
      Nom_i:=StrToInt(nn);
      S1:=StringGrid1.Cells[4,I+1];
      L:=Length(S1);
      Sh:=StrToInt64(S1);
      for Y := 0 to Kol_Klap-1 do
      Begin
         Sheet3 := XL2.WorkSheets.Item[1];
         Sheet3.Activate;
         clipboard.assign(img4.Picture);
            XL2.ActiveWorkBook.Worksheets[1].Range['G' + IntToStr(J+6)].Select;
            Memo1.Lines.Add('Paste  '+IntToStr(I)+'   '+IntToStr(Y));
            XL2.ActiveWorkBook.Worksheets[1].Paste;
            clipboard.Clear;
        //
        R1:=Pos('Канал-ЭКВ-',Nam);
        if (R1<>0) then
        Begin
            clipboard.assign(img2.Picture);
            XL2.ActiveWorkBook.Worksheets[1].Range['G' + IntToStr(J+3)].Select;
            Memo1.Lines.Add('Paste  '+IntToStr(I)+'   '+IntToStr(Y));
            XL2.ActiveWorkBook.Worksheets[1].Paste;
            clipboard.Clear;
            TU :='ТУ 4864-156-40149153-2010';
        End;
       { if EKO=0 then
        begin
            clipboard.assign(img3.Picture);
            XL2.ActiveWorkBook.Worksheets[1].Range['B' + IntToStr(J+5)].Select;
            XL2.ActiveWorkBook.Worksheets[1].Paste;
        end;
        if WZ<>'' then
        Begin
          clipboard.assign(img3.Picture);
            XL2.ActiveWorkBook.Worksheets[1].Range['B' + IntToStr(J+5)].Select;
            Memo1.Lines.Add('Paste  '+IntToStr(I)+'   '+IntToStr(Y));
            XL2.ActiveWorkBook.Worksheets[1].Paste;
            clipboard.Clear;
        End;}
        S1:=IntToStr(Sh);
        NN:=IntToStr(Nom_i);
        SSS:=chek(S1);

        img1.Picture.Bitmap.Assign(barcod_EKV.barcode(SSS, rgb(0,0,0), 8, 0, 0, true, true, false));
        {if EKO=0 then
         XL2.ActiveWorkBook.Worksheets[1].Range['E' + IntToStr(J+5)].Select
         Else }
        XL2.ActiveWorkBook.Worksheets[1].Range['F' + IntToStr(J+5)].Select;
        XL2.ActiveWorkBook.Worksheets[1].Paste;

        XL2.ActiveWorkBook.WorkSheets[1].Cells[j + 1, 2] :=Nam;//izd;
        XL2.ActiveWorkBook.WorkSheets[1].Cells[j + 2, 2] :=TU;
        XL2.ActiveWorkBook.WorkSheets[1].Cells[j + 2, 7] :=WZ;
        XL2.ActiveWorkBook.WorkSheets[1].Cells[j + 3, 4] :=Zak;
        if EKO=1 then
        Begin
         Res1:=Pos('Электрокалорифер ВЕНЭ-',Nam);
         { if Res1<>0 then

          XL2.ActiveWorkBook.WorkSheets[1].Cells[j + 3, 7] :=' 380 В, 3~50 Гц' //Moch+  кВт;
          else   }
         // XL2.ActiveWorkBook.WorkSheets[1].Cells[j + 3, 7] :=Moch+' кВт, 380 В, 3~50 Гц'; //Moch+  кВт;

        End;
        XL2.ActiveWorkBook.WorkSheets[1].Cells[j + 4, 4] :=BZ;

        XL2.ActiveWorkBook.WorkSheets[1].Cells[j + 6, 4] :=' Зав.ном '+NN;
        XL2.ActiveWorkBook.WorkSheets[1].Cells[j + 6, 3] :=God;

        J:=J+7;
        Inc(Sh);
        Inc(Nom_i);
      End;
    End;
 //
     J:=1;
 if EKO=1 then
 Begin
  Sheet3 := XL2.WorkSheets.Item[2];
  Sheet3.Activate;

   for I := 0 to Kol_Zak-1 do
    Begin
      Kol_Klap:= StrToInt(StringGrid1.Cells[0,I+1]);
      bz:= StringGrid1.Cells[1,I+1];
      Nam:= StringGrid1.Cells[2,I+1];
      Nam1:= StringGrid1.Cells[2,I+1];
      Izd:=  StringGrid1.Cells[2,I+1];
      Zak:= StringGrid1.Cells[8,I+1];
      Moch:= StringGrid1.Cells[9,I+1];
      SyStem.SysUtils.FormatSettings.DecimalSeparator :=(',');
      nn:= StringGrid1.Cells[6,I+1];
      Nom_i:=StrToInt(nn);
      S1:=StringGrid1.Cells[4,I+1];
      L:=Length(S1);
      Sh:=StrToInt64(S1);

      for Y := 0 to Kol_Klap-1 do
      Begin
         Sheet3 := XL2.WorkSheets.Item[2];
         Sheet3.Activate;
        //   ====================Поиск схемы Воздухонагреватель Канал-ЭКВ-K-200-6,0_ZI
        //Электрокалорифер ВЕНЭВ-А-612-720-700-70-IP66_201017057-5-КОМ
        res1:=Pos('_',izd);// Воздухонагреватель Канал-ЭКВ-K-200-6,0_ZI
         if Res1<>0 then
         Delete(izd,Res1,300); //Воздухонагреватель Канал-ЭКВ-K-200-6,0

        r1:=Pos('ЭКВ-К-',izd);
        R2:=Pos('ЭКВ-K-',Nam);//K Латинская
        if (R1<>0) or (R2<>0) then
        Begin
         res1:=Pos('-',izd);     //Воздухонагреватель Канал-ЭКВ-K-200-6,0
         if Res1<>0 then
         Delete(izd,1,Res1); //ЭКВ-K-200-6,0
         //
         res1:=Pos('-',izd);     //Воздухонагреватель Канал-ЭКВ-K-200-6,0
         if Res1<>0 then
         Delete(izd,1,Res1); //K-200-6,0
         //
         res1:=Pos('-',izd);     //Воздухонагреватель Канал-ЭКВ-K-200-6,0
         if Res1<>0 then
         Delete(izd,1,Res1); //200-6,0
         //
         res1:=Pos('-',izd);
         if Res1<>0 then
         begin
         Ver:=Copy(izd,1,Res1-1);//200
         Delete(izd,1,Res1);  //6,0
         end;
         //
         Front:=izd;

        End;

          //Moch_F:=StrToFloat(StringReplace(Moch,'.',',',[rfReplaceAll]));
          //Kol_T:=StrToFloat(StringReplace(StringGrid1.Cells[10,I+1],'.',',',[rfReplaceAll]));
         //
         //fileName :='V:\ОИТ\Ceh\Шильды ЭКО\Шаблон\'+Ver+'-'+Front+'-'+Nom_1+'.PNG';
         fileName :='V:\ОИТ\Ceh\Шильды ЭКО\Шаблон\ЭКВ-К'+ZI+'\ЭКВ-К-'+Ver+'-'+Front+'.xlsx';
          if FileExists(fileName) then
         begin
            XL1 := CreateOleObject('Excel.Application');
            Memo1.Lines.Add('XL2 := CreateOleObject(');
            XL1.Application.EnableEvents := false;
            XL1.Workbooks.Open(fileName);
            Sheet1 := XL1.WorkSheets.Item[2];
            Sheet1.Activate;
            Sheet1.Shapes.item['Рисунок 1'].Select;
            XL1.Selection.Copy;
    ///
            XL2.ActiveWorkBook.WorkSheets[2].Cells[j + 1, 2] :=Nam;//izd;
            XL2.ActiveWorkBook.Worksheets[2].Range['B' + IntToStr(J+4)].Select;
            Memo1.Lines.Add('Paste  '+IntToStr(I)+'   '+IntToStr(Y));
            XL2.ActiveWorkBook.Worksheets[2].Paste;
            Sheet1:= UnAssigned;
            XL1.ActiveWorkBook.Close;
            XL1 := UnAssigned;
             J:=J+5;
         end
         else
         begin

          fileName :='V:\ОИТ\Ceh\Шильды ЭКО\Шаблон\Нестандарт\'+Nam1+'.xlsx';
          if FileExists(fileName) then
          begin
            XL1 := CreateOleObject('Excel.Application');
            Memo1.Lines.Add('XL2 := CreateOleObject(');
            XL1.Application.EnableEvents := false;
            XL1.Workbooks.Open(fileName);
            Sheet1 := XL1.WorkSheets.Item[2];
            Sheet1.Activate;
            Sheet1.Shapes.item['Рисунок 1'].Select;
            XL1.Selection.Copy;

            XL2.ActiveWorkBook.WorkSheets[2].Cells[j + 1, 2] :=Nam;//izd;
            XL2.ActiveWorkBook.Worksheets[2].Range['B' + IntToStr(J+4)].Select;
            Memo1.Lines.Add('Paste  '+IntToStr(I)+'   '+IntToStr(Y));
            XL2.ActiveWorkBook.Worksheets[2].Paste;
            Sheet1:= UnAssigned;
            XL1.ActiveWorkBook.Close;
            XL1 := UnAssigned;
             J:=J+5;
          end;
         end;
      End;
    End;
 End;
    XL2.ActiveWorkBook.Worksheets[2].Range['A' + IntToStr(J)+':I' + IntToStr(J+5)].Delete;

  Memo1.Lines.Add('1');
 // XL2.Visible := True;
  Memo1.Lines.Add('XL2.Visible := True;');

  Sheet3:= UnAssigned;
  XL2 := UnAssigned;
  Memo1.Lines.Add('XL2 := UnAssigned;');
  Clear_StringGrid(StringGrid1);
  Close;
end;

procedure TFShtrih_EKV.Button2Click(Sender: TObject);
var
  XL2: Variant;
  i, j, y, PosTrud, l: integer;
  Color1: TColor;
        //trud:Float;
  TrudStr,dir,God,mes,Dat_S1,Dat,Zakaz,bz,izd,Nom: string;
begin
  Dat_S1:=FormatDateTime('mm.dd.YYYY',Now);
  Dat:=FormatDateTime('dd.mm.YYYY',Now);
  God:= FormatDateTime('YYYY',Now);
  Mes:= FormatDateTime('mmmm',Now);
  Dir := 'V:\ОИТ\CKlapana\Сопроводительные Клапана\' + God;
  CreateDir(Dir);

  Dir := 'V:\ОИТ\CKlapana\Сопроводительные Клапана\'  + God + '\' + mes + '\';
  CreateDir(Dir);

  XL2 := CreateOleObject('Excel.Application');
  CopyFile(PWideChar('V:\ОИТ\CKlapana2\2013\Штрих.xlsx'), PWideChar(Dir +
  '\'  + 'Штрих ' + Dat_S1 + '.xlsx'), False);

  XL2.Workbooks.Open(Dir + '\' + 'Штрих ' + Dat_S1 + '.xlsx');
  XL2.Application.EnableEvents := false;
  J:=1;
  XL2.ActiveWorkBook.Worksheets[1].Rows['1:7'].Select;
  XL2.Selection.Copy;
  for I := 0 to 4 do
    Begin

        XL2.ActiveWorkBook.Worksheets[1].Range['A' + IntToStr(J+7)].Select;
        XL2.ActiveWorkBook.Worksheets[1].Paste;
        J:=J+7;
       // XL2.ActiveWorkBook.WorkSheets[1].Cells[j + 1, i + 1] := TrudStr;
    End;
    J:=1;
  for I := 0 to 4 do
    Begin
        SSS:=chek(S1);
        img1.Picture.Bitmap.Assign(barcode(SSS, rgb(0,0,0), 8, 0, 0, true, true, false));
        XL2.ActiveWorkBook.Worksheets[1].Range['E' + IntToStr(J+5)].Select;
        XL2.ActiveWorkBook.Worksheets[1].Paste;
        //XL2.ActiveWorkBook.Worksheets[1].Selection.ShapeRange.ZOrder:= 1; на задний план- не работает не Select
        XL2.ActiveWorkBook.WorkSheets[1].Cells[j + 1, 2] :=izd;
        XL2.ActiveWorkBook.WorkSheets[1].Cells[j + 3, 4] :=Zakaz;
        XL2.ActiveWorkBook.WorkSheets[1].Cells[j + 4, 4] :=BZ;
        XL2.ActiveWorkBook.WorkSheets[1].Cells[j + 6, 4] :=' Зав.ном '+Nom;
        XL2.ActiveWorkBook.WorkSheets[1].Cells[j + 6, 3] :=God;
        J:=J+7;
    End;

    XL2.Visible := True;
    XL2 := UnAssigned;

end;

function TFShtrih_EKV.chek(var n:string):string;
var  L,A,A1,B,B1,C,C1,R,I,X:Integer;
S:string;
Begin
      A:=0;
      A1:=0;
      S:=N;
      for I := 1 to Length(S) do// Суммировать цифры на четных позициях
        Begin
             C:=I mod 2;
             if C=0 then
             Begin
                B:=StrToInt(S[i]);
                A:=A+B;
             End;
        End;
        // Результат пункта 1 умножить на 3;
        A:=A*3;
        // Суммировать цифры на нечетных позициях;
        for I := 1 to Length(S) do
        Begin

             C:=I mod 2;
             if C<>0 then
             Begin
                B1:=StrToInt(S[i]);
                A1:=A1+B1;
             End;
        End;
        //Суммировать результаты пунктов 2 и 3;
        C:=A+A1;
        C1:=C;
        // Контрольное число — разница между окончательной суммой и ближайшим к ней наибольшим числом, кратным 10-ти.
        X:=0;
        for I := 0 to 9 do
          Begin
              // C1:=C1+1;
               L:=C1 mod 10;
               if L=0  then
               Begin
                R:=C1-C;
                Result:=S+INtToStr(R);
                Break;
               End;
               Inc(C1);
          End;

End;

procedure TFShtrih_EKV.btn1Click(Sender: TObject);
begin
 SSS:=chek(S1);
 img1.Picture.Bitmap.Assign(barcode(SSS, rgb(0,0,0), 8, 0, 0, true, true, false));
 Clipboard.Assign(img1.Picture.Bitmap);
end;

procedure TFShtrih_EKV.medt1Change(Sender: TObject);
begin
 //image1.Picture.Bitmap.Assign(barcode(SSS, 0, 10, 0, 255, true, false, false))
end;

procedure TFShtrih_EKV.StringGrid1KeyPress(Sender: TObject; var Key: Char);
var
n:integer;
begin
  if Key =#8 then
  Begin
      if (StringGrid1.RowCount=2) then
        exit;

    for n:=StringGrid1.Row to StringGrid1.RowCount-2 do
    begin
      StringGrid1.Rows[n]:=StringGrid1.Rows[n+1];

    end;
    StringGrid1.Rowcount:=StringGrid1.Rowcount-1;
  End;
end;

procedure TFShtrih_EKV.StringGrid1MouseWheelDown(Sender: TObject;
  Shift: TShiftState; MousePos: TPoint; var Handled: Boolean);
begin
  Handled := true;
  StringGrid1.Perform(WM_VSCROLL, SB_LINEDOWN, 0);
end;

procedure TFShtrih_EKV.StringGrid1MouseWheelUp(Sender: TObject; Shift: TShiftState;
  MousePos: TPoint; var Handled: Boolean);
begin
  Handled := true;
  StringGrid1.Perform(WM_VSCROLL, SB_LINEUP, 0);
end;

procedure TFShtrih_EKV.FormCreate(Sender: TObject);
begin
 //image1.Picture.Bitmap.Assign(barcode(sss, rgb(0,0,0), 10, 0, 0, true, true, false))
end;

procedure TFShtrih_EKV.FormShow(Sender: TObject);
begin
//:= medt1.Text;
  Button1.SetFocus;
end;

procedure TFShtrih_EKV.img1DblClick(Sender: TObject);
begin
 //image1.Picture.Bitmap.Assign(barcode(SSS, 255, 10, 180, 255, true, false, false))
end;

procedure TFShtrih_EKV.Image2Click(Sender: TObject);
begin
 //отсылка на печать
 //image1.Picture.Bitmap.SaveToFile('screen.bmp');
 //ShellExecute(0,'print','screen.bmp','','',SW_hide)
end;

end.
