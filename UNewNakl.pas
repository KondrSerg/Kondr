unit UNewNakl;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, StdCtrls, Grids, ExtCtrls, ComObj, Math, Buttons, DB,
  ADODB;

type
  TFNewNakl = class(TForm)
    Panel2: TPanel;
    Label5: TLabel;
    Label6: TLabel;
    Label1: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Panel3: TPanel;
    Label9: TLabel;
    Label10: TLabel;
    SG2: TStringGrid;
    Button3: TButton;
    Edit1: TEdit;
    Edit2: TEdit;
    ComboBox1: TComboBox;
    Edit3: TEdit;
    Edit4: TEdit;
    Edit5: TEdit;
    Edit6: TEdit;
    Panel1: TPanel;
    Label3: TLabel;
    Label4: TLabel;
    SG1: TStringGrid;
    Button2: TButton;
    Memo1: TMemo;
    ProgressBar1: TProgressBar;
    SD1: TStringGrid;
    Button4: TButton;
    Label11: TLabel;
    Label15: TLabel;
    Button5: TButton;
    SD2: TStringGrid;
    SD3: TStringGrid;
    SD4: TStringGrid;
    SD5: TStringGrid;
    SD6: TStringGrid;
    SD7: TStringGrid;
    SpeedButton1: TSpeedButton;
    Edit7: TEdit;
    SD8: TStringGrid;
    SD9: TStringGrid;
    SD: TStringGrid;
    Stenki: TStringGrid;
    Memo2: TMemo;
    SSvarka: TStringGrid;
    MemVG: TMemo;
    RSG1: TStringGrid;
    RSG2: TStringGrid;
    RSG3: TStringGrid;
    RSG4: TStringGrid;
    chk1: TCheckBox;
    ADOQuery2: TADOQuery;
    SQLConnector1: TADOConnection;
    sg3: TStringGrid;
    Label2: TLabel;
    btn1: TButton;
    btn2: TButton;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    function Xls_To_StringGrid(AGrid: TStringGrid; AXLSFile: string;
                        List:
                        string; kpd: Integer; L: Integer; Coll: Integer; Roww:
                        Integer): Boolean;
    procedure Button3Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure SG2DblClick(Sender: TObject);
    procedure SG2SelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    function Xls_To_Integer(Str: string): Integer;
    procedure SpeedButton1Click(Sender: TObject);
    procedure Clear_StringGrid(SG:TStringGrid);
    function Float(var Str:String): String;
    procedure btn1Click(Sender: TObject);
    procedure btn2Click(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
                Summa: Integer;
                F_Rasf: Integer;
                Puty:String;
  end;

var
  FNewNakl: TFNewNakl;

implementation

uses Main;

{$R *.dfm}
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
procedure TFNewNakl.Clear_StringGrid(SG:TStringGrid);
Var I:Integer;
Begin
        for i:=0 To SG.RowCount Do
        Begin
            SG.Rows[i].Clear;
        End;

end;
function TFNewNakl.Xls_To_Integer(Str: string): Integer;
Var
I:Integer;
Begin
     try
     I:=StrToInt(Str);
     except
     I:=0;
     End;
     Result:=I;
end;
function TFNewNakl.Xls_To_StringGrid(AGrid: TStringGrid; AXLSFile: string; List:
        string; kpd: Integer; L: Integer; Coll: Integer; Roww: Integer):
        Boolean;
const
        xlCellTypeLastCell = $0000000B;
var
        XLApp, Sheet: OLEVariant;
        RangeMatrix: Variant;
        x, y, k, r, Res, I, J, Col, Posic, Kod, PosZ, F: Integer;
        Nam, Str, Str1, StrDat, Dat: string;
begin
        Result := False;
        Form1.Clear_StringGrid(SD1);
        XLApp := CreateOleObject('Excel.Application');
        try
                XLApp.Visible := False;
                XLApp.Workbooks.Open(AXLSFile);
                for Col := 1 to
                        XLApp.Workbooks[ExtractFileName(AXLSFile)].WorkSheets.Count do
                begin
                        Nam :=
                                XLApp.Workbooks[ExtractFileName(AXLSFile)].WorkSheets[Col].Name;
                        Res := AnsiCompareStr(Nam, List);
                        if Res = 0 then
                        begin
                                XLApp.Workbooks[ExtractFileName(AXLSFile)].WorkSheets.Item[Col].Activate;
                                Sheet :=
                                        XLApp.Workbooks[ExtractFileName(AXLSFile)].WorkSheets[Col];
                                Break;
                        end;
                end;
                Nam := Sheet.Name;
                Sheet.Cells.SpecialCells(xlCellTypeLastCell,
                        EmptyParam).Activate;
                if (kpd = 1) and (L = 1) then
                begin
                        x := Roww;
                        y := Coll;
                        AGrid.RowCount := x;
                        AGrid.ColCount := y + 2;
                        RangeMatrix := XLApp.Range['A1', XLApp.Cells.Item[X,
                                Y]].Value;
                        k := 1;
                        repeat
                                for r := 1 to y do
                                begin
                                        if r = 19 then
                                                Continue;
                                        Str := string(RangeMatrix[K, R]);
                                        AGrid.Cells[(r), (k - 1)] :=
                                                string(RangeMatrix[K, R]);
                                end;
                                Inc(k, 1);
                                AGrid.RowCount := k + 1;
                        until k > x;
                        RangeMatrix := Unassigned;
                end;

                if (kpd = 2) and (L = 2) then //Ножницы
                begin
                        x := Roww; //XLApp.ActiveCell.Row;
                        y := Coll;
                        AGrid.RowCount := x;
                        AGrid.ColCount := y + 2;
                        RangeMatrix := XLApp.Range['A1', XLApp.Cells.Item[X,
                                Y]].Value;
                        k := 1;
                        repeat
                                for r := 1 to y do
                                begin
                                        if (r = 8) or (R = 16) then
                                                Continue;
                                        if r = 9 then
                                                Continue;
                                        Str := string(RangeMatrix[K, R]);
                                        AGrid.Cells[(r), (k - 1)] :=
                                                string(RangeMatrix[K, R]);
                                end;
                                Inc(k, 1);
                                AGrid.RowCount := k + 1;
                        until k > x;
                        RangeMatrix := Unassigned;
                end;
                if (Kpd = 0) and (L = 0) then

                begin
                        x := Roww; //10;//XLApp.ActiveCell.Row;
                        y := Coll; //XLApp.ActiveCell.Column; //30;
                        AGrid.RowCount := x;
                        AGrid.ColCount := y + 2;
                        RangeMatrix := XLApp.Range['A1', XLApp.Cells.Item[X,
                                Y]].Value;
                        k := 1;
                        repeat
                                for r := 1 to y do
                                begin

                                        Str := string(RangeMatrix[K, R]);
                                        AGrid.Cells[(r), (k - 1)] :=
                                                string(RangeMatrix[K, R]);
                                end;
                                Inc(k, 1);
                                AGrid.RowCount := k + 1;
                        until k > x;
                        RangeMatrix := Unassigned;

                end;
                //=\===================================================Гибка
                if (Kpd = 0) and (L = 1) then

                begin
                        x := Roww; //10;//XLApp.ActiveCell.Row;
                        y := Coll; //XLApp.ActiveCell.Column; //30;
                        AGrid.RowCount := x;
                        AGrid.ColCount := y + 2;
                        RangeMatrix := XLApp.Range['A1', XLApp.Cells.Item[X,
                                Y]].Value;
                        k := 1;
                        repeat
                                for r := 1 to y do
                                begin
                                        if (R = 31) or (R = 19) or (R = 20) then
                                                Continue;
                                        Str := string(RangeMatrix[K, R]);
                                        AGrid.Cells[(r), (k - 1)] :=
                                                string(RangeMatrix[K, R]);
                                end;
                                Inc(k, 1);
                                AGrid.RowCount := k + 1;
                        until k > x;
                        RangeMatrix := Unassigned;

                end;
                //=\===================================================Программа
                if (Kpd = 0) and (L = 2) then

                begin
                        x := 16; //10;//XLApp.ActiveCell.Row;
                        y := 30; //XLApp.ActiveCell.Column; //30;
                        AGrid.RowCount := x;
                        AGrid.ColCount := y + 2;
                        for R := 0 to 16 do
                        begin
                                for K := 0 to 30 do
                                begin
                                        AGrid.Cells[(r), (k)] := Sheet.Cells[R +
                                                30, k];
                                end;
                        end;
                end;
        finally
                // Quit Excel
                if not VarIsEmpty(XLApp) then
                begin
                        // XLApp.DisplayAlerts := False;
                        XLApp.Quit;
                        XLAPP := Unassigned;
                        Sheet := Unassigned;
                        Result := True;
                end;
        end;
end;

procedure TFNewNakl.FormShow(Sender: TObject);
var
        I, Kol, J, y, Res, Nom1, Nom2: Integer;
        Nom_Poz, Zak, IDGP, Kol_Zap, Kol_Raz: string;
begin

        //Button1.Enabled := False;
        Summa := 0;
        SG1.Enabled := False;
        SG1.RowCount := 1;
        Clear_StringGrid(SG1);
        Clear_StringGrid(SG2);
        Clear_StringGrid(SD1);
        SG2.Cells[0, 0] := 'Заказ';
        SG2.Cells[1, 0] := 'Расчетная дата';
        SG2.Cells[2, 0] := 'Кол во запущенных';
        SG2.Cells[3, 0] := 'Номер';
        SG2.Cells[4, 0] := 'ID';
        SG2.Cells[5, 0] := 'IDgp';
        SG2.ColWidths[0] := 0;
        SG2.ColWidths[2] := 0;
        SG2.ColWidths[4] := 0;
        Edit1.Text := '0';
        Memo1.Lines.Clear;
        Nom2 := 0;
       SQLConnector1.Connected:=False;
       SQLConnector1.ConnectionString:=Form1.MSSQL_CONN_STR_Glob;
       SQLConnector1.Connected:=True;
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++  674
        if Form1.Zapusk2=0 Then
        Begin                                                                     //AND ([Технолог Готов] IS NULL)
        if not Form1.mkQuerySelect(Form1.ADOQuery2,
                'Select * from %s Where  ( [Номер]<>0)   AND ([Файл] IS NULL) AND ([Заготовка Запуск] IS NULL)    AND ([Отмена] IS NULL)  AND ( [Х] IS NULL) Order By Номер',
                ['Запуск']) then
                exit;

        end;
        if Form1.Zapusk2=1 Then
        Begin                                                                     //AND ([Файл] IS NULL) AND ([Заготовка Запуск] IS NULL) AND ([Технолог Готов] IS NULL)
        if not Form1.mkQuerySelect(Form1.ADOQuery2,
                'Select * from %s Where  ( [Номер]='+#39+Label2.Caption+#39+') ', //AND (NOT [IdГП] IS NULL)  AND ([IdГП]<>0)    AND ([Отмена] IS NULL)  AND ( [Х] IS NULL) Order By Номер',
                ['Запуск']) then
                exit;

        end;
        if Form1.Zapusk2=2 Then
        Begin
            if not Form1.mkQuerySelect(Form1.ADOQuery2,
                'Select * from %s  WHERE Планирование=' + #39
                + Form1.ConvertDat1(Form1.ZSG.Cells[I_FN_KOL_ZAP + 3,Form1.ZSG.Row ])+ #39, ['Запуск'])
                then
                exit;
                SG1.Enabled:=True;
                Clear_StringGrid(SG1);

                SG1.ColCount := 9;
                SG1.Cells[0, 0] := 'Заказ';
                SG1.Cells[1, 0] := 'Кол во запущенных';
                SG1.Cells[2, 0] := 'Номер';
                SG1.Cells[3, 0] := 'ID';
                SG1.Cells[4, 0] := 'Клапан';
                SG1.Cells[5, 0] := 'Сборка';
                SG1.Cells[6, 0] := 'Сварка';
                SG1.Cells[7, 0] := 'Привод';
                SG1.Cells[8, 0] := 'IDGP';
                SG1.RowCount:=1;;
                if Form1.ADOQuery2.RecordCount <> 0 then
                begin
                  SG1.RowCount := Form1.ADOQuery2.RecordCount + 1;

                  for i := 0 to Form1.ADOQuery2.RecordCount - 1 do
                  begin
                        SG1.Cells[0, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Заказ').AsString;
                        SG1.Cells[1, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Кол во запущенных').AsString;
                        SG1.Cells[2, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Номер').AsString;
                        SG1.Cells[3, i + 1] :=
                                Form1.ADOQuery2.FieldByName('ID').AsString;
                        SG1.Cells[4, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Изделие').AsString;
                        SG1.Cells[5, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Н\ч Сборка Клапана').AsString;
                        SG1.Cells[6, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Н\ч Сварка').AsString;
                        SG1.Cells[7, i + 1] :=
                                Form1.ADOQuery2.FieldByName('МодПривода').AsString;
                        SG1.Cells[8, i + 1] :=
                                Form1.ADOQuery2.FieldByName('IdГП').AsString;
                        Form1.ADOQuery2.Next;
                  end;
                end;
        end;
        if Form1.Zapusk2<>2 Then
        Begin
        Y := 1;
        j := 1;
        SG2.FixedRows:=1;
        SG2.RowCount := Form1.ADOQuery2.RecordCount + 1;
        Memo1.Lines.Clear;
        if Form1.ADOQuery2.RecordCount <> 0 then
        begin
                for I := 0 to Form1.ADOQuery2.RecordCount - 1 do
                begin
                        Nom1 := Form1.ADOQuery2.FieldByName('Номер').AsInteger;
                        if j > 1 then
                                Nom2 := StrToInt(SG2.Cells[3, j - 1]);
                        if Nom1 = Nom2 then
                        begin
                                Form1.ADOQuery2.Next;
                                Continue;
                        end;
                        SG2.Cells[0, j] :=
                                Form1.ADOQuery2.FieldByName('Заказ').AsString;
                        SG2.Cells[1, j] :=
                                Form1.ADOQuery2.FieldByName('Расчетная дата готовности').AsString;
                        SG2.Cells[2, j] :=
                                Form1.ADOQuery2.FieldByName('Кол во запущенных').AsString;
                        SG2.Cells[3, j] :=
                                Form1.ADOQuery2.FieldByName('Номер').AsString;
                        SG2.Cells[4, j] :=
                                Form1.ADOQuery2.FieldByName('ID').AsString;
                        SG2.Cells[5, j] :=
                                Form1.ADOQuery2.FieldByName('IDГП').AsString;
                        Nom_Poz := Form1.ADOQuery2.FieldByName('№Поз').AsString;
                        Inc(y);
                        Inc(j);
                        Form1.ADOQuery2.Next;
                end;
        end;
        SG2.RowCount := Y + 2;
        End;

end;

procedure TFNewNakl.FormClose(Sender: TObject; var Action: TCloseAction);
Var I:Integer;
begin
        Summa := 0;
        Label11.Caption := '0';
        Label9.Caption := '0';
        Label1.Caption := '0';
        Clear_StringGrid(SG1);;
end;

procedure TFNewNakl.Button3Click(Sender: TObject);
var
        I, Kol, Y, j, Pos, Pos1, x: Integer;
        Nom_Poz, Zak, IDGP, Briket, Briket1, Briket2, Br, Br1, Br2, Kol_Zap:
        string;
begin
        SG1.Enabled:=True;
        Clear_StringGrid(SG1);

        SG1.ColCount := 9;
        SG1.Cells[0, 0] := 'Заказ';
        SG1.Cells[1, 0] := 'Кол во запущенных';
        SG1.Cells[2, 0] := 'Номер';
        SG1.Cells[3, 0] := 'ID';
        SG1.Cells[4, 0] := 'Клапан';
        SG1.Cells[5, 0] := 'Сборка';
        SG1.Cells[6, 0] := 'Сварка';
        SG1.Cells[7, 0] := 'Привод';
        SG1.Cells[8, 0] := 'IDGP';
        //  ([Заказ]= '+#39+Zak+#39+')AND
        if not Form1.mkQuerySelect(Form1.ADOQuery2,
                'Select * from %s Where   (Номер='
                + #39 + Label15.Caption + #39 + ') AND (Отмена IS NULL) Order By Заказ ', ['Запуск']) then
                exit;
        SG1.RowCount:=1;;
        if Form1.ADOQuery2.RecordCount <> 0 then
        begin
                SG1.RowCount := Form1.ADOQuery2.RecordCount + 1;

                for i := 0 to Form1.ADOQuery2.RecordCount - 1 do
                begin
                        SG1.Cells[0, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Заказ').AsString;
                        SG1.Cells[1, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Кол во запущенных').AsString;
                        SG1.Cells[2, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Номер').AsString;
                        SG1.Cells[3, i + 1] :=
                                Form1.ADOQuery2.FieldByName('ID').AsString;
                        SG1.Cells[4, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Изделие').AsString;
                        SG1.Cells[5, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Н\ч Сборка Клапана').AsString;
                        SG1.Cells[6, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Н\ч Сварка').AsString;
                        SG1.Cells[7, i + 1] :=
                                Form1.ADOQuery2.FieldByName('МодПривода').AsString;
                        SG1.Cells[8, i + 1] :=
                                Form1.ADOQuery2.FieldByName('IdГП').AsString;
                        Form1.ADOQuery2.Next;
                end;
        end;
end;

procedure TFNewNakl.Button2Click(Sender: TObject);
begin
        Close;
end;

procedure TFNewNakl.SG2DblClick(Sender: TObject);
Var I:Integer;
begin
        for i:=0 To SG1.RowCount Do
        Begin
            SG1.Rows[i].Clear;
        End;
        Label15.Caption := SG2.Cells[3, SG2.Row];
        Button3.Click;
end;

procedure TFNewNakl.SG2SelectCell(Sender: TObject; ACol, ARow: Integer;
  var CanSelect: Boolean);
begin
        Label15.Caption := SG2.Cells[3, ARow];
        Button3.Click;
end;

procedure TFNewNakl.FormKeyPress(Sender: TObject; var Key: Char);
begin
     If Key=#27 Then
     Close;
end;

procedure TFNewNakl.Button4Click(Sender: TObject);
var
        Kol_ZAP, Kol, Kol_Zap_Ranee, Res, Nomer, Zak_Int, i, A, B, D, y, F_KPU,
        F_KPD,xy, Res1, e, Zag, Zap, j, g, k,Q,Res_Mat,Res_Proch,Res_Stand,qq,
        Res_Sbor,Res_Klap,Podstavka,Srav,Srav1:Integer;
        Res_VElem,Res_Elem,Res_Oboz,Flag,o,p,s,h,f,m,n,v,w,KolGib,AA,bb,Res_Tol,
        F2,F1,F0 //Фланцы
        ,Res_KPU,Res_KPD,Res_list,Res_Ger:Integer;
        Res_Lop,Sten_V,Sten_G,hh,pp,Res_Kog_Priv,GG,Res_Kompl_priv,ww,www,
        Res_Detal,Os,R25,Ugol,Lenta,Kol_Otv,ID,Dl_Slova,
        DlinaI:Integer;

        Zak, Dat, Plan_Dat, Vn_Dat, Nom, Pos_Vst, Pos_Ml, R, Pereh, Privod,
        Zag_S,Zap_S, Dir_main, God, Mes, Fil,Dir,DlinRaz,ShirRaz,Mat,
        Dlina,Shir,Kol_Gib,Razm,Tol_S,Pos_Flan1,Pos_Flan,Pos_Dop,
        Pos_Sn,Pos_Privod,Pos_Isp,Pos_Ram,SvarkaS,Diam: string;
        Svar, Sbor, Izdel,VElem,Elem,Oboz, IDGP,Tip,EI,NC_S,Pos1,Pos2,Pos3,Pos4,Pos5,OboznSh,NC_S2,N2,N_S: string;

       Res_Tol55, Kol_Ed,Kol_Oboz,Kol_Ed1,Kol_Oboz1,NC,NC2,ND22,C,DlinI,ShirI,Tol,Razmet_KPU,DlinD,A_Lenta,NC_OB
       ,L,Kol_Lop,T,NCB,DlinaI1,DlinII,NC_GER: Double;

        Dat1, Dat2, Dat3: TDate;
        Kanban,Trumph,Nog,Ugl,Gib,Prok,Pila,PilaLent,Obraziv,Svarka,HACO,VALCI,KROMKA,TOKARN,ROTOR,Zig:Boolean;
        XL, XL1, XL2, XL_Temp, V1: Variant;
        Res_Tol1,Res_Tol2,Res_Resh,Res_N2,Res_Svarka,PosTrud,Res_N1,Res_KED,
        Flag_Razb_Klapana,// 1 Клапан состоит из двух половинок
        Klap,Flag1
        ,S1,S2,S3,S44,S4,BM,
        ST_EB//ТЕКИ 07.239.01.00.012 всегда ТРУМПФ
        ,VG:Integer;
        Trumpf_S,Kanban_S,Gib_S,NC_GER_S,SH1,SH2,Str:string;
        VG1,VG2,VG3,VG4,Res_N,Detal,Lop,xxx:Integer;//Круглые КПУ //Диск обл ВГ 439.02.00.003-250
                                              //Обечайка ТЕКИ 07.112.01.00.002-250
                                              //Подстав Ниппель  ВГ 356.00.00.002-200
                                              //Фланец ТЕКИ 07.112.01.00.101-250
        Razmer,res_s:Integer;
        DiamD:Double;
begin

        F_KPU := 0;
        F_KPD := 0;
        y := 1;
        Xy := 1;
        g:=1;
        K:=1;
        O:=1;
        S:=1;
        P:=1;
        H:=1;
        N:=1;
        M:=1;
        v:=1;
        w:=1;
        ww:=1;
        www:=1;
        Q:=1;
        qq:=1;
        AA:=1;
        PP:=1;
        e := 30;
        S1:=1;
        S2:=1;
        S3:=1;
        S44:=1;
        S4:=1;

        //Vn_Dat := FormatDateTime('mm.dd.yyyy', DateTimePicker1.Date);

        Clear_StringGrid(SD1);
        Clear_StringGrid(SD2);
        Clear_StringGrid(SD3);
        Clear_StringGrid(SD4);
        Clear_StringGrid(SD5);
        Clear_StringGrid(SD6);
        Clear_StringGrid(SD7);
        Clear_StringGrid(SD8);
        Clear_StringGrid(SD9);
        Clear_StringGrid(SD);
        Clear_StringGrid(Stenki);
        Clear_StringGrid(RSG1);
        Clear_StringGrid(RSG2);
        Clear_StringGrid(RSG3);
        Clear_StringGrid(RSG4);
        Clear_StringGrid(SG3);
        SD1.RowCount:=2;
        SD2.RowCount:=2;
        SD3.RowCount:=2;
        SD4.RowCount:=2;
        SD5.RowCount:=2;
        SD6.RowCount:=2;
        SD7.RowCount:=2;
        SD8.RowCount:=2;
        SD9.RowCount:=2;
        SD.RowCount:=2;
        Stenki.RowCount:=2;

        Memo2.Lines.Clear;
        SD1.Cells[2,0]:='Kanban True';
        SD2.Cells[2,0]:='Kanban False';
        SD3.Cells[2,0]:='Trumph';
        SD4.Cells[2,0]:='Nognicy';
        SD5.Cells[2,0]:='Gibka';
        SD6.Cells[2,0]:='Prokat';
        SD7.Cells[2,0]:='Obraziv';

        SD8.Cells[2,0]:='Uglorub';

        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                Dir_main := ExtractFileDir(ParamStr(0));
        God := FormatDateTime('yyyy', Now);
        Mes := FormatDateTime('mmmm', Now);
        Dir :=Form1.Put_KTO+'\CKlapana\Накладные(пожар)\' + God;
        CreateDir(Dir);

        Dir := Form1.Put_KTO+'\CKlapana\Накладные(пожар)\' + God + '\' + Mes+'\';
        CreateDir(Dir);

        XL := CreateOleObject('Excel.Application');
        CopyFile(PWideChar(Form1.Put_KTO+'\CKlapana\2013\Пожарные.xlsx'), PWideChar(Dir +'\№ '+Label15.Caption+'.xlsx'), False);
        Puty:=Dir + '\№ '+Label15.Caption+'.xlsx';
        XL.Workbooks.Open(Dir + '\№ '+Label15.Caption+'.xlsx');
        XL.Application.EnableEvents := false;
        For i:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[1].Range['B'+IntToStr(I+4)+':E'+IntToStr(I+4)].Merge;
               XL.ActiveWorkBook.WorkSheets[1].Rows[i+4].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[1].Rows[i+4].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[1].Cells[I+4, 1] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[1].Cells[I+4, 2] := SG1.Cells[4,i+1];
               XL.ActiveWorkBook.WorkSheets[1].Cells[I+4, 6] := SG1.Cells[1,i+1];
               Inc(q);
        end;
  for xxx := 1 to SG1.RowCount do
  begin
    if SG1.Cells[2, xxx] <> '' then
    begin
      Kol_Oboz:=0;
      Kol_Zap:=0;
      NC_OB:=0;
      NC:=0;
      NC_GER:=0;
      Flag1:=0;
      NC_S:='';
      NC_GER_S:='';
      NC_S2:='';
      Tip:='';
      Nom := SG1.Cells[2, xxx];
      Izdel := SG1.Cells[4, xxx];
      Kol_Zap := StrToInt(SG1.Cells[1, xxx]);
      IDGP  := SG1.Cells[ 8, xxx];
      if (Form1.Zapusk2=2) OR (Form1.Zapusk2=1) Then
      Begin
                ADOQuery2.Close;
                ADOQuery2.SQL.Clear;
                ADOQuery2.SQL.Text :='Select * from Специф Where   (IdГП=' + #39 +
                IDGP+ #39 +') AND (ВидЭлемента<>' + #39 +'Сборочные единицы'+ #39 +
                ') AND (ВидЭлемента<>' + #39 +'Детали'+ #39 +') ORDER BY Длина ' ;
                ADOQuery2.Active := true;
                Flag_Razb_Klapana:=0;

                 ADOQuery2.First;
                 RSG1.RowCount:= RSG1.RowCount+ ADOQuery2.RecordCount ;
                 for J:=0 To  ADOQuery2.RecordCount-1 Do
                 Begin
                      Flag:=0;

                      KolGib:=0;
                      VElem:= ADOQuery2.FieldByName('ВидЭлемента').AsString;
                      Elem:= ADOQuery2.FieldByName('Элемент').AsString;
                      Oboz:= ADOQuery2.FieldByName('Обозначение').AsString;
                      //Kol_Ed:= ADOQuery2.FieldByName('КолНаЕд').AsFloat;
                      Str:=ADOQuery2.FieldByName('Количество').AsString;
                        If Str='' then
                        Kol_Ed:=0
                        else
                     Kol_Ed:= StrToFloat(StringReplace(Str,'.',',',[rfReplaceAll]));
                     If Str='' then
                        Kol_Oboz:=0
                      else
                      Kol_Oboz:= StrToFloat(StringReplace(Str,'.',',',[rfReplaceAll]));
                      //Trumph,Nog,Ugl,Gib,Prok,Pila
                      Dlina:= ADOQuery2.FieldByName('Ширина').AsString;
                      Shir:= ADOQuery2.FieldByName('Длина').AsString;
                      Diam:= ADOQuery2.FieldByName('Диаметр').AsString;
                      if Diam='' then
                      DiamD:=0
                      else
                      DiamD:=StrToFloat(StringReplace(Diam,'.',',',[rfReplaceAll]));
                      //Dlina:= Form1.ADOQuery2.FieldByName('Длина').AsString;
                      //Shir:= Form1.ADOQuery2.FieldByName('Ширина').AsString;
                      DlinRaz:= StringReplace(ADOQuery2.FieldByName('ШиринаРазв').AsString,'.',',',[rfReplaceAll]);
                      ShirRaz:= StringReplace(ADOQuery2.FieldByName('ДлинаРазв').AsString,'.',',',[rfReplaceAll]);
                      if DlinRaz ='' Then
                            DlinRaz:='0';
                      if ShirRaz ='' Then
                            ShirRaz:='0';
                      Kol_Gib:= ADOQuery2.FieldByName('КолГибов').AsString;
                      if Kol_Gib=''  Then
                      KolGib:=0
                      Else
                      KolGib:=StrToInt(Kol_Gib);

                      Mat:= ADOQuery2.FieldByName('Материал').AsString;
                      if Mat='' Then
                                        Mat:='0';
                      EI:= ADOQuery2.FieldByName('ЕИ').AsString;

                      Kanban:= ADOQuery2.FieldByName('Канбан').AsBoolean;
                      Trumph:= ADOQuery2.FieldByName('Trumph').AsBoolean;

                      SH1:=StringReplace(ADOQuery2.FieldByName('ШиринаПолки1').AsString,'.',',',[rfReplaceAll]);
                      if SH1='' Then
                      SH1:='0';
                      SH2:=StringReplace(ADOQuery2.FieldByName('ШиринаПолки2').AsString,'.',',',[rfReplaceAll]);
                      if SH2='' Then
                      SH2:='0';
                      RSG1.Cells[0,s]:=IntToStr(s);
                      RSG1.Cells[1,s]:=Tip;
                      RSG1.Cells[2,s]:=Elem;
                      RSG1.Cells[3,s]:=FloatToStr(Kol_Oboz*Kol_Zap);
                      RSG1.Cells[4,s]:=Diam;
                      RSG1.Cells[5,s]:=Oboz;
                      //RSG1.Cells[6,s]:=Dlina;
                      //RSG1.Cells[7,s]:=Shir;
                      RSG1.Cells[6,s]:=Kol_Gib;
                      RSG1.Cells[7,s]:=Mat;
                      RSG1.Cells[8,s]:=N_S;//Н\ч
                      inc(S);
                      ADOQuery2.Next;
                 End;


        Q:=Q+2;
        {XL.ActiveWorkBook.WorkSheets[1].Rows[q+2].Font.Bold := True;
        XL.ActiveWorkBook.WorkSheets[1].Rows[q+3].Font.Bold := True;
        XL.ActiveWorkBook.WorkSheets[1].Rows[q+3].Font.Size := 14;
        XL.ActiveWorkBook.WorkSheets[1].Rows[q+2].Font.Size := 14;
        XL.ActiveWorkBook.WorkSheets[1].Select;

        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 2] := '№';
        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 3] := 'Тип клапана';
        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 4] := 'Изделие';
        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 5] := 'Размер';
        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 6] := 'Кол на 1 клапан';
        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 7] := 'Общее кол-во';
        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 8] := 'ЕдИзм';
        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 9] := 'Обознач';
        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 10] := 'Мат';
       // XL.ActiveWorkBook.WorkSheets[1].Cells[q+2, 4] := 'Kanban';  }

        For j:=0 To RSG1.RowCount-1 Do
        Begin
            Razm:=RSG1.Cells[7,j+1]+'*'+RSG1.Cells[7,j+1];
            if RSG1.Cells[0,j+1]='' Then
            Continue;


            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 2] := RSG1.Cells[0,j+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 3] := RSG1.Cells[2,j+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 4] := RSG1.Cells[3,j+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 5] := Razm;//Razmer
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 6] := RSG1.Cells[4,j+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 7] := RSG1.Cells[5,j+1];

            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 8] := RSG1.Cells[1,j+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 9] := RSG1.Cells[6,j+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 10] := RSG1.Cells[12,j+1];
            XL.ActiveWorkBook.WorkSheets[1].Range['B'+IntToStr(Q+5), 'J' +
                IntToStr(q+5)].Borders.LineStyle := 1;
            XL.ActiveWorkBook.WorkSheets[1].Cells[2, 6] :=Label15.Caption;
            Inc(Q);
        end;
        Q:=Q+2;

    End;
    end;
  end;
  XL.Visible := True;
  XL := UnAssigned;
end;
function TFNewNakl.Float(var Str:String): String;
Var PosTrud:Integer;
Begin
        Result:='';
  PosTrud := Pos('.', Str); //Трудоемкость FLOAT
  if PosTrud <> 0 then
  begin
    Delete(Str, PosTrud, 1);
    Insert(',', Str, PosTrud);
  end;
  Result:=Str;

end;
procedure TFNewNakl.Button5Click(Sender: TObject);
Var I:Integer;
Res,Res1,Res2:Integer;
Nam,Oboz:String;
begin
              For I:=1 To SD1.RowCount Do
              Begin

              end;
              Memo1.Lines.Add(SG1.Cells[1,i]);
end;

procedure TFNewNakl.SpeedButton1Click(Sender: TObject);
begin
        Form1.ExportGridtoExcel2(SD1);
end;

procedure TFNewNakl.btn1Click(Sender: TObject);
var
        Kol_ZAP, Kol, Kol_Zap_Ranee, Res, Nomer, Zak_Int, i, A, B, D, y, F_KPU,
        F_KPD,xy, Res1, e, Zag, Zap, j, g, k,Q,Res_Mat,Res_Proch,Res_Stand,qq,
        Res_Sbor,Res_Klap,Podstavka,Srav,Srav1:Integer;
        Res_VElem,Res_Elem,Res_Oboz,Flag,o,p,s,h,f,m,n,v,w,KolGib,AA,bb,Res_Tol,
        F2,F1,F0 //Фланцы
        ,Res_KPU,Res_KPD,Res_list,Res_Ger:Integer;
        Res_Lop,Sten_V,Sten_G,hh,pp,Res_Kog_Priv,GG,Res_Kompl_priv,ww,www,
        Res_Detal,Os,R25,Ugol,Lenta,Kol_Otv,ID,Dl_Slova,
        DlinaI:Integer;

        Zak, Dat, Plan_Dat, Vn_Dat, Nom, Pos_Vst, Pos_Ml, R, Pereh, Privod,
        Zag_S,Zap_S, Dir_main, God, Mes, Fil,Dir,DlinRaz,ShirRaz,Mat,
        Dlina,Shir,Kol_Gib,Razm,Tol_S,Pos_Flan1,Pos_Flan,Pos_Dop,
        Pos_Sn,Pos_Privod,Pos_Isp,Pos_Ram,SvarkaS,Diam: string;
        Svar, Sbor, Izdel,VElem,Elem,Oboz, IDGP,Tip,EI,NC_S,Pos1,Pos2,Pos3,Pos4,Pos5,OboznSh,NC_S2,N2,N_S: string;

       Res_Tol55, Kol_Ed,Kol_Oboz,Kol_Ed1,Kol_Oboz1,NC,NC2,ND22,C,DlinI,ShirI,Tol,Razmet_KPU,DlinD,A_Lenta,NC_OB
       ,L,Kol_Lop,T,NCB,DlinaI1,DlinII,NC_GER: Double;

        Dat1, Dat2, Dat3: TDate;
        Kanban,Trumph,Nog,Ugl,Gib,Prok,Pila,PilaLent,Obraziv,Svarka,HACO,VALCI,KROMKA,TOKARN,ROTOR,Zig:Boolean;
        XL, XL1, XL2, XL_Temp, V1: Variant;
        Res_Tol1,Res_Tol2,Res_Resh,Res_N2,Res_Svarka,PosTrud,Res_N1,Res_KED,
        Flag_Razb_Klapana,// 1 Клапан состоит из двух половинок
        Klap,Flag1
        ,S1,S2,S3,S44,S4,BM,
        ST_EB//ТЕКИ 07.239.01.00.012 всегда ТРУМПФ
        ,VG:Integer;
        Trumpf_S,Kanban_S,Gib_S,NC_GER_S,SH1,SH2,Str:string;
        VG1,VG2,VG3,VG4,Res_N,Rec_Count:Integer;//Круглые КПУ //Диск обл ВГ 439.02.00.003-250
                                              //Обечайка ТЕКИ 07.112.01.00.002-250
                                              //Подстав Ниппель  ВГ 356.00.00.002-200
                                              //Фланец ТЕКИ 07.112.01.00.101-250
        Razmer,res_s:Integer;
        DiamD:Double;
begin

        F_KPU := 0;
        F_KPD := 0;
        y := 1;
        Xy := 1;
        g:=1;
        K:=1;
        O:=1;
        S:=1;
        P:=1;
        H:=1;
        N:=1;
        M:=1;
        v:=1;
        w:=1;
        ww:=1;
        www:=1;
        Q:=1;
        qq:=1;
        AA:=1;
        PP:=1;
        e := 30;
        S1:=1;
        S2:=1;
        S3:=1;
        S44:=1;
        S4:=1;

        //Vn_Dat := FormatDateTime('mm.dd.yyyy', DateTimePicker1.Date);

        Clear_StringGrid(SD1);
        Clear_StringGrid(SD2);
        Clear_StringGrid(SD3);
        Clear_StringGrid(SD4);
        Clear_StringGrid(SD5);
        Clear_StringGrid(SD6);
        Clear_StringGrid(SD7);
        Clear_StringGrid(SD8);
        Clear_StringGrid(SD9);
        Clear_StringGrid(SD);
        Clear_StringGrid(Stenki);
        Clear_StringGrid(RSG1);
        Clear_StringGrid(RSG2);
        Clear_StringGrid(RSG3);
        Clear_StringGrid(RSG4);
        Clear_StringGrid(SG3);
        SD1.RowCount:=2;
        SD2.RowCount:=2;
        SD3.RowCount:=2;
        SD4.RowCount:=2;
        SD5.RowCount:=2;
        SD6.RowCount:=2;
        SD7.RowCount:=2;
        SD8.RowCount:=2;
        SD9.RowCount:=2;
        SD.RowCount:=2;
        Stenki.RowCount:=2;

        Memo2.Lines.Clear;
        SD1.Cells[2,0]:='Kanban True';
        SD2.Cells[2,0]:='Kanban False';
        SD3.Cells[2,0]:='Trumph';
        SD4.Cells[2,0]:='Nognicy';
        SD5.Cells[2,0]:='Gibka';
        SD6.Cells[2,0]:='Prokat';
        SD7.Cells[2,0]:='Obraziv';

        SD8.Cells[2,0]:='Uglorub';


        for i := 1 to SG1.RowCount do
        begin
                Flag1:=0;
                if SG1.Cells[2, i] <> '' then
                begin
                        Kol_Oboz:=0;
                        Kol_Zap:=0;
                        NC_OB:=0;
                        NC:=0;
                        NC_GER:=0;
                        Flag1:=0;
                        NC_S:='';
                        NC_GER_S:='';
                        NC_S2:='';
                        Tip:='';
                        Nom := SG1.Cells[2, i];
                        Izdel := SG1.Cells[4, i];
                        Kol_Zap := StrToInt(SG1.Cells[1, i]);
                        Res := Pos('КПД', Izdel);
                        if Res <> 0 then
                        Tip:='КПД';
                        Res := Pos('КПУ', Izdel);
                        if Res <> 0 then
                        Tip:='КПУ';
                        //Клапан ГЕРМИК-ДУ-710*710-2*ф-ЭМП220-сн-0-0
                        res:=Pos('ГЕРМИК-ДУ',IZdel);
                        if Res<>0 Then
                        Tip:='ГЕРМИК-ДУ';

                      Izdel := SG1.Cells[4, i];
                      Res_Kpd:=Pos('КПД',Izdel);
                      Res_Kpu:=Pos('КПУ',Izdel);
                      Res_Ger:=Pos('ГЕРМИК-ДУ',Izdel);
                      Res_KED:=Pos('КЭД',Izdel);
                      //Клапан КЭД-03-900*350-2*ф-MВ220-сн-р25-0
                      if Res_KED <> 0 then
                        begin
                                Kol := StrToInt(SG1.Cells[1, i]);
                                Res := Pos(' ', Izdel);
                                Delete(Izdel, 1, Res);
                                //========================================
                                Res := Pos('-', Izdel);
                                Pos1 := Copy(Izdel, 1, Res - 1);
                                Delete(Izdel, 1, Res);
                                //========================================
                                Res := Pos('-', Izdel);
                                Pos2 := Copy(Izdel, 1, Res - 1);
                                Delete(Izdel, 1, Res);
                                //========================================
                                Res := Pos('*', Izdel);
                                A := StrToInt(Copy(Izdel, 1, Res - 1));
                                Delete(Izdel, 1, Res);
                                //========================================
                                Res := Pos('-', Izdel);
                                B := StrToInt(Copy(Izdel, 1, Res - 1));
                                Delete(Izdel, 1, Res);
                                //========================================
                                Res := Pos('*', Izdel);
                                Pos_Flan := Copy(Izdel, 1, Res - 1);
                                Delete(Izdel, 1, Res);
                                //========================================
                                Res := Pos('-', Izdel);
                                Pos_Flan1 := Copy(Izdel, 1, Res - 1);
                                Delete(Izdel, 1, Res);
                                //========================================
                                Res := Pos('-', Izdel);
                                Pos_Privod := Copy(Izdel, 1, Res - 1);
                                Delete(Izdel, 1, Res);
                                //========================================
                                Res := Pos('-', Izdel);
                                Pos_Sn := Copy(Izdel, 1, Res - 1);
                                Delete(Izdel, 1, Res);
                                //========================================
                                Res := Pos('-', Izdel);
                                Pos_Dop := Copy(Izdel, 1, Res - 1);
                                Delete(Izdel, 1, Res);
                                //========================================
                                Pos_Ram := Izdel;
                                //========================================

                        End;
                      //++++++++++++++++++++++++++++++++
                      if Res_Kpu<>0 Then
                      Begin
                      Res := Pos(' ', Izdel);
                        Delete(Izdel, 1, Res);
                        //======================================== KPU
                        Res := Pos('-', Izdel);
                        Pos1 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //======================================== 1 H
                        Res := Pos('-', Izdel);
                        N2:=Copy(Izdel,1,2);
                        Pos2 := Copy(Izdel, 1, 1);
                        Pos3 := Copy(Izdel, 2, 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        //Клапан КПУ-1Н-О-Н-600*400-2*ф-MB220-сн-0-0-0-0-0-0
                        //Клапан КПД-4-01-500*500-1*ф-ЭМП220-вн-0-0
                        //Клапан КЭД-03-900*350-2*ф-MВ220-сн-р25-0
                        //========================================
                        Res := Pos('-', Izdel);
                        Pos4 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Res := Pos('-', Izdel);
                        Pos5 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        //Клапан КПУ-2Н-О-Н-200-2*ф-MB220-T-сн-0-0-0-2*100-0-0
                        //Клапан КПУ-1Н-О-Н-200*200-2*ф-MB220-сн-0-с-0-0-0-0
                        Res := Pos('-', Izdel);//Проверка на круглый если -
                        If Res >5 Then //200*200- Квадрат
                        Begin
                                Res := Pos('*', Izdel);
                                A := StrToInt(Copy(Izdel, 1, Res - 1));
                                Delete(Izdel, 1, Res);
                                //========================================
                                Res := Pos('-', Izdel);
                                B := StrToInt(Copy(Izdel, 1, Res - 1));
                                Delete(Izdel, 1, Res);
                        end
                        else
                        begin
                                Res := Pos('-', Izdel);
                                A := StrToInt(Copy(Izdel, 1, Res - 1));
                                Delete(Izdel, 1, Res);
                                B:=0;
                        end;
                        {Res := Pos('*', Izdel);

                        A := StrToInt(Copy(Izdel, 1, Res - 1));
                        Delete(Izdel, 1, Res);
                        //========================================
                        Res := Pos('-', Izdel);
                        B := StrToInt(Copy(Izdel, 1, Res - 1));
                        Delete(Izdel, 1, Res); }

                      end;
                        //++++++++++++++++++++++++++++++++++++++++++++++++++++++
                        //Клапан ГЕРМИК-ДУ-710*710-2*ф-ЭМП220-сн-0-0-
                        //Клапан ГЕРМИК-ДУ-Д-710*710-2*ф-ЭМП220-сн-0-0-
                        if Res_Ger<>0 Then
                        Begin
                        Res := Pos(' ', Izdel);
                        Delete(Izdel, 1, Res);
                        //======================================== ГЕРМИК
                        Res := Pos('-', Izdel);
                        Pos1 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //======================================== ДУ
                        Res := Pos('-', Izdel);
                        Delete(Izdel, 1, Res);
                        //710*710-
                        //Д-710*710-
                        Res := Pos('*', Izdel);
                        Res1 := Pos('-', Izdel);
                        if (Res1<Res){Д-710*710-} then
                             Delete(Izdel, 1, Res1);
                        //========================================
                        Res := Pos('*', Izdel);
                        A := StrToInt(Copy(Izdel, 1, Res - 1));
                        Delete(Izdel, 1, Res);
                        //========================================
                        Res := Pos('-', Izdel);
                        B := StrToInt(Copy(Izdel, 1, Res - 1));
                        Delete(Izdel, 1, Res);
                        end;
                        //+++++++++++++++++++++++++++++++++++++++++++++++++++++
                        Res_KPD := Pos('КПД', Izdel);
                        if Res_KPD <> 0 then
                        begin

                                Kol := StrToInt(SG1.Cells[1, i]);
                                Res := Pos(' ', Izdel);
                                Delete(Izdel, 1, Res);
                                //========================================
                                Res := Pos('-', Izdel);
                                Pos1 := Copy(Izdel, 1, Res - 1);
                                Delete(Izdel, 1, Res);
                                //========================================
                                Res := Pos('-', Izdel);
                                Pos2 := Copy(Izdel, 1, Res - 1);
                                Delete(Izdel, 1, Res);
                                //========================================
                                Res := Pos('-', Izdel);
                                Pos_Isp := Copy(Izdel, 1, Res - 1);
                                Delete(Izdel, 1, Res);
                                //========================================
                                Res := Pos('*', Izdel);
                                A := StrToInt(Copy(Izdel, 1, Res - 1));
                                Delete(Izdel, 1, Res);
                                //========================================
                                Res := Pos('-', Izdel);
                                B := StrToInt(Copy(Izdel, 1, Res - 1));
                                Delete(Izdel, 1, Res);
                                //========================================
                                Res := Pos('*', Izdel);
                                Pos_Flan := Copy(Izdel, 1, Res - 1);
                                Delete(Izdel, 1, Res);
                                //========================================
                                Res := Pos('-', Izdel);
                                Pos_Flan1 := Copy(Izdel, 1, Res - 1);
                                Delete(Izdel, 1, Res);
                                //========================================
                                Res := Pos('-', Izdel);
                                Pos_Privod := Copy(Izdel, 1, Res - 1);
                                Delete(Izdel, 1, Res);
                                //========================================
                                Res := Pos('-', Izdel);
                                Pos_Sn := Copy(Izdel, 1, Res - 1);
                                Delete(Izdel, 1, Res);
                                //========================================
                                Res := Pos('-', Izdel);
                                Pos_Dop := Copy(Izdel, 1, Res - 1);
                                Delete(Izdel, 1, Res);
                                //========================================
                                Pos_Ram := Izdel;
                                //========================================

                        End;
                      Izdel := SG1.Cells[4, i];
                      Res_Kpd:=Pos('КПД',Izdel);
                      Res_Kpu:=Pos('КПУ',Izdel);
                      if Res_KED<>0 Then
                          Res_Kpd:=1;
                      F2:=Pos('2*ф',Izdel);
                      F1:=Pos('1*ф',Izdel);
                      F0:=Pos('0*ф',Izdel);
                      IDGP:=SG1.Cells[8, i];
                      if F2<>0 Then
                      Tip:=Tip+'-'+Pos2+Pos3+'-'+IntToStr(A)+'*'+IntToStr(B)+'-2*ф';
                      if F1<> 0 Then
                      Tip:=Tip+'-'+Pos2+Pos3+'-'+IntToStr(A)+'*'+IntToStr(B)+'-1*ф';
                      if F0<> 0 Then
                      Tip:=Tip+'-'+Pos2+Pos3+'-'+IntToStr(A)+'*'+IntToStr(B)+'-0*ф';

                        {if not Form1.mkQuerySelect(Form1.ADOQuery2,
                                'Select * from %s Where   (IdГП=' + #39 +
                                IDGP+ #39 +')  ORDER BY ДлинаИ', ['Специф']) then
                                // IDGP+ #39 +') AND (ВидЭлемента=' + #39 +'Детали'+ #39 +') ORDER BY Длина', ['Специф']) then
                                exit;  }
                      ADOQuery2.Close;
                      ADOQuery2.SQL.Clear;
                      ADOQuery2.SQL.Text :='Select * from Специф Where   (IdГП=' + #39 +
                                IDGP+ #39 +')  ORDER BY ДлинаИ ' ;
                        ADOQuery2.Active := true;
                Flag_Razb_Klapana:=0;
                 for J:=0 To  ADOQuery2.RecordCount-1 Do  //Проверка на Клапан разбитый на 2
                 Begin
                        Elem:= ADOQuery2.FieldByName('Элемент').AsString;
                        Str:=ADOQuery2.FieldByName('Количество').AsString;
                        If Str='' then
                        Kol_Ed:=0
                        else
                        Kol_Ed:= ADOQuery2.FieldByName('Количество').AsFloat;
                        Klap:=AnsiCompareStr('Клапан',Elem);
                        if (Klap=0) and (Kol_Ed>1) then
                        Begin
                                Flag_Razb_Klapana:=1;//Клапан разбит на 2
                                Break;
                        End;
                        Res_Lop:=AnsiCompareStr('Стенка лопатки',Elem);
                        if Res_Lop=0 Then
                        Begin
                                Kol_lop:= ADOQuery2.FieldByName('Количество').AsFloat;
                        End;
                        ADOQuery2.Next;
                 end;
                 ADOQuery2.First;
                 for J:=0 To  ADOQuery2.RecordCount-1 Do
                 Begin
                      Flag:=0;

                      KolGib:=0;
                      VElem:= ADOQuery2.FieldByName('ВидЭлемента').AsString;
                      Elem:= ADOQuery2.FieldByName('Элемент').AsString;
                      Oboz:= ADOQuery2.FieldByName('Обозначение').AsString;
                      //Kol_Ed:= ADOQuery2.FieldByName('КолНаЕд').AsFloat;
                      Str:=ADOQuery2.FieldByName('Количество').AsString;
                        If Str='' then
                        Kol_Ed:=0
                        else
                     Kol_Ed:= ADOQuery2.FieldByName('Количество').AsFloat;
                     If Str='' then
                        Kol_Oboz:=0
                      else
                      Kol_Oboz:= ADOQuery2.FieldByName('Количество').AsFloat;
                      //Trumph,Nog,Ugl,Gib,Prok,Pila
                      Dlina:= ADOQuery2.FieldByName('Ширина').AsString;
                      Shir:= ADOQuery2.FieldByName('Длина').AsString;
                      Diam:= ADOQuery2.FieldByName('Диаметр').AsString;
                      if Diam='' then
                      DiamD:=0
                      else
                      DiamD:=StrToFloat(Diam);
                      //Dlina:= Form1.ADOQuery2.FieldByName('Длина').AsString;
                      //Shir:= Form1.ADOQuery2.FieldByName('Ширина').AsString;
                      DlinRaz:= ADOQuery2.FieldByName('ШиринаРазв').AsString;
                      ShirRaz:= ADOQuery2.FieldByName('ДлинаРазв').AsString;
                      if DlinRaz ='' Then
                            DlinRaz:='0';
                      if ShirRaz ='' Then
                            ShirRaz:='0';
                      Kol_Gib:= ADOQuery2.FieldByName('КолГибов').AsString;
                      if Kol_Gib=''  Then
                      KolGib:=0
                      Else
                      KolGib:=StrToInt(Kol_Gib);

                      Mat:= ADOQuery2.FieldByName('Материал').AsString;
                      if Mat='' Then
                                        Mat:='0';
                      EI:= ADOQuery2.FieldByName('ЕИ').AsString;

                      Kanban:= ADOQuery2.FieldByName('Канбан').AsBoolean;
                      Trumph:= ADOQuery2.FieldByName('Trumph').AsBoolean;
                      Nog:= ADOQuery2.FieldByName('Ножницы').AsBoolean;
                      Ugl:= ADOQuery2.FieldByName('Углоруб').AsBoolean;
                      Gib:= ADOQuery2.FieldByName('Гибка').AsBoolean;
                      Prok:= ADOQuery2.FieldByName('Прокатка').AsBoolean;
                      Pila:= ADOQuery2.FieldByName('Пила').AsBoolean;
                      PilaLent:= ADOQuery2.FieldByName('Пила ленточная').AsBoolean;
                     { Obraziv:= ADOQuery2.FieldByName('Образив').AsBoolean;
                      HACO:= ADOQuery2.FieldByName('HACO').AsBoolean;
                      VALCI:= ADOQuery2.FieldByName('Вальцы').AsBoolean;
                      KROMKA:= ADOQuery2.FieldByName('Кромкогиб').AsBoolean;}
                      TOKARN:= ADOQuery2.FieldByName('Токарный').AsBoolean;
                      ROTOR:= ADOQuery2.FieldByName('Ротор').AsBoolean;
                      SH1:=ADOQuery2.FieldByName('ШиринаПолки1').AsString;
                      SH2:=ADOQuery2.FieldByName('ШиринаПолки2').AsString;

                      //==================================================


                      Res_Mat:=Pos('Материалы',VELem);
                     Res_Proch:=Pos('Прочие изделия',VELem);
                     Res_Stand:=Pos('Стандартные изделия',VELem);
                     Res_Sbor:=Pos('Сборочные единицы',VELem);
                     Res_Klap:=Pos('Клапан',ELem);
                     Res_Kompl_priv:=Pos('Комплект привода',ELem);
                     Os:= Pos('Ось',ELem);
                      hh:=A MOD 50;
                      gg:=B MOD 50;

                      //XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
                      Res_Detal:=AnsiCompareStr('Детали',VElem);
                      if Res_Detal=0 Then
                      MemVg.Lines.Add(Elem+'   '+Shir);
                      OboznSh:=Oboz;

                      {if B=0 then
                      begin
                      Res:=Pos(' - ',OboznSh);
                      if (Res<>0) Then                      //   AND ((Sten_G<>0) ANd (Sten_V<>0))
                        Delete(OboznSh,Res,30);
                      BM:=Pos('ВМ-',OboznSh);
                      Res1:=Pos('-',OboznSh);
                      if (BM=0) and (Res1<>0) then
                         Delete(OboznSh,Res1,30);
                      End
                      else
                      begin
                        Res:=Pos('-',OboznSh);
                        if (Res<>0) Then                      //   AND ((Sten_G<>0) ANd (Sten_V<>0))
                        Delete(OboznSh,Res,30);
                      end;}
                      if (Res_Detal=0) Then
                      Begin
                      if not Form1.mkQuerySelect66(Form1.ADOQuery4,
                                'Select * from %s Where   (Обозначение=' + #39 +
                                OboznSh + #39 +') AND (Технолог=1)  ORDER BY Длина ', ['СпецифОбщая']) then
                                //OboznSh + #39 +') ', ['Шаблон']) then
                                exit;
                      Rec_Count:= Form1.ADOQuery4.RecordCount;
                      Kanban:= Form1.ADOQuery4.FieldByName('Канбан').AsBoolean;
                      Trumph:= Form1.ADOQuery4.FieldByName('Trumph').AsBoolean;
                      Nog:= Form1.ADOQuery4.FieldByName('Ножницы').AsBoolean;
                      Ugl:= Form1.ADOQuery4.FieldByName('Углоруб').AsBoolean;
                      Gib:= Form1.ADOQuery4.FieldByName('Гибка').AsBoolean;
                      Prok:= Form1.ADOQuery4.FieldByName('Прокатка').AsBoolean;
                      Pila:= Form1.ADOQuery4.FieldByName('Пила').AsBoolean;
                      PilaLent:= Form1.ADOQuery4.FieldByName('ПилаЛенточная').AsBoolean;
                      //Obraziv:= Form1.ADOQuery4.FieldByName('Образив').AsBoolean;
                      Svarka:=Form1.ADOQuery4.FieldByName('Сварка').AsBoolean;
                      //Ugol:= Form1.ADOQuery4.FieldByName('Угол').AsInteger;
                      //Kol_Gib:= Form1.ADOQuery4.FieldByName('КолГибов').AsString;
                      //HACO:= Form1.ADOQuery4.FieldByName('HACO').AsBoolean;
                      VALCI:= Form1.ADOQuery4.FieldByName('Вальцы').AsBoolean;
                      KROMKA:= Form1.ADOQuery4.FieldByName('Кромкогиб').AsBoolean;
                      TOKARN:= Form1.ADOQuery4.FieldByName('Токарный').AsBoolean;
                      ROTOR:= Form1.ADOQuery4.FieldByName('Ротор').AsBoolean;
                      //Zig:= Form1.ADOQuery4.FieldByName('Зиг').AsBoolean;
                      N_S:= Form1.ADOQuery4.FieldByName('Н\чГибка').AsString;
                        if (Res_Kpu<>0) AND (B=0) Then
                        Begin

                                if VALCI =True then
                                begin
                                RSG1.Cells[0,s1]:=IntToStr(s1);
                                RSG1.Cells[1,s1]:=Tip;
                                RSG1.Cells[2,s1]:=Elem;
                                RSG1.Cells[3,s1]:=FloatToStr(Kol_Oboz*Kol_Zap);
                                RSG1.Cells[4,s1]:=Diam;
                                RSG1.Cells[5,s1]:=Oboz;
                                //RSG1.Cells[6,s1]:=Dlina;
                                RSG1.Cells[6,s1]:=Kol_Gib;
                                RSG1.Cells[7,s1]:=Mat;
                                RSG1.Cells[8,s1]:=N_S;//Н\ч
                                Inc(S1);
                                RSG1.RowCount:= RSG1.RowCount+1;
                                end;
                                if KROMKA =True then
                                begin
                                RSG2.Cells[0,s2]:=IntToStr(s2);
                                RSG2.Cells[1,s2]:=Tip;
                                RSG2.Cells[2,s2]:=Elem;
                                RSG2.Cells[3,s2]:=FloatToStr(Kol_Oboz*Kol_Zap);
                                RSG2.Cells[4,s2]:=Diam;
                                RSG2.Cells[5,s2]:=Oboz;
                                //RSG2.Cells[6,s2]:=Dlina;
                                RSG2.Cells[6,s2]:=Kol_Gib;
                                RSG2.Cells[7,s2]:=Mat;
                                RSG2.Cells[8,s2]:=N_S;//Н\ч
                                Inc(S2);
                                RSG2.RowCount:= RSG2.RowCount+1;
                                end;
                                if ROTOR =True then
                                begin
                                RSG3.Cells[0,s3]:=IntToStr(s3);
                                RSG3.Cells[1,s3]:=Tip;
                                RSG3.Cells[2,s3]:=Elem;
                                RSG3.Cells[3,s3]:=FloatToStr(Kol_Oboz*Kol_Zap);
                                RSG3.Cells[4,s3]:=Diam;
                                RSG3.Cells[5,s3]:=Oboz;
                                //RSG3.Cells[6,s3]:=Dlina;
                                RSG3.Cells[6,s3]:=Kol_Gib;
                                RSG3.Cells[7,s3]:=Mat;
                                RSG3.Cells[8,s3]:=N_S;//Н\ч
                                Inc(S3);
                                RSG3.RowCount:= RSG3.RowCount+1;
                                end;
                                if TOKARN =True then
                                begin
                                RSG4.Cells[0,s4]:=IntToStr(s4);
                                RSG4.Cells[1,s4]:=Tip;
                                RSG4.Cells[2,s4]:=Elem;
                                RSG4.Cells[3,s4]:=FloatToStr(Kol_Oboz*Kol_Zap);
                                RSG4.Cells[4,s4]:=Diam;
                                RSG4.Cells[5,s4]:=Oboz;
                                //RSG4.Cells[6,s]:=Dlina;
                                RSG4.Cells[6,s4]:=Kol_Gib;
                                RSG4.Cells[7,s4]:=Mat;
                                RSG4.Cells[8,s4]:=N_S;//Н\ч
                                Inc(S4);
                                RSG4.RowCount:= RSG4.RowCount+1;
                                end;
                                if Zig =True then
                                begin
                                SG3.Cells[0,s44]:=IntToStr(s44);
                                SG3.Cells[1,s44]:=Tip;
                                SG3.Cells[2,s44]:=Elem;
                                SG3.Cells[3,s44]:=FloatToStr(Kol_Oboz*Kol_Zap);
                                SG3.Cells[4,s44]:=Diam;
                                SG3.Cells[5,s44]:=Oboz;
                                //RSG4.Cells[6,s]:=Dlina;
                                SG3.Cells[6,s44]:=Kol_Gib;
                                SG3.Cells[7,s44]:=Mat;
                                SG3.Cells[8,s44]:=N_S;//Н\ч
                                Inc(S44);
                                SG3.RowCount:= SG3.RowCount+1;
                                end;
                        end;
                       if Kol_Gib=''  Then
                       KolGib:=0
                       Else
                       KolGib:=StrToInt(Kol_Gib);
                      end;
                      //XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
                      If (Svarka=True) Then
                      Begin
                            Dl_Slova:=Length(OboznSh);
                            Delete(OboznSh,Dl_Slova,1);
                      end;
        //SSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSS    STENKI
                      Sten_V:=Pos('Стенка вертикальная',Elem);
                      Sten_G:=Pos('Стенка горизонтальная',Elem);

                      NC:=0;
                      NC_S:='';
                      Tol:=0;
                      Tol_S:='';
                      //==============================================
                      Res_Stand:=Pos('Стандартные изделия',VELem);
                      if (Kanban=True) AND (REs_Stand=0) AND (Res_Sbor=0) AND (Res_Mat=0)   Then
                      Begin
                      For y:=1 To SD1.RowCount Do
                      Begin
                           Res_Elem:=AnsiCompareStr(Elem,SD1.Cells[3,y]);
                           Res_Oboz:=AnsiCompareStr(Oboz,SD1.Cells[6,y]);
                           if (Res_Elem=0) AND (Res_Oboz=0) Then
                           Begin

                               Kol_Ed:=Kol_Ed+StrToFloat(SD1.Cells[4,y]);
                               SD1.Cells[4,y]:=FloatToStr(Kol_Ed);

                               Kol_Oboz:=(Kol_Oboz*Kol_Zap)+StrToFloat(SD1.Cells[5,y]);
                               SD1.Cells[5,y]:=FloatToStr(Kol_Oboz);
                               Flag:=1;
                               Break;
                           end ;
                     end;
                     If Flag=0 Then
                     Begin
                                SD1.Cells[2,k]:=Tip;
                                SD1.Cells[3,k]:=Elem;
                                SD1.Cells[4,k]:=FloatToStr(Kol_Ed);
                                SD1.Cells[5,k]:=FloatToStr(Kol_Ed*Kol_Zap);
                                SD1.Cells[6,k]:=Oboz;
                                SD1.Cells[0,k]:=IntToStr(k);
                                SD1.Cells[1,k]:=EI;
                                SD1.Cells[7,k]:=Dlina;
                                SD1.Cells[8,k]:=Shir;
                                SD1.Cells[9,k]:=DlinRaz;
                                SD1.Cells[10,k]:=ShirRaz;
                                SD1.Cells[11,k]:=Kol_Gib;
                                SD1.Cells[12,k]:=Mat;
                                SD1.Cells[13,k]:=VElem;
                                Inc(k);
                                SD1.RowCount:= SD1.RowCount+1;
                                //Break;
                     end;
                     end;

                     //Заготовка (Kanban=False) AND (Res_Mat=0) AND (Res_Proch=0) AND (Res_Stand=0) AND (Res_Sbor=0) AND (Res_Kompl_priv=0) AND (Res_Klap=0)
                      if (Kanban=False) AND (Res_Detal=0) Then
                      //For y:=1 To SD1.RowCount Do Заготовка
                      Begin
                      Kol_Ed:= ADOQuery2.FieldByName('Количество').AsFloat;
                      Kol_Oboz:= ADOQuery2.FieldByName('Количество').AsFloat;
                        For O:=1 To SD2.RowCount Do
                      Begin
                           Res_Elem:=AnsiCompareStr(Elem,SD2.Cells[3,o]);
                           Res_Oboz:=AnsiCompareStr(Oboz,SD2.Cells[6,o]);
                           if (Res_Elem=0) AND (Res_Oboz=0) Then
                           Begin
                               Kol_Ed:=Kol_Ed+StrToFloat(SD2.Cells[4,o]);
                               SD2.Cells[4,o]:=FloatToStr(Kol_Ed);

                               Kol_Oboz:=(Kol_Oboz*Kol_Zap)+StrToFloat(SD2.Cells[5,o]);
                               SD2.Cells[5,o]:=FloatToStr(Kol_Oboz);
                               Flag:=1;
                               Break;
                           end ;
                     end;
                     If Flag=0 Then
                     Begin
                                SD2.Cells[2,p]:=Tip;
                                SD2.Cells[3,p]:=Elem;
                                SD2.Cells[4,p]:=FloatToStr(Kol_Ed);
                                SD2.Cells[5,p]:=FloatToStr(Kol_Ed*Kol_Zap);
                                SD2.Cells[6,p]:=Oboz;
                                SD2.Cells[0,p]:=IntToStr(p);
                                SD2.Cells[1,p]:=EI;
                                SD2.Cells[7,p]:=Dlina;
                                SD2.Cells[8,p]:=Shir;
                                SD2.Cells[9,p]:=DlinRaz;
                                SD2.Cells[10,p]:=ShirRaz;
                                SD2.Cells[11,p]:=Kol_Gib;
                                SD2.Cells[12,p]:=Mat;
                                SD2.Cells[13,p]:=VElem;
                                Inc(p);
                                SD2.RowCount:= SD2.RowCount+1;
                                //Break;
                     end;
                      end;
                      //+++++++++++++++++++++++++++++++++++++++++++++++++++++ (Kanban=False)  AND
                      if (Res_Sbor<>0)   Then       //Сборка
                      Begin
                      Kol_Ed:= ADOQuery2.FieldByName('Количество').AsFloat;
                      Kol_Oboz:= ADOQuery2.FieldByName('Количество').AsFloat;

                        For O:=1 To SD.RowCount Do
                      Begin
                           Res_Elem:=AnsiCompareStr(Elem,SD.Cells[3,o]);
                           Res_Oboz:=AnsiCompareStr(Oboz,SD.Cells[6,o]);
                           if (Res_Elem=0) AND (Res_Oboz=0) Then
                           Begin
                               Kol_Ed:=Kol_Ed+StrToFloat(SD.Cells[4,o]);
                               SD.Cells[4,o]:=FloatToStr(Kol_Ed);

                               Kol_Oboz:=(Kol_Oboz*Kol_Zap)+StrToFloat(SD.Cells[5,o]);
                               SD.Cells[5,o]:=FloatToStr(Kol_Oboz);
                               Flag:=1;
                               Break;
                           end ;
                     end;
                     If Flag=0 Then
                     Begin
                                SD.Cells[2,aa]:=Tip;
                                SD.Cells[3,aa]:=Elem;
                                SD.Cells[4,aa]:=FloatToStr(Kol_Ed);
                                SD.Cells[5,aa]:=FloatToStr(Kol_Ed*Kol_Zap);
                                SD.Cells[6,aa]:=Oboz;
                                SD.Cells[0,aa]:=IntToStr(aa);
                                SD.Cells[1,aa]:=EI;
                                SD.Cells[7,aa]:=Dlina;
                                SD.Cells[8,aa]:=Shir;
                                SD.Cells[9,aa]:=DlinRaz;
                                SD.Cells[10,aa]:=ShirRaz;
                                SD.Cells[11,aa]:=Kol_Gib;
                                SD.Cells[12,aa]:=Mat;
                                SD.Cells[13,aa]:=VElem;
                                Inc(aa);
                                SD.RowCount:= SD.RowCount+1;
                                //Break;
                     end;
                      end;
//------------------------------------------------------------------------
                      if (Kanban=True) AND ((Sten_g<>0) OR (Sten_V<>0)) AND (Res_KPD=0) Then       //Стенка   AND ((Sten_g=0) OR (Sten_V=0))
                      Begin
                      Kol_Ed:= ADOQuery2.FieldByName('Количество').AsFloat;
                      Kol_Oboz:= ADOQuery2.FieldByName('Количество').AsFloat;

                        For O:=1 To Stenki.RowCount Do
                      Begin
                           Res_Elem:=AnsiCompareStr(Elem,Stenki.Cells[3,o]);
                           Res_Oboz:=AnsiCompareStr(Oboz,Stenki.Cells[6,o]);
                           if (Res_Elem=0) AND (Res_Oboz=0) Then
                           Begin
                               Kol_Ed:=Kol_Ed+StrToFloat(Stenki.Cells[4,o]);
                               Stenki.Cells[4,o]:=FloatToStr(Kol_Ed);

                               Kol_Oboz:=(Kol_Oboz*Kol_Zap)+StrToFloat(Stenki.Cells[5,o]);
                               Stenki.Cells[5,o]:=FloatToStr(Kol_Oboz);
                               Flag:=1;
                               Break;
                           end ;
                     end;
                     If Flag=0 Then
                     Begin
                                Stenki.Cells[2,pp]:=Tip;
                                Stenki.Cells[3,pp]:=Elem;
                                Stenki.Cells[4,pp]:=FloatToStr(Kol_Ed);
                                Stenki.Cells[5,pp]:=FloatToStr(Kol_Ed*Kol_Zap);
                                Stenki.Cells[6,pp]:=Oboz;
                                Stenki.Cells[0,pp]:=IntToStr(pp);
                                Stenki.Cells[1,pp]:=EI;
                                if Sten_G<>0 Then
                                Stenki.Cells[7,pp]:=IntToStr(A);
                                if Sten_V<>0 Then
                                Stenki.Cells[7,pp]:=IntToStr(B);
                                Stenki.Cells[8,pp]:=Shir;
                                Stenki.Cells[9,pp]:=DlinRaz;
                                Stenki.Cells[10,pp]:=ShirRaz;
                                Stenki.Cells[11,pp]:=Kol_Gib;
                                Stenki.Cells[12,pp]:=Mat;
                                Stenki.Cells[13,pp]:=VElem;
                                Inc(pp);
                                Stenki.RowCount:= Stenki.RowCount+1;
                                //Break;
                     end;
                      end;
//------------------------------------------------------------------------
                      if (Kanban=True) AND ( (Sten_G<>0) OR (Sten_V<>0)) AND (Res_KPD<>0) Then       //Стенка   AND ((Sten_g=0) OR (Sten_V=0))
                      Begin
                      Kol_Ed:= ADOQuery2.FieldByName('Количество').AsFloat;
                      Kol_Oboz:= ADOQuery2.FieldByName('Количество').AsFloat;

                        For O:=1 To Stenki.RowCount Do
                      Begin
                           Res_Elem:=AnsiCompareStr(Elem,Stenki.Cells[3,o]);
                           Res_Oboz:=AnsiCompareStr(Oboz,Stenki.Cells[6,o]);
                           if (Res_Elem=0) AND (Res_Oboz=0) Then
                           Begin
                               Kol_Ed:=Kol_Ed+StrToFloat(Stenki.Cells[4,o]);
                               Stenki.Cells[4,o]:=FloatToStr(Kol_Ed);

                               Kol_Oboz:=(Kol_Oboz*Kol_Zap)+StrToFloat(Stenki.Cells[5,o]);
                               Stenki.Cells[5,o]:=FloatToStr(Kol_Oboz);
                               Flag:=1;
                               Break;
                           end ;
                     end;
                     If Flag=0 Then
                     Begin
                                Stenki.Cells[2,pp]:=Tip;
                                Stenki.Cells[3,pp]:=Elem;
                                Stenki.Cells[4,pp]:=FloatToStr(Kol_Ed);
                                Stenki.Cells[5,pp]:=FloatToStr(Kol_Ed*Kol_Zap);
                                Stenki.Cells[6,pp]:=Oboz;
                                Stenki.Cells[0,pp]:=IntToStr(pp);
                                Stenki.Cells[1,pp]:=EI;
                                if Sten_G<>0 Then
                                Stenki.Cells[7,pp]:=IntToStr(A);
                                if Sten_V<>0 Then
                                Stenki.Cells[7,pp]:=IntToStr(B);
                                Stenki.Cells[8,pp]:=Shir;
                                Stenki.Cells[9,pp]:=DlinRaz;
                                Stenki.Cells[10,pp]:=ShirRaz;
                                Stenki.Cells[11,pp]:=Kol_Gib;
                                Stenki.Cells[12,pp]:=Mat;
                                Stenki.Cells[13,pp]:=VElem;
                                Inc(pp);
                                Stenki.RowCount:= Stenki.RowCount+1;
                                //Break;
                     end;
                      end;
                      //+++++++++++++++++++++++++++++++++++++++++++++++++++++(Kanban=False) AND
                      if  ((Res_Stand<>0) OR (Res_Mat<>0) OR (Res_Proch<>0) OR  (Res_Kompl_priv<>0)) Then       //СCklad
                      Begin
                      Kol_Ed:= ADOQuery2.FieldByName('Количество').AsFloat;
                      Kol_Oboz:= ADOQuery2.FieldByName('Количество').AsFloat;

                        For O:=1 To SD9.RowCount Do
                      Begin
                           Res_Elem:=AnsiCompareStr(Elem,SD9.Cells[3,o]);
                           Res_Oboz:=AnsiCompareStr(Oboz,SD9.Cells[6,o]);
                           if (Res_Elem=0) AND (Res_Oboz=0) Then
                           Begin
                               Kol_Ed:=Kol_Ed+StrToFloat(SD9.Cells[4,o]);
                               SD9.Cells[4,o]:=FloatToStr(Kol_Ed);

                               Kol_Oboz:=(Kol_Oboz*Kol_Zap)+StrToFloat(SD9.Cells[5,o]);
                               SD9.Cells[5,o]:=FloatToStr(Kol_Oboz);
                               Flag:=1;
                               Break;
                           end ;
                     end;
                     If Flag=0 Then
                     Begin
                                SD9.Cells[2,qq]:=Tip;
                                SD9.Cells[3,qq]:=Elem;
                                SD9.Cells[4,qq]:=FloatToStr(Kol_Ed);
                                SD9.Cells[5,qq]:=FloatToStr(Kol_Ed*Kol_Zap);
                                SD9.Cells[6,qq]:=Oboz;
                                SD9.Cells[0,qq]:=IntToStr(qq);
                                SD9.Cells[1,qq]:=EI;
                                SD9.Cells[7,qq]:=Dlina;
                                SD9.Cells[8,qq]:=Shir;
                                SD9.Cells[9,qq]:=DlinRaz;
                                SD9.Cells[10,qq]:=ShirRaz;
                                SD9.Cells[11,qq]:=Kol_Gib;
                                SD9.Cells[12,qq]:=Mat;
                                SD9.Cells[13,qq]:=VElem;
                                Inc(qq);
                                SD9.RowCount:= SD9.RowCount+1;
                                //Break;
                     end;
                      end;
                     Res_N2:=AnsiCompareStr('2Н',N2);
                     Res_N1:=AnsiCompareStr('1Н',N2);
                      //+++++++++++++++++++++++++++++++++++++++++++++++++++++
                      C:=0;
                      if (Kanban=False) AND (Trumph=True) Then
                      Begin
                      Kol_Ed:= ADOQuery2.FieldByName('КолНаЕд').AsFloat;
                      Kol_Oboz:= ADOQuery2.FieldByName('Количество').AsFloat;
                      For O:=1 To SD3.RowCount Do
                        Begin
                           Res_Elem:=AnsiCompareStr(Elem,SD3.Cells[3,o]);
                           Res_Oboz:=AnsiCompareStr(Oboz,SD3.Cells[6,o]);
                           if (Res_Elem=0) AND (Res_Oboz=0) Then
                           Begin
                               Kol_Oboz:=(Kol_Oboz*Kol_Zap)+StrToFloat(SD3.Cells[4,o]);
                               SD3.Cells[4,o]:=FloatToStr(Kol_Oboz);
                               //+++++++++++++++++
                              { NC_OB:= StrToFloat(SD3.Cells[16,o])*Kol_Oboz;
                               SD3.Cells[15,o]:=FloatToStr(NC_OB);  }
                               Flag:=1;
                               Break;
                           end ;
                           //end;
                        end;

                     If Flag=0 Then
                     Begin
                        if Res_KPD<>0 Then
                        Begin
                                SD3.Cells[2,s]:=Tip;
                                SD3.Cells[3,s]:=Elem;
                                SD3.Cells[4,s]:=FloatToStr(Kol_Oboz*Kol_Zap);
                                //SD3.Cells[5,s]:=FloatToStr(Kol_Oboz);
                                SD3.Cells[6,s]:=Oboz;
                                SD3.Cells[0,s]:=IntToStr(s);
                                SD3.Cells[1,s]:=EI;
                                SD3.Cells[7,s]:=Dlina;
                                SD3.Cells[8,s]:=Shir;
                                SD3.Cells[9,s]:=DlinRaz;
                                SD3.Cells[10,s]:=ShirRaz;
                                SD3.Cells[11,s]:=Kol_Gib;
                                SD3.Cells[12,s]:=Mat;
                                if DlinRaz='' Then
                                DlinRaz:='0';
                                {if (Res_KPU<>0) AND (A>100) AND (B>100) AND (F2<>0) Then
                                C := StrToFloat(DlinRaz)-(((A - 24) /(Kol_Ed/2) - 28)/2) - 34.14;
                                if (Res_KPU<>0) AND ((A=100) OR (B=100))  Then
                                C := StrToFloat(DlinRaz)-((A /(Kol_Ed/2)- 14)/2-13) - 24.3;
                                if (Res_KPU<>0) AND (F1<>0) AND (Res_N2=0)  Then 
                                C := StrToFloat(DlinRaz)-(StrToFloat(DlinRaz)/2-34.3) - 24.3;}
                                C := StrToFloat(SH1);
                                SD3.Cells[14,s]:=FloatToStr(C);
                                SD3.Cells[13,s]:=VElem;
                                Inc(s);
                                SD3.RowCount:= SD3.RowCount+1;
                                //Break;
                                //end;

                        end;
                        if Res_KPU<>0 Then
                        Begin
                                SD3.Cells[2,s]:=Tip;
                                SD3.Cells[3,s]:=Elem;
                                SD3.Cells[4,s]:=FloatToStr(Kol_Oboz*Kol_Zap);
                                //SD3.Cells[5,s]:=FloatToStr(Kol_Oboz);
                                SD3.Cells[6,s]:=Oboz;
                                SD3.Cells[0,s]:=IntToStr(s);
                                SD3.Cells[1,s]:=EI;
                                SD3.Cells[7,s]:=Dlina;
                                SD3.Cells[8,s]:=Shir;
                                SD3.Cells[9,s]:=DlinRaz;
                                SD3.Cells[10,s]:=ShirRaz;
                                SD3.Cells[11,s]:=Kol_Gib;
                                SD3.Cells[12,s]:=Mat;
                                if DlinRaz='' Then
                                        DlinRaz:='0';
                               { if (Res_KPU<>0) AND (A>100) AND (B>100) AND (F2<>0) Then
                                        C := StrToFloat(DlinRaz)-(((A - 24) /(Kol_Ed/2) - 28)/2) - 34.14;
                                if (Res_KPU<>0) AND ((A=100) OR (B=100))  Then
                                        C := StrToFloat(DlinRaz)-((A /(Kol_Ed/2)- 14)/2-13) - 24.3;
                                if (Res_KPU<>0) AND (F1<>0) AND (Res_N2=0)  Then
                                        C := StrToFloat(DlinRaz)-(StrToFloat(DlinRaz)/2-34.3) - 24.3;}
                                if SH1='' Then
                                C:=0
                                Else
                                C := StrToFloat(SH1);
                                SD3.Cells[14,s]:=FloatToStr(C);
                                SD3.Cells[13,s]:=VElem;
                                Inc(s);
                                SD3.RowCount:= SD3.RowCount+1;
                        end;
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                        if (Res_KPU=0) AND (Res_KPD=0) Then  //Гермик
                        Begin
                                SD3.Cells[2,s]:=Tip;
                                SD3.Cells[3,s]:=Elem;
                                SD3.Cells[4,s]:=FloatToStr(Kol_Oboz*Kol_Zap);
                                //SD3.Cells[5,s]:=FloatToStr(Kol_Oboz);
                                SD3.Cells[6,s]:=Oboz;
                                SD3.Cells[0,s]:=IntToStr(s);
                                SD3.Cells[1,s]:=EI;
                                SD3.Cells[7,s]:=Dlina;
                                SD3.Cells[8,s]:=Shir;
                                SD3.Cells[9,s]:=DlinRaz;
                                SD3.Cells[10,s]:=ShirRaz;
                                SD3.Cells[11,s]:=Kol_Gib;
                                SD3.Cells[12,s]:=Mat;
                                if DlinRaz='' Then
                                DlinRaz:='0';
                                {if (Res_KPU<>0) AND (A>100) AND (B>100) AND (F2<>0) Then
                                        C := StrToFloat(DlinRaz)-(((A - 24) /(Kol_Ed/2) - 28)/2) - 34.14;
                                if (Res_KPU<>0) AND ((A=100) OR (B=100))  Then
                                        C := StrToFloat(DlinRaz)-((A /(Kol_Ed/2)- 14)/2-13) - 24.3;
                                if (Res_KPU<>0) AND (F1<>0)  Then
                                        C := StrToFloat(DlinRaz)-((A /(Kol_Ed/2)- 14)/2-13) - 24.3;}
                                if SH1='' Then
                                C:=0
                                Else
                                C := StrToFloat(SH1);
                                SD3.Cells[14,s]:=FloatToStr(C);
                                SD3.Cells[13,s]:=VElem;
                                Inc(s);
                                SD3.RowCount:= SD3.RowCount+1;

                        end;
                        End;
                     end;
                      //+++++++++++++++++++++++++++++++++++++++++++++++++++++ Ножницы
                      if (Kanban=False) AND (Nog=True) Then
                      Begin
                      Kol_Ed:= ADOQuery2.FieldByName('Количество').AsFloat;
                      Kol_Oboz:= ADOQuery2.FieldByName('Количество').AsFloat;
                       { For O:=1 To SD4.RowCount Do
                        Begin
                           Res_Elem:=AnsiCompareStr(Elem,SD4.Cells[3,o]);
                           Res_Oboz:=AnsiCompareStr(Oboz,SD4.Cells[6,o]);
                           if (Res_Elem=0) AND (Res_Oboz=0) Then
                           Begin
                               Kol_Oboz:=(Kol_Oboz*Kol_Zap)+StrToFloat(SD4.Cells[4,o]);
                               SD4.Cells[4,o]:=FloatToStr(Kol_Oboz);
                               //+++++++++++++++++
                               NC_OB:= StrToFloat(SD4.Cells[16,o])*Kol_Oboz;
                               SD4.Cells[15,o]:=FloatToStr(NC_OB);
                               Flag:=1;
                               Break;
                           end ;
                           //end;
                        end;

                     If Flag=0 Then
                     Begin }

                                SD4.Cells[2,h]:=Tip;
                                SD4.Cells[3,h]:=Elem;
                                SD4.Cells[4,h]:=FloatToStr(Kol_Oboz*Kol_Zap);
                                SD4.Cells[6,h]:=Oboz;
                                SD4.Cells[0,h]:=IntToStr(h);
                                SD4.Cells[1,h]:=EI;
                                SD4.Cells[7,h]:=Shir;
                                SD4.Cells[8,h]:=Dlina;
                                SD4.Cells[9,h]:=ShirRaz;
                                SD4.Cells[10,h]:=DlinRaz;
                                if Res_KPD<>0 Then
                                Begin
                                        SD4.Cells[7,h]:=Dlina;
                                        SD4.Cells[8,h]:=Shir;
                                        SD4.Cells[9,h]:=DlinRaz;
                                        SD4.Cells[10,h]:=ShirRaz;
                                end;

                                SD4.Cells[11,h]:=Kol_Gib;
                                SD4.Cells[12,h]:=Mat;
                                if DlinRaz='' Then
                                DlinRaz:='0';
                                //+ ++++++++++++++++++++++++++++++++++++++++++
                                //Лист ДПРНМ 1,0х600х1500 Л63 ГОСТ 931-90
                                 Res_List:=Pos('Лист ДПРНМ 1,0х600х1500 Л63 ГОСТ 931-90',Mat);
                                 if Res_List<>0 Then
                                 Mat:='1';
                                 //Лента на шестьдесят НЕРЖ 0,3
                                 Res_List:=Pos('НЕРЖ',Mat);
                                 if Res_List<>0 Then
                                 Delete(Mat,1,Res_List+4);
                                Res_List:=Pos('Л',Mat);
                                        if Res_List<>0 Then
                                        Begin
                                                Res_Tol:=Pos(' ',Mat);// Лента
                                                Delete(Mat,1,Res_Tol);
                                                //
                                                Res_Tol:=Pos(' ',Mat);//ОЦ
                                                Delete(Mat,1,Res_Tol);

                                                Res_Tol:=Pos(' ',Mat);//08pc
                                                Delete(Mat,1,Res_Tol);

                                                Res_Tol:=Pos('х',Mat);//x
                                                Delete(Mat,Res_Tol,50);

                                                Res_Tol:=Pos('*',Mat);//*
                                                Delete(Mat,Res_Tol,50);

                                                Res_Tol:=Pos(' ',Mat);//08pc
                                                Delete(Mat,1,Res_Tol);
                                        end
                                        Else
                                        Begin
                                                Res_Tol:=Pos(' ',Mat);
                                                Delete(Mat,1,Res_Tol);
                                        end;
                                        Res_List:=Pos('ОЦ',Mat);
                                        if Res_List<>0 Then
                                        Delete(Mat,1,2);
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                Tol:=StrToFloat(Mat);
                                Srav:=CompareValue(Tol,0.55);
                                if (Srav=0) OR (Srav<0) Then
                                Begin
                                        Tol_S:='0.55';
                                        Res_Tol55:=0.55;
                                End;
                                Srav:=CompareValue(Tol,0.7);
                                if ((Srav>0) OR (Srav=0)) AND (Tol<=1) Then
                                Begin
                                        Tol_S:='0.7';
                                        Res_Tol55:=7;
                                End;

                                Srav:=CompareValue(Tol,1.2);
                                //Srav1:= CompareValue(Tol,0.7);
                                if ((Srav=0) OR (Srav>0)) AND (Tol<=2) Then
                                Begin
                                        Tol_S:='1.2';
                                        Res_Tol55:=2;
                                End;
                                DlinI:=StrToFloat(ShirRAz);
                                ShirI:=StrToFloat(DlinRaz);

                                if DlinI<=500 Then
                                        DlinI:=500;
                                //---------------
                                if (DlinI>500) AND (DlinI<=1000) Then
                                        DlinI:=1000;
                                //---------------
                                if (DlinI>1000) AND (DlinI<=1500) Then
                                        DlinI:=1500;
                                //---------------
                                if (DlinI>1500) AND (DlinI<=2000) Then
                                        DlinI:=2000;
                                //---------------
                                if (DlinI>2000) AND (DlinI<=2500) Then
                                        DlinI:=2500;
                                //+====================================
                                if Res_Tol55=0.55 Then
                                Begin
                                if ShirI<=50 Then
                                        ShirI:=50;
                                //---------------
                                if (ShirI>50) AND (ShirI<=100) Then
                                        ShirI:=100;
                                //---------------
                                if (ShirI>100) AND (ShirI<=200) Then
                                        ShirI:=200;
                                //---------------
                                if (ShirI>200) AND (ShirI<=500) Then
                                        ShirI:=500;
                                //---------------
                                if (ShirI>500) AND (ShirI<=600) Then
                                        ShirI:=600;
                                //---------------
                                if (ShirI>600) AND (ShirI<=1250) Then
                                        ShirI:=1250;
                                End;
                                //+====================================
                                if Res_Tol55=7 Then
                                Begin
                                if ShirI<=200 Then
                                        ShirI:=200;
                                //---------------
                                if (ShirI>200) AND (ShirI<=500) Then
                                        ShirI:=500;
                                //---------------
                                if (ShirI>500) Then
                                        ShirI:=1250;

                                End;
                                //+====================================
                                if Res_Tol55=2 Then
                                Begin
                                if ShirI<=200 Then
                                        ShirI:=200;
                                //---------------
                                if (ShirI>200) AND (ShirI<=500) Then
                                        ShirI:=500;
                                //---------------
                                if (ShirI>500) Then
                                        ShirI:=1250;

                                End;

                                if not Form1.mkQuerySelect1(Form1.ADOQuery1,
                                'Select * from %s Where   ([№]=' + #39 +'1'+ #39 +') And ([Толщина]='+Tol_S+') And ([Длина]=' + #39 +FloatToStr(DlinI)+ #39 +') And ([Высота]=' + #39 +FloatToStr(ShirI)+ #39 +')', ['[Резка Гибка]']) then
                                exit;
                                NC_S:= (Form1.ADOQuery1.FieldByName('Норма').AsString);
                                if NC_S='' Then
                                        NC_S:='0';
                                //SD4.Cells[13,h]:=NC_S;
                                NC:=StrToFloat(NC_S)*Kol_Oboz*Kol_Zap;

//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++=
                               { if (Res_KPU<>0) AND (A>100) AND (B>100) AND (F2<>0) Then
                                        C := StrToFloat(ShirRaz)-(((A - 24) /(Kol_Ed/2) - 28)/2) - 34.14;
                                if (Res_KPU<>0) AND (A>100) AND (B>100) AND (F2<>0)AND (Flag_Razb_Klapana =1) Then
                                        C := StrToFloat(ShirRaz)-(((A - 24) /(Kol_Ed/4) - 28)/2) - 34.14;
                                if (Res_KPU<>0) AND ((A=100) OR   (B=100))  Then
                                        C := StrToFloat(ShirRaz)-((A /(Kol_Ed/2) - 14)/2-13) - 24.3;
                                if (Res_KPU<>0) AND (F1<>0)  AND (Res_N2=0)  Then
                                        C := StrToFloat(ShirRaz)-(StrToFloat(ShirRaz) /(Kol_Ed/2)- 34.3) - 24.3;  }
                                if SH1='' Then
                                C:=0
                                Else
                                C := StrToFloat(SH1);
                                if Flag_Razb_Klapana =1 Then
                                        Kol_Lop:=Kol_Ed/4
                                else
                                        Kol_Lop:=Kol_Ed/2;
                                {L :=(A/Kol_LOp-14)/2 -13;
                                if (Res_KPU<>0) AND (F1<>0)  AND (Res_N1=0)  Then
                                        C := StrToFloat(ShirRaz)-(L) - 24.3; }
                                //+ ++++++++++++++++++++++++++++++++++++++++++
                                Res_Lop:=AnsiCompareStr('Корпус лопатки',SD4.Cells[3,h]);
                                 if (Res_KPU<>0) AND (Res_Lop=0) AND (C<220)  Then //4 Ugla
                                 Begin
                                       Ugol:=4;
                                 end;
                                 if (Res_KPU<>0) AND (Res_Lop=0) AND ((C>=220) AND (C<440))  Then //6 Uglov
                                 Begin
                                       Ugol:=6;
                                 end;
                                 if (Res_KPU<>0) AND (Res_Lop=0) AND (C>=400)   Then //8 Uglov
                                 Begin
                                       Ugol:=8;
                                 end;
//+++++++++++++++++++++++++++++++++++++++++++++
                                //if (Res_KPD<>0) AND (Res_Lop=0) Then
                                //Begin
                                 //+====================================
                                if Res_Tol55=0.55 Then
                                Begin
                                if ShirI<=600 Then
                                        ShirI:=600;
                                //---------------
                                if (ShirI>600)  Then
                                        ShirI:=1250;
                                End;
                                //+====================================
                                if Res_Tol55=7 Then
                                Begin
                                if ShirI<=500 Then
                                        ShirI:=500;
                                //---------------
                                if (ShirI>500)  Then
                                        ShirI:=1250;

                                End;
                                //+====================================
                                if Res_Tol55=2 Then
                                Begin
                                if ShirI<=500 Then
                                        ShirI:=500;
                                //---------------
                                if (ShirI>500)  Then
                                        ShirI:=500;
                                //---------------
                                if (ShirI>500) Then
                                        ShirI:=1250;

                                End;
                                if not Form1.mkQuerySelect1(Form1.ADOQuery1,
                                'Select * from %s Where   ([№]=' + #39 +'11'+ #39 +') And ([Толщина]='+Tol_S+') And ([Длина]=' + #39 +FloatToStr(DlinI)+ #39 +') And ([Высота]=' + #39 +FloatToStr(ShirI)+ #39 +')', ['[Резка Гибка]']) then
                                exit;
                                NC_S2:= (Form1.ADOQuery1.FieldByName('Норма').AsString);
                                if NC_S2='' Then
                                        NC_S2:='0';
                                if not Form1.mkQuerySelect1(Form1.ADOQuery1,
                                'Select * from %s Where   ([№]=' + #39 +'12'+ #39 +') ', ['[Резка Гибка]']) then
                                exit;
                                if C>200 Then
                                Razmet_KPU:=Form1.ADOQuery1.FieldByName('Норма').AsFloat
                                Else
                                Razmet_KPU:=0;
                                SD4.Cells[16,h]:=FloatToStr(StrToFloat(NC_S)+(StrToFloat(NC_S2)+Razmet_KPU)*Ugol);
                                NC:=NC+(((StrToFloat(NC_S2)*Ugol+Razmet_KPU))*Kol_Oboz*Kol_Zap);
                                //end;
                                SD4.Cells[13,h]:=FloatToStr(C);
                                SD4.Cells[14,h]:=VElem;
                                SD4.Cells[15,h]:=FloatToStr(NC);
                                if not Form1.mkQueryUpdate2(Form1.ADOQuery4,
                                        'UPDATE %s SET [Н\ч Ножницы]=' + #39
                                                + Form1.ConvertFloat(FloatToStr(NC)) + #39 +
                                        ' WHERE ([IdГП]=' + #39 + IDGP + #39 +
                                                ') AND (Элемент=' + #39 + Elem + #39 +') AND (Обозначение=' + #39 + Oboz + #39 +')  ', ['Специф']) then
                                        Exit;
                                Inc(h);
                                SD4.RowCount:= SD4.RowCount+1;
                                //Break;
                       // End;
                     //end;
                     end;
                      //+++++++++++++++++++++++++++++++++++++++++++++++++++++
                      NC_GER_S:='';
                      NC_GER:=0;   //(Kanban=False) AND
                      if ((Gib=True) OR (HACO=True)) Then
                      Begin
                      Kol_Ed:= ADOQuery2.FieldByName('КолНаЕд').AsFloat;
                      Kol_Oboz:= ADOQuery2.FieldByName('Количество').AsFloat;

                                SD5.Cells[2,n]:=Tip;
                                SD5.Cells[3,n]:=Elem;
                                SD5.Cells[4,n]:=FloatToStr(Kol_Oboz*Kol_Zap);
                                //SD5.Cells[5,n]:=FloatToStr(Kol_Oboz);
                                SD5.Cells[6,n]:=Oboz;
                                SD5.Cells[0,n]:=IntToStr(n);
                                SD5.Cells[1,n]:=EI;
                                SD5.Cells[7,n]:=Dlina;
                                SD5.Cells[8,n]:=Shir;
                                SD5.Cells[9,n]:=DlinRaz;
                                SD5.Cells[10,n]:=ShirRaz;
                                SD5.Cells[11,n]:=IntToStr(KolGib);
                                SD5.Cells[12,n]:=Mat;
                                SD5.Cells[13,n]:=VElem;
                                Res_Lop:=AnsiCompareStr('Корпус лопатки',Elem);
                                if Res_Lop=0 Then
                                Begin
                                        SD5.Cells[14,n]:=FloatToStr(0.034*Kol_Oboz*Kol_Zap);
                                        SD5.Cells[19,n]:=FloatToStr(0.034);
                                End
                                Else
                                Begin
                                        ShirI:=StrToFloat(ShirRaz);
                                        DlinI:=StrToFloat(DlinRaz);
                                        DlinII:=StrToFloat(DlinRaz);
                                        //Лента Оц 1,5х264
                                        //Лист ОЦ 08пс 1,5х1250х2500
                                        //ОЦ 1,5
                                        //Лента на шестьдесят НЕРЖ 0,3
                                        //Лист ДПРНМ 1,0х600х1500 Л63 ГОСТ 931-90
                                 Res_List:=Pos('Лист ДПРНМ 1,0х600х1500 Л63 ГОСТ 931-90',Mat);
                                 if Res_List<>0 Then
                                 Mat:='1';
                                 //Лист ДПРНМ 1,0х600х1500 Л63 ГОСТ 931-90
                                 Res_List:=Pos('НЕРЖ',Mat);
                                 if Res_List<>0 Then
                                 Delete(Mat,1,Res_List+4);
                                        Res_List:=Pos('Л',Mat);
                                        if Res_List<>0 Then
                                        Begin
                                                Res_Tol:=Pos(' ',Mat);// Лента
                                                Delete(Mat,1,Res_Tol);
                                                //
                                                Res_Tol:=Pos(' ',Mat);//ОЦ
                                                Delete(Mat,1,Res_Tol);

                                                Res_Tol:=Pos(' ',Mat);//08pc
                                                Delete(Mat,1,Res_Tol);

                                                Res_Tol:=Pos('х',Mat);//x
                                                Delete(Mat,Res_Tol,50);

                                                Res_Tol:=Pos('*',Mat);//*
                                                Delete(Mat,Res_Tol,50);

                                                Res_Tol:=Pos(' ',Mat);//08pc
                                                Delete(Mat,1,Res_Tol);
                                        end
                                        Else
                                        Begin
                                                Res_Tol:=Pos(' ',Mat);
                                                Delete(Mat,1,Res_Tol);
                                        end;
                                        Res_List:=Pos('ОЦ',Mat);
                                        if Res_List<>0 Then
                                        Delete(Mat,1,2);
                                        
                                        if Mat='' Then
                                        Mat:='0';
                                        Tol:=StrToFloat(Mat);
                                        if Tol<=1 Then
                                        Tol_S:='0.55';
                                        if (Tol>1) AND (Tol<=2) Then
                                        Tol_S:='2';
                                        if Tol>2 Then
                                        Tol_S:='3';
                                        if DlinI<=2000 Then
                                        DlinI:=2000
                                        Else
                                        DlinI:=3000;
                                        //
                                        if DlinII<=1000 Then
                                        DlinII:=1000
                                        Else
                                        DlinII:=1100;
                                if  (HACO=True) Then
                                Begin
                                 if not Form1.mkQuerySelect1(Form1.ADOQuery1,
                                'Select * from %s Where   ([№]=' + #39 +'66'+ #39 +') ' +
                                ' And ([Длина]=' + #39 +FloatToStr(DlinII)+ #39 +') ', ['[Резка Гибка]']) then
                                exit;
                                NC_GER_S:= (Form1.ADOQuery1.FieldByName('Норма').AsString);
                                end;
                                if  (GIB=True) Then
                                Begin
                                 if not Form1.mkQuerySelect1(Form1.ADOQuery1,
                                'Select * from %s Where   ([№]=' + #39 +'3'+ #39 +') And ([Толщина]=' + #39 +(Tol_s)+ #39 +
                                ') And ([Длина]=' + #39 +FloatToStr(DlinI)+ #39 +') And ([Высота]=' + #39 +IntToStr(KolGib)+ #39 +')', ['[Резка Гибка]']) then
                                exit;
                                NC_S:= (Form1.ADOQuery1.FieldByName('Норма').AsString);
                                End;

                                If NC_S='' Then
                                NC:=0 Else
                                NC:=StrToFloat(NC_S);

                                If NC_GER_S='' Then
                                NC_GER:=0 Else
                                NC_GER:=StrToFloat(NC_GER_S);

                                SD5.Cells[14,n]:=FloatToStr((NC*Kol_Oboz)*Kol_Zap);
                                SD5.Cells[19,n]:=NC_S;
                                SD5.Cells[21,n]:=FloatToStr((NC_GER*(Kol_Oboz/2))*Kol_Zap);
                                end;
                                //======================
                                Res_Resh:=Pos('Решетка жалюзийная',Elem);
                                if Res_Resh<>0 Then
                                Begin
                                       if not Form1.mkQuerySelect1(Form1.ADOQuery1,
                                        'Select * from %s Where   ([№]=' + #39 +'9'+ #39 +') And ([Толщина]=' + #39 +IntToStr(KolGib)+ #39 +
                                        ')', ['[Резка Гибка]']) then
                                        exit;
                                        NC_S:= (Form1.ADOQuery1.FieldByName('Высота').AsString);
                                        If NC_S='' Then
                                                NC:=0 Else
                                                NC:=StrToFloat(NC_S);
                                        //++++++++++++
                                        NC_S2:= (Form1.ADOQuery1.FieldByName('Длина').AsString);
                                        If NC_S2='' Then
                                                NC2:=0 Else
                                                NC2:=StrToFloat(NC_S2);
                                        SD5.Cells[19,n]:=FloatToStr((NC)+(NC2));
                                        SD5.Cells[14,n]:=FloatToStr(((NC*Kol_Zap)+(NC2/Kol_Zap))*Kol_Oboz);
                                end;
                                Res_Lop:=AnsiCompareStr('Стенка лопатки',Elem);
                                if Res_Lop=0 Then
                                Begin
                                        Kol_lop:= ADOQuery2.FieldByName('Количество').AsFloat;
                                End;
                                Lenta:=Pos('Лента уплотнительная',Elem);
                                //Лента уплотнительная
                                if Lenta <>0 Then
                                Begin
                                      DlinD:=B-1.5;
                                      if B<335 Then
                                        DlinI:=170;
                                      if (B>=335) AND (B<500) Then
                                      Begin
                                        DlinI:=335;
                                      end;

                                      if (B>=500) AND (B<665) Then
                                      Begin
                                        DlinI:=500;
                                      end;

                                      if (B>=665) AND (B<830) Then
                                      Begin
                                        DlinI:=665;
                                      end;

                                      if (B>=830) AND (B<995) Then
                                      Begin
                                        DlinI:=830;
                                      end;

                                      if (B>=995) AND (B<1160) Then
                                      Begin
                                        DlinI:=995;
                                      end;

                                      if (B>=1160) AND (B<1325) Then
                                      Begin
                                        DlinI:=1160;
                                      end;

                                      if (B>=1325) AND (B<1490) Then
                                      Begin
                                        DlinI:=1325;
                                      end;

                                      if (B>=1490) AND (B<1655) Then
                                      Begin
                                        DlinI:=1490;
                                      end;

                                      if (B>=1655) AND (B<1820) Then
                                      Begin
                                        DlinI:=1655;
                                      end;

                                      if (B>=1820) AND (B<1985) Then
                                      Begin
                                        DlinI:=1820;
                                      end;

                                      if (B>=1985) AND (B<2150) Then
                                      Begin
                                        DlinI:=1985;
                                      end;

                                      if (B>=2150) AND (B<2315) Then
                                      Begin
                                        DlinI:=2150;
                                      end;

                                      if (B>=2315) AND (B<2440) Then
                                      Begin
                                        DlinI:=2315;
                                      end;


                                      if not Form1.mkQuerySelect1(Form1.ADOQuery1,
                                        'Select * from %s Where   ([№]=' + #39 +'12'+ #39 +') And ([Высота]=' + #39 +FloatToStr(DlinI)+ #39 +
                                        ')', ['[Резка Гибка]']) then
                                        exit;
                                      //Kol_otv:= KOl_Lop;//(Form1.ADOQuery1.FieldByName('Норма').AsInteger);
                                      A_Lenta:=(B-1.5-(KOl_Lop/2-1)*150)/2;
                                      SD5.Cells[15,n]:=FloatToStr(DlinD);//В Клапана-1,5
                                      SD5.Cells[16,n]:= FloatToStr(A_Lenta);
                                      SD5.Cells[17,n]:=FloatToStr(KOl_Lop/2);
                                      SD5.Cells[18,n]:='150';
                                      T:=0.008*(KOl_Lop/2)*Kol_Oboz*Kol_Zap+0.008*Kol_Ed*Kol_Zap; //Пробивка отверстий в одной ленте
                                      SD5.Cells[20,n]:= FloatToStr(T);
                                end;
                                if SD5.Cells[20,n]='' Then
                                NCB:=(StrToFloat(SD5.Cells[14,n])+0)
                                Else NCB:=(StrToFloat(SD5.Cells[14,n])+StrToFloat(SD5.Cells[20,n]));

                                if not Form1.mkQueryUpdate2(Form1.ADOQuery4,
                                        'UPDATE %s SET [Н\ч Гибка]=' + #39
                                                + Form1.ConvertFloat(FloatToStr(NCB)) + #39 +
                                        ' WHERE ([IdГП]=' + #39 + IDGP + #39 +
                                                ') AND (Элемент=' + #39 + Elem + #39 +') AND (Обозначение=' + #39 + Oboz + #39 +')  ', ['Специф']) then
                                        Exit;
                                Inc(n);
                                SD5.RowCount:= SD5.RowCount+1;
                                //Break;
                     //end;
                     end;
                     //+++++++++++++++++++++++++++++++++++++++++++++++++++++
                      if (Kanban=False) AND (Obraziv=True) Then
                      Begin
                      Kol_Ed:= ADOQuery2.FieldByName('Количество').AsFloat;
                      Kol_Oboz:= ADOQuery2.FieldByName('Количество').AsFloat;
                                SD7.Cells[2,w]:=Tip;
                                SD7.Cells[3,w]:=Elem;
                                SD7.Cells[4,w]:=FloatToStr(Kol_Ed*Kol_Zap);
                               // SD7.Cells[5,w]:=FloatToStr(Kol_Oboz);
                                SD7.Cells[6,w]:=Oboz;
                                SD7.Cells[0,w]:=IntToStr(w);
                                SD7.Cells[1,w]:=EI;
                                SD7.Cells[7,w]:=Shir;
                                SD7.Cells[8,w]:=Dlina;
                                SD7.Cells[9,w]:=ShirRaz;
                                SD7.Cells[10,w]:=DlinRaz ;
                                //SD7.Cells[11,w]:=Kol_Gib;
                                SD7.Cells[12,w]:=Mat;
                                SD7.Cells[13,w]:=VElem;
                                if (DlinI>750) Then
                                          DlinI:=1000;
                                if not Form1.mkQuerySelect1(Form1.ADOQuery1,
                                'Select * from %s Where   ([№]=' + #39 +'8'+ #39 +') And  ([Длина]=' + #39 +FloatToStr(DlinI)+ #39 +')', ['[Резка Гибка]']) then
                                          exit;
                                NC_S:= (Form1.ADOQuery1.FieldByName('Норма').AsString);
                                If NC_S='' Then
                                NC:=0 Else
                                NC:=StrToFloat(NC_S);
                                SD7.Cells[14,w]:=FloatToStr(NC*Kol_Ed*Kol_Zap);
                                if not Form1.mkQueryUpdate2(Form1.ADOQuery4,
                                'UPDATE %s SET [Н\ч Пила]=' + #39
                                + Form1.ConvertFloat(FloatToStr(NC*Kol_Ed*Kol_Zap)) + #39 +
                                ' WHERE ([IdГП]=' + #39 + IDGP + #39 +
                                ') AND (Элемент=' + #39 + Elem + #39 +') AND (Обозначение=' + #39 + Oboz + #39 +')  ', ['Специф']) then
                                        Exit;
                                Inc(w);
                                SD7.RowCount:= SD7.RowCount+1;
                      End;
                      //+++++++++++++++++++++++++++++++++++++++++++++++++++++
                      if (Kanban=False) AND (Pila=True) Then
                      Begin
                      Kol_Ed:= ADOQuery2.FieldByName('Количество').AsFloat;
                      Kol_Oboz:= ADOQuery2.FieldByName('Количество').AsFloat;
                                SD6.Cells[2,w]:=Tip;
                                SD6.Cells[3,w]:=Elem;
                                SD6.Cells[4,w]:=FloatToStr(Kol_Ed*Kol_Zap);
                               // SD7.Cells[5,w]:=FloatToStr(Kol_Oboz);
                                SD6.Cells[6,w]:=Oboz;
                                SD6.Cells[0,w]:=IntToStr(w);
                                SD6.Cells[1,w]:=EI;
                                SD6.Cells[7,w]:=Shir;
                                SD6.Cells[8,w]:=Dlina;
                                SD6.Cells[9,w]:=ShirRaz;
                                SD6.Cells[10,w]:=DlinRaz ;
                                //SD6.Cells[11,w]:=Kol_Gib;
                                SD6.Cells[12,w]:=Mat;
                                SD6.Cells[13,w]:=VElem;

                                 Res_List:=Pos('Лист ДПРНМ 1,0х600х1500 Л63 ГОСТ 931-90',Mat);
                                 if Res_List<>0 Then
                                 Mat:='1';
                                 Res_List:=Pos('НЕРЖ',Mat);
                                 if Res_List<>0 Then
                                 Delete(Mat,1,Res_List+4);
                                        Res_List:=Pos('Л',Mat);
                                        if Res_List<>0 Then
                                        Begin
                                                Res_Tol:=Pos(' ',Mat);// Лента
                                                Delete(Mat,1,Res_Tol);

                                                Res_Tol:=Pos(' ',Mat);//ОЦ
                                                Delete(Mat,1,Res_Tol);

                                                Res_Tol:=Pos(' ',Mat);//08pc
                                                Delete(Mat,1,Res_Tol);

                                                Res_Tol:=Pos('х',Mat);//x
                                                Delete(Mat,Res_Tol,50);

                                                Res_Tol:=Pos('*',Mat);//*
                                                Delete(Mat,Res_Tol,50);

                                                Res_Tol:=Pos(' ',Mat);//08pc
                                                Delete(Mat,1,Res_Tol);
                                        end
                                        Else
                                        Begin
                                                Res_Tol:=Pos(' ',Mat);
                                                Delete(Mat,1,Res_Tol);
                                        end;
                                        Res_List:=Pos('ОЦ',Mat);
                                        if Res_List<>0 Then
                                        Delete(Mat,1,2);
                                        if Mat='' Then
                                        Mat:='0';
                                        //++++++++++++++++++++++++++++++++==
                                        Res_Detal:=AnsiCompareStr('Детали',VElem);
                                        Res_Lop:=Pos('Лопатка',Elem);
                                        Res_Ger:=Pos('ГЕРМИК',Izdel);
                                        R25:=Pos('р25',Izdel);
                                        Res_List:=Pos('20001',Mat);//Лопатка Р25
                                        if Res_Ger=0 Then
                                          SD6.Cells[9,w]:=ShirRaz;                         //  AND (Res_List=0)

                                        Ugol:=Pos('Стенка',Elem);
                                        //++++++++++++++++++++++++++++++++++++++++++++++++
                                        if (Res_Detal=0) AND (Ugol<>0) AND (R25<>0) Then
                                        Begin
                                          DlinI:=StrToFloat(Shir)+78;
                                          SD6.Cells[9,w]:=FloatToStr(DlinI);
                                          if DlinI<=350 Then
                                          DlinI:=350;
                                          if (DlinI>350) AND (DlinI<=500) Then
                                          DlinI:=500;
                                          if (DlinI>500) AND (DlinI<=750) Then
                                          DlinI:=750;
                                          if (DlinI>750) Then
                                          DlinI:=1000;
                                          if not Form1.mkQuerySelect1(Form1.ADOQuery1,
                                          'Select * from %s Where   ([№]=' + #39 +'7'+ #39 +') And  ([Длина]=' + #39 +FloatToStr(DlinI)+ #39 +')', ['[Резка Гибка]']) then
                                          exit;
                                          NC_S:= (Form1.ADOQuery1.FieldByName('Норма').AsString);
                                          If NC_S='' Then
                                          NC:=0 Else
                                          NC:=StrToFloat(NC_S)+0.007;
                                          SD6.Cells[14,w]:=FloatToStr(NC*Kol_Ed*Kol_Zap);
                                          if not Form1.mkQueryUpdate2(Form1.ADOQuery4,
                                        'UPDATE %s SET [Н\ч Пила]=' + #39
                                                + Form1.ConvertFloat(FloatToStr(NC*Kol_Ed*Kol_Zap)) + #39 +
                                        ' WHERE ([IdГП]=' + #39 + IDGP + #39 +
                                                ') AND (Элемент=' + #39 + Elem + #39 +') AND (Обозначение=' + #39 + Oboz + #39 +')  ', ['Специф']) then
                                        Exit;
                                          Inc(w);
                                SD6.RowCount:= SD6.RowCount+1;
                                        End;

                                        //++++++++++++++++++++++++++++++++++++++++++++++++
                                        if (Res_Detal=0) AND (Res_Lop<>0) AND (R25<>0) AND (Res_List<>0) Then
                                        Begin
                                          DlinI:=StrToFloat(Shir);
                                          SD6.Cells[9,w]:=FloatToStr(DlinI);
                                          if DlinI<=250 Then
                                          DlinI:=250;
                                          if (DlinI>250) AND (DlinI<=500) Then
                                          DlinI:=500;
                                          if (DlinI>500) AND (DlinI<=750) Then
                                          DlinI:=750;
                                          if (DlinI>750) Then
                                          DlinI:=1000;
                                          if not Form1.mkQuerySelect1(Form1.ADOQuery1,
                                          'Select * from %s Where   ([№]=' + #39 +'8'+ #39 +') And  ([Длина]=' + #39 +FloatToStr(DlinI)+ #39 +')', ['[Резка Гибка]']) then
                                          exit;
                                          NC_S:= (Form1.ADOQuery1.FieldByName('Норма').AsString);
                                          If NC_S='' Then
                                          NC:=0 Else
                                          NC:=StrToFloat(NC_S);
                                          SD6.Cells[14,w]:=FloatToStr(NC*Kol_Ed*Kol_Zap);
                                          if not Form1.mkQueryUpdate2(Form1.ADOQuery4,
                                        'UPDATE %s SET [Н\ч Пила]=' + #39
                                                + Form1.ConvertFloat(FloatToStr(NC*Kol_Ed*Kol_Zap)) + #39 +
                                        ' WHERE ([IdГП]=' + #39 + IDGP + #39 +
                                                ') AND (Элемент=' + #39 + Elem + #39 +') AND (Обозначение=' + #39 + Oboz + #39 +')  ', ['Специф']) then
                                        Exit;
                                          Inc(w);
                                                SD6.RowCount:= SD6.RowCount+1;

                                        End;
                                        Ugol:=Pos('ФЭЗ.П.0593',Mat);
                                        //++++++++++++++++++++++++++++++++++++++++++++++++
                                        if (Res_Detal=0) AND (Res_Lop<>0) AND (Ugol<>0) Then
                                        Begin
                                          DlinI:=StrToFloat(Shir);
                                          SD6.Cells[9,w]:=FloatToStr(DlinI);
                                          if DlinI<=250 Then
                                          DlinI:=250;
                                          if (DlinI>250) AND (DlinI<=500) Then
                                          DlinI:=500;
                                          if (DlinI>500) AND (DlinI<=750) Then
                                          DlinI:=750;
                                          if (DlinI>750) Then
                                          DlinI:=1000;
                                          if not Form1.mkQuerySelect1(Form1.ADOQuery1,
                                          'Select * from %s Where   ([№]=' + #39 +'8'+ #39 +') And  ([Длина]=' + #39 +FloatToStr(DlinI)+ #39 +')', ['[Резка Гибка]']) then
                                          exit;
                                          NC_S:= (Form1.ADOQuery1.FieldByName('Норма').AsString);
                                          If NC_S='' Then
                                          NC:=0 Else
                                          NC:=StrToFloat(NC_S);
                                          SD6.Cells[14,w]:=FloatToStr(NC*Kol_Ed*Kol_Zap);
                                          if not Form1.mkQueryUpdate2(Form1.ADOQuery4,
                                        'UPDATE %s SET [Н\ч Пила]=' + #39
                                                + Form1.ConvertFloat(FloatToStr(NC*Kol_Ed*Kol_Zap)) + #39 +
                                        ' WHERE ([IdГП]=' + #39 + IDGP + #39 +
                                                ') AND (Элемент=' + #39 + Elem + #39 +') AND (Обозначение=' + #39 + Oboz + #39 +')  ', ['Специф']) then
                                        Exit;
                                          Inc(w);
                                                SD6.RowCount:= SD6.RowCount+1;

                                        End;
                                //Break;
                      end;
                     //end;
                      //+++++++++++++++++++++++++++++++++++++++++++++++++++++

                     if Trumph=True Then
                     Trumpf_S:='True'
                     else Trumpf_S:='False';

                     if Kanban=True Then
                     Kanban_S:='True'
                     else Kanban_S:='False';

                     if Gib=True Then
                     Gib_S:='True'
                     else Gib_S:='False';

                     {if not Form1.mkQueryUpdate2(Form1.ADOQuery1,
                     'UPDATE %s SET [Trumph]=' +
                     #39+ Trumpf_S + #39 +
                     ', Гибка='+#39+ Gib_S+ #39 +
                     ', Канбан='+#39+ Kanban_S + #39 +
                     ' WHERE (Обозначение=' + #39 + Oboz + #39 +') AND ([IdГП]=' + #39 + IDGP + #39 +
                     ') ', ['Специф']) then
                     Exit; }
                       ADOQuery2.Next;

                 end;
        End;

        End;
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        Dir_main := ExtractFileDir(ParamStr(0));
        God := FormatDateTime('yyyy', Now);
        Mes := FormatDateTime('mmmm', Now);
        Dir :=Form1.Put_KTO+'\CKlapana\Накладные(пожар)1\' + God;
        CreateDir(Dir);

        Dir := Form1.Put_KTO+'\CKlapana\Накладные(пожар)1\' + God + '\' + Mes+'\';
        CreateDir(Dir);
        Puty:=(Dir +'\№ '+Label15.Caption+'.xlsx');
        XL := CreateOleObject('Excel.Application');
        CopyFile(PWideChar(Form1.Put_KTO+'\CKlapana\2013\Пожарные.xlsx'), PWideChar(Dir +'\№ '+Label15.Caption+'.xlsx'), False);
        XL.Workbooks.Open(Dir + '\№ '+Label15.Caption+'.xlsx');
        XL.Application.EnableEvents := false;
        {                       RSG1.Cells[0,s]:=IntToStr(s);
                                RSG1.Cells[1,s]:=Tip;
                                RSG1.Cells[2,s]:=Elem;
                                RSG1.Cells[3,s]:=FloatToStr(Kol_Oboz*Kol_Zap);
                                RSG1.Cells[4,s]:=Diam;
                                RSG1.Cells[5,s]:=Oboz;
                                //RSG1.Cells[6,s]:=Dlina;
                                //RSG1.Cells[7,s]:=Shir;
                                RSG1.Cells[6,s]:=Kol_Gib;
                                RSG1.Cells[7,s]:=Mat;
                                RSG1.Cells[8,s]:=N_S;//Н\ч}
                //+++++++++++++++++++++++++++++++++++++++
        For I:=0 To RSG1.RowCount-1 Do //Вальцы
        Begin

            XL.ActiveWorkBook.WorkSheets[10].Cells[(i)+(5), 2] := RSG1.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[10].Cells[(i)+(5), 3] := RSG1.Cells[1,i+1];//Тип
            XL.ActiveWorkBook.WorkSheets[10].Cells[(i)+(5), 4] := RSG1.Cells[2,i+1]; //Элемент
            XL.ActiveWorkBook.WorkSheets[10].Cells[(i)+(5), 5] := RSG1.Cells[4,i+1]; //Dlin
            //XL.ActiveWorkBook.WorkSheets[10].Cells[(i)+(5), 6] := RSG1.Cells[10,i+1]; //Shir
            //XL.ActiveWorkBook.WorkSheets[10].Cells[(i)+(5), 7] := RSG1.Cells[13,i+1]; //Virub
            XL.ActiveWorkBook.WorkSheets[10].Cells[(i)+(5), 8] := RSG1.Cells[3,i+1]; //KolEd
            XL.ActiveWorkBook.WorkSheets[10].Cells[(i)+(5), 9] := RSG1.Cells[8,i+1];//N\C
            XL.ActiveWorkBook.WorkSheets[10].Cells[(i)+(5), 10] := RSG1.Cells[7,i+1]; //Mat
            XL.ActiveWorkBook.WorkSheets[10].Cells[(i)+(5), 11] := RSG1.Cells[5,i+1]; //Obozn
            XL.ActiveWorkBook.WorkSheets[10].Range['B4', 'K' +
                IntToStr(I+5)].Borders.LineStyle := 1;
            XL.ActiveWorkBook.WorkSheets[10].Cells[2, 5] :=Label15.Caption;
            XL.ActiveWorkBook.WorkSheets[10].Cells[4, 1] :=Label15.Caption;
            NC_S:= RSG1.Cells[8,i+1];
            if NC_S='' Then
            NC_S:='0';
            NC:=NC+StrToFloat(NC_S);
            q:=I+5;
        end;
        NC_S:= FloatToStr(NC);
        NC_S:=Float(NC_S);
        XL.ActiveWorkBook.WorkSheets[10].Cells[q+1, 8]:='Итого';
        XL.ActiveWorkBook.WorkSheets[10].Cells[q+1, 9]:=NC_S;
        For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[10].Range['B'+IntToStr(q+4)+':E'+IntToStr(q+4)].Merge;
               XL.ActiveWorkBook.WorkSheets[10].Rows[q+4].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[10].Rows[q+4].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[10].Cells[q+4, 1] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[10].Cells[q+4, 2] := SG1.Cells[4,i+1];
               XL.ActiveWorkBook.WorkSheets[10].Cells[q+4, 6] := SG1.Cells[1,i+1];
               Inc(q);
        end;
        XL.ActiveWorkBook.WorkSheets[10].Cells[Q+7, 3] :='Вальцовщик';
         XL.ActiveWorkBook.WorkSheets[10].Cells[Q+7, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[10].Cells[Q+7, 9] :='Дата';
         XL.ActiveWorkBook.WorkSheets[10].Cells[Q+7, 11] :='____________________';
         XL.ActiveWorkBook.WorkSheets[10].Rows[q+7].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[10].Rows[q+7].Font.Size := 12;
                //+++++++++++++++++++++++++++++++++++++++
        For I:=0 To RSG2.RowCount-1 Do //Кромкогиб
        Begin

            XL.ActiveWorkBook.WorkSheets[11].Cells[(i)+(5), 2] := RSG2.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[11].Cells[(i)+(5), 3] := RSG2.Cells[1,i+1];//Тип
            XL.ActiveWorkBook.WorkSheets[11].Cells[(i)+(5), 4] := RSG2.Cells[2,i+1]; //Элемент
            XL.ActiveWorkBook.WorkSheets[11].Cells[(i)+(5), 5] := RSG2.Cells[4,i+1]; //Dlin
            //XL.ActiveWorkBook.WorkSheets[11].Cells[(i)+(5), 6] := RSG2.Cells[10,i+1]; //Shir
            //XL.ActiveWorkBook.WorkSheets[11].Cells[(i)+(5), 7] := RSG2.Cells[13,i+1]; //Virub
            XL.ActiveWorkBook.WorkSheets[11].Cells[(i)+(5), 8] := RSG2.Cells[3,i+1]; //KolEd
            XL.ActiveWorkBook.WorkSheets[11].Cells[(i)+(5), 9] := RSG2.Cells[8,i+1];//N\C
            XL.ActiveWorkBook.WorkSheets[11].Cells[(i)+(5), 10] := RSG2.Cells[7,i+1]; //Mat
            XL.ActiveWorkBook.WorkSheets[11].Cells[(i)+(5), 11] := RSG2.Cells[5,i+1]; //Obozn
            XL.ActiveWorkBook.WorkSheets[11].Range['B4', 'K' +
                IntToStr(I+5)].Borders.LineStyle := 1;
            XL.ActiveWorkBook.WorkSheets[11].Cells[2, 5] :=Label15.Caption;
            XL.ActiveWorkBook.WorkSheets[11].Cells[4, 1] :=Label15.Caption;
            NC_S:= RSG2.Cells[8,i+1];
            if NC_S='' Then
            NC_S:='0';
            NC:=NC+StrToFloat(NC_S);
            q:=I+5;
        end;
        NC_S:= FloatToStr(NC);
        NC_S:=Float(NC_S);
        XL.ActiveWorkBook.WorkSheets[11].Cells[q+1, 8]:='Итого';
        XL.ActiveWorkBook.WorkSheets[11].Cells[q+1, 9]:=NC_S;
        For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[11].Range['B'+IntToStr(q+4)+':E'+IntToStr(q+4)].Merge;
               XL.ActiveWorkBook.WorkSheets[11].Rows[q+4].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[11].Rows[q+4].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[11].Cells[q+4, 1] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[11].Cells[q+4, 2] := SG1.Cells[4,i+1];
               XL.ActiveWorkBook.WorkSheets[11].Cells[q+4, 6] := SG1.Cells[1,i+1];
               Inc(q);
        end;
        XL.ActiveWorkBook.WorkSheets[11].Cells[Q+7, 3] :='Кромкогибец';
         XL.ActiveWorkBook.WorkSheets[11].Cells[Q+7, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[11].Cells[Q+7, 9] :='Дата';
         XL.ActiveWorkBook.WorkSheets[11].Cells[Q+7, 11] :='____________________';
         XL.ActiveWorkBook.WorkSheets[11].Rows[q+7].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[11].Rows[q+7].Font.Size := 12;
                         //+++++++++++++++++++++++++++++++++++++++
        For I:=0 To RSG4.RowCount-1 Do //Токарный
        Begin

            XL.ActiveWorkBook.WorkSheets[12].Cells[(i)+(5), 2] := RSG4.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[12].Cells[(i)+(5), 3] := RSG4.Cells[1,i+1];//Тип
            XL.ActiveWorkBook.WorkSheets[12].Cells[(i)+(5), 4] := RSG4.Cells[2,i+1]; //Элемент
            XL.ActiveWorkBook.WorkSheets[12].Cells[(i)+(5), 5] := RSG4.Cells[4,i+1]; //Dlin
            //XL.ActiveWorkBook.WorkSheets[12].Cells[(i)+(5), 6] := RSG4.Cells[10,i+1]; //Shir
            //XL.ActiveWorkBook.WorkSheets[12].Cells[(i)+(5), 7] := RSG4.Cells[13,i+1]; //Virub
            XL.ActiveWorkBook.WorkSheets[12].Cells[(i)+(5), 8] := RSG4.Cells[3,i+1]; //KolEd
            XL.ActiveWorkBook.WorkSheets[12].Cells[(i)+(5), 9] := RSG4.Cells[8,i+1];//N\C
            XL.ActiveWorkBook.WorkSheets[12].Cells[(i)+(5), 10] := RSG4.Cells[7,i+1]; //Mat
            XL.ActiveWorkBook.WorkSheets[12].Cells[(i)+(5), 11] := RSG4.Cells[5,i+1]; //Obozn
            XL.ActiveWorkBook.WorkSheets[12].Range['B4', 'K' +
                IntToStr(I+5)].Borders.LineStyle := 1;
            XL.ActiveWorkBook.WorkSheets[12].Cells[2, 5] :=Label15.Caption;
            XL.ActiveWorkBook.WorkSheets[12].Cells[4, 1] :=Label15.Caption;
            NC_S:= RSG4.Cells[8,i+1];
            if NC_S='' Then
            NC_S:='0';
            NC:=NC+StrToFloat(NC_S);
            q:=I+5;
        end;
        NC_S:= FloatToStr(NC);
        NC_S:=Float(NC_S);
        XL.ActiveWorkBook.WorkSheets[12].Cells[q+1, 8]:='Итого';
        XL.ActiveWorkBook.WorkSheets[12].Cells[q+1, 9]:=NC_S;
        For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[12].Range['B'+IntToStr(q+4)+':E'+IntToStr(q+4)].Merge;
               XL.ActiveWorkBook.WorkSheets[12].Rows[q+4].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[12].Rows[q+4].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[12].Cells[q+4, 1] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[12].Cells[q+4, 2] := SG1.Cells[4,i+1];
               XL.ActiveWorkBook.WorkSheets[12].Cells[q+4, 6] := SG1.Cells[1,i+1];
               Inc(q);
        end;
        XL.ActiveWorkBook.WorkSheets[12].Cells[Q+7, 3] :='Токарь';
         XL.ActiveWorkBook.WorkSheets[12].Cells[Q+7, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[12].Cells[Q+7, 9] :='Дата';
         XL.ActiveWorkBook.WorkSheets[12].Cells[Q+7, 11] :='____________________';
         XL.ActiveWorkBook.WorkSheets[12].Rows[q+7].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[12].Rows[q+7].Font.Size := 12;
         //+++++++++++++++++++++++++++++++++++++++
        For I:=0 To RSG3.RowCount-1 Do //Ротор
        Begin

            XL.ActiveWorkBook.WorkSheets[13].Cells[(i)+(5), 2] := RSG3.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[13].Cells[(i)+(5), 3] := RSG3.Cells[1,i+1];//Тип
            XL.ActiveWorkBook.WorkSheets[13].Cells[(i)+(5), 4] := RSG3.Cells[2,i+1]; //Элемент
            XL.ActiveWorkBook.WorkSheets[13].Cells[(i)+(5), 5] := RSG3.Cells[4,i+1]; //Dlin
            //XL.ActiveWorkBook.WorkSheets[13].Cells[(i)+(5), 6] := RSG3.Cells[10,i+1]; //Shir
            //XL.ActiveWorkBook.WorkSheets[13].Cells[(i)+(5), 7] := RSG3.Cells[13,i+1]; //Virub
            XL.ActiveWorkBook.WorkSheets[13].Cells[(i)+(5), 8] := RSG3.Cells[3,i+1]; //KolEd
            XL.ActiveWorkBook.WorkSheets[13].Cells[(i)+(5), 9] := RSG3.Cells[8,i+1];//N\C
            XL.ActiveWorkBook.WorkSheets[13].Cells[(i)+(5), 10] := RSG3.Cells[7,i+1]; //Mat
            XL.ActiveWorkBook.WorkSheets[13].Cells[(i)+(5), 11] := RSG3.Cells[5,i+1]; //Obozn
            XL.ActiveWorkBook.WorkSheets[13].Range['B4', 'K' +
                IntToStr(I+5)].Borders.LineStyle := 1;
            XL.ActiveWorkBook.WorkSheets[13].Cells[2, 5] :=Label15.Caption;
            XL.ActiveWorkBook.WorkSheets[13].Cells[4, 1] :=Label15.Caption;
            NC_S:= RSG3.Cells[8,i+1];
            if NC_S='' Then
            NC_S:='0';
            NC:=NC+StrToFloat(NC_S);
            q:=I+5;
        end;
        NC_S:= FloatToStr(NC);
        NC_S:=Float(NC_S);
        XL.ActiveWorkBook.WorkSheets[13].Cells[q+1, 8]:='Итого';
        XL.ActiveWorkBook.WorkSheets[13].Cells[q+1, 9]:=NC_S;
        For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[13].Range['B'+IntToStr(q+4)+':E'+IntToStr(q+4)].Merge;
               XL.ActiveWorkBook.WorkSheets[13].Rows[q+4].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[13].Rows[q+4].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[13].Cells[q+4, 1] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[13].Cells[q+4, 2] := SG1.Cells[4,i+1];
               XL.ActiveWorkBook.WorkSheets[13].Cells[q+4, 6] := SG1.Cells[1,i+1];
               Inc(q);
        end;
        XL.ActiveWorkBook.WorkSheets[13].Cells[Q+7, 3] :='Роторист';
         XL.ActiveWorkBook.WorkSheets[13].Cells[Q+7, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[13].Cells[Q+7, 9] :='Дата';
         XL.ActiveWorkBook.WorkSheets[13].Cells[Q+7, 11] :='____________________';
         XL.ActiveWorkBook.WorkSheets[13].Rows[q+7].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[13].Rows[q+7].Font.Size := 12;
                  //+++++++++++++++++++++++++++++++++++++++
        For I:=0 To SG3.RowCount-1 Do //ЗигМашина
        Begin

            XL.ActiveWorkBook.WorkSheets[14].Cells[(i)+(5), 2] := SG3.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[14].Cells[(i)+(5), 3] := SG3.Cells[1,i+1];//Тип
            XL.ActiveWorkBook.WorkSheets[14].Cells[(i)+(5), 4] := SG3.Cells[2,i+1]; //Элемент
            XL.ActiveWorkBook.WorkSheets[14].Cells[(i)+(5), 5] := SG3.Cells[4,i+1]; //Dlin
            //XL.ActiveWorkBook.WorkSheets[13].Cells[(i)+(5), 6] := SG3.Cells[10,i+1]; //Shir
            //XL.ActiveWorkBook.WorkSheets[13].Cells[(i)+(5), 7] := SG3.Cells[13,i+1]; //Virub
            XL.ActiveWorkBook.WorkSheets[14].Cells[(i)+(5), 8] := SG3.Cells[3,i+1]; //KolEd
            XL.ActiveWorkBook.WorkSheets[14].Cells[(i)+(5), 9] := SG3.Cells[8,i+1];//N\C
            XL.ActiveWorkBook.WorkSheets[14].Cells[(i)+(5), 10] := SG3.Cells[7,i+1]; //Mat
            XL.ActiveWorkBook.WorkSheets[14].Cells[(i)+(5), 11] := SG3.Cells[5,i+1]; //Obozn
            XL.ActiveWorkBook.WorkSheets[14].Range['B4', 'K' +
                IntToStr(I+5)].Borders.LineStyle := 1;
            XL.ActiveWorkBook.WorkSheets[14].Cells[2, 5] :=Label15.Caption;
            XL.ActiveWorkBook.WorkSheets[14].Cells[4, 1] :=Label15.Caption;
            NC_S:= SG3.Cells[8,i+1];
            if NC_S='' Then
            NC_S:='0';
            NC:=NC+StrToFloat(NC_S);
            q:=I+5;
        end;
        NC_S:= FloatToStr(NC);
        NC_S:=Float(NC_S);
        XL.ActiveWorkBook.WorkSheets[14].Cells[q+1, 8]:='Итого';
        XL.ActiveWorkBook.WorkSheets[14].Cells[q+1, 9]:=NC_S;
        For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[14].Range['B'+IntToStr(q+4)+':E'+IntToStr(q+4)].Merge;
               XL.ActiveWorkBook.WorkSheets[14].Rows[q+4].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[14].Rows[q+4].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[14].Cells[q+4, 1] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[14].Cells[q+4, 2] := SG1.Cells[4,i+1];
               XL.ActiveWorkBook.WorkSheets[14].Cells[q+4, 6] := SG1.Cells[1,i+1];
               Inc(q);
        end;
         XL.ActiveWorkBook.WorkSheets[14].Cells[Q+7, 3] :='Роторист';
         XL.ActiveWorkBook.WorkSheets[14].Cells[Q+7, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[14].Cells[Q+7, 9] :='Дата';
         XL.ActiveWorkBook.WorkSheets[14].Cells[Q+7, 11] :='____________________';
         XL.ActiveWorkBook.WorkSheets[14].Rows[q+7].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[14].Rows[q+7].Font.Size := 12;
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++  Kanban
        For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[1].Range['B'+IntToStr(I+4)+':E'+IntToStr(I+4)].Merge;
               XL.ActiveWorkBook.WorkSheets[1].Rows[i+4].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[1].Rows[i+4].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[1].Cells[I+4, 1] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[1].Cells[I+4, 2] := SG1.Cells[4,i+1];
               XL.ActiveWorkBook.WorkSheets[1].Cells[I+4, 6] := SG1.Cells[1,i+1];
               Inc(q);
        end;
        Q:=Q+2;
        XL.ActiveWorkBook.WorkSheets[1].Rows[q+2].Font.Bold := True;
        XL.ActiveWorkBook.WorkSheets[1].Rows[q+3].Font.Bold := True;
        XL.ActiveWorkBook.WorkSheets[1].Rows[q+3].Font.Size := 14;
        XL.ActiveWorkBook.WorkSheets[1].Rows[q+2].Font.Size := 14;
        XL.ActiveWorkBook.WorkSheets[1].Select;

        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 2] := '№';
        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 3] := 'Тип клапана';
        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 4] := 'Изделие';
        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 5] := 'Размер';
        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 6] := 'Кол на 1 клапан';
        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 7] := 'Общее кол-во';
        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 8] := 'ЕдИзм';
        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 9] := 'Обознач';
        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 10] := 'Мат';
        XL.ActiveWorkBook.WorkSheets[1].Cells[q+2, 4] := 'Kanban';
        For I:=0 To SD1.RowCount-1 Do
        Begin
            Razm:=SD1.Cells[7,i+1]+'*'+SD1.Cells[7,i+1];
            if SD1.Cells[0,i+1]='' Then
            Continue;


            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 2] := SD1.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 3] := SD1.Cells[2,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 4] := SD1.Cells[3,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 5] := Razm;//Razmer
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 6] := SD1.Cells[4,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 7] := SD1.Cells[5,i+1];

            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 8] := SD1.Cells[1,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 9] := SD1.Cells[6,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 10] := SD1.Cells[12,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Range['B'+IntToStr(Q+5), 'J' +
                IntToStr(q+5)].Borders.LineStyle := 1;
            XL.ActiveWorkBook.WorkSheets[1].Cells[2, 6] :=Label15.Caption;
            Inc(Q);
        end;
        Q:=Q+2;

        XL.ActiveWorkBook.WorkSheets[1].Rows[q+4].Font.Bold := True;
        XL.ActiveWorkBook.WorkSheets[1].Rows[q+4].Font.Size := 14;
        XL.ActiveWorkBook.WorkSheets[1].Cells[(Q)+(4), 4] :='Заготовка';
        For I:=0 To SD2.RowCount-1 Do
        Begin
            Razm:=SD2.Cells[7,i+1]+'*'+SD2.Cells[8,i+1];
            if SD2.Cells[0,i+1]='' Then
            Continue;
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 2] := SD2.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 3] := SD2.Cells[2,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 4] := SD2.Cells[3,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 5] := Razm;
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 6] := SD2.Cells[4,i+1];//KolEd
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 7] := SD2.Cells[5,i+1];//KolOb

            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 8] := SD2.Cells[1,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 9] := SD2.Cells[6,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 10] := SD2.Cells[12,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Range['B'+IntToStr(Q+3), 'J' +
                IntToStr(I+5)].Borders.LineStyle := 1;
            Inc(Q);
        end;
        Q:=Q+2;
        XL.ActiveWorkBook.WorkSheets[1].Rows[q+4].Font.Bold := True;
        XL.ActiveWorkBook.WorkSheets[1].Rows[q+4].Font.Size := 14;
        XL.ActiveWorkBook.WorkSheets[1].Cells[(Q)+(4), 4] :='Сбор. Единицы';
        For I:=0 To SD.RowCount-1 Do
        Begin
            Razm:=SD.Cells[7,i+1]+'*'+SD.Cells[8,i+1];
            if SD.Cells[0,i+1]='' Then
            Continue;
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 2] := SD.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 3] := SD.Cells[2,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 4] := SD.Cells[3,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 5] := Razm;
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 6] := SD.Cells[4,i+1];//KolEd
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 7] := SD.Cells[5,i+1];//KolOb

            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 8] := SD.Cells[1,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 9] := SD.Cells[6,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 10] := SD.Cells[12,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Range['B'+IntToStr(Q+5), 'J' +
                IntToStr(I+5)].Borders.LineStyle := 1;
            Inc(Q);
        end;
         XL.ActiveWorkBook.WorkSheets[1].Cells[2, 6] :=Label15.Caption;
         XL.ActiveWorkBook.WorkSheets[1].Cells[Q+6, 4] :='Комплектовщик (Заготовка)';
         XL.ActiveWorkBook.WorkSheets[1].Cells[Q+8, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[1].Cells[Q+8, 8] :='Дата';
         XL.ActiveWorkBook.WorkSheets[1].Cells[Q+8, 9] :='____________________';
         XL.ActiveWorkBook.WorkSheets[1].Rows[q+8].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[1].Rows[q+8].Font.Size := 12;
         XL.ActiveWorkBook.WorkSheets[1].Rows[q+6].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[1].Rows[q+6].Font.Size := 12;
        //++++++++++++++++++++++++++++++++++++++++++++++++  Склад
        Q:=1;
        For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[2].Range['B'+IntToStr(I+3)+':D'+IntToStr(I+3)].Merge;
               XL.ActiveWorkBook.WorkSheets[2].Rows[i+3].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[2].Rows[i+3].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[2].Cells[I+3, 1] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[2].Cells[I+3, 2] := SG1.Cells[4,i+1];
               XL.ActiveWorkBook.WorkSheets[2].Cells[I+3, 5] := SG1.Cells[1,i+1];
               Inc(q);
        end;
        Q:=Q+2;
        XL.ActiveWorkBook.WorkSheets[2].Rows[q+2].Font.Bold := True;
        XL.ActiveWorkBook.WorkSheets[2].Rows[q+3].Font.Bold := True;
        XL.ActiveWorkBook.WorkSheets[2].Rows[q+3].Font.Size := 14;
        XL.ActiveWorkBook.WorkSheets[2].Rows[q+2].Font.Size := 14;
        XL.ActiveWorkBook.WorkSheets[2].Select;

        XL.ActiveWorkBook.WorkSheets[2].Cells[q+3, 2] := '№';
        XL.ActiveWorkBook.WorkSheets[2].Cells[q+3, 3] := 'Тип клапана';
        XL.ActiveWorkBook.WorkSheets[2].Cells[q+3, 4] := 'Изделие';
        XL.ActiveWorkBook.WorkSheets[2].Cells[q+3, 5] := 'Размер';
        XL.ActiveWorkBook.WorkSheets[2].Cells[q+3, 6] := 'Кол на 1 клапан';
        XL.ActiveWorkBook.WorkSheets[2].Cells[q+3, 7] := 'Общее кол-во';
        XL.ActiveWorkBook.WorkSheets[2].Cells[q+3, 8] := 'ЕдИзм';
        XL.ActiveWorkBook.WorkSheets[2].Cells[q+3, 9] := 'Обознач';
        XL.ActiveWorkBook.WorkSheets[2].Cells[q+3, 10] := 'Мат';
        XL.ActiveWorkBook.WorkSheets[2].Cells[q+2, 4] := 'Склад';
        For I:=0 To SD9.RowCount-1 Do
        Begin
            Razm:=SD9.Cells[7,i+1]+'*'+SD9.Cells[8,i+1];
            if SD9.Cells[0,i+1]='' Then
            Continue;
            XL.ActiveWorkBook.WorkSheets[2].Cells[(q)+(5), 2] := SD9.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[2].Cells[(q)+(5), 3] := SD9.Cells[2,i+1];
            XL.ActiveWorkBook.WorkSheets[2].Cells[(q)+(5), 4] := SD9.Cells[3,i+1];
            XL.ActiveWorkBook.WorkSheets[2].Cells[(q)+(5), 5] := Razm;
            XL.ActiveWorkBook.WorkSheets[2].Cells[(q)+(5), 6] := SD9.Cells[4,i+1];//KolEd
            XL.ActiveWorkBook.WorkSheets[2].Cells[(q)+(5), 7] := SD9.Cells[5,i+1];//KolOb

            XL.ActiveWorkBook.WorkSheets[2].Cells[(q)+(5), 8] := SD9.Cells[1,i+1];
            XL.ActiveWorkBook.WorkSheets[2].Cells[(q)+(5), 9] := SD9.Cells[6,i+1];
            XL.ActiveWorkBook.WorkSheets[2].Cells[(q)+(5), 10] := SD9.Cells[12,i+1];
            XL.ActiveWorkBook.WorkSheets[2].Range['B'+IntToStr(q+3), 'J' +
                IntToStr(q+5)].Borders.LineStyle := 1;
            Inc(Q);
        end;
         XL.ActiveWorkBook.WorkSheets[2].Cells[2, 6] :=Label15.Caption;
         XL.ActiveWorkBook.WorkSheets[2].Cells[Q+8, 3] :='Кладовщик';
         XL.ActiveWorkBook.WorkSheets[2].Cells[Q+8, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[2].Cells[Q+8, 8] :='Дата';
         XL.ActiveWorkBook.WorkSheets[2].Cells[Q+8, 9] :='____________________';
         XL.ActiveWorkBook.WorkSheets[2].Rows[q+8].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[2].Rows[q+8].Font.Size := 12;
        //++++++++++++++++++++++++++++++++++++++++++++++++  Склад Стенки
        Q:=1;
        For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[3].Range['B'+IntToStr(I+3)+':D'+IntToStr(I+3)].Merge;
               XL.ActiveWorkBook.WorkSheets[3].Rows[i+3].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[3].Rows[i+3].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[3].Cells[I+3, 1] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[3].Cells[I+3, 2] := SG1.Cells[4,i+1];
               XL.ActiveWorkBook.WorkSheets[3].Cells[I+3, 5] := SG1.Cells[1,i+1];
               Inc(q);
        end;
        Q:=Q+2;
        XL.ActiveWorkBook.WorkSheets[3].Rows[q+2].Font.Bold := True;
        XL.ActiveWorkBook.WorkSheets[3].Rows[q+3].Font.Bold := True;
        XL.ActiveWorkBook.WorkSheets[3].Rows[q+3].Font.Size := 14;
        XL.ActiveWorkBook.WorkSheets[3].Rows[q+2].Font.Size := 14;
        XL.ActiveWorkBook.WorkSheets[3].Select;

        XL.ActiveWorkBook.WorkSheets[3].Cells[q+3, 2] := '№';
        XL.ActiveWorkBook.WorkSheets[3].Cells[q+3, 3] := 'Тип клапана';
        XL.ActiveWorkBook.WorkSheets[3].Cells[q+3, 4] := 'Изделие';
        XL.ActiveWorkBook.WorkSheets[3].Cells[q+3, 5] := 'Размер';
        XL.ActiveWorkBook.WorkSheets[3].Cells[q+3, 6] := 'Кол на 1 клапан';
        XL.ActiveWorkBook.WorkSheets[3].Cells[q+3, 7] := 'Общее кол-во';
        XL.ActiveWorkBook.WorkSheets[3].Cells[q+3, 8] := 'ЕдИзм';
        XL.ActiveWorkBook.WorkSheets[3].Cells[q+3, 9] := 'Обознач';
        XL.ActiveWorkBook.WorkSheets[3].Cells[q+3, 10] := 'Мат';
        XL.ActiveWorkBook.WorkSheets[3].Cells[q+2, 4] := 'Склад Стенки';
        NC:=0;
        NC_S:='';
        For I:=0 To Stenki.RowCount-1 Do
        Begin
            Razm:=Stenki.Cells[7,i+1];
            if Stenki.Cells[0,i+1]='' Then
            Continue;
            XL.ActiveWorkBook.WorkSheets[3].Cells[(q)+(5), 2] := Stenki.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[3].Cells[(q)+(5), 3] := Stenki.Cells[2,i+1];
            XL.ActiveWorkBook.WorkSheets[3].Cells[(q)+(5), 4] := Stenki.Cells[3,i+1];
            XL.ActiveWorkBook.WorkSheets[3].Cells[(q)+(5), 5] := Razm;
            XL.ActiveWorkBook.WorkSheets[3].Cells[(q)+(5), 6] := Stenki.Cells[4,i+1];//KolEd
            XL.ActiveWorkBook.WorkSheets[3].Cells[(q)+(5), 7] := Stenki.Cells[5,i+1];//KolOb

            XL.ActiveWorkBook.WorkSheets[3].Cells[(q)+(5), 8] := Stenki.Cells[1,i+1];
            XL.ActiveWorkBook.WorkSheets[3].Cells[(q)+(5), 9] := Stenki.Cells[6,i+1];
            XL.ActiveWorkBook.WorkSheets[3].Cells[(q)+(5), 10] := Stenki.Cells[12,i+1];
            XL.ActiveWorkBook.WorkSheets[3].Range['B'+IntToStr(q+3), 'J' +
                IntToStr(q+5)].Borders.LineStyle := 1;
            Inc(Q);
        end;
         XL.ActiveWorkBook.WorkSheets[3].Cells[2, 6] :=Label15.Caption;
         XL.ActiveWorkBook.WorkSheets[3].Cells[Q+8, 3] :='Кладовщик';
         XL.ActiveWorkBook.WorkSheets[3].Cells[Q+8, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[3].Cells[Q+8, 8] :='Дата';
         XL.ActiveWorkBook.WorkSheets[3].Cells[Q+8, 9] :='____________________';
         XL.ActiveWorkBook.WorkSheets[3].Rows[q+8].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[3].Rows[q+8].Font.Size := 12;
        //+++++++++++++++++++++++++++++++++++++++
        For I:=0 To SD4.RowCount-1 Do //Ножницы
        Begin

            XL.ActiveWorkBook.WorkSheets[4].Cells[(i)+(5), 2] := SD4.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[4].Cells[(i)+(5), 3] := SD4.Cells[2,i+1];//Тип
            XL.ActiveWorkBook.WorkSheets[4].Cells[(i)+(5), 4] := SD4.Cells[3,i+1]; //Элемент
            XL.ActiveWorkBook.WorkSheets[4].Cells[(i)+(5), 5] := SD4.Cells[9,i+1]; //Dlin
            XL.ActiveWorkBook.WorkSheets[4].Cells[(i)+(5), 6] := SD4.Cells[10,i+1]; //Shir
            XL.ActiveWorkBook.WorkSheets[4].Cells[(i)+(5), 7] := SD4.Cells[13,i+1]; //Virub
            XL.ActiveWorkBook.WorkSheets[4].Cells[(i)+(5), 8] := SD4.Cells[4,i+1]; //KolEd
            XL.ActiveWorkBook.WorkSheets[4].Cells[(i)+(5), 9] := SD4.Cells[15,i+1];//N\C
            XL.ActiveWorkBook.WorkSheets[4].Cells[(i)+(5), 10] := SD4.Cells[12,i+1]; //Mat
            XL.ActiveWorkBook.WorkSheets[4].Cells[(i)+(5), 11] := SD4.Cells[6,i+1]; //Obozn
            XL.ActiveWorkBook.WorkSheets[4].Range['B4', 'K' +
                IntToStr(I+5)].Borders.LineStyle := 1;
            XL.ActiveWorkBook.WorkSheets[4].Cells[2, 5] :=Label15.Caption;
            XL.ActiveWorkBook.WorkSheets[4].Cells[4, 1] :=Label15.Caption;
            NC_S:= SD4.Cells[15,i+1];
            if NC_S='' Then
            NC_S:='0';
            NC:=NC+StrToFloat(NC_S);
            q:=I+5;
        end;
        NC_S:= FloatToStr(NC);
        NC_S:=Float(NC_S);
        XL.ActiveWorkBook.WorkSheets[4].Cells[q+1, 8]:='Итого';
        XL.ActiveWorkBook.WorkSheets[4].Cells[q+1, 9]:=NC_S;
        For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[4].Range['B'+IntToStr(q+4)+':E'+IntToStr(q+4)].Merge;
               XL.ActiveWorkBook.WorkSheets[4].Rows[q+4].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[4].Rows[q+4].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[4].Cells[q+4, 1] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[4].Cells[q+4, 2] := SG1.Cells[4,i+1];
               XL.ActiveWorkBook.WorkSheets[4].Cells[q+4, 6] := SG1.Cells[1,i+1];
               Inc(q);
        end;
        XL.ActiveWorkBook.WorkSheets[4].Cells[Q+7, 3] :='Резчик';
         XL.ActiveWorkBook.WorkSheets[4].Cells[Q+7, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[4].Cells[Q+7, 9] :='Дата';
         XL.ActiveWorkBook.WorkSheets[4].Cells[Q+7, 11] :='____________________';
         XL.ActiveWorkBook.WorkSheets[4].Rows[q+7].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[4].Rows[q+7].Font.Size := 12;
        //+++++++++++++++++++++++++++++++++++++++
        Q:=1;
        For I:=0 To SD3.RowCount-1 Do  //Trumph
        Begin

            XL.ActiveWorkBook.WorkSheets[5].Cells[(i)+(13), 5] := SD3.Cells[0,i];
            XL.ActiveWorkBook.WorkSheets[5].Cells[(i)+(13), 6] := SD3.Cells[2,i];
            XL.ActiveWorkBook.WorkSheets[5].Cells[(i)+(13), 7] := SD3.Cells[3,i];
            XL.ActiveWorkBook.WorkSheets[5].Cells[(i)+(13), 9] := SD3.Cells[7,i];
            XL.ActiveWorkBook.WorkSheets[5].Cells[(i)+(13), 10] := SD3.Cells[8,i];
            XL.ActiveWorkBook.WorkSheets[5].Cells[(i)+(13), 11] := SD3.Cells[14,i];
            XL.ActiveWorkBook.WorkSheets[5].Cells[(i)+(13), 8] := SD3.Cells[4,i];
            XL.ActiveWorkBook.WorkSheets[5].Cells[(i)+(13), 12] := SD3.Cells[12,i];
            XL.ActiveWorkBook.WorkSheets[5].Cells[(i)+(13), 13] := SD3.Cells[6,i];
            XL.ActiveWorkBook.WorkSheets[5].Range['E21', 'M' +
                IntToStr(I+13)].Borders.LineStyle := 1;
            //XL.ActiveWorkBook.WorkSheets[5].Cells[18, 9] :=Label15.Caption;
            //XL.ActiveWorkBook.WorkSheets[5].Cells[33, 2] :='СОПРОВОДИТЕЛЬНЫЙ ЛИСТ №  '+Label15.Caption;
            Q:=I+13;
        end;
        For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[5].Range['F'+IntToStr(Q+4)+':G'+IntToStr(Q+4)].Merge;
               XL.ActiveWorkBook.WorkSheets[5].Rows[q+4].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[5].Rows[q+4].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[5].Cells[q+4, 5] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[5].Cells[q+4, 6] := SG1.Cells[4,i+1];
               XL.ActiveWorkBook.WorkSheets[5].Cells[q+4, 8] := SG1.Cells[1,i+1];
               Inc(q);
        end;
        XL.ActiveWorkBook.WorkSheets[5].Cells[Q+7, 6] :='Оператор';
         XL.ActiveWorkBook.WorkSheets[5].Cells[Q+7, 7] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[5].Cells[Q+7, 8] :='Дата';
         XL.ActiveWorkBook.WorkSheets[5].Cells[Q+7, 9] :='____________________';
         XL.ActiveWorkBook.WorkSheets[5].Rows[q+7].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[5].Rows[q+7].Font.Size := 12;
         XL.ActiveWorkBook.WorkSheets[5].Cells[12, 2] :=Label15.Caption;
        //+++++++++++++++++++++++++++++++++++++++
        NC:=0;
        NC_S:='';
        NC_S2:='';
        NC_GER_S:='';
        For I:=0 To SD5.RowCount-1 Do //Гибка
        Begin

            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 2] := SD5.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 3] := SD5.Cells[2,i+1];//Тип
            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 4] := SD5.Cells[3,i+1]; //Элемент
            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 5] := SD5.Cells[4,i+1]; //  Kol
            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 6] := SD5.Cells[11,i+1]; // KolGib
            if SD5.Cells[14,i+1]<>'' Then
            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 7] :=FloatToStr(StrToFloat(SD5.Cells[14,i+1])); // N\C
            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 8] := SD5.Cells[9,i+1]; //A
            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 9] := SD5.Cells[10,i+1]; //B

            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 10] := SD5.Cells[15,i+1]; //B
            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 11] := SD5.Cells[17,i+1]; //B
            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 12] := SD5.Cells[18,i+1]; //B
            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 13] := SD5.Cells[16,i+1]; //B

            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 14] := SD5.Cells[12,i+1]; //Mat
            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 15] := SD5.Cells[6,i+1]; //Chert
            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 16] := SD5.Cells[20,i+1]; // N\C Выруб
            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 17] := SD5.Cells[21,i+1]; // HACO
            XL.ActiveWorkBook.WorkSheets[6].Range['B4', 'Q' +
                IntToStr(I+5)].Borders.LineStyle := 1;
            XL.ActiveWorkBook.WorkSheets[6].Cells[2, 5] :=Label15.Caption;
            NC_S:= SD5.Cells[14,i+1];
            if NC_S='' Then
            NC_S:='0';
            //============
            NC_S2:= SD5.Cells[20,i+1];//Выруб
            if NC_S2='' Then
            NC_S2:='0';
            //============
            NC_GER_S:= SD5.Cells[21,i+1];//HACO
            if NC_GER_S='' Then
            NC_GER_S:='0';
            NC:=NC+StrToFloat(NC_S)+StrToFloat(NC_GER_S)+StrToFloat(NC_S2) ;
            Q:=I+5;
        end;
        XL.ActiveWorkBook.WorkSheets[6].Cells[(Q)+1, 6] :='Итого';
        NC_S:= FloatToStr(NC);
        XL.ActiveWorkBook.WorkSheets[6].Cells[(q)+1, 7] :=NC_S;
        For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[6].Range['C'+IntToStr(q+2)+':D'+IntToStr(q+2)].Merge;
               XL.ActiveWorkBook.WorkSheets[6].Rows[q+2].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[6].Rows[q+2].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[6].Cells[q+2, 2] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[6].Cells[q+2, 3] := SG1.Cells[4,i+1];
               XL.ActiveWorkBook.WorkSheets[6].Cells[q+2, 6] := SG1.Cells[1,i+1];

               Inc(q);
        end;
        XL.ActiveWorkBook.WorkSheets[6].Cells[Q+5, 3] :='Оператор';
         XL.ActiveWorkBook.WorkSheets[6].Cells[Q+5, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[6].Cells[Q+5, 8] :='Дата';
         XL.ActiveWorkBook.WorkSheets[6].Cells[Q+5, 9] :='____________________';
         XL.ActiveWorkBook.WorkSheets[6].Rows[q+5].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[6].Rows[q+5].Font.Size := 12;
        //+++++++++++++++++++++++++++++++++++++++
         NC:=0;
        For I:=0 To SD6.RowCount-1 Do //Pila
        Begin

            XL.ActiveWorkBook.WorkSheets[7].Cells[(i)+(5), 2] := SD6.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[7].Cells[(i)+(5), 3] := SD6.Cells[2,i+1];//Тип
            XL.ActiveWorkBook.WorkSheets[7].Cells[(i)+(5), 4] := SD6.Cells[3,i+1]; //Элемент
            XL.ActiveWorkBook.WorkSheets[7].Cells[(i)+(5), 5] := SD6.Cells[9,i+1]; //Dlin
            NC_S2:= SD6.Cells[14,i+1];
            NC_S2:=Float(NC_S2);
            XL.ActiveWorkBook.WorkSheets[7].Cells[(i)+(5), 6] := NC_S2; //

            XL.ActiveWorkBook.WorkSheets[7].Cells[(i)+(5), 7] := SD6.Cells[4,i+1]; //KolEd
            XL.ActiveWorkBook.WorkSheets[7].Cells[(i)+(5), 8] := SD6.Cells[12,i+1];//Mat
            XL.ActiveWorkBook.WorkSheets[7].Cells[(i)+(5), 9] := SD6.Cells[6,i+1]; //Obozn
            XL.ActiveWorkBook.WorkSheets[7].Range['B4', 'J' +
                IntToStr(I+5)].Borders.LineStyle := 1;
            XL.ActiveWorkBook.WorkSheets[7].Cells[2, 5] :=Label15.Caption;
            XL.ActiveWorkBook.WorkSheets[7].Cells[4, 1] :=Label15.Caption;
            NC_S:= SD6.Cells[14,i+1];
            if NC_S='' Then
            NC_S:='0';
            NC:=NC+StrToFloat(NC_S);
            Q:=I+5;
        end;
        XL.ActiveWorkBook.WorkSheets[7].Cells[(Q)+1, 5] :='Итого';
        NC_S:= FloatToStr(NC);
        NC_S:=Float(NC_S);
        PosTrud := Pos(',', NC_S); //Трудоемкость FLOAT
                                if PosTrud <> 0 then
                                begin
                                        Delete(NC_S, PosTrud, 1);
                                        Insert('.', NC_S, PosTrud);
                                end;
        XL.ActiveWorkBook.WorkSheets[7].Cells[(q)+1, 6] :=NC_S;
                For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[7].Range['B'+IntToStr(q+4)+':E'+IntToStr(q+4)].Merge;
               XL.ActiveWorkBook.WorkSheets[7].Rows[q+4].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[7].Rows[q+4].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[7].Cells[q+4, 1] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[7].Cells[q+4, 2] := SG1.Cells[4,i+1];
               XL.ActiveWorkBook.WorkSheets[7].Cells[q+4, 6] := SG1.Cells[1,i+1];
               Inc(q);
        end;
        XL.ActiveWorkBook.WorkSheets[7].Cells[Q+7, 3] :='Оператор';
         XL.ActiveWorkBook.WorkSheets[7].Cells[Q+7, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[7].Cells[Q+7, 8] :='Дата';
         XL.ActiveWorkBook.WorkSheets[7].Cells[Q+7, 9] :='____________________';
         XL.ActiveWorkBook.WorkSheets[7].Rows[q+7].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[7].Rows[q+7].Font.Size := 12;
         NC:=0;
         //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        For I:=0 To SD7.RowCount-1 Do //Pila Lent
        Begin

            XL.ActiveWorkBook.WorkSheets[8].Cells[(i)+(5), 2] := SD7.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[8].Cells[(i)+(5), 3] := SD7.Cells[2,i+1];//Тип
            XL.ActiveWorkBook.WorkSheets[8].Cells[(i)+(5), 4] := SD7.Cells[3,i+1]; //Элемент
            XL.ActiveWorkBook.WorkSheets[8].Cells[(i)+(5), 5] := SD7.Cells[9,i+1]; //Dlin
            NC_S2:= SD7.Cells[14,i+1];
            NC_S2:=Float(NC_S2);
            XL.ActiveWorkBook.WorkSheets[8].Cells[(i)+(5), 6] := NC_S2; //

            XL.ActiveWorkBook.WorkSheets[8].Cells[(i)+(5), 7] := SD7.Cells[4,i+1]; //KolEd
            XL.ActiveWorkBook.WorkSheets[8].Cells[(i)+(5), 8] := SD7.Cells[12,i+1]; //Mat
            XL.ActiveWorkBook.WorkSheets[8].Cells[(i)+(5), 9] := SD7.Cells[6,i+1]; //Obozn
            XL.ActiveWorkBook.WorkSheets[8].Range['B4', 'J' +
                IntToStr(I+5)].Borders.LineStyle := 1;
            XL.ActiveWorkBook.WorkSheets[8].Cells[2, 5] :=Label15.Caption;
            XL.ActiveWorkBook.WorkSheets[8].Cells[4, 1] :=Label15.Caption;
            NC_S:= SD7.Cells[14,i+1];
            if NC_S='' Then
            NC_S:='0';
            NC:=NC+StrToFloat(NC_S);
            Q:=I+5;
        end;
        XL.ActiveWorkBook.WorkSheets[8].Cells[(Q)+1, 5] :='Итого';
        NC_S:= FloatToStr(NC);
        NC_S:=Float(NC_S);
        PosTrud := Pos(',', NC_S); //Трудоемкость FLOAT
                                if PosTrud <> 0 then
                                begin
                                        Delete(NC_S, PosTrud, 1);
                                        Insert('.', NC_S, PosTrud);
                                end;
        XL.ActiveWorkBook.WorkSheets[8].Cells[(q)+1, 6] :=NC_S;
                For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[8].Range['B'+IntToStr(q+4)+':E'+IntToStr(q+4)].Merge;
               XL.ActiveWorkBook.WorkSheets[8].Rows[q+4].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[8].Rows[q+4].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[8].Cells[q+4, 1] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[8].Cells[q+4, 2] := SG1.Cells[4,i+1];
               XL.ActiveWorkBook.WorkSheets[8].Cells[q+4, 6] := SG1.Cells[1,i+1];
               Inc(q);
        end;
         XL.ActiveWorkBook.WorkSheets[8].Cells[Q+7, 3] :='Оператор';
         XL.ActiveWorkBook.WorkSheets[8].Cells[Q+7, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[8].Cells[Q+7, 8] :='Дата';
         XL.ActiveWorkBook.WorkSheets[8].Cells[Q+7, 9] :='____________________';
         XL.ActiveWorkBook.WorkSheets[8].Rows[q+7].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[8].Rows[q+7].Font.Size := 12;
         NC:=0;
         //+++++++++++++++++++++++++++++++++++++++++++++++
         Q:=6;
        For I:=0 To SD3.RowCount-1 Do //Сопроводитр
        Begin

            XL.ActiveWorkBook.WorkSheets[9].Cells[(i)+(7), 2] := SD3.Cells[0,i];
            XL.ActiveWorkBook.WorkSheets[9].Cells[(i)+(7), 3] := SD3.Cells[2,i];//Тип
            XL.ActiveWorkBook.WorkSheets[9].Cells[(i)+(7), 4] := SD3.Cells[3,i]; //Элемент
            XL.ActiveWorkBook.WorkSheets[9].Cells[(i)+(7), 5] := SD3.Cells[7,i]; //Dlin
            XL.ActiveWorkBook.WorkSheets[9].Cells[(i)+(7), 6] := SD3.Cells[8,i]; //Shir

            XL.ActiveWorkBook.WorkSheets[9].Cells[(i)+(7), 7] := SD3.Cells[4,i]; //KolEd
            XL.ActiveWorkBook.WorkSheets[9].Cells[(i)+(7), 8] := SD3.Cells[12,i]; //Mat
            XL.ActiveWorkBook.WorkSheets[9].Cells[(i)+(7), 9] := SD3.Cells[6,i]; //Obozn
            XL.ActiveWorkBook.WorkSheets[9].Range['B4', 'J' +
                IntToStr(I+5)].Borders.LineStyle := 1;


            Inc(Q);
        end;
        XL.ActiveWorkBook.WorkSheets[9].Cells[5, 4] :='Trumpf';
            XL.ActiveWorkBook.WorkSheets[9].Cells[2, 5] :=Label15.Caption;
            XL.ActiveWorkBook.WorkSheets[9].Cells[4, 1] :=Label15.Caption;
        Q:=Q+4;
        //Nog
        XL.ActiveWorkBook.WorkSheets[9].Cells[q-2, 4] :='Ножницы';
        For I:=0 To SD4.RowCount-1 Do //Сопроводитр
        Begin

            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 2] := SD4.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 3] := SD4.Cells[2,i+1];//Тип
            XL.ActiveWorkBook.WorkSheets[9].Cells[q, 4] := SD4.Cells[3,i+1]; //Элемент
            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 5] := SD4.Cells[7,i+1]; //Dlin
            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 6] := SD4.Cells[8,i+1]; //Shir

            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 7] := SD4.Cells[4,i+1]; //KolEd
            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 8] := SD4.Cells[12,i+1]; //Mat
            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 9] := SD4.Cells[6,i+1]; //Obozn
            XL.ActiveWorkBook.WorkSheets[9].Range['B4', 'J' +
                IntToStr(q+5)].Borders.LineStyle := 1;

           Inc(Q);
        end;
        Q:=Q+4;
        //Nog
        XL.ActiveWorkBook.WorkSheets[9].Cells[q-2, 4] :='Гибка';
        For I:=0 To SD5.RowCount-1 Do //Сопроводитр
        Begin

            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 2] := SD5.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 3] := SD5.Cells[2,i+1];//Тип
            XL.ActiveWorkBook.WorkSheets[9].Cells[q, 4] := SD5.Cells[3,i+1]; //Элемент
            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 5] := SD5.Cells[7,i+1]; //Dlin
            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 6] := SD5.Cells[8,i+1]; //Shir

            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 7] := SD5.Cells[4,i+1]; //KolEd
            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 8] := SD5.Cells[12,i+1]; //Mat
            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 9] := SD5.Cells[6,i+1]; //Obozn
            XL.ActiveWorkBook.WorkSheets[9].Range['B4', 'J' +
                IntToStr(q+5)].Borders.LineStyle := 1;

           Inc(Q);
        end;
        For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[9].Range['B'+IntToStr(q+4)+':E'+IntToStr(q+4)].Merge;
               XL.ActiveWorkBook.WorkSheets[9].Rows[q+4].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[9].Rows[q+4].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[9].Cells[q+4, 1] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[9].Cells[q+4, 2] := SG1.Cells[4,i+1];
               XL.ActiveWorkBook.WorkSheets[9].Cells[q+4, 6] := SG1.Cells[1,i+1];
               Inc(q);
        end;
        XL.ActiveWorkBook.WorkSheets[9].Cells[Q+7, 3] :='Оператор';
         XL.ActiveWorkBook.WorkSheets[9].Cells[Q+7, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[9].Cells[Q+7, 8] :='Дата';
         XL.ActiveWorkBook.WorkSheets[9].Cells[Q+7, 9] :='____________________';
         XL.ActiveWorkBook.WorkSheets[9].Rows[q+7].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[9].Rows[q+7].Font.Size := 12;
        XL.Visible := True;
        XL := UnAssigned;
end;

procedure TFNewNakl.btn2Click(Sender: TObject);
Var S,Dir:String;
begin
        S := ExtractFileDir(ParamStr(0));
        //Puty:='\\192.168.7.1\Kto\CKlapana\Накладные(пожар)\2016\Декабрь\№ 6571.xlsx';
        Dir:=S +'\Пожарные\';
        CreateDir(Dir);
        CopyFile(PWideChar(Puty),PWideChar(Dir+'\№ ' + Label15.Caption + '.xlsx'), False);
end;

end.
