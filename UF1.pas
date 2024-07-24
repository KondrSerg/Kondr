unit UF1;

interface

uses
        Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls,
        Forms,
        Dialogs, Grids, StdCtrls, ComCtrls,ExcelXP, ExtCtrls, Math, ComObj, Buttons,UConnCeh,UConnKlap,UOsnova,
  AdvObj, BaseGrid, AdvGrid;

type
        TFV1 = class(TForm)
                Panel1: TPanel;
                Label1: TLabel;
                Label2: TLabel;
                Label3: TLabel;
                Label4: TLabel;
                Label5: TLabel;
                Label6: TLabel;
                Label7: TLabel;
                DateTimePicker1: TDateTimePicker;
                Memo1: TMemo;
                Panel2: TPanel;
                Button1: TButton;
                Button2: TButton;
                Button3: TButton;
    SGS: TStringGrid;
    SGSBrik: TStringGrid;
    btn2: TBitBtn;
    btn3: TBitBtn;
    btn1: TButton;
    Edt1: TEdit;
    Label3lbl1: TLabel;
    SG: TAdvStringGrid;
    Memommo1: TMemo;
    Label16: TLabel;
    Label17: TLabel;
    Button4: TButton;
                procedure FormShow(Sender: TObject);
                procedure SGSSelectCell(Sender: TObject; ACol, ARow: Integer;
                        var CanSelect: Boolean);
                procedure SGSSetEditText(Sender: TObject; ACol, ARow: Integer;
                        const Value: string);
                procedure Button2Click(Sender: TObject);
                procedure Button1Click(Sender: TObject);
                procedure FormClose(Sender: TObject; var Action: TCloseAction);
                procedure DateTimePicker1Enter(Sender: TObject);
    procedure SGSBrikSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure SGSBrikSetEditText(Sender: TObject; ACol, ARow: Integer;
      const Value: String);
    procedure Edt1Change(Sender: TObject);
    procedure btn2Click(Sender: TObject);
    procedure btn3Click(Sender: TObject);
    procedure btn1Click(Sender: TObject);
    procedure DateTimePicker1Change(Sender: TObject);
        private
                { Private declarations }
        public
                { Public declarations }
                Order:string;
                Conn_Ceh:Connect_Miass_Ceh;
                Conn_Klap:Connect_Miass_Klap;
                UOsnova_Main:Osnova_Main ;
                Cvet,Cvet1,Stat:Integer;
        end;

var
        FV1: TFV1;

implementation

uses Main;

{$R *.dfm}

procedure TFV1.FormShow(Sender: TObject);
var
        I, Kol, Kol_Zap, Max: Integer;
        S,B: string;
begin
        Button1.Enabled := False;
        if Form1.FlagDolg=1 Then
        Begin
            Button3.Visible := True;
            Edt1.Visible:=True;
        end;
        //Enabled_Ok1 := 0;
        //Enabled_Ok2 := 0;
        Label2.Caption := '0';
        Label4.Caption := '0';
        Label6.Caption := '0';
        DateTimePicker1.DateTime:=Now;
        for i:=0 To SGS.RowCount Do
        Begin
            SGS.Rows[i].Clear;
        End;
        btn1.Click;

end;

procedure TFV1.SGSSelectCell(Sender: TObject; ACol, ARow: Integer;
        var CanSelect: Boolean);
begin
        if ACol = 5 then
        begin
                SGS.Options := SGS.Options + [goEditing];

        end
        else
                SGS.Options := SGS.Options - [goEditing];
end;

procedure TFV1.SGSSetEditText(Sender: TObject; ACol, ARow: Integer;
        const Value: string);
var
        I, Kol_Zap, Kol_Zap_R, Kol: Integer;
        Svar, Sbor, Svar_o, Sbor_o: Double;
        Pdo,Kto:String;
begin
        I := ACol;
        if SGS.Cells[5, ARow] = '' then
                Exit;
        PDO:=SGS.Cells[27, ARow];
        KTO:=SGS.Cells[28, ARow];
        Kol_Zap := StrToInt(SGS.Cells[5, ARow]);
        Kol_Zap_R := StrToInt(SGS.Cells[11, ARow]);
        Kol := StrToInt(SGS.Cells[3, ARow]);
        if Kol<Kol_Zap then
        Begin
         MessageDlg('Запустить больше чем есть, не возможно!',
                                        mtError, [mbOk], 0);
         SGS.Cells[5, ARow]:='0';
         Exit;
        End;
        if SGS.Cells[6, ARow] = '' then
                Svar := 0
        else
                Svar := StrToFloat(SGS.Cells[6, ARow]);
        if SGS.Cells[7, ARow] = '' then
                Sbor := 0
        else
                Sbor := StrToFloat(SGS.Cells[7, ARow]);
        Svar := RoundTo(Svar * Kol_Zap, -2);
        Sbor := RoundTo(Sbor * Kol_Zap, -2);
        SGS.Cells[8, ARow] := FloatToStr(Svar);
        SGS.Cells[9, ARow] := FloatToStr(Sbor);
        Svar_o := 0;
        Sbor_o := 0;
        for I := 1 to SGS.RowCount do
        begin
                if SGS.Cells[8, I] <> '' then
                begin
                        Kol_Zap := StrToInt(SGS.Cells[5, ARow]);
                        Kol_Zap_R := StrToInt(SGS.Cells[4, ARow]);
                        //++++++++++++++++++++++++++++++++++++++++
                        Kol := Kol_Zap_R - Kol_Zap;
                        //++++++++++++++++++++++++++++++++++++++++
                        if SGS.Cells[5, i] = '' then
                                Kol_Zap := 0
                        else
                                Kol_Zap := StrToInt(SGS.Cells[5, i]);
                        if Kol_Zap = 0 then
                        begin
                                Svar_o := Svar_o - StrToFloat(SGS.Cells[8, I]);
                                Sbor_o := Sbor_o - StrToFloat(SGS.Cells[9, I]);
                        end;
                        if Kol_Zap <> 0 then
                        begin
                                Svar_o := RoundTo(Svar_o +
                                        StrToFloat(SGS.Cells[8, I]), -2);
                                Sbor_o := RoundTo(Sbor_o +
                                        StrToFloat(SGS.Cells[9, I]), -2);
                        end;
                end;
        end;
        Label2.Caption := FloatToStr(Svar_o);
        Label6.Caption := FloatToStr(Sbor_o);
end;

procedure TFV1.Button2Click(Sender: TObject);
begin
        Close;
end;

procedure TFV1.Button1Click(Sender: TObject);
var
        I, J,  Pos1, Kol_Zap, Kol_Zap_R, Kol, Rec_Count, Kol_Ob, Res, e:
        Integer;
        Svar, Sbor, Svar_o, Sbor_o,EL: Double;
        Dat, Dat1, Kol_Raz, Nomer_Raz, IdGP, Zak, S_Svar, S_Sbor, Dat_zak,
        Dat_Z, ss,Privod, Vn_Dat, Dir, S, Teh_Gotov, Teh_Obrab, Plan_Dat,God,
        mes,Shtrih_Str,Shtrih2_Str,Priv_Str,PDO,KTO,BZ,Dir_V,Kol_Lop,IzdelGP,A,B: string;
        XL2: Variant;
        Ii,Kol_Sh,Kol_pr,NachNom,KonNom,NomPos,Shtrih2,Shtrih,u,uu,ff,ff1: Integer;
        NachKod,KonKOd,KodPos:Int64;
Doc: OleVariant;
Appl, wdDoc, wdRang,Sheet2,Sheet1,Rang:variant;
IDKO,Briket,Brik_ned1,Brik_Den1,Brik_Ned2,Brik_Den2,Izdel,Nam,fileName,Dir1,R_Dat:string;
Elem,Vid,Oboz,Kol_S,Mat,Stan,D,bz1,UP: string;
Res_Nog,Res_Trumph,Res_Gib,Res_Pila,Kanban,Val ,Svark,Rotor,Pokr :Integer;
Sb,Sv,nom,Sb_T,Sb_L,Raskl,Sb_R,Sb_Obv,Sb_T_O,Sb_L_O,Raskl_O,Sb_R_O,Sb_Obv_O,El_o,EL_S:string;
begin
        e := 1;
        Nom:=Label4.Caption;
     if (MessageDlg('Проверь, все ли клапана? Потом не добавишь!',mtInformation,[mbYes,MbNo],0)=mrYes) Then
     Begin
        SyStem.SysUtils.FormatSettings.DecimalSeparator :=('.');
        Kol_Ob := 0;
        Dat1 := FormatDateTime('mm.dd.yyyy',
                                        DateTimePicker1.Date);
        Vn_Dat := FormatDateTime('mm.dd.yyyy', DateTimePicker1.Date);
        God := FormatDateTime('yyyy', Now);
        Mes := FormatDateTime('mmmm', Now);
        Dir1:='Суточные задания (Возд)';

        UOsnova_Main:=Osnova_Main.Create() ;
        S_Sbor:='0';
        S_Svar:='0';
        Sbor_o:=0;
        Svar_o:=0;
        //Label4.Caption:='6581';
         if Form1.Flag_Briket =0 Then  //C брикетом БЗ
         Begin
          for I := 1 to SGSBrik.RowCount do
          begin
                if SGSBrik.Cells[16, i] <> '' then
                begin
                        if (SGSBrik.Cells[9, I] <> '') and (SGSBrik.Cells[16, i] <> '') and
                                (SGSBrik.Cells[7, i] <> '') then
                        begin
                                BZ:=SGSBrik.Cells[4, I];
                                Izdel:= SGSBrik.Cells[13, i];
                                Nam:= SGSBrik.Cells[13, i];
                                IzdelGP:=SGSBrik.Cells[29, i];
                                A :=SGSBrik.Cells[20, i];
                                B :=SGSBrik.Cells[21, i];
                                Res:=Pos(' ',IzdelGP);
                                Delete(IzdelGP,1,Res+1);
                                KTO:=SGSBrik.Cells[30, i];
                                idgp:=SGSBrik.Cells[16, i];
                                idko:=SGSBrik.Cells[35, i];
                                Briket:=SGSBrik.Cells[1, i];
                                Brik_ned1:=SGSBrik.Cells[31, I];
                                Brik_Den1:=SGSBrik.Cells[32, I];
                                Brik_Ned2:=SGSBrik.Cells[33, I];
                                Brik_Den2:=SGSBrik.Cells[34, I];
                                Kol_lop:=SGSBrik.Cells[36, I];
                                bz1:=SGSBrik.Cells[37, I];
                                UP:= SGSBrik.Cells[40, I];
                                Nomer_Raz := '';
                                Kol_Raz := '';
                                Kol_Zap := StrToInt(SGSBrik.Cells[7, i]);
                                if  SGSBrik.Cells[15, i]='' Then
                                Kol_Pr:=0
                                else
                                Kol_pr := StrToInt(SGSBrik.Cells[15, i]);
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                NachKod:= StrToInt64(SGSBrik.Cells[25, i]);// по умолчанию =0
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                KonKod:= StrToInt64(SGSBrik.Cells[26, i]); // по умолчанию =0
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                KodPos:= StrToInt64(SGSBrik.Cells[28, i]); // по умолчанию =0
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                NachNom:= StrToInt(SGSBrik.Cells[23, i]); // по умолчанию =0
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                KonNom:= StrToInt(SGSBrik.Cells[24, i]); // по умолчанию =0
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                NomPos:= StrToInt(SGSBrik.Cells[27, i]); //НомерПос по умолчанию =1
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                if NomPos=1 Then
                                Begin
                                  NomPos:=NachNom+Kol_Zap-1;
                                  KodPos:=NachKod +Kol_Zap-1;
                                end
                                Else
                                Begin
                                  NomPos:=NomPos+Kol_Zap;
                                  KodPos:=KodPos+Kol_Zap;
                                end;
                                KonNom:=NomPos;
                                KonKod:=KodPos;
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                Plan_Dat := Form1.ConvertDat1(SGSBrik.Cells[2, i]);
                                if SGSBrik.Cells[39, i]='' then
                                R_Dat:='NULL'
                                else
                                R_Dat := #39+Form1.ConvertDat1(SGSBrik.Cells[39, i])+#39;
                               // R_Dat := Form1.ConvertDat1(SGSBrik.Cells[39, i]);
                                Teh_Gotov := Form1.ConvertDat1(SGSBrik.Cells[22,
                                        i]);
                                //Teh_Obrab := Form1.ConvertDat1(SGS.Cells[20,
                                //        i]);
                                if Kol_Zap = 0 then
                                        Continue;
                                Kol_Ob := Kol_Ob + Kol_Zap;
                                Kol_Zap_R := StrToInt(SGSBrik.Cells[17, i]);
                                Kol := StrToInt(SGSBrik.Cells[5, i]);
                                Privod := (SGSBrik.Cells[14, i]);

                                Dat_Z := SGSBrik.Cells[12, i];
                                res := Pos('.', Dat_z);
                                if Res <> 0 then
                                begin
                                        SS := Copy(Dat_Z, 1, Res);
                                        Delete(Dat_Z, 1, Res);
                                        Insert(SS, Dat_z, 4);
                                end;
                                Zak := SGSBrik.Cells[3, i];
                                Dat_Zak := Dat_z;
                                Dat := FormatDateTime('mm.dd.yyyy', Now);
                                Dat1 := FormatDateTime('mm.dd.yyyy',
                                        DateTimePicker1.Date);

                                if Kol <> Kol_Zap then
                                begin
                                        if SGSBrik.Cells[18, i] = '' then
                                                Kol_Raz := IntToStr(Kol_Zap)
                                        else
                                                Kol_Raz := SGSBrik.Cells[18, i] + ','
                                                        + IntToStr(Kol_Zap);
                                        if SGSBrik.Cells[19, i] = '' then
                                                Nomer_Raz := Nom
                                        else
                                                Nomer_Raz := SGSBrik.Cells[19, i] +
                                                        ',' + Nom;
                                end;
                                if Kol = Kol_Zap then
                                begin
                                        Kol_Raz := IntToStr(Kol_Zap);
                                        Nomer_Raz := Nom;
                                end;
                                Form1.Memo2.Lines.Add('1');

                                SV:=StringReplace(SGSBrik.Cells[8, I], ',', '.',[rfReplaceAll, rfIgnoreCase]);
                                Sb:=StringReplace(SGSBrik.Cells[9, I], ',', '.',[rfReplaceAll, rfIgnoreCase]);
                                S_Svar :=FloatToStr(StrToFloat(SV)*KOL_Zap);
                                S_Sbor :=FloatToStr(StrToFloat( Sb)*KOL_Zap);

                                //  Sb_T,Sb_L,Raskl,Sb_R,Sb_T_O,Sb_L_O,Raskl_O,Sb_R_O
                                Raskl:=StringReplace(SGSBrik.Cells[42, I], ',', '.',[rfReplaceAll, rfIgnoreCase]);
                                Raskl_O :=FloatToStr(StrToFloat(Raskl)*KOL_Zap);
                                //
                                Sb_L:=StringReplace(SGSBrik.Cells[43, I], ',', '.',[rfReplaceAll, rfIgnoreCase]);
                                Sb_L_O :=FloatToStr(StrToFloat(Sb_L)*KOL_Zap);
                                //
                                Sb_t:=StringReplace(SGSBrik.Cells[44, I], ',', '.',[rfReplaceAll, rfIgnoreCase]);
                                Sb_T_O :=FloatToStr(StrToFloat(Sb_t)*KOL_Zap);
                                //
                                Sb_R:=StringReplace(SGSBrik.Cells[45, I], ',', '.',[rfReplaceAll, rfIgnoreCase]);
                                Sb_R_O :=FloatToStr(StrToFloat(Sb_R)*KOL_Zap);
                                //
                                Sb_Obv:=StringReplace(SGSBrik.Cells[46, I], ',', '.',[rfReplaceAll, rfIgnoreCase]);
                                Sb_Obv_O :=FloatToStr(StrToFloat(Sb_Obv)*KOL_Zap);
                                //
                                EL_S:=StringReplace(SGSBrik.Cells[47, I], ',', '.',[rfReplaceAll, rfIgnoreCase]);
                                EL_O :=FloatToStr(StrToFloat(EL_S)*KOL_Zap);
                                //
                                Res := Pos(',', S_Svar);
                                Delete(S_Svar, Res, 1);
                                if Res <> 0 then
                                        Insert('.', S_Svar, Res);

                                Res := Pos(',', S_Sbor);
                                Delete(S_Sbor, Res, 1);
                                if Res <> 0 then
                                        Insert('.', S_Sbor, Res);

                                Priv_Str:=IntToStr(Kol_pr*Kol_Zap);

                                KonNom:=NomPos;
                                KonKod:=KodPos;

                                if not Form1.mkQueryInsert(Form1.ADOQuery2,
                                'Insert Into %s ' +
                                '([Дата],[план Дата],Заказ,Изделие,[Кол во],'+
                                '[Кол во запущенных],[' + FN_RAS_DATA_GOTOVN+ '],Номер,[Н\ч Сварка],[Н\ч Сборка Клапана],'+
                                '[МодПривода],НачНомер,КонНомер,НачКод,КонКод,Привод,'+
                                'IdГП,ПДО,КТО,БЗ,IdКО,'+
                                'Брик_Нед1,Брик_Ден1,Брик_Нед2,Брик_Ден2,Брикет,'+
                                'КолЛопаток,Планирование,Cvet,bz,[Сборка Примечание],Исполнение,Заказчик,'+
                                'Расключение,[Сборка лопаток],[Сборка тяг],СборкаРеш,Обварка,Электро) ' +
                                'Values (%s,%s,%s,%s,%s,' +
                                '%s,%s,%s,%s ,%s,'+
                                '%s,%s,%s,%s,%s,%s,'+
                                '%s,%s,%s,%s ,%s,'+
                                '%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s ,%s,%s,%s,%s ,%s,%s,%s)',
                                ['ЗапускВозд',
                                                #39 + Dat_Zak + #39,
                                                #39 + Plan_dat + #39,
                                                #39 + SGSBrik.Cells[3, i] +#39,
                                                #39 + SGSBrik.Cells[13, i] +#39,
                                                #39 + SGSBrik.Cells[5, i] +#39,

                                                #39 + IntToStr(Kol_Zap)+ #39,
                                                 R_Dat ,
                                                #39 + nom +#39,
                                                #39 + S_Svar + #39,
                                                #39 + S_Sbor + #39,

                                                #39 + Privod + #39,
                                                #39 + IntToStr(NachNom) +#39,
                                                #39 + IntToStr(KonNom) +#39,
                                                #39 + IntToStr(NachKod) +#39,
                                                #39 + IntToStr(KonKod) +#39,
                                                #39 + Priv_Str +#39,

                                                #39 + IdGP +#39,
                                                #39 + IzdelGP +#39,
                                                #39 + KTO +#39,
                                                #39 + BZ +#39,
                                                #39 + IDKO +#39,
                                                #39 + Brik_Ned1 +#39,
                                                #39 + Brik_Den1 +#39,
                                                #39 + Brik_Ned2 +#39,
                                                #39 + Brik_Den2 +#39,
                                                #39 + Briket +#39,
                                                #39 + Kol_Lop +#39,#39 + Dat1 +#39,#39 + IntToStr(Cvet1) +#39,#39 + bz1 +#39,
                                                #39 + SGSBrik.Cells[38, i] +#39,
                                                #39 + SGSBrik.Cells[40, i] +#39,
                                                #39 + SGSBrik.Cells[41, i] +#39,

                                                #39 + Raskl_O+#39,
                                                #39 + Sb_L_O +#39,
                                                #39 + Sb_T_O +#39,
                                                #39 + Sb_R_O +#39,
                                                #39 + Sb_Obv_O +#39,
                                                #39 + El_O +#39

                                                ])
                                then
                                exit;
                                Form1.Memo2.Lines.Add('2');
                                Svar_o := RoundTo(Svar_o +
                                        StrToFloat(S_Svar), -2);
                                Sbor_o := RoundTo(Sbor_o +
                                        StrToFloat(S_Sbor), -2);
                                //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

                                Inc(e);
                                //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                        end;
                end;
          end;
         end
         Else
         Begin
          for I := 1 to SGS.RowCount do
          begin
                IDGP:=SGS.Cells[10, i];
                if SGS.Cells[10, i] <> '' then
                begin
                        if (SGS.Cells[7, I] <> '') and (IDGP <> '') and
                                (SGS.Cells[5, i] <> '') then
                        begin
                                BZ:=SGS.Cells[29, i];
                                BZ1:=SGS.Cells[31, i];
                                Kol_Lop:=SGS.Cells[30, i];
                                IzdelGP:=SGS.Cells[27, i];
                                Res:=Pos(' ',IzdelGP);
                                Delete(IzdelGP,1,Res+1);
                                KTO:=SGS.Cells[28, i];

                                Izdel:= SGS.Cells[16, i];
                                Nam:= SGS.Cells[16, i];
                                A :=SGS.Cells[17, i];
                                B :=SGS.Cells[18, i];
                                idgp:=SGS.Cells[10, i];
                                Nomer_Raz := '';
                                Kol_Raz := '';
                                Kol_Zap := StrToInt(SGS.Cells[5, i]);
                                if  SGS.Cells[24, i]='' Then
                                Kol_Pr:=0
                                else
                                Kol_pr := StrToInt(SGS.Cells[24, i]);
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                if  SGS.Cells[22, i]='' Then
                                NachKod:=0
                                Else
                                NachKod:= StrToInt64(SGS.Cells[22, i]);
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                if  SGS.Cells[23, i]='' Then
                                KonKod:=0
                                Else
                                KonKod:= StrToInt64(SGS.Cells[23, i]);
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                if  SGS.Cells[20, i]='' Then
                                NachNom:=0
                                Else
                                NachNom:= StrToInt(SGS.Cells[20, i]);
                                //+++++++++++++++++++++++++++++++++++++++++++++

                                if  SGS.Cells[21, i]='' Then
                                KonNom:=0
                                Else
                                KonNom:= StrToInt(SGS.Cells[21, i]);
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                if  SGS.Cells[25, i]='' Then
                                NomPos:=0
                                Else
                                NomPos:= StrToInt(SGS.Cells[25, i]);
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                if NomPos=1 Then
                                Begin
                                  NomPos:=NachNom+Kol_Zap-1;
                                  KodPos:=NachKod +Kol_Zap-1;
                                end
                                Else
                                Begin
                                  NomPos:=NomPos+Kol_Zap;
                                  KodPos:=KodPos+Kol_Zap;
                                end;
                                KonNom:=NomPos;
                                KonKod:=KodPos;
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                idgp:=SGS.Cells[10, i];
                                Plan_Dat := Form1.ConvertDat1(SGS.Cells[1, i]);
                                if SGS.Cells[33, i]='' then
                                R_Dat:='NULL'
                                else
                                R_Dat := #39+Form1.ConvertDat1(SGS.Cells[33, i])+#39;
                                Teh_Gotov := Form1.ConvertDat1(SGS.Cells[19, i]);
                                if Kol_Zap = 0 then
                                        Continue;
                                Kol_Ob := Kol_Ob + Kol_Zap;
                                Kol_Zap_R := StrToInt(SGS.Cells[11, i]);
                                Kol := StrToInt(SGS.Cells[3, i]);
                                Privod := (SGS.Cells[14, i]);

                                Dat_Z := SGS.Cells[15, i];
                                res := Pos('.', Dat_z);
                                if Res <> 0 then
                                begin
                                        SS := Copy(Dat_Z, 1, Res);
                                        Delete(Dat_Z, 1, Res);
                                        Insert(SS, Dat_z, 4);
                                end;
                                Zak := SGS.Cells[2, i];
                                Dat_Zak := Dat_z;
                                Dat := FormatDateTime('mm.dd.yyyy', Now);
                                Dat1 := FormatDateTime('mm.dd.yyyy',
                                        DateTimePicker1.Date);

                                if Kol <> Kol_Zap then
                                begin
                                        if SGS.Cells[12, i] = '' then
                                                Kol_Raz := IntToStr(Kol_Zap)
                                        else
                                                Kol_Raz := SGS.Cells[12, i] + ','
                                                        + IntToStr(Kol_Zap);
                                        if SGS.Cells[13, i] = '' then
                                                Nomer_Raz := Nom
                                        else
                                                Nomer_Raz := SGS.Cells[13, i] +
                                                        ',' + Nom;
                                end;
                                if Kol = Kol_Zap then
                                begin
                                        Kol_Raz := IntToStr(Kol_Zap);
                                        Nomer_Raz := Nom;
                                end;
                                Form1.Memo2.Lines.Add('3');
                                SV:=StringReplace(SGS.Cells[8, I], ',', '.',[rfReplaceAll, rfIgnoreCase]);
                                Sb:=StringReplace(SGS.Cells[9, I], ',', '.',[rfReplaceAll, rfIgnoreCase]);
                                S_Svar :=FloatToStr(StrToFloat(SV));
                                S_Sbor :=FloatToStr(StrToFloat(Sb));
                                //  Sb_T,Sb_L,Raskl,Sb_R,Sb_T_O,Sb_L_O,Raskl_O,Sb_R_O
                                Raskl:=StringReplace(SGS.Cells[42, I], ',', '.',[rfReplaceAll, rfIgnoreCase]);
                                if Raskl='' then
                                   Raskl:='0';
                                Raskl_O :=FloatToStr(StrToFloat(Raskl)*KOL_Zap);
                                //
                                Sb_L:=StringReplace(SGS.Cells[43, I], ',', '.',[rfReplaceAll, rfIgnoreCase]);
                                if Sb_l='' then
                                   Sb_l:='0';
                                Sb_L_O :=FloatToStr(StrToFloat(Sb_L)*KOL_Zap);
                                //
                                Sb_t:=StringReplace(SGS.Cells[44, I], ',', '.',[rfReplaceAll, rfIgnoreCase]);
                                if Sb_t='' then
                                   Sb_t:='0';
                                Sb_T_O :=FloatToStr(StrToFloat(Sb_t)*KOL_Zap);
                                //
                                Sb_R:=StringReplace(SGS.Cells[45, I], ',', '.',[rfReplaceAll, rfIgnoreCase]);
                                if Sb_R='' then
                                   Sb_R:='0';
                                Sb_R_O :=FloatToStr(StrToFloat(Sb_R)*KOL_Zap);
                                //
                                //
                                Sb_Obv:=StringReplace(SGS.Cells[46, I], ',', '.',[rfReplaceAll, rfIgnoreCase]);
                                if Sb_Obv='' then
                                   Sb_Obv:='0';
                                Sb_Obv_O :=FloatToStr(StrToFloat(Sb_Obv)*KOL_Zap);
                                //
                                EL_S:=StringReplace(SGS.Cells[47, I], ',', '.',[rfReplaceAll, rfIgnoreCase]);
                                EL_O :=FloatToStr(StrToFloat(EL_S)*KOL_Zap);
                                //
                                Res := Pos(',', S_Svar);
                                Delete(S_Svar, Res, 1);
                                if Res <> 0 then
                                        Insert('.', S_Svar, Res);

                                Res := Pos(',', S_Sbor);
                                Delete(S_Sbor, Res, 1);
                                if Res <> 0 then
                                        Insert('.', S_Sbor, Res);

                                Priv_Str:=IntToStr(Kol_pr*Kol_Zap);
                                if not Form1.mkQueryInsert(Form1.ADOQuery2,
                                                'Insert Into %s ' +
                                                '([Дата],[план Дата],Заказ,Изделие,[Кол во],'+
                                                '[Кол во запущенных],[' + FN_RAS_DATA_GOTOVN+ '],Номер,[Н\ч Сварка],[Н\ч Сборка Клапана],'+
                                                '[МодПривода],НачНомер,КонНомер,НачКод,КонКод,'+
                                                'Привод,IdГП,ПДО,КТО,БЗ,КолЛопаток,Планирование,Cvet,bz,[Сборка Примечание],Исполнение,Заказчик,'+
                                                'Расключение,[Сборка лопаток],[Сборка тяг],СборкаРеш,Обварка,Электро) ' +
                                                'Values (%s,%s,%s,%s,' +
                                                '%s,%s,%s,%s,%s ,%s,%s,%s,%s,%s,%s,'+
                                                '%s,%s,%s,%s,%s,%s,%s,%s ,%s,%s,%s ,%s,%s,%s,%s,%s,%s ,%s)',
                                                ['ЗапускВозд',
                                                #39 + Dat_Zak + #39,
                                                #39 + Plan_dat + #39,
                                                #39 + SGS.Cells[2, i] +#39,
                                                #39 + SGS.Cells[16, i] +#39,
                                                #39 + SGS.Cells[3, i] +#39,
                                                #39 + IntToStr(Kol_Zap)+ #39,
                                                R_Dat ,
                                                #39 + Nom +#39,
                                                #39 + S_Svar + #39,
                                                #39 + S_Sbor + #39,
                                                #39 + Privod + #39,
                                                #39 + IntToStr(NachNom) +#39,
                                                #39 + IntToStr(KonNom) +#39,
                                                #39 + IntToStr(NachKod) +#39,
                                                #39 + IntToStr(KonKod) +#39,
                                                #39 + Priv_Str +#39,
                                                #39 + IdGP +#39,
                                                #39 + IzdelGP +#39,
                                                #39 + KTO +#39,
                                                #39 + BZ +#39,  #39 + Kol_lop +#39 ,
                                                #39 + Dat1 +#39,#39 + IntToStr(Cvet1) +#39,#39 + bz1 +#39,
                                                #39 + SGS.Cells[32, i] +#39 ,
                                                #39 + SGS.Cells[34, i] +#39 ,
                                                #39 + SGS.Cells[35, i] +#39,
                                                #39 + Raskl_O+#39,
                                                #39 + Sb_L_O +#39,
                                                #39 + Sb_T_O +#39,
                                                #39 + Sb_R_O +#39,
                                                #39 + Sb_Obv_O +#39,
                                                #39 + El_O +#39])
                                then
                                exit;

                                        //++++++++++++++++++++++++++++++++++++++++++++++++++++
                                Form1.Memo2.Lines.Add('4');
                                Svar_o := RoundTo(Svar_o +
                                        StrToFloat(SV), -2);
                                Sbor_o := RoundTo(Sbor_o +
                                        StrToFloat(Sb), -2);
                                //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

                                Inc(e);
                                //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                        end;
                end;
          end;
         end;
        // Form1.ADOQuery1.FieldByName('РаскрЛопаток').AsString;
         UOsnova_Main:=Osnova_Main.Create() ;
         fileName := '';//ExtractFileDir(ParamStr(0));//+'\Klapan.EXE';
         Vn_Dat := FormatDateTime('dd.mm.yyyy', DateTimePicker1.Date);
         UOsnova_Main.Flag_Error :=0;
        { if not Form1.mkQuerySelect(Form1.ADOQuery2,
                                        'Select * from %s Where  (Заказ= ' + #39
                                        + Zak + #39 + ') AND ([IdГП]= ' +
                                        #39 + (IDgp) + #39 + ') AND ([IdКО]= ' +
                                        #39 + (KO) + #39 + ')',
                                        ['KlapanaZap']) then
                                        exit;    }
                Stat:=1;// StrToInt(Form1.ADOQuery2.FieldByName('Статус').AsString);
         if Stat=1 then
         begin
            UOsnova_Main.Flag_Error := 0;                            //  Edit1.Text
            UOsnova_Main.Tab3:='СпецифОбщая';
            Form1.Memo4.Lines.Add('СпецифОбщая');
            if not UOsnova_Main.Osnova(Vn_Dat, 'ЗапускВозд', 'СпецифВозд', #39 + Label4.Caption + #39, fileName, 1, Cvet1) then
            Exit;

         end

         else
         begin
            UOsnova_Main.Flag_Error := 0;                            //  Edit1.Text
            UOsnova_Main.Tab3:='СпецифФлекс';
            Form1.Memo4.Lines.Add('СпецифФлекс');
            UOsnova_Main.Osnova_Flex( Vn_Dat,'[ЗапускВозд]','СпецифВозд',
            #39+Nom+#39,fileName, 1,Cvet1);
         end;
        // UOsnova_Main.Osnova( Vn_Dat,'ЗапускВозд','СпецифВозд',#39+Nom+#39,fileName,1,Cvet1 ) ;
        //then
        //Exit;
          if  UOsnova_Main.Flag_Error =2 Then //ДЕТАЛИ НЕ ВСЕ
          //ЗАНЕСТИ В ЗАПУСК И Поставить флаг ФлагЗаготовки
          Begin
            if not Form1.mkQueryDelete( Form1.ADOQuery1, 'DELETE FROM %s Where (Номер= '
                        +#39+Nom+#39+') '  ,
                        ['ЗапускВозд'] )
            Then
            Exit;
            //++++++++++++++++++++++++++++++++++++++++++++++++++++
            Form1.Memo2.Lines.Add('UOsnova_Main.Osnova( ДЕТАЛИ НЕ ВСЕ,ЗапускВозд,СпецифВозд,'+Nom+','+fileName+' )UOsnova_Main.Flag_Error =2') ;
            Exit;
          end;
        //++++++++++++++++++++++++++++++++++++++++++++++++++
          {if  UOsnova_Main.Flag_Error =1 Then //ДАТА ЛЕВАЯ УДАЛИТЬ Из ЗАПУСКА
                                                //И UPDATE КЛАПАНА НЕ ДЕЛАТЬ
          Begin
                        if not Form1.mkQueryDelete( Form1.ADOQuery1, 'DELETE FROM %s Where (Номер= '
                        +#39+Nom+#39+') '  ,
                        ['ЗапускВозд'] )
                        Then
                        Exit;
                        Form1.Memo2.Lines.Add('UOsnova_Main.Osnova( ДАТА НЕ ТА - '+Vn_Dat+',ЗапускВозд,СпецифВозд,'+Nom+','+fileName+' )UOsnova_Main.Flag_Error =1') ;
                        //Exit;
          end;}
          D := FormatDateTime('dd.mm.yyyy', DateTimePicker1.Date);
          KTO:='';
         Form1.SUT(D,Nom,Dir1,1,Cvet1,kto);//ZZZZZZZZZZZZZZZZZZZZZZ

         if  (UOsnova_Main.Flag_Error =0) OR (UOsnova_Main.Flag_Error =2) Then
         Begin
          if Form1.Flag_Briket =0 Then  //C брикетом БЗ
          Begin
           for I := 1 to SGSBrik.RowCount do
           begin
                if SGSBrik.Cells[16, i] <> '' then
                begin
                        if (SGSBrik.Cells[9, I] <> '') and (SGSBrik.Cells[16, i] <> '') and
                                (SGSBrik.Cells[7, i] <> '') then
                        begin
                                BZ:=SGSBrik.Cells[4, I];
                                Izdel:= SGSBrik.Cells[13, i];
                                Nam:= SGSBrik.Cells[13, i];
                                IzdelGP:=SGSBrik.Cells[29, i];
                                Res:=Pos(' ',IzdelGP);
                                Delete(IzdelGP,1,Res+1);
                                KTO:=SGSBrik.Cells[30, i];
                                idgp:=SGSBrik.Cells[16, i];
                                idko:=SGSBrik.Cells[35, i];
                                Briket:=SGSBrik.Cells[1, i];
                                Brik_ned1:=SGSBrik.Cells[31, I];
                                Brik_Den1:=SGSBrik.Cells[32, I];
                                Brik_Ned2:=SGSBrik.Cells[33, I];
                                Brik_Den2:=SGSBrik.Cells[34, I];
                                Kol_lop:=SGSBrik.Cells[36, I];

                                Nomer_Raz := '';
                                Kol_Raz := '';
                                Kol_Zap := StrToInt(SGSBrik.Cells[7, i]);
                                if  SGSBrik.Cells[15, i]='' Then
                                Kol_Pr:=0
                                else
                                Kol_pr := StrToInt(SGSBrik.Cells[15, i]);
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                NachKod:= StrToInt64(SGSBrik.Cells[25, i]);// по умолчанию =0
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                KonKod:= StrToInt64(SGSBrik.Cells[26, i]); // по умолчанию =0
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                KodPos:= StrToInt64(SGSBrik.Cells[28, i]); // по умолчанию =0
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                NachNom:= StrToInt(SGSBrik.Cells[23, i]); // по умолчанию =0
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                KonNom:= StrToInt(SGSBrik.Cells[24, i]); // по умолчанию =0
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                NomPos:= StrToInt(SGSBrik.Cells[27, i]); //НомерПос по умолчанию =1
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                if NomPos=1 Then
                                Begin
                                  NomPos:=NachNom+Kol_Zap-1;
                                  KodPos:=NachKod +Kol_Zap-1;
                                end
                                Else
                                Begin
                                  NomPos:=NomPos+Kol_Zap;
                                  KodPos:=KodPos+Kol_Zap;
                                end;

                                //+++++++++++++++++++++++++++++++++++++++++++++
                                Plan_Dat := Form1.ConvertDat1(SGSBrik.Cells[2, i]);
                                Teh_Gotov := Form1.ConvertDat1(SGSBrik.Cells[22,
                                        i]);

                                if Kol_Zap = 0 then
                                        Continue;
                                Kol_Ob := Kol_Ob + Kol_Zap;
                                Kol_Zap_R := StrToInt(SGSBrik.Cells[17, i]);
                                Kol := StrToInt(SGSBrik.Cells[5, i]);
                                Privod := (SGSBrik.Cells[14, i]);

                                Dat_Z := SGSBrik.Cells[12, i];
                                res := Pos('.', Dat_z);
                                if Res <> 0 then
                                begin
                                        SS := Copy(Dat_Z, 1, Res);
                                        Delete(Dat_Z, 1, Res);
                                        Insert(SS, Dat_z, 4);
                                end;
                                Zak := SGSBrik.Cells[3, i];
                                Dat_Zak := Dat_z;
                                Dat := FormatDateTime('mm.dd.yyyy', Now);
                                Dat1 := FormatDateTime('mm.dd.yyyy',
                                        DateTimePicker1.Date);

                                if Kol <> Kol_Zap then
                                begin
                                        if SGSBrik.Cells[18, i] = '' then
                                                Kol_Raz := IntToStr(Kol_Zap)
                                        else
                                                Kol_Raz := SGSBrik.Cells[18, i] + ','
                                                        + IntToStr(Kol_Zap);
                                        if SGSBrik.Cells[19, i] = '' then
                                                Nomer_Raz := Nom
                                        else
                                                Nomer_Raz := SGSBrik.Cells[19, i] +
                                                        ',' + Nom;
                                end;
                                if Kol = Kol_Zap then
                                begin
                                        Kol_Raz := IntToStr(Kol_Zap);
                                        Nomer_Raz := Nom;
                                end;
                                Form1.Memo2.Lines.Add('5');
                                KonNom:=NomPos;
                                KonKod:=KodPos;
                                //++++++++++++++++++++++++++++++++++++++++++++++++++++
                                if not  Form1.mkQueryUpdate(Form1.ADOQuery1,
                                'UPDATE %s SET [Кол во запущенных]=' + #39
                                + IntToStr(Kol_Zap +  Kol_Zap_R) + #39
                                + ',[Кол во запущенных пос]=' +#39 + IntToStr(Kol_Zap) + #39 +
                                ',Номер=' + #39 + Nom + #39 +
                                ',[' + FN_KOL_RAZ + ']='+ #39 + Kol_Raz + #39 + ',[' +
                                FN_NOMER_RAZ + ']=' + #39 + Nomer_Raz +#39 +
                                ',[' + 'НомерПос' + ']='+#39 + IntToStr(NomPos) + #39 +
                                ',[' + 'КодПос' + ']='+#39 + IntToStr(KodPos) + #39 + ' WHERE ([IDГП]=' + #39 +
                                IDGP + #39+') And ([IDКО]=' + #39 +IDKO + #39+')' , ['KlapanaZap'])
                                then
                                Exit;

                                //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                        end;
                end;
           end;
          end
          Else
          Begin
           for I := 1 to SGS.RowCount do
           begin
                IDGP:=SGS.Cells[10, i];
                if SGS.Cells[10, i] <> '' then
                begin
                        if (SGS.Cells[7, I] <> '') and (IDGP <> '') and
                                (SGS.Cells[5, i] <> '')
                        then
                        begin
                                BZ:=SGS.Cells[29, i];
                                Kol_Lop:=SGS.Cells[30, i];
                                IzdelGP:=SGS.Cells[27, i];
                                Res:=Pos(' ',IzdelGP);
                                Delete(IzdelGP,1,Res+1);
                                KTO:=SGS.Cells[28, i];

                                Izdel:= SGS.Cells[16, i];
                                Nam:= SGS.Cells[16, i];

                                idgp:=SGS.Cells[10, i];
                                Nomer_Raz := '';
                                Kol_Raz := '';
                                Kol_Zap := StrToInt(SGS.Cells[5, i]);
                                if  SGS.Cells[24, i]='' Then
                                Kol_Pr:=0
                                else
                                Kol_pr := StrToInt(SGS.Cells[24, i]);
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                if  SGS.Cells[22, i]='' Then
                                NachKod:=0
                                Else
                                NachKod:= StrToInt64(SGS.Cells[22, i]);
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                if  SGS.Cells[23, i]='' Then
                                KonKod:=0
                                Else
                                KonKod:= StrToInt64(SGS.Cells[23, i]);
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                if  SGS.Cells[20, i]='' Then
                                NachNom:=0
                                Else
                                NachNom:= StrToInt(SGS.Cells[20, i]);
                                //+++++++++++++++++++++++++++++++++++++++++++++

                                if  SGS.Cells[21, i]='' Then
                                KonNom:=0
                                Else
                                KonNom:= StrToInt(SGS.Cells[21, i]);
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                if  SGS.Cells[25, i]='' Then
                                NomPos:=0
                                Else
                                NomPos:= StrToInt(SGS.Cells[25, i]);
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                if NomPos=1 Then
                                Begin
                                  NomPos:=NachNom+Kol_Zap-1;
                                  KodPos:=NachKod +Kol_Zap-1;
                                end
                                Else
                                Begin
                                  NomPos:=NomPos+Kol_Zap;
                                  KodPos:=KodPos+Kol_Zap;
                                end;
                                KonNom:=NomPos;
                                KonKod:=KodPos;
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                idgp:=SGS.Cells[10, i];
                                Plan_Dat := Form1.ConvertDat1(SGS.Cells[1, i]);
                                Teh_Gotov := Form1.ConvertDat1(SGS.Cells[19,
                                        i]);
                                //Teh_Obrab := Form1.ConvertDat1(SGS.Cells[20,
                                //        i]);
                                if Kol_Zap = 0 then
                                        Continue;
                                Kol_Ob := Kol_Ob + Kol_Zap;
                                Kol_Zap_R := StrToInt(SGS.Cells[11, i]);
                                Kol := StrToInt(SGS.Cells[3, i]);
                                Privod := (SGS.Cells[14, i]);
                                SGS.Cells[11, i] := IntToStr(Kol_Zap +
                                        Kol_Zap_R);
                                SGS.Cells[4, i] := IntToStr(Kol - (Kol_Zap +
                                        Kol_Zap_R));
                                Dat_Z := SGS.Cells[15, i];
                                res := Pos('.', Dat_z);
                                if Res <> 0 then
                                begin
                                        SS := Copy(Dat_Z, 1, Res);
                                        Delete(Dat_Z, 1, Res);
                                        Insert(SS, Dat_z, 4);
                                end;
                                Zak := SGS.Cells[2, i];
                                Dat_Zak := Dat_z;
                                Dat := FormatDateTime('mm.dd.yyyy', Now);
                                Dat1 := FormatDateTime('mm.dd.yyyy',
                                        DateTimePicker1.Date);

                                if Kol <> Kol_Zap then
                                begin
                                        if SGS.Cells[12, i] = '' then
                                                Kol_Raz := IntToStr(Kol_Zap)
                                        else
                                                Kol_Raz := SGS.Cells[12, i] + ','
                                                        + IntToStr(Kol_Zap);
                                        if SGS.Cells[13, i] = '' then
                                                Nomer_Raz := Nom
                                        else
                                                Nomer_Raz := SGS.Cells[13, i] +
                                                        ',' + Nom;
                                end;
                                if Kol = Kol_Zap then
                                begin
                                        Kol_Raz := IntToStr(Kol_Zap);
                                        Nomer_Raz :=Nom;
                                end;
                                Form1.Memo2.Lines.Add('6');

                                //++++++++++++++++++++++++++++++++++++++++++++++++++++
                                if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                                                'UPDATE %s SET [Кол во запущенных]=' + #39
                                                + IntToStr(Kol_Zap +
                                                Kol_Zap_R) + #39
                                                + ',[Кол во запущенных пос]=' +
                                                #39 + IntToStr(Kol_Zap) + #39 +
                                               ',Номер=' + #39 +
                                                Nom + #39 +
                                                ',[' + FN_KOL_RAZ + ']='
                                                +
                                                #39 + Kol_Raz + #39 + ',[' +
                                                FN_NOMER_RAZ + ']=' + #39 +
                                                Nomer_Raz +#39 +
                                                ',[' + 'НомерПос' + ']='
                                                +#39 + IntToStr(NomPos) + #39 +
                                                ',[' + 'КодПос' + ']='
                                                +#39 + IntToStr(KodPos) + #39 + ' WHERE [IDГП]=' + #39 +
                                                (IDGP) + #39, ['KlapanaZap'])
                                then
                                Exit;

                                //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                        end;
                end;  
           end;
          end;
         End;

        Form1.Button28.Click;
        Close;
     End;
end;

procedure TFV1.FormClose(Sender: TObject; var Action: TCloseAction);
var
        I: Integer;
        S: string;
begin
        S := ExtractFileDir(ParamStr(0));
        Memo1.Lines.Clear;
        for I := 0 to SGS.ColCount - 1 do
                Memo1.Lines.Add(IntToStr(SGS.ColWidths[i]));
        Memo1.Lines.SaveToFile(S + '\SGS.ini');
        //+++++++++++++++++++
        Memo1.Lines.Clear;
        for I := 0 to SGSBrik.ColCount - 1 do
                Memo1.Lines.Add(IntToStr(SGSBrik.ColWidths[i]));
        Memo1.Lines.SaveToFile(S + '\SGSBrik.ini');
end;

procedure TFV1.DateTimePicker1Change(Sender: TObject);
var S,S1:string;
D,D1:TDateTime;

I:Integer;
begin
         S:=FormatDateTime('mm.dd.YYYY',DateTimePicker1.Date);
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select  max(планирование) as S from %s Where планирование<'+#39+S+#39 ,
                ['ЗапускВозд']) then
                exit;
        D:= Form1.ADOQuery1.FieldByName('S').AsDateTime;
//+++++++++++++++++++++++++++++++
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select  max(планирование) as S from %s Where планирование<'+#39+S+#39 ,
                ['Запуск']) then
                exit;
        D1:= Form1.ADOQuery1.FieldByName('S').AsDateTime;

        if D>D1 then
        Begin
        S1:=FormatDateTime('mm.dd.YYYY',D);
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select  * from %s Where планирование='+#39+S1+#39 ,
                ['ЗапускВозд']) then
                exit;
        End
        Else
        Begin
        S1:=FormatDateTime('mm.dd.YYYY',D1);
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select  * from %s Where планирование='+#39+S1+#39 ,
                ['Запуск']) then
                exit;
        End;

        I:=Form1.ADOQuery1.RecordCount;
        Cvet:= Form1.ADOQuery1.FieldByName('Cvet').AsInteger;
        case Cvet of
            0:
            begin
              Label16.Font.Color:=RGB(255,255,255);//Белый
              Label17.Font.Color:=RGB(234,239,45); //Желтый
              Cvet1:=1;
            End;
            1:
            begin
              Label16.Font.Color:=RGB(234,239,45); //Желтый
              Label17.Font.Color:=RGB(41,243,104);//Зеленый
              Cvet1:=2;
            End;
            2:
            begin
              Label16.Font.Color:=RGB(41,243,104);//Зеленый
              Label17.Font.Color:=RGB(122,156,250); //Голубой
              Cvet1:=3;
            End;
            3:
            begin
              Label16.Font.Color:=RGB(122,156,250); //Голубой
              Label17.Font.Color:=RGB(234,239,45); //Желтый
              Cvet1:=1;
            End;
        end;
end;

procedure TFV1.DateTimePicker1Enter(Sender: TObject);
begin
        Button1.Enabled := True;
end;

procedure TFV1.SGSBrikSelectCell(Sender: TObject; ACol, ARow: Integer;
  var CanSelect: Boolean);
begin
        if ACol = 7 then
        begin
                SGSBrik.Options := SGSBrik.Options + [goEditing];

        end
        else
                SGSBrik.Options := SGSBrik.Options - [goEditing];
end;

procedure TFV1.SGSBrikSetEditText(Sender: TObject; ACol, ARow: Integer;
  const Value: String);
var
        I, Kol_Zap, Kol_Zap_R, Kol: Integer;
        Svar1, Sbor1, Svar_o1, Sbor_o1: Double;
        Pdo,Kto:String;
begin
        I := ACol;

        if SGSBrik.Cells[7, ARow] = '' then
                Exit;
        //PDO:=SGSBrik.Cells[29, ARow];
        //KTO:=SGSBrik.Cells[30, ARow];
        Kol_Zap := StrToInt(SGSBrik.Cells[7, ARow]);
        Kol_Zap_R := StrToInt(SGSBrik.Cells[6, ARow]);
        Kol := StrToInt(SGSBrik.Cells[5, ARow]);
        if Kol<Kol_Zap then
        Begin
         MessageDlg('Запустить больше чем есть, не возможно!',
                                        mtError, [mbOk], 0);
         SGSBrik.Cells[7, ARow]:='0';
         Exit;
        End;
        if SGSBrik.Cells[8, ARow] = '' then
                Svar1 := 0
        else
                Svar1 := StrToFloat(SGSBrik.Cells[8, ARow]);
        if SGSBrik.Cells[9, ARow] = '' then
                Sbor1 := 0
        else
                Sbor1 := StrToFloat(SGSBrik.Cells[9, ARow]);
        Svar1 := RoundTo(Svar1 * Kol_Zap, -2);
        Sbor1 := RoundTo(Sbor1 * Kol_Zap, -2);
        SGSBrik.Cells[10, ARow] := FloatToStr(Svar1);
        SGSBrik.Cells[11, ARow] := FloatToStr(Sbor1);
        Svar_o1 := 0;
        Sbor_o1 := 0;
        for I := 1 to SGSBrik.RowCount do
        begin
                if SGSBrik.Cells[10, I] <> '' then
                begin
                        Kol_Zap := StrToInt(SGSBrik.Cells[7, ARow]);
                        Kol_Zap_R := StrToInt(SGSBrik.Cells[6, ARow]);
                        //++++++++++++++++++++++++++++++++++++++++
                      {  if ((PDO<>'') AND (KTO='')) Then
                        Begin
                                 If Kol_Zap>1 then
                                 MessageDlg('Можно запустить только 1 клапан (Первый Прогон) !',
                                        mtError, [mbOk], 0);
                                //SGS.Cells[5, ARow] := '1';
                                //Exit;
                        end;
                        //++++++++++++++++++++++++++++++++++++++++
                        if ((PDO<>'') AND (KTO<>'')) Then
                        Begin
                                 If Kol_Zap>1 then
                                 MessageDlg('Можно запустить только 1 клапан (Второй Прогон) !',
                                        mtError, [mbOk], 0);
                        end;
                        //++++++++++++++++++++++++++++++++++++++++
                        if ((PDO='') AND (KTO<>'')) Then
                        Begin
                                 MessageDlg('Спроси мастера, клапан косячный!',
                                        mtError, [mbOk], 0);
                                SGS.Cells[5, ARow] := '0';
                                Exit;
                        end; }
                        Kol := Kol_Zap_R - Kol_Zap;
                        //++++++++++++++++++++++++++++++++++++++++
                       { if Kol < 0 then
                        begin
                                MessageDlg('Можно запустить только ' +
                                        SGSBrik.Cells[6, ARow] + ' !',
                                        mtError, [mbOk], 0);
                                SGSBrik.Cells[7, ARow] := '';
                                Exit;
                        end;   }
                        if SGSBrik.Cells[7, i] = '' then
                                Kol_Zap := 0
                        else
                                Kol_Zap := StrToInt(SGSBrik.Cells[7, i]);
                        if Kol_Zap = 0 then
                        begin
                                Svar_o1 := Svar_o1 - StrToFloat(SGSBrik.Cells[8, I])*Kol_Zap;
                                Sbor_o1 := Sbor_o1 - StrToFloat(SGSBrik.Cells[9, I])*Kol_Zap;
                        end;
                        if Kol_Zap <> 0 then
                        begin
                                Svar_o1 := RoundTo(Svar_o1 +
                                        StrToFloat(SGSBrik.Cells[8, I])*Kol_Zap, -2);
                                Sbor_o1 := RoundTo(Sbor_o1 +
                                        StrToFloat(SGSBrik.Cells[9, I])*Kol_Zap, -2);
                        end;
                end;
        end;
        Label2.Caption := FloatToStr(Svar_o1);
        Label6.Caption := FloatToStr(Sbor_o1);

end;

procedure TFV1.Edt1Change(Sender: TObject);
begin
        Label4.Caption:=edt1.Text;
end;

procedure TFV1.btn2Click(Sender: TObject);
begin
        Order := ' ORDER BY [' + SGSBrik.Cells[SGSBrik.Col, 0] +
                '] ';
        btn1.Click;
end;

procedure TFV1.btn3Click(Sender: TObject);
begin
        Order := ' ORDER BY [' + SGSBrik.Cells[SGSBrik.Col, 0] +
                '] Desc';
        btn1.Click;
end;

procedure TFV1.btn1Click(Sender: TObject);
var
        I, Kol, Kol_Zap, Max,R: Integer;
        S,B,bz,UP,priv: string;
begin
        S := ExtractFileDir(ParamStr(0));
        Memo1.Lines.Clear;
        Memo1.Lines.LoadFromFile(S + '\SGS.ini');
        SGS.ColCount := Memo1.Lines.Count;
        for I := 0 to Memo1.Lines.Count - 2 do
        begin
                SGS.ColWidths[i] := StrToInt(Memo1.Lines.Strings[i]);
        end;
        for i:=0 To SGS.RowCount Do
        Begin
            SGS.Rows[i].Clear;
        End;
        //+++++++++++++++++++++++++++++++++++++++++++++++++++
        Memo1.Lines.Clear;
        Memo1.Lines.LoadFromFile(S + '\SGSBrik.ini');
        SGSBrik.ColCount := Memo1.Lines.Count;
        for I := 0 to Memo1.Lines.Count - 2 do
        begin
                SGSBrik.ColWidths[i] := StrToInt(Memo1.Lines.Strings[i]);
        end;
        for i:=0 To SGSBrik.RowCount Do
        Begin
            SGSBrik.Rows[i].Clear;
        End;
        SGSBrik.ColCount :=52;
        if Form1.Flag_Briket =0 Then  //C брикетом БЗ
        begin
                SGSBrik.Visible:=True;
                SGS.Visible:=False;
                SGSBrik.Cells[0, 0] := '№';
                SGSBrik.ColWidths[0] := 60;
                SGSBrik.Cells[1, 0] := 'Брикет';
                SGSBrik.Cells[2, 0] :=FN_PLAN_DATA;
                SGSBrik.Cells[3, 0] := FN_ZAK;
                SGSBrik.Cells[4, 0] := 'Б\з';
                SGSBrik.Cells[5, 0] := 'В заказе';
                SGSBrik.Cells[6, 0] := 'Свободно';
                SGSBrik.Cells[7, 0] := 'Кол-во на запуск';
                SGSBrik.Cells[8, 0] := FN_SVARKA_NC + ' на единицу';
                SGSBrik.Cells[9, 0] := FN_SBOR_KLAP_NC + ' на единицу';
                SGSBrik.Cells[10, 0] := 'Всего ' + FN_SVARKA_NC;
                SGSBrik.Cells[11, 0] := 'Всего ' + FN_SBOR_KLAP_NC;
                SGSBrik.Cells[12, 0] := 'Дата';
                SGSBrik.Cells[13, 0] := 'Изделие';
                SGSBrik.Cells[14, 0] := 'Privod';
                SGSBrik.Cells[15, 0] := 'Кол приводов(общее)';
                SGSBrik.Cells[16, 0] := 'IdГП';
                SGSBrik.Cells[17, 0] := 'Кол запущенных';
                SGSBrik.Cells[18, 0] := 'Кол Раз';
                SGSBrik.Cells[19, 0] := 'Номер Раз';

                SGSBrik.Cells[20, 0] := 'A';
                SGSBrik.Cells[21, 0] := 'B';
                SGSBrik.Cells[22, 0] := FN_TEHNOLOG;
                SGSBrik.Cells[23, 0] := 'НачНомер';
                SGSBrik.Cells[24, 0] := 'КонНомер';
                SGSBrik.Cells[25, 0] := 'НачКод';
                SGSBrik.Cells[26, 0] := 'КонКод';//Заводск номер из 20 ячейкй

                SGSBrik.Cells[27, 0] := 'НомерПос';
                SGSBrik.Cells[28, 0] := 'КодПос';
                SGSBrik.Cells[29, 0] := 'ИзделиеГП';
                SGSBrik.Cells[30, 0] := 'КТО';
                SGSBrik.Cells[31, 0] := 'Брик_Нед1';
                SGSBrik.Cells[32, 0] := 'Брик_Ден1';
                SGSBrik.Cells[33, 0] := 'Брик_Нед2';
                SGSBrik.Cells[34, 0] := 'Брик_Ден1';
                SGSBrik.Cells[35, 0] := 'IdКО';
                SGSBrik.Cells[36, 0] := 'РаскрЛопаток';
                SGSBrik.Cells[37, 0] := 'bz';
                SGSBrik.Cells[38, 0] := 'Сборка Примечание';
                SGSBrik.Cells[39, 0] := 'Расчетная дата готовности';
                SGSBrik.Cells[40, 0] := 'Исполнение';
                SGSBrik.Cells[41, 0] := 'Заказчик';
                //SGSBrik.Cells[42, 0] := 'Stat';
               // SGSBrik.ColWidths[17] := 0;
                SGSBrik.ColWidths[18] := 0;
                SGSBrik.ColWidths[19] := 0;
                SGSBrik.ColWidths[20] := 0;
                SGSBrik.ColWidths[21] := 0;
                SGSBrik.ColWidths[22] := 0;
                SGSBrik.ColWidths[23] := 0;
                SGSBrik.ColWidths[24] := 0;
                SGSBrik.ColWidths[25] := 0;
                SGSBrik.ColWidths[26] := 0;
                SGSBrik.ColWidths[27] := 0;
                SGSBrik.ColWidths[28] := 0;
                //SGSBrik.ColWidths[29] := 0;
                SGSBrik.ColWidths[30] := 0;
                SGSBrik.ColWidths[35] := 0;
                if Order='' then
                        Order:= ' Order By Заказ ';
                if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s Where'+
                ' ([' + FN_KOL_ZAP + ']<[' + FN_KOL + ']) AND'+
                '  ((Статус='+#39+'1'+#39+')or (СтатусФ='+#39+'1'+#39+') or (СтатусФлекс='+#39+'1'+#39+'))'+
                ' AND ([' + FN_SBOR_KLAP_NC +']<>0)'+
                ' AND (NOT [' + FN_SBOR_KLAP_NC +'] IS NULL)'+
                ' AND (NOT [Брикет] IS NULL)'+
                ' AND ([Брикет] <>'+#39+#39+')'+
                ' AND (IdКО <>'+#39+'0'+#39+')'+
                //' AND ([БЗ] <>'+#39+#39+')'                +
                //' AND (NOT [БЗ] IS NULL)'+
                ' AND ([Отмена] IS NULL) AND (Х IS NULL)'+
                ' AND (NOT [' + FN_TEHNOLOG  +'] IS NULL)'+
                        Order, ['KlapanaZap']) then
                exit;
                
                if  Form1.ADOQuery1.RecordCount=0 Then
                    SGSBrik.RowCount:=2
                  Else
                  SGSBrik.RowCount := Form1.ADOQuery1.RecordCount + 1;
                for I := 0 to Form1.ADOQuery1.RecordCount - 1 do
                begin
                        Kol := Form1.ADOQuery1.FieldByName(FN_KOL).AsInteger;
                        Kol_Zap := Form1.ADOQuery1.FieldByName(FN_KOL_ZAP).AsInteger;
                        Stat:= Form1.ADOQuery1.FieldByName('Статус').AsInteger;
                        SGSBrik.Cells[0, i + 1] :=IntToStr(I+1);
                        //++++++++++++++++++++++++++++++++++++++++++++Briket
                        SGSBrik.Cells[1, i + 1] :=
                        Form1.ADOQuery1.FieldByName('Брикет').AsString;
                        //++++++++++++++++++++++++++++++++++++++++++++Briket
                        UP:=Form1.ADOQuery1.FieldByName('Исполнение').AsString;
                        SGSBrik.Cells[2, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_PLAN_DATA).AsString;
                        SGSBrik.Cells[3, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_ZAK).AsString;
                        SGSBrik.Cells[4, i + 1] :=
                        Form1.ADOQuery1.FieldByName('БЗ').AsString;
                        SGSBrik.Cells[5, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_KOL).AsString;
                        Kol := Kol - Kol_Zap;
                        SGSBrik.Cells[6, i + 1] := IntToStr(Kol);

                        SGSBrik.Cells[8, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_SVARKA_NC).AsString;
                        SGSBrik.Cells[9, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_SBOR_KLAP_NC).AsString;
                        //+++++++++++++++++++++++++++++++++++++++++
                        SGSBrik.Cells[12, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_DAT).AsString;
                        SGSBrik.Cells[13, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_NAM).AsString;
                        Priv:= Form1.ADOQuery1.FieldByName(FN_MOD_PRIVOD).AsString ;
                        R:=Pos('Электропривод',Priv);
                        if R<>0 then
                        Priv := StringReplace(Priv, 'Электропривод ', '', [rfReplaceAll]);
                        SGSBrik.Cells[14, i + 1] := Priv;

                        SGSBrik.Cells[15, i + 1] :=
                        Form1.ADOQuery1.FieldByName('Привод').AsString;

                        SGSBrik.Cells[16, i + 1] :=
                        Form1.ADOQuery1.FieldByName('IDГП').AsString;
                        SGSBrik.Cells[17, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_KOL_ZAP).AsString;
                        SGSBrik.Cells[18, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_KOL_RAZ).AsString;
                        SGSBrik.Cells[19, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_NOMER_RAZ).AsString;

                        SGSBrik.Cells[20, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_A).AsString;
                        SGSBrik.Cells[21, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_B).AsString;
                        SGSBrik.Cells[22, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_TEHNOLOG).AsString;
                        SGSBrik.Cells[23, i + 1] :=
                        Form1.ADOQuery1.FieldByName('НачНомер').AsString;
                        SGSBrik.Cells[24, i + 1] :=
                        Form1.ADOQuery1.FieldByName('КонНомер').AsString;
                        SGSBrik.Cells[25, i + 1] :=
                        Form1.ADOQuery1.FieldByName('НачКод').AsString;
                        SGSBrik.Cells[26, i + 1] :=
                        Form1.ADOQuery1.FieldByName('КонКод').AsString;

                        SGSBrik.Cells[27, i + 1] :=
                        Form1.ADOQuery1.FieldByName('НомерПос').AsString;
                        SGSBrik.Cells[28, i + 1] :=
                        Form1.ADOQuery1.FieldByName('КодПос').AsString;

                        SGSBrik.Cells[29, I + 1] :=
                        Form1.ADOQuery1.FieldByName('ИзделиеГП').AsString;
                        SGSBrik.Cells[30 , I + 1] :=
                        Form1.ADOQuery1.FieldByName('№Поз').AsString;
                        SGSBrik.Cells[31, I + 1] := Form1.ADOQuery1.FieldByName('Брик_Нед1').AsString;
                        SGSBrik.Cells[32, I + 1] := Form1.ADOQuery1.FieldByName('Брик_Ден1').AsString;
                        SGSBrik.Cells[33, I + 1] := Form1.ADOQuery1.FieldByName('Брик_Нед2').AsString;
                        SGSBrik.Cells[34, I + 1] := Form1.ADOQuery1.FieldByName('Брик_Ден1').AsString;
                        SGSBrik.Cells[35, I + 1] := Form1.ADOQuery1.FieldByName('IdКО').AsString;
                        SGSBrik.Cells[36, I + 1] := Form1.ADOQuery1.FieldByName('РаскрЛопаток').AsString;
                        SGSBrik.Cells[37, I + 1] := Form1.ADOQuery1.FieldByName('bz').AsString;
                        SGSBrik.Cells[38, I+1] := Form1.ADOQuery1.FieldByName('Технолог Прим').AsString+' '+
                        Form1.ADOQuery1.FieldByName(FN_SBOR_PRIM).AsString;;
                        SGSBrik.Cells[39, I+1] := Form1.ADOQuery1.FieldByName('Расчетная дата готовности').AsString;
                        SGSBrik.Cells[40, I+1] :=UP;
                        SGSBrik.Cells[41, I+1] :=Form1.ADOQuery1.FieldByName('Заказчик').AsString;
                         if Form1.ADOQuery1.FieldByName('Расключение').AsString='' then
                            SGSBrik.Cells[42, I+1] :='0'
                         else
                        SGSBrik.Cells[42, I+1] :=Form1.ADOQuery1.FieldByName('Расключение').AsString;

                        if Form1.ADOQuery1.FieldByName('Сборка лопаток').AsString='' then
                            SGSBrik.Cells[43, I+1] :='0'
                         else
                        SGSBrik.Cells[43, I+1] :=Form1.ADOQuery1.FieldByName('Сборка лопаток').AsString;

                        if Form1.ADOQuery1.FieldByName('Сборка тяг').AsString='' then
                            SGSBrik.Cells[44, I+1] :='0'
                         else
                        SGSBrik.Cells[44, I+1] :=Form1.ADOQuery1.FieldByName('Сборка тяг').AsString;

                        if Form1.ADOQuery1.FieldByName('СборкаРеш').AsString='' then
                            SGSBrik.Cells[45, I+1] :='0'
                         else
                        SGSBrik.Cells[45, I+1] :=Form1.ADOQuery1.FieldByName('СборкаРеш').AsString;

                        if Form1.ADOQuery1.FieldByName('Обварка').AsString='' then
                            SGSBrik.Cells[46, I+1] :='0'
                         else
                        SGSBrik.Cells[46, I+1] :=Form1.ADOQuery1.FieldByName('Обварка').AsString;

                        if Form1.ADOQuery1.FieldByName('Электро').AsString='' then
                            SGSBrik.Cells[47, I+1] :='0'
                         else
                        SGSBrik.Cells[47, I+1] :=Form1.ADOQuery1.FieldByName('Электро').AsString;

                        Form1.ADOQuery1.Next;

                end;
        end;
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        if Form1.Flag_Briket =1 Then
        begin
                SGSBrik.Visible:=False;
                SGS.Visible:=True;
                SGS.ColCount := 52;
                SGS.Cells[0, 0] := '№';
                SGS.ColWidths[0] := 60;
                SGS.Cells[1, 0] := FN_PLAN_DATA;
                SGS.Cells[2, 0] := FN_ZAK;
                SGS.Cells[3, 0] := FN_KOL;
                SGS.Cells[4, 0] := 'Свободно';
                SGS.Cells[5, 0] := 'Кол-во на запуск';
                SGS.Cells[6, 0] := FN_SVARKA_NC + ' на единицу';
                SGS.Cells[7, 0] := FN_SBOR_KLAP_NC + ' на единицу';
                SGS.Cells[8, 0] := 'Всего ' + FN_SVARKA_NC;
                SGS.Cells[9, 0] := 'Всего ' + FN_SBOR_KLAP_NC;
                SGS.Cells[10, 0] := 'IdГП';
                SGS.Cells[11, 0] := 'Кол запущенных';
                SGS.Cells[12, 0] := 'Кол Раз';
                SGS.Cells[13, 0] := 'Номер Раз';
                SGS.Cells[14, 0] := 'Privod';
                SGS.Cells[15, 0] := 'Дата';
                SGS.Cells[16, 0] := 'Изделие';
                SGS.Cells[17, 0] := 'A';
                SGS.Cells[18, 0] := 'B';
                SGS.Cells[19, 0] := FN_TEHNOLOG;
                SGS.Cells[20, 0] := 'НачНомер';
                SGS.Cells[21, 0] := 'КонНомер';
                SGS.Cells[22, 0] := 'НачКод';
                SGS.Cells[23, 0] := 'КонКод';//Заводск номер из 20 ячейкй
                SGS.Cells[24, 0] := 'Кол приводов(общее)';
                SGS.Cells[25, 0] := 'НомерПос';
                SGS.Cells[26, 0] := 'КодПос';

                SGS.Cells[27, 0] := 'ИзделиеГП';
                SGS.Cells[28, 0] := 'КТО';
                SGS.Cells[29, 0] := 'БЗ';
                SGS.Cells[30, 0] := 'РаскрЛопаток';
                SGS.Cells[31, 0] := 'bz';
                SGS.Cells[32, 0] := 'Сборка Примечание';
                SGS.Cells[33, 0] := 'Расчетная дата готовности';
                SGS.Cells[34, 0] := 'Исполнение';
                SGS.Cells[35, 0] := 'Заказчик';
                //SGS.ColWidths[11] := 0;
               // SGS.ColWidths[12] := 0;
                //SGS.ColWidths[14] := 0;
                SGS.ColWidths[17] := 0;
                SGS.ColWidths[18] := 0;
                SGS.ColWidths[19] := 0;
                SGS.ColWidths[20] := 0;
                SGS.ColWidths[21] := 0;
                SGS.ColWidths[22] := 0;
                SGS.ColWidths[23] := 0;
                SGS.ColWidths[25] := 0;
                SGS.ColWidths[26] := 0;
               // SGS.ColWidths[27] := 0;
                SGS.ColWidths[28] := 0;

                if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s Where'+
                ' ([' + FN_KOL_ZAP + ']<[' + FN_KOL + ']) AND'+
                //' AND (NOT [' + FN_Mod_Privod +'] IS NULL)'+
                '  ((Статус='+#39+'1'+#39+') or (СтатусФ='+#39+'1'+#39+') or (СтатусФлекс='+#39+'1'+#39+'))'+
                ' AND ([' + FN_SBOR_KLAP_NC +']<>0)'+
                ' AND (NOT [' + FN_SBOR_KLAP_NC +'] IS NULL)'+
               // ' AND ([БЗ] ='+#39+#39+')'+
                ' AND ([IdКО] =0)'+
                ' AND ([Отмена] IS NULL) AND (Х IS NULL)'+
                ' AND (NOT [' + FN_TEHNOLOG  +'] IS NULL)'+
                        ' Order By Заказ ', ['KlapanaZap']) then
                exit;
                if  Form1.ADOQuery1.RecordCount=0 Then
                     SGS.RowCount :=2
                    Else
                SGS.RowCount := Form1.ADOQuery1.RecordCount + 1;

                for I := 0 to Form1.ADOQuery1.RecordCount - 1 do
                begin
                        Kol := Form1.ADOQuery1.FieldByName(FN_KOL).AsInteger;
                        Kol_Zap := Form1.ADOQuery1.FieldByName(FN_KOL_ZAP).AsInteger;
                        Stat:= Form1.ADOQuery1.FieldByName('Статус').AsInteger;
                        SGS.Cells[0, i + 1] :=IntToStr(I+1);
                        SGS.Cells[1, i + 1] := Form1.ADOQuery1.FieldByName(FN_PLAN_DATA).AsString;
                        SGS.Cells[2, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_ZAK).AsString;
                        SGS.Cells[3, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_KOL).AsString;

                        Kol := Kol - Kol_Zap;
                        SGS.Cells[4, i + 1] := IntToStr(Kol);
                        SGS.Cells[6, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_SVARKA_NC).AsString;
                        SGS.Cells[7, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_SBOR_KLAP_NC).AsString;
                        SGS.Cells[10, i + 1] :=
                        Form1.ADOQuery1.FieldByName('IdГП').AsString;
                        SGS.Cells[11, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_KOL_ZAP).AsString;
                        SGS.Cells[12, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_KOL_RAZ).AsString;
                        SGS.Cells[13, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_NOMER_RAZ).AsString;

                        Priv:= Form1.ADOQuery1.FieldByName(FN_MOD_PRIVOD).AsString ;
                        R:=Pos('Электропривод',Priv);
                        if R<>0 then
                        Priv := StringReplace(Priv, 'Электропривод ', '', [rfReplaceAll]);
                        SGS.Cells[14, i + 1] := Priv;
                       { SGS.Cells[14, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_MOD_PRIVOD).AsString; }
                        SGS.Cells[15, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_DAT).AsString;
                        SGS.Cells[16, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_NAM).AsString;

                        SGS.Cells[17, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_A).AsString;
                        SGS.Cells[18, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_B).AsString;
                        SGS.Cells[19, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_TEHNOLOG).AsString;
                        SGS.Cells[20, i + 1] :=
                        Form1.ADOQuery1.FieldByName('НачНомер').AsString;
                        SGS.Cells[21, i + 1] :=
                        Form1.ADOQuery1.FieldByName('КонНомер').AsString;
                        SGS.Cells[22, i + 1] :=
                        Form1.ADOQuery1.FieldByName('НачКод').AsString;
                        SGS.Cells[23, i + 1] :=
                        Form1.ADOQuery1.FieldByName('КонКод').AsString;
                        SGS.Cells[24, i + 1] :=
                        Form1.ADOQuery1.FieldByName('Привод').AsString;

                        SGS.Cells[25, i + 1] :=
                        Form1.ADOQuery1.FieldByName('НомерПос').AsString;
                        SGS.Cells[26, i + 1] :=
                        Form1.ADOQuery1.FieldByName('КодПос').AsString;

                        SGS.Cells[27, I + 1] :=
                        Form1.ADOQuery1.FieldByName('ИзделиеГП').AsString;
                        SGS.Cells[28 , I + 1] :=
                        Form1.ADOQuery1.FieldByName('№Поз').AsString;
                        SGS.Cells[29 , I + 1] :=
                        Form1.ADOQuery1.FieldByName('БЗ').AsString;
                        SGS.Cells[30, I + 1] := Form1.ADOQuery1.FieldByName('РаскрЛопаток').AsString;
                        SGS.Cells[31, I + 1] := Form1.ADOQuery1.FieldByName('bz').AsString;
                        SGS.Cells[32, I+1] := Form1.ADOQuery1.FieldByName('Технолог Прим').AsString+' '+
                        Form1.ADOQuery1.FieldByName(FN_SBOR_PRIM).AsString;;
                        SGS.Cells[33, I+1] := Form1.ADOQuery1.FieldByName('Расчетная дата готовности').AsString;
                        SGS.Cells[34, I+1] := Form1.ADOQuery1.FieldByName('Исполнение').AsString;
                        SGS.Cells[35, I+1] := Form1.ADOQuery1.FieldByName('Заказчик').AsString;

                        if Form1.ADOQuery1.FieldByName('Расключение').AsString='' then
                            SGS.Cells[42, I+1] :='0'
                         else
                        SGS.Cells[42, I+1] :=Form1.ADOQuery1.FieldByName('Расключение').AsString;

                        if Form1.ADOQuery1.FieldByName('Сборка лопаток').AsString='' then
                            SGS.Cells[43, I+1] :='0'
                         else
                        SGS.Cells[43, I+1] :=Form1.ADOQuery1.FieldByName('Сборка лопаток').AsString;

                        if Form1.ADOQuery1.FieldByName('Сборка тяг').AsString='' then
                            SGS.Cells[44, I+1] :='0'
                         else
                        SGS.Cells[44, I+1] :=Form1.ADOQuery1.FieldByName('Сборка тяг').AsString;

                        if Form1.ADOQuery1.FieldByName('СборкаРеш').AsString='' then
                            SGS.Cells[45, I+1] :='0'
                         else
                        SGS.Cells[45, I+1] :=Form1.ADOQuery1.FieldByName('СборкаРеш').AsString;

                        if Form1.ADOQuery1.FieldByName('Обварка').AsString='' then
                            SGS.Cells[46, I+1] :='0'
                         else
                        SGS.Cells[46, I+1] :=Form1.ADOQuery1.FieldByName('Обварка').AsString;

                        if Form1.ADOQuery1.FieldByName('Электро').AsString='' then
                            SGS.Cells[47, I+1] :='0'
                         else
                        SGS.Cells[47, I+1] :=Form1.ADOQuery1.FieldByName('Электро').AsString;
                        Form1.ADOQuery1.Next;

                end;
        end;

        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select MAX(Номер) as a from %s ',
                ['KlapanaZap']) then
                exit;
        Max := Form1.ADOQuery1.FieldValues['a'];    // Max
        Label4.Caption :=IntToStr(Max+1); // edt1.Text;
                if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select * from %s Where Номер='+#39+IntToStr(Max)+#39 ,
                ['ЗапускВозд']) then
                exit;
       { Cvet:= Form1.ADOQuery1.FieldByName('Cvet').AsInteger;
        case Cvet of
            0:
            begin
              Label16.Font.Color:=RGB(255,255,255);//Белый
              Label17.Font.Color:=RGB(234,239,45); //Желтый
              Cvet1:=1;
            End;
            1:
            begin
              Label16.Font.Color:=RGB(234,239,45); //Желтый
              Label17.Font.Color:=RGB(41,243,104);//Зеленый
              Cvet1:=2;
            End;
            2:
            begin
              Label16.Font.Color:=RGB(41,243,104);//Зеленый
              Label17.Font.Color:=RGB(122,156,250); //Голубой
              Cvet1:=3;
            End;
            3:
            begin
              Label16.Font.Color:=RGB(122,156,250); //Голубой
              Label17.Font.Color:=RGB(234,239,45); //Желтый
              Cvet1:=1;
            End;
        end; }

end;

end.
