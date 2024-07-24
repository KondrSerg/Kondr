program Klapan;

uses
  Forms,
  Main in 'Main.pas' {Form1},
  Reg in 'Reg.pas' {Regist},
  UNomer in 'UNomer.pas' {FNomer},
  UBriket in 'UBriket.pas' {FBriket},
  USpec in 'USpec.pas' {FSpec},
  UDetal in 'UDetal.pas' {FDetal},
  UNomerProizv in 'UNomerProizv.pas' {Form2},
  UPrim in 'UPrim.pas' {FPrim},
  Unit3 in 'Unit3.pas' {Form3},
  UZapusk in 'UZapusk.pas' {FZapusk},
  UV in 'UV.pas' {FV},
  URasform in 'URasform.pas' {FRasform},
  UKokZap in 'UKokZap.pas' {FKolZap},
  UNew in 'UNew.pas' {FVozKl},
  UF1 in 'UF1.pas' {FV1},
  UZapuskVozd in 'UZapuskVozd.pas' {FZapuskVozd},
  UNorm in 'UNorm.pas' {FNorm},
  UNewSbor in 'UNewSbor.pas' {FNewSbor},
  USborAll in 'USborAll.pas' {FSborAll},
  UOtk in 'UOtk.pas' {FOTK},
  UPlus in 'UPlus.pas' {FPlus},
  UNewNakl in 'UNewNakl.pas' {FNewNakl},
  uThread in 'uThread.pas',
  uSockets in 'uSockets.pas',
  UServer in 'UServer.pas' {FServer},
  UTSP in 'UTSP.pas' {FTSP},
  UPrivod in 'UPrivod.pas' {FPrivod},
  UNewNaklVoz in 'UNewNaklVoz.pas' {FNAklVoz},
  Unit4 in 'Unit4.pas' {Form4},
  UProgres in 'UProgres.pas' {FProgres},
  US in 'US.pas' {FS},
  USpecSTAM in 'USpecSTAM.pas' {FSTAM},
  UKolLop in 'UKolLop.pas' {FKolLop},
  UKolZap in 'UKolZap.pas' {Form5},
  Unit1 in 'Pusk\Unit1.pas' {FPusk},
  UNewNaklSTAM in 'UNewNaklSTAM.pas' {fNewNaklSTAM},
  USTAMTime in 'USTAMTime.pas' {FStamTime},
  UPDO in 'UPDO.pas' {FPDO},
  UProgrammist in 'UProgrammist.pas' {FProgramm},
  Unit2 in 'Unit2.pas',
  UNorma1 in 'UNorma1.pas' {FNorma},
  UNormaVoz in 'UNormaVoz.pas' {FNormaVoz},
  U710 in 'U710.pas' {F710},
  UConnCeh in '..\Osnova\UConnCeh.pas',
  UConnKlap in '..\Osnova\UConnKlap.pas',
  UOsnova in '..\Osnova\UOsnova.pas',
  USpecOb in 'USpecOb.pas' {FSpecOb},
  USborkaNCH in 'USborkaNCH.pas' {FSborNCH},
  UnewTip in 'UnewTip.pas' {FNewTip},
  UFile in 'UFile.pas' {FFile},
  mainBRAK in '..\BRAK\mainBRAK.pas' {FBRAK},
  UnewBRAK in '..\BRAK\UnewBRAK.pas' {FNewBRAK},
  UGibka in 'UGibka.pas' {FGibka},
  UNCH in 'UNCH.pas' {Fnch},
  UPochta in 'UPochta.pas' {FPochta},
  UKol in '..\BRAK\UKol.pas' {FKol},
  UPolsov in 'UPolsov.pas' {FPolsov},
  UpolsRed in 'UpolsRed.pas' {FPolsRed},
  UShtrih in '..\ШтрихКод_Рабочий\UShtrih.pas' {FShtrih},
  barcod in '..\ШтрихКод_Рабочий\barcod.pas',
  UBrigada in 'UBrigada.pas' {FBrigada},
  URabota in 'URabota.pas' {FRabota},
  USborKan in 'USborKan.pas' {FSborKan},
  barcod_EKV in 'barcod_EKV.pas',
  UShtrih_EKV in 'UShtrih_EKV.pas' {FShtrih_EKV},
  UPoisk in 'UPoisk.pas' {FPoisk};

{$R *.res}

begin
        Application.Initialize;
        Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TRegist, Regist);
  Application.CreateForm(TFNomer, FNomer);
  Application.CreateForm(TFBriket, FBriket);
  Application.CreateForm(TFSpec, FSpec);
  Application.CreateForm(TFDetal, FDetal);
  Application.CreateForm(TForm2, Form2);
  Application.CreateForm(TFPrim, FPrim);
  Application.CreateForm(TForm3, Form3);
  Application.CreateForm(TFZapusk, FZapusk);
  Application.CreateForm(TFV, FV);
  Application.CreateForm(TFRasform, FRasform);
  Application.CreateForm(TFKolZap, FKolZap);
  Application.CreateForm(TFVozKl, FVozKl);
  Application.CreateForm(TFV1, FV1);
  Application.CreateForm(TFZapuskVozd, FZapuskVozd);
  Application.CreateForm(TFNorm, FNorm);
  Application.CreateForm(TFNewSbor, FNewSbor);
  Application.CreateForm(TFSborAll, FSborAll);
  Application.CreateForm(TFOTK, FOTK);
  Application.CreateForm(TFPlus, FPlus);
  Application.CreateForm(TFNewNakl, FNewNakl);
  Application.CreateForm(TFServer, FServer);
  Application.CreateForm(TFTSP, FTSP);
  Application.CreateForm(TFPrivod, FPrivod);
  Application.CreateForm(TFNAklVoz, FNAklVoz);
  Application.CreateForm(TForm4, Form4);
  Application.CreateForm(TFProgres, FProgres);
  Application.CreateForm(TFS, FS);
  Application.CreateForm(TFSTAM, FSTAM);
  Application.CreateForm(TFKolLop, FKolLop);
  Application.CreateForm(TForm5, Form5);
  Application.CreateForm(TFPusk, FPusk);
  Application.CreateForm(TfNewNaklSTAM, fNewNaklSTAM);
  Application.CreateForm(TFStamTime, FStamTime);
  Application.CreateForm(TFPDO, FPDO);
  Application.CreateForm(TFProgramm, FProgramm);
  Application.CreateForm(TFNorma, FNorma);
  Application.CreateForm(TFNormaVoz, FNormaVoz);
  Application.CreateForm(TF710, F710);
  Application.CreateForm(TFSpecOb, FSpecOb);
  Application.CreateForm(TFSborNCH, FSborNCH);
  Application.CreateForm(TFNewTip, FNewTip);
  Application.CreateForm(TFFile, FFile);
  Application.CreateForm(TFBRAK, FBRAK);
  Application.CreateForm(TFNewBRAK, FNewBRAK);
  Application.CreateForm(TFGibka, FGibka);
  Application.CreateForm(TFGibka, FGibka);
  Application.CreateForm(TFnch, Fnch);
  Application.CreateForm(TFPochta, FPochta);
  Application.CreateForm(TFPochta, FPochta);
  Application.CreateForm(TFKol, FKol);
  Application.CreateForm(TFPolsov, FPolsov);
  Application.CreateForm(TFPolsRed, FPolsRed);
  Application.CreateForm(TFPolsRed, FPolsRed);
  Application.CreateForm(TFShtrih, FShtrih);
  Application.CreateForm(TFBrigada, FBrigada);
  Application.CreateForm(TFRabota, FRabota);
  Application.CreateForm(TFRabota, FRabota);
  Application.CreateForm(TFSborKan, FSborKan);
  Application.CreateForm(TFShtrih_EKV, FShtrih_EKV);
  Application.CreateForm(TFPoisk, FPoisk);
  Application.CreateForm(TFPoisk, FPoisk);
  Application.Run;
end.

