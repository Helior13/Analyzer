unit unMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, ShellApi, Grids, jpeg, Buttons, XPMan;

type
  TForm1 = class(TForm)
    Image1: TImage;
    btn_openfile: TButton;
    XPManifest1: TXPManifest;
    OpenDialog: TOpenDialog;
    Panel1: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    StringGrid: TStringGrid;
    Label10: TLabel;
    Shape1: TShape;
    Memo: TMemo;
    procedure FormCreate(Sender: TObject);
    procedure LoadMatrix(filepath:string);
    procedure Calculation;
    procedure Analisys;
    procedure btn_openfileClick(Sender: TObject);
    procedure FormMouseWheelDown(Sender: TObject; Shift: TShiftState;
      MousePos: TPoint; var Handled: Boolean);
    procedure FormMouseWheelUp(Sender: TObject; Shift: TShiftState;
      MousePos: TPoint; var Handled: Boolean);
    procedure StringGridDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure MemoClick(Sender: TObject);
  private
    { Private declarations }
  public
  procedure wmdropfiles(var message: tmessage); message wm_dropfiles;
    { Public declarations }
  end;
const MM = 99;

var
  Form1: TForm1;
  Matrix: array [0..MM,0..MM] of real; // Матрица ответов испытуемых
  n,m: integer;                        // Количество испытуемых и заданий
  R_ave,D,SKO: real;                   // Среднее значение достижений Дисперсия и СКО
  K_difficulty: array [0..MM] of real; // Коэффициенты сложности
  K_relyability: real;                 // Коэффициент надежности расщепленных частей теста
  K_korr_sum: array [0..MM] of real;   // Суммарные коэффициенты корреляции заданий
  Dif_t: real;                         // Дифференцирующая способность теста
  Dif: array [0..MM] of real;          // Дифференцирующая способность rf;ljuj pflfybz


implementation

{$R *.dfm}
procedure TForm1.wmdropfiles(var message: tmessage);
 var
 hdrop:thandle;
 i,c:longword;
 s:string;
 begin
 hdrop:=message.wparam;
 c:=dragqueryfile(hdrop,longword(-1),pchar(nil),0);
 for i:=0 to c-1 do
 begin
 setlength(s,dragqueryfile(hdrop,i,pchar(nil),0));
 dragqueryfile(hdrop,i,pchar(s),length(s)+1);
 //ShowMessage(s);
 If Pos('.csv',LowerCase(s))<>0 then
    Loadmatrix(s) else
                  Showmessage('Неправильный формат файла!');
 end;
 dragfinish(hdrop);
 end;

procedure TForm1.FormCreate(Sender: TObject);
begin
DragAcceptFiles(Handle, true);
Height:=310;
Width:=370;
end;


Procedure Tform1.LoadMatrix(Filepath:string);
var
  sdata, srow: TStrings;
  i,j:integer;
  MaxMark: array [0..MM] of real;
begin
// Заполнение матрицы ответов
  sdata:=TStringList.Create;
  srow:=TStringList.Create;
  srow.Delimiter:=',';
  sdata.LoadFromFile(Filepath);
  n:=sdata.Count-4;

  // на случай если максимальный балл не не равен 1
  srow.DelimitedText:=sdata[0];
  m:=srow.Count-9;
  For j:=0 to m-1 do MaxMark[j]:=StrToFloat(Copy(srow[j+8],Pos('/',srow[j+8])+1,4));

  for i:=0 to n-1 do begin
      srow.DelimitedText:=sdata[i+1];
      For j:=0 to m-1 do   // заполнение и нормализация
          If srow[j+8]='-' then Matrix[i,j]:=0 else
             If StrToFloat(srow[j+8])<0 then Matrix[i,j]:=0 else Matrix[i,j]:=StrToFloat(srow[j+8])/MaxMark[j];
  end;

// Вычисление достижения тестируемого
 for i:=0 to n-1 do begin
                    for j:=0 to m-1 do Matrix[i,m]:=Matrix[i,m]+Matrix[i,j];
                    Matrix[i,m]:=Matrix[i,m]/m;
                    end;
// Вычисление среднего значения по заданию
 for j:=0 to m-1 do begin
                    for i:=0 to n-1 do Matrix[n,j]:=Matrix[n,j]+Matrix[i,j];
                    Matrix[n,j]:=Matrix[n,j]/n;
                    end;
  srow.Free;
  sdata.Free;


Calculation;
Label2.Caption:='Анализируемый файл результатов тестирования: '+ExtractFileName(Filepath);
Analisys;
end;

Procedure TForm1.Calculation;
Type TSum = (E ,O ,ExO);
Var i,j,k:integer;
    S_even,S_odd,Sort: array [0..MM] of real;
    K_korr: array [0..MM,0..MM] of real;
    Px,Py,Pxy:Real;


function ArrSum(Summ: TSum; degree: byte = 1):real;
Var i:integer;
begin
Result:=0;
Case Summ of
     E: For i:=0 to n-1 do If degree = 1 then Result:=Result+S_even[i] else
                                              Result:=Result+S_even[i]*S_even[i];
     O: For i:=0 to n-1 do If degree = 1 then Result:=Result+S_odd[i] else
                                              Result:=Result+S_odd[i]*S_odd[i];
   ExO: For i:=0 to n-1 do Result:=Result+S_even[i]*S_odd[i];
     end;
end;

Begin
// Инициализация переменных
R_ave:=0;D:=0;SKO:=0;
For i:=0 to MM do begin S_even[i]:=0; S_odd[i]:=0; Sort[i]:=0 end;
For j:=0 to m-1 do K_korr_sum[j]:=0;

// Вычисление статистических параметров
For i:=0 to n-1 do R_ave:=R_ave+Matrix[i,m];
R_ave:=R_ave/n;
For i:=0 to n-1 do D:=D+(Matrix[i,m]-R_ave)*(Matrix[i,m]-R_ave);
D:=D/(n-1);
SKO:=Sqrt(D);

// Вычисление коэффициентов сложности
For j:=0 to m do K_difficulty[j]:=1-Matrix[n,j];

// Вычисление коэффициента надежности расщепленных частей теста

For i:=0 to n-1 do begin
                   j:=0;
                   While j<(m-2)  do begin
                                     S_odd[i]:=S_odd[i]+Matrix[i,j];
                                     S_even[i]:=S_even[i]+Matrix[i,j+1];
                                     j:=j+2;
                                     end;
                   If Odd(m) then  begin
                                   S_odd[i]:=S_odd[i]+Matrix[i,j];
                                   S_odd[i]:=S_odd[i]/((m div 2)+1);
                                   end else
                                           S_odd[i]:=S_odd[i]/(m div 2);
                   S_even[i]:=S_even[i]/(m div 2);
                   end;
K_relyability:=(n*ArrSum(ExO)-ArrSum(E)*ArrSum(O))/Sqrt((n*ArrSum(E,2)-ArrSum(E)*ArrSum(E))*(n*ArrSum(O,2)-ArrSum(O)*ArrSum(O)));
K_relyability:=2*K_relyability/(1+K_relyability);


// Определение степени корреляции тестовых заданий между собой
  // Проведение дефузификации
For i:=0 to n-1 do For j:=0 to m-1 do If Matrix[i,j]<0.5 then Matrix[i,j]:=0 else Matrix[i,j]:=1;
  // Вычисление коэффициентов корреляции заданий
For k:=0 to m-1 do
    For j:=0 to m-1 do If k=j then K_korr[j,k]:=1 else
        begin
        Px:=0;Py:=0;Pxy:=0;
        For i:=0 to n-1 do begin
                         If (Matrix[i,j]=Matrix[i,k]) and (Matrix[i,j]=1) then Pxy:=Pxy+1;
                         If Matrix[i,j]=1 then Px:=Px+1;
                         If Matrix[i,k]=1 then Py:=Py+1;
                         end;
        Pxy:=Pxy/n;Px:=Px/n;Py:=Py/n;
        If (Pxy-Px*Py)=0 then K_korr[j,k]:=0 else K_korr[j,k]:=(Pxy-Px*Py)/Sqrt(Px*Py*(1-Px)*(1-Py));
        end;
  // Вычисление суммарных коэффициентов корреляции
For j:=0 to m-1 do For k:=0 to m-1 do K_korr_sum[j]:=K_korr_sum[j]+K_korr[j,k];


// Расчет дифференцирующей способности теста
For i:=0 to n-1 do Sort[i]:=Matrix[i,m];
For i:=0 to n-1 do                                      // Сортировка по достижениям
     for j:=i to n-1 do If Sort[j]>Sort[i] then begin
                                          Pxy:=Sort[j];
                                          Sort[j]:=Sort[i];
                                          Sort[i]:=Pxy;
                                          end;
Px:=0;Py:=0;j:=n div 2;
For i:=0 to j-1 do Px:=Px+Sort[i];
Px:=Px/j;
For i:=j to n-1 do Py:=Py+Sort[i];
Py:=Py/(n-j);
Dif_t:=Px-Py;

// Расчет дифференцирующей способности каждого задания

For j:=0 to m-1 do begin
    Px:=0;Py:=0;k:=0;
    For i:=0 to n-1 do If Matrix[i,j]=1 then begin
                          Px:=Px+Matrix[i,m];
                          k:=k+1;
                          end else begin
                                  Py:=Py+Matrix[i,m];
                                  end;
    If (k=n) or (k=0) then Dif[j]:=0 else
        begin
        Px:=Px/k;
        Py:=Py/(n-k);
        Dif[j]:=((Px-Py)/SKO)*sqrt(k*(n-k)/(n*(n-1)));
        end;
    end;

end;

Procedure TForm1.Analisys;
Var i,j: integer;
    sh,sl: string;
    pp:real; //процент попадания достижений в интервал

function LinesVisible(Memo: TMemo): integer;
var
  OldFont : HFont;
  Hand : THandle;
  TM : TTextMetric;
  Rect : TRect;
  tempint : integer;
begin
  Hand := GetDC(Memo.Handle);
  try
    OldFont := SelectObject(Hand, Memo.Font.Handle);
    try
      GetTextMetrics(Hand, TM);
      Memo.Perform(EM_GETRECT, 0, longint(@Rect));
      tempint := (Rect.Bottom - Rect.Top) div
      (TM.tmHeight + TM.tmExternalLeading);
    finally
      SelectObject(Hand, OldFont);
    end;
  finally
    ReleaseDC(Memo.Handle, Hand);
  end;
  Result := tempint;
end;

begin
Label4.Caption:='Среднее значение достижений испытуемых:   '+FloatToStrF(R_ave,ffFixed,6,3);
Label5.Caption:='Дисперсия достижений испытуемых:   '+FloatToStrF(D,ffFixed,6,3);
Label6.Caption:='Среднеквадратическое отклонение достижений испытуемых:   '+FloatToStrF(SKO,ffFixed,6,3);
Label7.Caption:='Показатель дифференцирующей способности теста:   '+FloatToStrF(Dif_t,ffFixed,6,3);
Label8.Caption:='Коэффициент надежности теста:   '+FloatToStrF(K_relyability,ffFixed,6,3);

StringGrid.Cells[0,0]:='№ Задания';
StringGrid.Cells[1,0]:='K сложности';
StringGrid.Cells[2,0]:='K корреляции';
StringGrid.Cells[3,0]:='Дифф. способность';
StringGrid.RowCount:=m+1;
StringGrid.Height:=StringGrid.RowCount*(StringGrid.DefaultRowHeight+1)+1;
For j:=1 to m do StringGrid.Cells[0,j]:=IntToStr(j);
For j:=0 to m-1 do StringGrid.Cells[1,j+1]:=FloatToStrF(K_difficulty[j],ffFixed,6,3);
For j:=0 to m-1 do StringGrid.Cells[2,j+1]:=FloatToStrF(K_korr_sum[j],ffFixed,6,3);
For j:=0 to m-1 do StringGrid.Cells[3,j+1]:=FloatToStrF(Dif[j],ffFixed,7,3);

Label10.Top:=StringGrid.Height+440;
Memo.Top:=Label10.Top+30;

//Анализ распределения
pp:=0;
For i:=0 to n-1 do If ((R_ave-3*SKO)<Matrix[i,m]) and (Matrix[i,m]<(R_ave+3*SKO)) then pp:=pp+1;
pp:=Round(pp/n*100);
Memo.Lines.Add( '   '+FloatToStr(pp)+'% значений достижений испытуемых попадают в интервал от -3 СКО до +3 СКО, следовательно распределение достижений ');
If pp>75 then Memo.Lines.Strings[Memo.Lines.Count-1]:=Memo.Lines.Strings[Memo.Lines.Count-1]+'близко к нормальному, что свидетельствует о достоверности исследования.'
   else Memo.Lines.Strings[Memo.Lines.Count-1]:=Memo.Lines.Strings[Memo.Lines.Count-1]+'не является нормальным, что свидетельствует о недостоверности исследований.';
Memo.Lines.Add('');

//Анализ дифференцирующей способности
Memo.Lines.Add('   Дифференцирующая способность всего теста ');
If Dif_t>0.35 then Memo.Lines.Strings[Memo.Lines.Count-1]:=Memo.Lines.Strings[Memo.Lines.Count-1]+'отличная.' else
   If Dif_t>0.25 then Memo.Lines.Strings[Memo.Lines.Count-1]:=Memo.Lines.Strings[Memo.Lines.Count-1]+'хорошая.' else
      If Dif_t>0.15 then Memo.Lines.Strings[Memo.Lines.Count-1]:=Memo.Lines.Strings[Memo.Lines.Count-1]+'удовлетворительная.' else
         Memo.Lines.Strings[Memo.Lines.Count-1]:=Memo.Lines.Strings[Memo.Lines.Count-1]+'неудовлетворительная.';

Memo.Lines.Add('');

//Анализ надежности теста
Memo.Lines.Add('   Надежность теста ');
If K_relyability>0.9 then Memo.Lines.Strings[Memo.Lines.Count-1]:=Memo.Lines.Strings[Memo.Lines.Count-1]+'отличная.' else
   If K_relyability>0.8 then Memo.Lines.Strings[Memo.Lines.Count-1]:=Memo.Lines.Strings[Memo.Lines.Count-1]+'хорошая.' else
      If K_relyability>0.7 then Memo.Lines.Strings[Memo.Lines.Count-1]:=Memo.Lines.Strings[Memo.Lines.Count-1]+'удовлетворительная.' else
         Memo.Lines.Strings[Memo.Lines.Count-1]:=Memo.Lines.Strings[Memo.Lines.Count-1]+'неудовлетворительная.';

Memo.Lines.Add('');

//Анализ заданий
With Memo.Lines do
   begin
   Add('   Следует обратить внимание на следующие задания:');
   Add('');
   sh:='';sl:='';
   For j:=0 to m-1 do If K_difficulty[j]<0.17 then sl:=sl+IntToStr(j+1)+', ' else If K_difficulty[j]>0.83 then sh:=sh+IntToStr(j+1)+', ';
   If sl<>'' then
      begin
      Add('   С низким коэффициентом сложности: '+Copy(sl,0,length(sl)-2)+'.');
      Add('Рекомендуется усложнить вышеуказанные задания, возможно они слишком просты и не позволяют адекватно оценивать знания.');
      Add('');
      end;
   If sh<>'' then
      begin
      Add('   C высоким коэффициентом сложности: '+Copy(sh,0,length(sh)-2)+'.');
      Add('Возможно, вышеуказанные задания составлены неверно или сложны для понимания. Рекомендуется их упростить.');
      Add('');
      end;
   sl:='';
   For j:=0 to m-1 do If K_korr_sum[j]<1.2 then sl:=sl+IntToStr(j+1)+', ';
   If sl<>'' then
      begin
      Add('   Задания: '+Copy(sl,0,length(sl)-2));
      Add('имеют низкий коэффициент корреляции. Они слабо связаны с другими заданиями теста. Возможно, при их разработке были допущены просчеты или они относятся к другой предметной области.');
      Strings[Count-1]:=' Рекомендуется исключить вышеуказанные задания из теста или кардинально их изменить.';
      Add('');
      end;
   sl:='';sh:='';
   For j:=0 to m-1 do If Dif[j]<0 then sh:=sh+IntToStr(j+1)+', ' else If Dif[j]<0.3 then sl:=sl+IntToStr(j+1)+', ';
   If sl<>'' then
      begin
      Add('   Задания: '+Copy(sl,0,length(sl)-2));
      Add('имеют низкую дифференцирующую способность. Рекомендуется их изменить.');
      Add('');
      end;
   If sh<>'' then
      begin
      Add('   Задания '+Copy(sh,0,length(sh)-2));
      Add('полностью неспособны дифференцировать испытуемых на сильных и слабых. Рекомендуется исключить их из теста.');
      Add('');
      end;

   If Pos('Следует',Strings[Count-2])<>0 then Strings[Count-2]:='';
   end;



//ShowMessage(IntToStr(LinesVisible(Memo))+IntToStr(Memo.Lines.Count));
While LinesVisible(Memo)<Memo.Lines.Count+1 do Memo.Height:=Memo.Height+10;
Panel1.Height:=Memo.Top+Memo.Height+30;

Panel1.Show;
Image1.Hide;
VertScrollBar.Range:=Panel1.Height;
Perform(WM_SETREDRAW, 0, 0);
Height:=840;
Width:=820;
Panel1.Width:=805;
Constraints.MaxHeight:=Panel1.Height+Height-ClientHeight;
Position:=poScreenCenter;
Perform(WM_SETREDRAW, 1, 0);
end;


procedure TForm1.btn_openfileClick(Sender: TObject);
begin
If OpenDialog.Execute then LoadMatrix(OpenDialog.FileName);
end;

procedure TForm1.FormMouseWheelDown(Sender: TObject; Shift: TShiftState;
  MousePos: TPoint; var Handled: Boolean);
begin
VertScrollBar.Position:=VertScrollBar.Position+10;
end;

procedure TForm1.FormMouseWheelUp(Sender: TObject; Shift: TShiftState;
  MousePos: TPoint; var Handled: Boolean);
begin
VertScrollBar.Position:=VertScrollBar.Position-10;
end;



procedure TForm1.StringGridDrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
var

  Buf: array[byte] of char;
begin
  if gdFixed in State then
    Exit;

  with StringGrid do

  begin
    Canvas.Font := Font;
    Canvas.Font.Color := clWindowText;
    Canvas.Brush.Color := clWindow;

    Case Acol of                          // Граничные значения параметров заданий
       1:If (StrToFloat(Cells[ACol,ARow])<0.17) or (StrToFloat(Cells[ACol,ARow])>0.83) then Canvas.Brush.Color:=clRed;
       2:If (StrToFloat(Cells[ACol,ARow])<1.2) then Canvas.Brush.Color:=clRed;
       3:If (StrToFloat(Cells[ACol,ARow])<0.3) then Canvas.Brush.Color:=clRed;
       end;
    Canvas.FillRect(Rect);
    StrPCopy(Buf, Cells[ACol, ARow]);
    DrawText(Canvas.Handle, Buf, -1, Rect,
      DT_SINGLELINE or DT_VCENTER or DT_NOCLIP or DT_LEFT);

  end;
end;



procedure TForm1.MemoClick(Sender: TObject);
begin
HideCaret(Memo.Handle);
end;

end.

