unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, ksTLB, ksConstTLB, ComObj;

type
  TForm1 = class(TForm)
    BtnStartCompas: TButton;
    Label1: TLabel;
    BtnExitCompas: TButton;
    BtnSetMod: TButton;
    BtnClear: TButton;
    Label4: TLabel;
    Label5: TLabel;
    MComboBox: TComboBox;
    Label6: TLabel;
    MEdit: TEdit;
    Label8: TLabel;
    D1Edit: TEdit;
    Label9: TLabel;
    D2Edit: TEdit;
    Label10: TLabel;
    Label11: TLabel;
    I1CB: TComboBox;
    I2CB: TComboBox;
    Label12: TLabel;
    Label13: TLabel;
    IP1Edit: TEdit;
    IP2Edit: TEdit;
    Label14: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    Label17: TLabel;
    DP1Edit: TEdit;
    DP2Edit: TEdit;
    Label18: TLabel;
    Label19: TLabel;
    Label20: TLabel;
    Label21: TLabel;
    HP2Edit: TEdit;
    HP1Edit: TEdit;
    procedure BtnStartCompasClick(Sender: TObject);
    procedure BtnExitCompasClick(Sender: TObject);
    procedure BtnSetModClick(Sender: TObject);
    procedure BtnClearClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  kompas: KompasObject;
  Document3D: ksDocument3D;
  Part: ksPart;
  VariableCollection: ksVariableCollection;
  VariableM: ksVariable;
  VariableD1: ksVariable;
  VariableD2: ksVariable;
  VariableI1: ksVariable;
  VariableI2: ksVariable;
  VariableIP1: ksVariable;
  VariableIp2: ksVariable;
  VariableDP1: ksVariable;
  VariableDP2: ksVariable;
  VariableHP1: ksVariable;
  VariableHP2: ksVariable;
  VariableD: ksVariable;
  EditM: Double;
  EditD1: Double;
  EditD2: Double;
  EditI1: Double;
  EditI2: Double;
  EditIP1: Double;
  EditIP2: Double;
  EditDP1: Double;
  EditDP2: Double;
  EditHP1: Double;
  EditHP2: Double;

implementation

{$R *.dfm}

procedure TForm1.BtnClearClick(Sender: TObject);
begin
    VariableCollection:=ksVariableCollection(Part.VariableCollection());
    VariableM:=ksVariable(VariableCollection.GetByName('M', TRUE, TRUE));
    VariableD1:=ksVariable(VariableCollection.GetByName('d1', TRUE, TRUE));
    VariableD2:=ksVariable(VariableCollection.GetByName('d2', TRUE, TRUE));
    VariableI1:=ksVariable(VariableCollection.GetByName('i1', TRUE, TRUE));
    VariableI2:=ksVariable(VariableCollection.GetByName('i2', TRUE, TRUE));
    VariableIP1:=ksVariable(VariableCollection.GetByName('lp1', TRUE, TRUE));
    VariableIP2:=ksVariable(VariableCollection.GetByName('lp2', TRUE, TRUE));
    VariableDP1:=ksVariable(VariableCollection.GetByName('dp1', TRUE, TRUE));
    VariableDP2:=ksVariable(VariableCollection.GetByName('dp2', TRUE, TRUE));
    VariableHP1:=ksVariable(VariableCollection.GetByName('Hp1', TRUE, TRUE));
    VariableHP2:=ksVariable(VariableCollection.GetByName('Hp2', TRUE, TRUE));
   
    EditM := 8000;
    Form1.MComboBox.ItemIndex := 11;

    EditD1 := 110;
    Form1.D1Edit.Text := '110';
    EditD2 := 110;
    Form1.D2Edit.Text := '110';

    EditI1 := 1;
    Form1.I1CB.ItemIndex := 0;
    EditI2 := 1;
    Form1.I2CB.ItemIndex := 0;
    
    EditIP1 := 20;
    Form1.IP1Edit.Text := '20';
    EditIP2 := 20;
    Form1.IP2Edit.Text := '20';

    EditDP1 := 0;
    Form1.DP1Edit.Text := '0';
    EditDP2 := 0;
    Form1.DP2Edit.Text := '0';

    EditHP1 := 0;
    Form1.HP1Edit.Text := '0';
    EditHP2 := 0;
    Form1.HP2Edit.Text := '0';
    
    VariableM.value := EditM;
    VariableD1.value := EditD1;
    VariableD2.value := EditD2;
    VariableI1.value := EditI1;
    VariableI2.value := EditI2;
    VariableIP1.value := EditIP1;
    VariableIP2.value := EditIP2;
    VariableDP1.value := EditDP1;
    VariableDP2.value := EditDP2;
    VariableHP1.value := EditHP1;
    VariableHP2.value := EditHP2;
    
    Part.RebuildModel();
    Document3D.RebuildDocument();
end;

procedure TForm1.BtnExitCompasClick(Sender: TObject);
begin
  kompas.Quit();
  Form1.Label1.Caption:= 'Компас выключен';
end;

procedure TForm1.BtnSetModClick(Sender: TObject);
begin
    VariableCollection:=ksVariableCollection(Part.VariableCollection());
    VariableM:=ksVariable(VariableCollection.GetByName('M', TRUE, TRUE));
    VariableD1:=ksVariable(VariableCollection.GetByName('d1', TRUE, TRUE));
    VariableD2:=ksVariable(VariableCollection.GetByName('d2', TRUE, TRUE));
    VariableI1:=ksVariable(VariableCollection.GetByName('i1', TRUE, TRUE));
    VariableI2:=ksVariable(VariableCollection.GetByName('i2', TRUE, TRUE));
    VariableIP1:=ksVariable(VariableCollection.GetByName('lp1', TRUE, TRUE));
    VariableIP2:=ksVariable(VariableCollection.GetByName('lp2', TRUE, TRUE));
    VariableDP1:=ksVariable(VariableCollection.GetByName('dp1', TRUE, TRUE));
    VariableDP2:=ksVariable(VariableCollection.GetByName('dp2', TRUE, TRUE));
    VariableHP1:=ksVariable(VariableCollection.GetByName('Hp1', TRUE, TRUE));
    VariableHP2:=ksVariable(VariableCollection.GetByName('Hp2', TRUE, TRUE));
   
    EditM := StrToFloat(Form1.MComboBox.Text);
    begin
    if not (StrToFloat(Form1.MEdit.Text) = 0) then
    begin
        EditM := StrToFloat(Form1.MEdit.Text);
    end
    end;

    EditD1 := StrToFloat(Form1.D1Edit.Text);
    EditD2 := StrToFloat(Form1.D2Edit.Text);

    begin
    if (Form1.I1CB.ItemIndex = -1) or (Form1.I1CB.ItemIndex = 0) then
    begin
        EditI1 := 1;
    end
    else if Form1.I1CB.ItemIndex = 1 then
    begin
        EditI1 := 2;
    end
    else if Form1.I1CB.ItemIndex = 2 then
    begin
        EditI1 := 3;
    end
    else if Form1.I1CB.ItemIndex = 3 then
    begin
        EditI1 := 4;
    end
    else 
    begin
        EditI1 := 1;
    end
    end;


    begin
    if (Form1.I2CB.ItemIndex = -1) or (Form1.I2CB.ItemIndex = 0) then
    begin
        EditI2 := 1;
    end
    else if Form1.I2CB.ItemIndex = 1 then
    begin
        EditI2 := 2;
    end
    else if Form1.I2CB.ItemIndex = 2 then
    begin
        EditI2 := 3;
    end
    else if Form1.I2CB.ItemIndex = 3 then
    begin
        EditI2 := 4;
    end
    else 
    begin
        EditI2 := 1;
    end
    end;
    
    EditIP1 := StrToFloat(Form1.IP1Edit.Text);
    EditIP2 := StrToFloat(Form1.IP2Edit.Text);

    EditDP1 := StrToFloat(Form1.DP1Edit.Text);
    EditDP2 := StrToFloat(Form1.DP2Edit.Text);

    EditHP1 := StrToFloat(Form1.HP1Edit.Text);
    EditHP2 := StrToFloat(Form1.HP2Edit.Text);
    
    VariableM.value := EditM;
    VariableD1.value := EditD1;
    VariableD2.value := EditD2;
    VariableI1.value := EditI1;
    VariableI2.value := EditI2;
    VariableIP1.value := EditIP1;
    VariableIP2.value := EditIP2;
    VariableDP1.value := EditDP1;
    VariableDP2.value := EditDP2;
    VariableHP1.value := EditHP1;
    VariableHP2.value := EditHP2;
    
    Part.RebuildModel();
    Document3D.RebuildDocument();
end;

procedure TForm1.BtnStartCompasClick(Sender: TObject);
begin
  try
    kompas := KompasObject(CreateOleObject('Kompas.Application.5'));
    kompas.Visible := true;
    Form1.Label1.Caption := 'Подключение произошло';
    Document3D:=ksDocument3D(kompas.Document3D());
    Document3D.Open('H:\Work\Projects\sapr\МУВП.a3d', FALSE);
    Part:=ksPart(Document3D.GetPart(-1));
  except
    Form1.Label1.Caption := 'Ошибка подключения к компасу';
  end;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  Form1.MComboBox.AddItem('6,3', nil);
  Form1.MComboBox.AddItem('16', nil);
  Form1.MComboBox.AddItem('31,5', nil);
  Form1.MComboBox.AddItem('63', nil);
  Form1.MComboBox.AddItem('125', nil);
  Form1.MComboBox.AddItem('250', nil);
  Form1.MComboBox.AddItem('500', nil);
  Form1.MComboBox.AddItem('710', nil);
  Form1.MComboBox.AddItem('1000', nil);
  Form1.MComboBox.AddItem('2000', nil);
  Form1.MComboBox.AddItem('4000', nil);
  Form1.MComboBox.AddItem('8000', nil);
  Form1.MComboBox.AddItem('16000', nil);

  Form1.I1CB.AddItem('Цилиндрический длинный конец', nil);
  Form1.I1CB.AddItem('Цилиндрический короткий конец', nil);
  Form1.I1CB.AddItem('Конический длинный конец', nil);
  Form1.I1CB.AddItem('Конический короткий конец', nil);

  Form1.I2CB.AddItem('Цилиндрический длинный конец', nil);
  Form1.I2CB.AddItem('Цилиндрический короткий конец', nil);
  Form1.I2CB.AddItem('Конический длинный конец', nil);
  Form1.I2CB.AddItem('Конический короткий конец', nil);
end;

end.




