unit bgMain;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,
  ComCtrls, ImgList, Buttons, ShellApi, StdCtrls, ExtCtrls;

const
  WM_TRAYNOTIFY = WM_USER + 8012;

type

  LONG = Longword;

  PChrAttrs = ^TChrAttrs;
  TChrAttrs = record
    Str: BYTE;
    Dex: BYTE;
    Con: BYTE;
    Int: BYTE;
    Wis: BYTE;
    Cha: BYTE;
    Cri: BYTE;
  end;

  PBgData = ^TBgData;
  TBgData = record
    GameType: BYTE;
    Points: BYTE;
    CriPoints: BYTE;
    CriStop: BYTE;
    Delay: LONG;
    Address: LONG;
  end;

  TNotifyIconData = record
    cbSize: DWORD;
    hWnd: HWND;
    uID: UINT;
    uFlags: UINT;
    uCallbackMessage: UINT;
    hIcon: HICON;
    szTip: array[0..63] of AnsiChar;
  end;

  TFrmBg = class(TForm)
    ImageListT: TImageList;
    GbSettings: TGroupBox;
    LabelPoints: TLabel;
    LabelTimes: TLabel;
    LabelCurPoints: TLabel;
    EditPoints: TEdit;
    CbBTimes: TComboBox;
    EditCurPoints: TEdit;
    PanelAbout: TPanel;
    ImageTips: TImage;
    lblAbout: TLabel;
    bevelSplit: TBevel;
    gbGt: TGroupBox;
    BtnFind: TBitBtn;
    BtinOk: TBitBtn;
    BtnExit: TBitBtn;
    lblTips: TLabel;
    lblVersion: TLabel;
    imgNote: TImage;
    lvGt: TListView;
    StatusBar: TStatusBar;
    procedure FormCreate(Sender: TObject);
    procedure BtnExitClick(Sender: TObject);
    procedure EditChange(Sender: TObject);
    procedure EditPress(Sender: TObject; var Key: Char);
    procedure EditDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure BtnFindClick(Sender: TObject);
    procedure BtnOkClick(Sender: TObject);
    procedure BtnNoteClick(Sender: TObject);
    procedure lvGtSelectItem(Sender: TObject; Item: TListItem;
      Selected: Boolean);
  private
    { Private declarations }
    { attributes }
    BgData: TBgData;
    Nid: TNotifyIconData;
    procedure AddNotifyIcon;
    procedure IconNotify(var Msg: TMessage); message WM_TRAYNOTIFY;
    { operations }
  end;

var
  FrmBg: TFrmBg;

function EnableHotKeyHook: BOOL; external 'Bgrun.dll';
function DisableHotKeyHook: BOOL; external 'Bgrun.dll';
function SetData(pData: PBgData): TBgData; stdcall; external 'Bgrun.dll';
function Search(pAttrs: PChrAttrs): TChrAttrs; stdcall; external 'Bgrun.dll';
implementation

{$R *.DFM}

procedure TFrmBg.FormCreate(Sender: TObject);
begin
  EditPoints.Tag := 108;
  BgData.GameType := 2;
  BgData.Points := 0;
  BgData.CriPoints := 0;
  BgData.Delay := 1;
  BgData.Address := 0;
  BgData.CriStop := 0;
  EditPoints.Text := '0';
  EditCurPoints.Text := '0';
  CbBTimes.ItemIndex := 0;
  ImageListT.GetIcon(BgData.GameType, Self.Icon);
  ImageListT.GetIcon(BgData.GameType, ImageTips.Picture.Icon);
end;

procedure TFrmBg.BtnExitClick(Sender: TObject);
begin
  Self.Close;
end;

procedure TFrmBg.EditChange(Sender: TObject);
var
Edit: TEdit;
begin
  if not (Sender is TEdit) then
    exit;
  Edit := (Sender as TEdit);
  if(Length(Edit.Text) > 0) then
  begin
    if(StrToInt(Edit.Text) > Edit.Tag) then
      Edit.Text := IntToStr(Edit.Tag);
  end;
end;

procedure TFrmBg.EditPress(Sender: TObject; var Key: Char);
begin
  if not (Key in ['0'..'9']) and (Key <> #8) then
    Key := #0;
end;

procedure TFrmBg.EditDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
Edit: TEdit;
begin
  if not (Sender is TEdit) then
    exit;
  Edit := (Sender as TEdit);
  if(Length(Edit.Text) > 0) then
  begin
    if(StrToInt(Edit.Text) >= Edit.Tag) then
      Edit.Text := '';
  end;
end;

procedure TFrmBg.BtnFindClick(Sender: TObject);
var
  ca: TChrAttrs;
  attrs: TChrAttrs;
  al: TStrings;
  s: String;
begin
  StatusBar.Panels[0].Text := '查找中...';
  StatusBar.Panels[1].Text := '空';
  attrs.Str := 0;
  s := '';
  al := TStringList.Create;
  try
    ExtractStrings([','], [], PChar(EditCurPoints.Text), al);
    if al.Count = 7 then
    begin
      attrs.Str := StrToInt(al[0]);
      attrs.Cri := StrToInt(al[1]);
      attrs.Dex := StrToInt(al[2]);
      attrs.Con := StrToInt(al[3]);
      attrs.Int := StrToInt(al[4]);
      attrs.Wis := StrToInt(al[5]);
      attrs.Cha := StrToInt(al[6]);
      s := '* '
    end;
  finally
    al.Free;
  end;
  Screen.Cursor := crHourglass;
  SetData(@BgData);
  ca := Search(@attrs);
  Screen.Cursor := crDefault;
  StatusBar.Panels[0].Text := '查找结果: ';
  StatusBar.Panels[1].Text := s + '力量: ' + IntToStr(ca.Str) +
    ', 敏捷: ' + IntToStr(ca.Dex) + ', 体格: ' + IntToStr(ca.Con) +
    ', 智力: ' + IntToStr(ca.Int) + ', 智慧: ' + IntToStr(ca.Wis) +
    ', 魅力: ' + IntToStr(ca.Cha);
end;

procedure TFrmBg.BtnOkClick(Sender: TObject);
begin
  AddNotifyIcon;
  if Length(EditPoints.Text) > 0 then
  begin
    BgData.Points := StrToInt(EditPoints.Text);
    BgData.Delay := 1000 div StrToInt(CbBTimes.Items[CbBTimes.ItemIndex]);
    BgData.CriStop := 0;
    if EnableHotKeyHook then
      SetData(@BgData);
  end;
end;

procedure TFrmBg.BtnNoteClick(Sender: TObject);
var
NoteFName: String;
begin
  NoteFName := ExtractFilePath(Application.ExeName);
  if(Length(NoteFName) > 0) then
  begin
    if(NoteFName[Length(NoteFName)] <> '\') then
      NoteFName := NoteFName + '\';
    NoteFName := NoteFName + '\' + 'Readme.txt';
    if FileExists(NoteFName) then
      ShellExecute(Self.Handle, PChar('open'), PChar(NoteFName), nil, nil,
        SW_NORMAL);
  end;
end;

procedure TFrmBg.lvGtSelectItem(Sender: TObject; Item: TListItem;
  Selected: Boolean);
begin
  ImageListT.GetIcon(Item.Index, ImageTips.Picture.Icon);
  BgData.GameType := Item.Index;
  StatusBar.Panels[0].Text := '当前游戏:';
  StatusBar.Panels[1].Text := Item.Caption;
end;

procedure TFrmBg.IconNotify(var Msg: TMessage);
begin
  if(Msg.LParam = wm_lbuttondown) then
  begin
    if(Self.Visible = false) then
    begin
      Shell_NotifyIcon(NIM_DELETE, @Nid);
      Self.Visible := true;
      DisableHotKeyHook;
    end;
  end;
end;

procedure TFrmBg.AddNotifyIcon;
begin
  Nid.cbSize := sizeof(NotifyIconData);
  Nid.hWnd := Self.Handle;
  Nid.uID := 0;
  Nid.uFlags := NIF_MESSAGE or NIF_ICON or NIF_TIP;
  Nid.uCallbackMessage := WM_TRAYNOTIFY;
  Nid.hIcon := ImageTips.Picture.Icon.Handle;
  strPLCopy(Nid.szTip,'Buldur''s Woods...?',63);
  Shell_NotifyIcon(NIM_ADD,@Nid);
  Self.Visible := false;
end;

end.

