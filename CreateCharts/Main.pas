unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics,
  Controls, Forms, Dialogs, StdCtrls, XLSReadWriteII5, xpgParseChart,
  XLSRelCells5, ExtCtrls, FormFormatChartItem, Xc12DataStyleSheet5, XLSDrawing5,
  Grids, XLSUtils5, Xc12Utils5;

type TSampleChartItemProps = class(TObject)
protected
     FShape   : TXLSDrwShapeProperies;
     FFont    : TXc12Font;
public
     constructor Create(ADrawing: TXLSDrawing; ADefaultFont: TXc12Font);
     destructor Destroy; override;

     property Shape   : TXLSDrwShapeProperies read FShape;
     property Font    : TXc12Font read FFont;
     end;

type
  TfrmMain = class(TForm)
    Label1: TLabel;
    edFilename: TEdit;
    Button3: TButton;
    dlgSave: TSaveDialog;
    Button1: TButton;
    Button4: TButton;
    shpChart: TShape;
    shpPlotArea: TShape;
    shpTitle: TShape;
    cbTitle: TCheckBox;
    btnTitleStyle: TButton;
    btnTitleText: TButton;
    Button5: TButton;
    Label3: TLabel;
    Label4: TLabel;
    Shape1: TShape;
    cbLegend: TCheckBox;
    btnLegendStyle: TButton;
    btnLegendText: TButton;
    lblLegendPos: TLabel;
    cbLegendPos: TComboBox;
    dlgFont: TFontDialog;
    cbChart: TComboBox;
    Label5: TLabel;
    Shape2: TShape;
    Shape3: TShape;
    Button2: TButton;
    Button6: TButton;
    Button7: TButton;
    Button8: TButton;
    Label2: TLabel;
    Label6: TLabel;
    sgData: TStringGrid;
    Label7: TLabel;
    btnAddSerie: TButton;
    BtnRemoveSerie: TButton;
    Button9: TButton;
    Label8: TLabel;
    cbMarkers: TComboBox;
    lblTitleRef: TLabel;
    edTitleRef: TEdit;
    Button10: TButton;
    Label9: TLabel;
    edValAxisNumFmt: TEdit;
    cbValAxisLog: TCheckBox;
    btbPos: TButton;
    cbCatAxValues: TCheckBox;
    procedure Button3Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure cbTitleClick(Sender: TObject);
    procedure btnTitleStyleClick(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure cbLegendClick(Sender: TObject);
    procedure btnLegendStyleClick(Sender: TObject);
    procedure btnTitleTextClick(Sender: TObject);
    procedure btnLegendTextClick(Sender: TObject);
    procedure btnAddSerieClick(Sender: TObject);
    procedure BtnRemoveSerieClick(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button8Click(Sender: TObject);
    procedure Button9Click(Sender: TObject);
    procedure Button10Click(Sender: TObject);
    procedure btbPosClick(Sender: TObject);
  private
    XLS           : TXLSReadWriteII5;
    FPropsChart   : TSampleChartItemProps;
    FPropsValAxis : TSampleChartItemProps;
    FPropsCatAxis : TSampleChartItemProps;
    FPropsPlotArea: TSampleChartItemProps;
    FPropsLegend  : TSampleChartItemProps;
    FPropsTitle   : TSampleChartItemProps;

    procedure SetupDefaultProps;

    procedure FillSheetValues;
    procedure MakeChart;

    procedure AddSerie;
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

{$R *.dfm}

procedure TfrmMain.AddSerie;
var
  i,j: integer;
begin
  j := sgData.ColCount;

  sgData.ColCount := sgData.ColCount + 1;
  sgData.Cells[j,0] := 'Ser ' + IntToStr(j);

  for i := 0 to 5 do
    sgData.Cells[j,i + 1] := IntToStr(Random(10000));

  btnRemoveSerie.Enabled := sgData.ColCount > 2;
end;

procedure TfrmMain.btnLegendStyleClick(Sender: TObject);
begin
  TfrmFormatChartItem.Create(Self).Execute(FPropsLegend.Shape);
end;

procedure TfrmMain.btnLegendTextClick(Sender: TObject);
begin
  FPropsLegend.Font.CopyToTFont(dlgFont.Font);
  if dlgFont.Execute then
    FPropsLegend.Font.Assign(dlgFont.Font);
end;

procedure TfrmMain.btnTitleStyleClick(Sender: TObject);
begin
  TfrmFormatChartItem.Create(Self).Execute(FPropsTitle.Shape);
end;

procedure TfrmMain.btnTitleTextClick(Sender: TObject);
begin
  FPropsTitle.Font.CopyToTFont(dlgFont.Font);
  if dlgFont.Execute then
    FPropsTitle.Font.Assign(dlgFont.Font);
end;

procedure TfrmMain.BtnRemoveSerieClick(Sender: TObject);
begin
  sgData.ColCount := sgData.ColCount - 1;
  btnRemoveSerie.Enabled := sgData.ColCount > 2;
end;

procedure TfrmMain.Button10Click(Sender: TObject);
begin
  TfrmFormatChartItem.Create(Self).Execute(FPropsChart.Shape);
end;

procedure TfrmMain.Button1Click(Sender: TObject);
begin
  Close;
end;

procedure TfrmMain.Button2Click(Sender: TObject);
begin
  TfrmFormatChartItem.Create(Self).Execute(FPropsValAxis.Shape);
end;

procedure TfrmMain.Button3Click(Sender: TObject);
begin
  dlgSave.InitialDir := ExtractFilePath(Application.ExeName);
  dlgSave.FileName := edFilename.Text;
  if dlgSave.Execute then
    edFilename.Text := dlgSave.FileName;
end;

procedure TfrmMain.Button4Click(Sender: TObject);
begin
  XLS.Filename := edFilename.Text;
  XLS.Write;
end;

procedure TfrmMain.Button5Click(Sender: TObject);
begin
  TfrmFormatChartItem.Create(Self).Execute(FPropsPlotArea.Shape);
end;

procedure TfrmMain.Button6Click(Sender: TObject);
begin
  FPropsValAxis.Font.CopyToTFont(dlgFont.Font);
  if dlgFont.Execute then
    FPropsValAxis.Font.Assign(dlgFont.Font);
end;

procedure TfrmMain.Button7Click(Sender: TObject);
begin
  TfrmFormatChartItem.Create(Self).Execute(FPropsCatAxis.Shape);
end;

procedure TfrmMain.Button8Click(Sender: TObject);
begin
  FPropsCatAxis.Font.CopyToTFont(dlgFont.Font);
  if dlgFont.Execute then
    FPropsCatAxis.Font.Assign(dlgFont.Font);
end;

procedure TfrmMain.Button9Click(Sender: TObject);
begin
  MakeChart;
end;

procedure TfrmMain.btbPosClick(Sender: TObject);
begin
  if XLS[0].Drawing.Charts.Count > 0 then begin
    XLS[0].Drawing.Charts[0].Col1 := 5;
    XLS[0].Drawing.Charts[0].Col2 := 15;
  end;
end;

procedure TfrmMain.btnAddSerieClick(Sender: TObject);
begin
  AddSerie;
end;

procedure TfrmMain.cbTitleClick(Sender: TObject);
begin
  btnTitleStyle.Enabled := cbTitle.Checked;
  btnTitleText.Enabled := cbTitle.Checked;
  lblTitleRef.Enabled := cbTitle.Checked;
  edTitleRef.Enabled := cbTitle.Checked;
end;

procedure TfrmMain.FillSheetValues;
var
  r,c: integer;
begin
  for c := 0 to sgData.ColCount - 1 do
    XLS[0].AsString[c,0] := sgData.Cells[c,0];

  for r := 1 to sgData.RowCount - 1 do
    XLS[0].AsString[0,r] := sgData.Cells[0,r];

  for c := 1 to sgData.ColCount - 1 do begin
    for r := 1 to sgData.RowCount - 1 do
      XLS[0].AsFloat[c,r] := StrToFloatDef(sgData.Cells[c,r],0);
  end;
end;

procedure TfrmMain.FormCreate(Sender: TObject);
var
  i: integer;
begin
  XLS := TXLSReadWriteII5.Create(Nil);

  cbLegendPos.ItemIndex := 0;

  cbChart.ItemIndex := 0;
  cbMarkers.ItemIndex := 0;

  FPropsChart    := TSampleChartItemProps.Create(XLS[0].Drawing,XLS.Font);
  FPropsValAxis  := TSampleChartItemProps.Create(XLS[0].Drawing,XLS.Font);
  FPropsCatAxis  := TSampleChartItemProps.Create(XLS[0].Drawing,XLS.Font);
  FPropsPlotArea := TSampleChartItemProps.Create(XLS[0].Drawing,XLS.Font);
  FPropsLegend   := TSampleChartItemProps.Create(XLS[0].Drawing,XLS.Font);
  FPropsTitle    := TSampleChartItemProps.Create(XLS[0].Drawing,XLS.Font);

  SetupDefaultProps;

  sgData.Cells[0,0] := 'Month';
  for i := 1 to 5 do
    sgData.Cells[0,i] := FormatSettings.ShortMonthNames[i];
end;

procedure TfrmMain.FormDestroy(Sender: TObject);
begin
  FPropsChart.Free;
  FPropsValAxis.Free;
  FPropsCatAxis.Free;
  FPropsPlotArea.Free;
  FPropsLegend.Free;
  FPropsTitle.Free;

  XLS.Free;
end;

procedure TfrmMain.MakeChart;
var
  i         : integer;
  S         : AxUCString;
  ChartSpace: TCT_ChartSpace;
  RCells    : TXLSRelCells;
  Text      : TXLSDrwText;
begin
  if sgData.ColCount <= 1 then begin
    ShowMessage('Add at least one serie');
    Exit;
  end;

  FillSheetValues;

  RCells := XLS[0].CreateRelativeCells;
  RCells.SetArea(1,0,sgData.ColCount - 1,sgData.RowCount - 1);

  case cbChart.ItemIndex of
    0: ChartSpace := XLS[0].Drawing.Charts.MakeBarChart(RCells,1,sgData.RowCount,True);
    1: ChartSpace := XLS[0].Drawing.Charts.MakeLineChart(RCells,1,sgData.RowCount);
    2: ChartSpace := XLS[0].Drawing.Charts.MakeAreaChart(RCells,1,sgData.RowCount);
    3: ChartSpace := XLS[0].Drawing.Charts.MakeBubbleChart(RCells,1,sgData.RowCount);
    4: ChartSpace := XLS[0].Drawing.Charts.MakeDoughnutChart(RCells,1,sgData.RowCount);
    5: ChartSpace := XLS[0].Drawing.Charts.MakePieChart(RCells,1,sgData.RowCount);
    6: ChartSpace := XLS[0].Drawing.Charts.MakeRadarChart(RCells,1,sgData.RowCount);
    7: ChartSpace := XLS[0].Drawing.Charts.MakeScatterChart(RCells,1,sgData.RowCount);
    else
       ChartSpace := XLS[0].Drawing.Charts.MakeBarChart(RCells,1,sgData.RowCount,True);
  end;

  if FPropsChart.Shape.Assigned then begin
    ChartSpace.Create_SpPr;
    ChartSpace.SpPr.Assign(FPropsChart.Shape.SpPr);
  end;
  if not FPropsChart.Font.Equal(XLS.Font) then
    SetChartFont(FPropsValAxis.Font,ChartSpace);
  if FPropsPlotArea.Shape.Assigned then begin
    ChartSpace.Chart.PlotArea.Create_SpPr;
    ChartSpace.Chart.PlotArea.SpPr.Assign(FPropsPlotArea.Shape.SpPr);
  end;
  if FPropsValAxis.Shape.Assigned then begin
    for i := 0 to ChartSpace.Chart.PlotArea.ValAxis.Count - 1 do begin
      ChartSpace.Chart.PlotArea.ValAxis[i].Shared.Create_SpPr;
      ChartSpace.Chart.PlotArea.ValAxis[i].Shared.SpPr.Assign(FPropsValAxis.Shape.SpPr);
    end;
  end;
  if not FPropsValAxis.Font.Equal(XLS.Font) then begin
    for i := 0 to ChartSpace.Chart.PlotArea.ValAxis.Count - 1 do
      SetChartFont(FPropsValAxis.Font,ChartSpace.Chart.PlotArea.ValAxis[0]);
  end;
  if edValAxisNumFmt.Text <> '' then begin
    for i := 0 to ChartSpace.Chart.PlotArea.ValAxis.Count - 1 do begin
      ChartSpace.Chart.PlotArea.ValAxis[i].Shared.Create_NumFmt;
      ChartSpace.Chart.PlotArea.ValAxis[i].Shared.NumFmt.FormatCode := edValAxisNumFmt.Text;
      ChartSpace.Chart.PlotArea.ValAxis[i].Shared.NumFmt.SourceLinked := False;
    end;
  end;
  if cbValAxisLog.Checked then begin
    for i := 0 to ChartSpace.Chart.PlotArea.ValAxis.Count - 1 do begin
      ChartSpace.Chart.PlotArea.ValAxis[i].Shared.Create_Scaling;
      ChartSpace.Chart.PlotArea.ValAxis[i].Shared.Scaling.Create_LogBase;
      ChartSpace.Chart.PlotArea.ValAxis[i].Shared.Scaling.LogBase.Val := 10;
      ChartSpace.Chart.PlotArea.ValAxis[i].Shared.Scaling.Create_Orientation;
      ChartSpace.Chart.PlotArea.ValAxis[i].Shared.Scaling.Orientation.Val := stoMaxMin;
    end;
  end;

  if cbCatAxValues.Checked then begin
    ChartSpace.Chart.PlotArea.BarChart.Shared.Series[0].Create_Cat;
    ChartSpace.Chart.PlotArea.BarChart.Shared.Series[0].Cat.Create_strRef;
    ChartSpace.Chart.PlotArea.BarChart.Shared.Series[0].Cat.strRef.F := 'Sheet1!' + AreaToRefStr(0,1,0,sgData.RowCount - 1,True,True,True,True);
  end;

  if FPropsCatAxis.Shape.Assigned then begin
    ChartSpace.Chart.PlotArea.CatAxis[0].Shared.Create_SpPr;
    ChartSpace.Chart.PlotArea.CatAxis[0].Shared.SpPr.Assign(FPropsCatAxis.Shape.SpPr);
  end;
  if not FPropsCatAxis.Font.Equal(XLS.Font) then
    SetChartFont(FPropsValAxis.Font,ChartSpace.Chart.PlotArea.CatAxis[0]);
  if cbLegend.Checked then begin
    ChartSpace.Chart.CreateDefaultLegend;
    ChartSpace.Chart.Legend.Create_LegendPos;
    case cbLegendPos.ItemIndex of
      0: ChartSpace.Chart.Legend.LegendPos.Val := stlpR;
      1: ChartSpace.Chart.Legend.LegendPos.Val := stlpL;
      2: ChartSpace.Chart.Legend.LegendPos.Val := stlpT;
      3: ChartSpace.Chart.Legend.LegendPos.Val := stlpB;
    end;
    if FPropsLegend.Shape.Assigned then begin
      ChartSpace.Chart.Legend.Create_SpPr;
      ChartSpace.Chart.Legend.SpPr.Assign(FPropsLegend.Shape.SpPr);
    end;
    if not FPropsLegend.Font.Equal(XLS.Font) then
      SetChartFont(FPropsValAxis.Font,ChartSpace.Chart.Legend);
  end;
  if cbTitle.Checked then begin
    ChartSpace.Chart.Create_Title;
    Text := TXLSDrwText.Create(ChartSpace.Chart.Title.Create_Tx);
    try
      i := Pos('!',edTitleRef.Text);

      S := '';
      if i > 1 then
        S := Copy(edTitleRef.Text,i + 1,MAXINT)
      else
        S := '';

      if IsAreaStr(S) then
        Text.Ref := XLS[0].CreateRelativeCells(S)
      else
        Text.PlainText := 'Chart Test';
    finally
      Text.Free;
    end;
    if not FPropsTitle.Font.Equal(XLS.Font) then
      SetChartFont(FPropsValAxis.Font,ChartSpace.Chart.Title);
  end;
end;

procedure TfrmMain.SetupDefaultProps;
begin
  FPropsTitle.Font.Size := 18;
  FPropsTitle.Font.Style := [xfsBold];
end;

procedure TfrmMain.cbLegendClick(Sender: TObject);
begin
  btnLegendStyle.Enabled := cbLegend.Checked;
  btnLegendText.Enabled := cbLegend.Checked;
  lblLegendPos.Enabled := cbLegend.Checked;
  cbLegendPos.Enabled := cbLegend.Checked;
end;

{ TSampleChartItemProps }

constructor TSampleChartItemProps.Create(ADrawing: TXLSDrawing; ADefaultFont: TXc12Font);
begin
  FShape := ADrawing.CreateShapeProps;
  FFont := TXc12Font.Create(Nil);
  FFont.Assign(ADefaultFont);
end;

destructor TSampleChartItemProps.Destroy;
begin
  FShape.Free;
  FFont.Free;

  inherited;
end;

end.
