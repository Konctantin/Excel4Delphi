unit Excel4Delphi.Utils;

interface

uses
  System.SysUtils, System.UITypes, System.Types, System.Classes, System.Math,
  Excel4Delphi, Excel4Delphi.Xml, Excel4Delphi.Common;

/// <summary>
/// Сохраняет страницу TZWorkBook в поток в формате HTML
/// </summary>
function SaveXmlssToHtml(sheet: TZSheet; CodePageName: string = 'UTF-8'): string;

implementation

uses
  Excel4Delphi.NumberFormats, System.StrUtils, System.AnsiStrings;

function SaveXmlssToHtml(sheet: TZSheet; CodePageName: string = 'UTF-8'): string;
var
  Xml: TZsspXMLWriterH;
  i, j, t, l, r: integer;
  NumTopLeft, NumArea: integer;
  s, value, numformat: string;
  Att: TZAttributesH;
  max_width: Real;
  strArray: TArray<string>;
  Stream: TStringStream;

  function HTMLStyleTable(name: string; const Style: TZStyle): string;
  var
    s: string;
    i, l: integer;
  begin
    result := #13#10 + ' .' + name + '{'#13#10;
    for i := 0 to 3 do
    begin
      s := 'border-';
      l := 0;
      case i of
        0:
          s := s + 'left:';
        1:
          s := s + 'top:';
        2:
          s := s + 'right:';
        3:
          s := s + 'bottom:';
      end;
      s := s + '#' + ColorToHTMLHex(Style.Border[TZBordersPos(i)].Color);
      if Style.Border[TZBordersPos(i)].Weight <> 0 then
        s := s + ' ' + IntToStr(Style.Border[TZBordersPos(i)].Weight) + 'px'
      else
        inc(l);
      case Style.Border[TZBordersPos(i)].LineStyle of
        ZEContinuous:
          s := s + ' ' + 'solid';
        ZEHair:
          s := s + ' ' + 'solid';
        ZEDot:
          s := s + ' ' + 'dotted';
        ZEDashDotDot:
          s := s + ' ' + 'dotted';
        ZEDash:
          s := s + ' ' + 'dashed';
        ZEDashDot:
          s := s + ' ' + 'dashed';
        ZESlantDashDot:
          s := s + ' ' + 'dashed';
        ZEDouble:
          s := s + ' ' + 'double';
      else
        inc(l);
      end;
      s := s + ';';
      if l <> 2 then
        result := result + s + #13#10;
    end;
    result := result + 'background:#' + ColorToHTMLHex(Style.BGColor) + ';}';
  end;

  function HTMLStyleFont(name: string; const Style: TZStyle): string;
  begin
    result := #13#10 + ' .' + name + '{'#13#10;
    result := result + 'color:#' + ColorToHTMLHex(Style.Font.Color) + ';';
    result := result + 'font-size:' + FloatToStr(Style.Font.Size, TFormatSettings.Invariant) + 'px;';
    result := result + 'font-family:' + Style.Font.name + ';}';
  end;

begin
  result := '';
  Stream := TStringStream.Create('', TEncoding.UTF8);
  Xml := TZsspXMLWriterH.Create(Stream);
  try
    Xml.TabLength := 1;
    // start
    Xml.Attributes.Clear();
    Xml.WriteRaw
      ('<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">',
      true, false);
    Xml.WriteTagNode('HTML', true, true, false);
    Xml.WriteTagNode('HEAD', true, true, false);
    Xml.WriteTag('TITLE', sheet.Title, true, false, false);

    // styles
    s := 'body {';
    s := s + 'background:#' + ColorToHTMLHex(sheet.WorkBook.Styles.DefaultStyle.BGColor) + ';';
    s := s + 'color:#' + ColorToHTMLHex(sheet.WorkBook.Styles.DefaultStyle.Font.Color) + ';';
    s := s + 'font-size:' + FloatToStr(sheet.WorkBook.Styles.DefaultStyle.Font.Size, TFormatSettings.Invariant) + 'px;';
    s := s + 'font-family:' + sheet.WorkBook.Styles.DefaultStyle.Font.name + ';}';

    s := s + HTMLStyleTable('T19', sheet.WorkBook.Styles.DefaultStyle);
    s := s + HTMLStyleFont('F19', sheet.WorkBook.Styles.DefaultStyle);

    for i := 0 to sheet.WorkBook.Styles.Count - 1 do
    begin
      s := s + HTMLStyleTable('T' + IntToStr(i + 20), sheet.WorkBook.Styles[i]);
      s := s + HTMLStyleFont('F' + IntToStr(i + 20), sheet.WorkBook.Styles[i]);
    end;

    Xml.WriteTag('STYLE', s, true, true, false);
    Xml.Attributes.Add('HTTP-EQUIV', 'CONTENT-TYPE');

    s := '';
    if trim(CodePageName) > '' then
      s := '; CHARSET=' + CodePageName;

    Xml.Attributes.Add('CONTENT', 'TEXT/HTML' + s);
    Xml.WriteTag('META', '', true, false, false);
    Xml.WriteEndTagNode(); // HEAD

    max_width := 0.0;
    for i := 0 to sheet.ColCount - 1 do
      max_width := max_width + sheet.ColWidths[i];

    // BODY
    Xml.Attributes.Clear();
    Xml.WriteTagNode('BODY', true, true, false);

    // Table
    Xml.Attributes.Clear();
    Xml.Attributes.Add('cellSpacing', '0');
    Xml.Attributes.Add('border', '0');
    Xml.Attributes.Add('width', FloatToStr(max_width).Replace(',', '.'));
    Xml.WriteTagNode('TABLE', true, true, false);

    Att := TZAttributesH.Create();
    Att.Clear();
    for i := 0 to sheet.RowCount - 1 do
    begin
      Xml.Attributes.Clear();
      Xml.Attributes.Add('height', FloatToStr(sheet.RowHeights[i]).Replace(',', '.'));
      Xml.WriteTagNode('TR', true, true, true);
      Xml.Attributes.Clear();
      for j := 0 to sheet.ColCount - 1 do
      begin
        NumTopLeft := sheet.MergeCells.InLeftTopCorner(j, i);
        NumArea := sheet.MergeCells.InMergeRange(j, i);
        // если ячейка входит в объединённые области и не является
        // верхней левой ячейкой в этой области - пропускаем её
        if not((NumArea >= 0) and (NumTopLeft = -1)) then
        begin
          Xml.Attributes.Clear();
          if NumTopLeft >= 0 then
          begin
            t := sheet.MergeCells.Items[NumTopLeft].Right - sheet.MergeCells.Items[NumTopLeft].Left;
            if t > 0 then
              Xml.Attributes.Add('colspan', IntToStr(t + 1));
            t := sheet.MergeCells.Items[NumTopLeft].Bottom - sheet.MergeCells.Items[NumTopLeft].Top;
            if t > 0 then
              Xml.Attributes.Add('rowspan', IntToStr(t + 1));
          end;
          t := sheet.Cell[j, i].CellStyle;
          if sheet.WorkBook.Styles[t].Alignment.Horizontal = ZHCenter then
            Xml.Attributes.Add('align', 'center')
          else if sheet.WorkBook.Styles[t].Alignment.Horizontal = ZHRight then
            Xml.Attributes.Add('align', 'right')
          else if sheet.WorkBook.Styles[t].Alignment.Horizontal = ZHJustify then
            Xml.Attributes.Add('align', 'justify');
          numformat := sheet.WorkBook.Styles[t].NumberFormat;
          Xml.Attributes.Add('class', 'T' + IntToStr(t + 20));
          Xml.Attributes.Add('width', IntToStr(sheet.Columns[j].WidthPix) + 'px');

          Xml.WriteTagNode('TD', true, false, false);
          Xml.Attributes.Clear();
          Att.Clear();
          Att.Add('class', 'F' + IntToStr(t + 20));
          if TFontStyle.fsbold in sheet.WorkBook.Styles[t].Font.Style then
            Xml.WriteTagNode('B', false, false, false);
          if TFontStyle.fsItalic in sheet.WorkBook.Styles[t].Font.Style then
            Xml.WriteTagNode('I', false, false, false);
          if TFontStyle.fsUnderline in sheet.WorkBook.Styles[t].Font.Style then
            Xml.WriteTagNode('U', false, false, false);
          if TFontStyle.fsStrikeOut in sheet.WorkBook.Styles[t].Font.Style then
            Xml.WriteTagNode('S', false, false, false);

          l := Length(sheet.Cell[j, i].Href);
          if l > 0 then
          begin
            Xml.Attributes.Add('href', sheet.Cell[j, i].Href);
              // target?
            Xml.WriteTagNode('A', false, false, false);
            Xml.Attributes.Clear();
          end;

          value := sheet.Cell[j, i].Data;

          // value := value.Replace(#13#10, '<br>');
          case sheet.Cell[j, i].CellType of
            TZCellType.ZENumber:
              begin
                r := numformat.IndexOf('.');
                if r > -1 then
                begin
                  value := FloatToStrF(sheet.Cell[j, i].AsDouble, ffNumber, 12,
                    Min(4, Max(0, numformat.Substring(r).Length - 1)));
                end
                else
                begin
                  value := FloatToStr(sheet.Cell[j, i].AsDouble);
                end;
              end;
            TZCellType.ZEDateTime:
              begin
                // todo: make datetimeformat from cell NumberFormat
                value := FormatDateTime('dd.mm.yyyy', sheet.Cell[j, i].AsDateTime);
              end;
          end;
          strArray := value.Split([#13, #10], TStringSplitOptions.ExcludeEmpty);
          for r := 0 to Length(strArray) - 1 do
          begin
            if r > 0 then
              Xml.WriteTag('BR', '');
            Xml.WriteTag('FONT', strArray[r], Att, false, false, true);
          end;

          if l > 0 then
            Xml.WriteEndTagNode(); // A

          if TFontStyle.fsbold in sheet.WorkBook.Styles[t].Font.Style then
            Xml.WriteEndTagNode(); // B
          if TFontStyle.fsItalic in sheet.WorkBook.Styles[t].Font.Style then
            Xml.WriteEndTagNode(); // I
          if TFontStyle.fsUnderline in sheet.WorkBook.Styles[t].Font.Style then
            Xml.WriteEndTagNode(); // U
          if TFontStyle.fsStrikeOut in sheet.WorkBook.Styles[t].Font.Style then
            Xml.WriteEndTagNode(); // S
          Xml.WriteEndTagNode(); // TD
        end;

      end;
      Xml.WriteEndTagNode(); // TR
    end;

    Xml.WriteEndTagNode(); // BODY
    Xml.WriteEndTagNode(); // HTML
    Xml.EndSaveTo();
    result := Stream.DataString;
    FreeAndNil(Att);
  finally
    Xml.Free();
    Stream.Free();
  end;
end;

end.
