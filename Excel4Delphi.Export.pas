unit Excel4Delphi.Export;

interface

uses
  Windows,
  SysUtils,
  UITypes,
  Types,
  Classes,
  Math,
  Graphics,
  AnsiStrings,
  Excel4Delphi,
  Excel4Delphi.Xml;

/// <summary>
/// Сохраняет страницу TZWorkBook в поток в формате HTML
/// </summary>
function SaveXmlssToHtml(sheet: TZSheet; CodePageName: string = 'UTF-8'): string;

implementation

function SaveXmlssToHtml(sheet: TZSheet; CodePageName: string = 'UTF-8'): string;
var xml: TZsspXMLWriterH;
  i, j, t, l, r: integer;
  NumTopLeft, NumArea: integer;
  s, value, numformat: string;
  Att: TZAttributesH;
  max_width: Real;
  strArray: TArray<string>;
  Stream: TStringStream;

  function HTMLStyleTable(name: string; const Style: TZStyle): string;
  var s: string; i, l: integer;
  begin
    result := #13#10 + ' .' + name + '{'#13#10;
    for i := 0 to 3 do begin
      s := 'border-';
      l := 0;
      case i of
        0: s := s + 'left:';
        1: s := s + 'top:';
        2: s := s + 'right:';
        3: s := s + 'bottom:';
      end;
      s := s + '#' + ColorToHTMLHex(Style.Border[TZBordersPos(i)].Color);
      if Style.Border[TZBordersPos(i)].Weight <> 0 then
        s := s + ' ' + IntToStr(Style.Border[TZBordersPos(i)].Weight) + 'px'
      else
        inc(l);
      case Style.Border[TZBordersPos(i)].LineStyle of
        ZEContinuous:    s := s + ' ' + 'solid';
        ZEHair:          s := s + ' ' + 'solid';
        ZEDot:           s := s + ' ' + 'dotted';
        ZEDashDotDot:    s := s + ' ' + 'dotted';
        ZEDash:          s := s + ' ' + 'dashed';
        ZEDashDot:       s := s + ' ' + 'dashed';
        ZESlantDashDot:  s := s + ' ' + 'dashed';
        ZEDouble:        s := s + ' ' + 'double';
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
    result := result + 'color:#'      + ColorToHTMLHex(Style.Font.Color) + ';';
    result := result + 'font-size:'   + FloatToStr(Style.Font.Size, TFormatSettings.Invariant) + 'px;';
    result := result + 'font-family:' + Style.Font.Name + ';}';
  end;

begin
  result := '';
  Stream := TStringStream.Create('', TEncoding.UTF8);
  xml := TZsspXMLWriterH.Create(Stream);
  try
    xml.TabLength := 1;
    // start
    xml.Attributes.Clear();
    xml.WriteRaw('<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">', true, false);
    xml.WriteTagNode('HTML', true, true, false);
    xml.WriteTagNode('HEAD', true, true, false);
    xml.WriteTag('TITLE', sheet.Title, true, false, false);

    //styles
    s := 'body {';
    s := s + 'background:#' + ColorToHTMLHex(sheet.WorkBook.Styles.DefaultStyle.BGColor) + ';';
    s := s + 'color:#'      + ColorToHTMLHex(sheet.WorkBook.Styles.DefaultStyle.Font.Color) + ';';
    s := s + 'font-size:'   + FloatToStr(sheet.WorkBook.Styles.DefaultStyle.Font.Size, TFormatSettings.Invariant) + 'px;';
    s := s + 'font-family:' + sheet.WorkBook.Styles.DefaultStyle.Font.Name + ';}';

    s := s + HTMLStyleTable('T19', sheet.WorkBook.Styles.DefaultStyle);
    s := s +  HTMLStyleFont('F19', sheet.WorkBook.Styles.DefaultStyle);

    for i := 0 to sheet.WorkBook.Styles.Count - 1 do begin
      s := s + HTMLStyleTable('T' + IntToStr(i + 20), sheet.WorkBook.Styles[i]);
      s := s +  HTMLStyleFont('F' + IntToStr(i + 20), sheet.WorkBook.Styles[i]);
    end;

    xml.WriteTag('STYLE', s, true, true, false);
    xml.Attributes.Add('HTTP-EQUIV', 'CONTENT-TYPE');

    s := '';
    if trim(CodePageName) > '' then
      s := '; CHARSET=' + CodePageName;

    xml.Attributes.Add('CONTENT', 'TEXT/HTML' + s);
    xml.WriteTag('META', '', true, false, false);
    xml.WriteEndTagNode(); // HEAD

    max_width := 0.0;
    for i := 0 to sheet.ColCount - 1 do
      max_width := max_width + sheet.ColWidths[i];

    //BODY
    xml.Attributes.Clear();
    xml.WriteTagNode('BODY', true, true, false);

    //Table
    xml.Attributes.Clear();
    xml.Attributes.Add('cellSpacing', '0');
    xml.Attributes.Add('border', '0');
    xml.Attributes.Add('width', FloatToStr(max_width).Replace(',', '.'));
    xml.WriteTagNode('TABLE', true, true, false);

    Att := TZAttributesH.Create();
    Att.Clear();
    for i := 0 to sheet.RowCount - 1 do begin
      xml.Attributes.Clear();
      xml.Attributes.Add('height', floattostr(sheet.RowHeights[i]).Replace(',', '.'));
      xml.WriteTagNode('TR', true, true, true);
      xml.Attributes.Clear();
      for j := 0 to sheet.ColCount - 1 do begin
        NumTopLeft := sheet.MergeCells.InLeftTopCorner(j, i);
        NumArea    := sheet.MergeCells.InMergeRange(j, i);
        // если ячейка входит в объединённые области и не является
        // верхней левой ячейкой в этой области - пропускаем её
        if not ((NumArea >= 0) and (NumTopLeft = -1)) then begin
          xml.Attributes.Clear();
          if NumTopLeft >= 0 then begin
            t := sheet.MergeCells.Items[NumTopLeft].Right - sheet.MergeCells.Items[NumTopLeft].Left;
            if t > 0 then
              xml.Attributes.Add('colspan', InttOstr(t + 1));
            t := sheet.MergeCells.Items[NumTopLeft].Bottom - sheet.MergeCells.Items[NumTopLeft].Top;
            if t > 0 then
              xml.Attributes.Add('rowspan', InttOstr(t + 1));
          end;
          t := sheet.Cell[j, i].CellStyle;
          if sheet.WorkBook.Styles[t].Alignment.Horizontal = ZHCenter then
            xml.Attributes.Add('align', 'center')
          else if sheet.WorkBook.Styles[t].Alignment.Horizontal = ZHRight then
            xml.Attributes.Add('align', 'right')
          else if sheet.WorkBook.Styles[t].Alignment.Horizontal = ZHJustify then
            xml.Attributes.Add('align', 'justify');
          numformat := sheet.WorkBook.Styles[t].NumberFormat;
          xml.Attributes.Add('class', 'T' + IntToStr(t + 20));
          xml.Attributes.Add('width', inttostr(sheet.Columns[j].WidthPix) + 'px');

          xml.WriteTagNode('TD', true, false, false);
          xml.Attributes.Clear();
          Att.Clear();
          Att.Add('class', 'F' + IntToStr(t + 20));
          if fsbold in sheet.WorkBook.Styles[t].Font.Style then
            xml.WriteTagNode('B', false, false, false);
          if fsItalic in sheet.WorkBook.Styles[t].Font.Style then
            xml.WriteTagNode('I', false, false, false);
          if fsUnderline in sheet.WorkBook.Styles[t].Font.Style then
            xml.WriteTagNode('U', false, false, false);
          if fsStrikeOut in sheet.WorkBook.Styles[t].Font.Style then
            xml.WriteTagNode('S', false, false, false);

          l := Length(sheet.Cell[j, i].Href);
          if l > 0 then begin
            xml.Attributes.Add('href', sheet.Cell[j, i].Href);
              //target?
            xml.WriteTagNode('A', false, false, false);
            xml.Attributes.Clear();
          end;

          value := sheet.Cell[j, i].Data;

          //value := value.Replace(#13#10, '<br>');
          case sheet.Cell[j, i].CellType of
            TZCellType.ZENumber:
              begin
                r := numformat.IndexOf('.');
                if r > -1 then begin
                  value := FloatToStrF(sheet.Cell[j, i].AsDouble, ffNumber, 12, Min(4, Max(0, numformat.Substring(r).Length - 1)));
                end
                else begin
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
          for r := 0 to Length(strArray) - 1 do begin
            if r > 0 then
              xml.WriteTag('BR', '');
            xml.WriteTag('FONT', strArray[r], Att, false, false, true);
          end;

          if l > 0 then
            xml.WriteEndTagNode(); // A

          if fsbold in sheet.WorkBook.Styles[t].Font.Style then
            xml.WriteEndTagNode(); // B
          if fsItalic in sheet.WorkBook.Styles[t].Font.Style then
            xml.WriteEndTagNode(); // I
          if fsUnderline in sheet.WorkBook.Styles[t].Font.Style then
            xml.WriteEndTagNode(); // U
          if fsStrikeOut in sheet.WorkBook.Styles[t].Font.Style then
            xml.WriteEndTagNode(); // S
          xml.WriteEndTagNode(); // TD
        end;

      end;
      xml.WriteEndTagNode(); // TR
    end;

    xml.WriteEndTagNode(); // BODY
    xml.WriteEndTagNode(); // HTML
    xml.EndSaveTo();
    Result := Stream.DataString;
    FreeAndNil(Att);
  finally
    xml.Free();
    Stream.Free();
  end;
end;

end.

