using ClosedXML.Excel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using static ExcelUtils.TableInfo;

namespace ExcelUtils
{
    public class ExcelTable
    {
        /// <summary>Конструктор</summary>
        /// <param name="ti">Описание таблицы</param>
        /// <param name="FirstRowIndex">Индекс первой строки (с нуля)</param>
        /// <param name="FirstColumnIndex">Индекс первого столбца (с нуля)</param>
        /// <param name="ShowHeaders">Показать заголовки столбцов</param>
        /// <param name="SheetName">Имя листа книги</param>
        public ExcelTable(DataTable Table, TableInfo ti, Int32 FirstRowIndex, Int32 FirstColumnIndex, bool ShowHeaders, string SheetName)
        {
            this.Table = Table;
            this.ti = ti;
            this.FirstRowIndex = FirstRowIndex;
            this.FirstColumnIndex = FirstColumnIndex;
            this.ShowHeaders = ShowHeaders;
            this.SheetName = SheetName;
            if (SheetName == null) this.SheetName = Table.TableName;
            //CloumnPosition = null;
            //ColumnDelete = null;
            //CloumnPosition = null;
            //RowFilter = null;
        }

        private IXLWorksheet? Sheet;

        /// <summary>Автофильтр в заголовке таблицы</summary>
        public bool AutoFilter { get; set; } = true;

        ///// <summary>Порядок столбцов. Названия столбцов через ","</summary>
        //public string? CloumnPosition;

        ///// <summary>Список столбцов для удаления через ","</summary>
        //public string? ColumnDelete;

        ///// <summary>Сортировка столбцов. Названия столбцов через "," (после имени можно указать desc или asc)</summary>
        //public string? ColumnSort;

        /// <summary>Индекс первого столбца (с единицы)</summary>
        public Int32 FirstColumnIndex { get; set; } = 1;

        /// <summary>Индекс первой строки (с единицы)</summary>
        public Int32 FirstRowIndex { get; set; } = 1;

        ///// <summary>Фильтрация строк. Выражение как в DataTable.RowFilter</summary>
        //public string? RowFilter;

        /// <summary>Описание таблицы</summary>
        public TableInfo ti { get; set; }

        /// <summary>Имя листа книги</summary>
        public string SheetName { get; set; } = "Лист1";

        /// <summary>Показать заголовки столбцов</summary>
        public bool ShowHeaders { get; set; } = true;

        /// <summary>Таблица с данными</summary>
        public DataTable Table { get; set; }


        /// <summary>Генерация документа Excel на основе описания</summary>
        public byte[] ExcelGenegate()
        {
            // пересоздать таблицу на основе данных о сортировке и порядке столбцов
            string ColList = string.Empty;
            foreach (var ci in ti.Columns)
            {
                ColList += ci.ColumnName + ",";
            }
            if (!string.IsNullOrEmpty(ti.ColumnSort)) TablePrepare(ColList.TrimEnd(",".ToCharArray()), ti.ColumnSort, "", "");

            using var wb = new XLWorkbook();
            Sheet = wb.AddWorksheet();
            Sheet.Outline.SummaryVLocation = XLOutlineSummaryVLocation.Top;
            Sheet.Name = SheetName;
            Sheet.Cell(FirstRowIndex, FirstColumnIndex).InsertTable(Table, false);

            // автофильтр
            if (AutoFilter) Sheet.SetAutoFilter(true);

            // границы
            GetFullTableRange()?.Cells().Style.Border.SetInsideBorder(XLBorderStyleValues.Thin);
            GetFullTableRange()?.Cells().Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);

            // Столбцы
            int ColIndex = FirstColumnIndex;
            var widthOverrides = new Dictionary<int, double>();
            foreach (var ci in ti.Columns)
            {
                IXLRange? ColRange = GetTableColumnRange(ColIndex);

                // заголовки
                if (ShowHeaders && !string.IsNullOrEmpty(ci.ColumnCaption))
                    Sheet.Cell(FirstRowIndex, ColIndex).Value = ci.ColumnCaption;

                // форматирование
                FormatingAplly(ColRange, ci.Formating);

                // форматирование заголовков
                if (ShowHeaders)
                {
                    IXLRange ColHeaderRange = Sheet.Cell(FirstRowIndex, ColIndex).AsRange();
                    FormatingAplly(ColHeaderRange, ci.FormatingHeader);
                }

                // итоги столбцов
                int scount = 0;
                foreach (var s in ci.SummatyList)
                {
                    string stext;
                    if (s == SummaryTypes.Sum) stext = "=\"Сумма: \"& SUM({0})";
                    else if (s == SummaryTypes.Min) stext = "=\"Минимум: \"& MIN({0})";
                    else if (s == SummaryTypes.Max) stext = "=\"Максимум: \"& MAX({0})";
                    else if (s == SummaryTypes.Count) stext = "=\"Строки: \"& ROWS({0})";
                    else if (s == SummaryTypes.Average) stext = "=\"Среднее: \"& AVERAGE({0})";
                    else continue;
                    IXLRange? xLRange = GetTableColumnRange(ColIndex);
                    if (xLRange != null)
                    {
                        Sheet.Cell(LastRowIndex + scount + 1, ColIndex).FormulaA1 = string.Format(stext, xLRange.ToString());
                        Sheet.Cell(LastRowIndex + scount + 1, ColIndex).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                    }
                    scount++;
                }

                // размер
                if (ci.Width > 0)
                    widthOverrides.Add(ColIndex, ci.Width);

                // видимость
                if (!ci.Visible) Sheet.Columns(ColIndex, ColIndex).Hide();

                ColIndex += 1;
            }

            // ширина столбцов
            Sheet.Columns(FirstColumnIndex, LastColumnIndex).AdjustToContents(1, 5, 150);
            foreach (var w in widthOverrides)
                Sheet.Column(w.Key).Width = w.Value;

            // условное форматирование таблицы
            foreach (ConditionFormattingTableInfo cf in ti.ConditionFormattings)
            {
                if (!Table.Columns.Contains(cf.SourceColumn))
                    continue;

                int dataStart = ShowHeaders ? FirstRowIndex + 1 : FirstRowIndex;
                int dataEnd = LastRowIndex;
                int sourceIdx = Table.Columns.IndexOf(cf.SourceColumn) + FirstColumnIndex;

                int targetIdx = cf.TargetColumn != null ? Table.Columns.IndexOf(cf.TargetColumn) : -1;
                IXLRange targetRange;
                if (targetIdx >= 0)
                    targetRange = Sheet.Range(dataStart, targetIdx + FirstColumnIndex, dataEnd, targetIdx + FirstColumnIndex);
                else
                    targetRange = Sheet.Range(dataStart, FirstColumnIndex, dataEnd, LastColumnIndex);

                var cond = targetRange.AddConditionalFormat();
                string sourceLetter = Sheet.Column(sourceIdx).ColumnLetter();
                string cellRef = "$" + sourceLetter + dataStart.ToString();

                string value1 = FormatConditionValue(cf.ConditionValue1);
                string value2 = FormatConditionValue(cf.ConditionValue2);

                string formula = cf.ConditionOperator switch
                {
                    Conditions.Equal => $"{cellRef}={value1}",
                    Conditions.NotEqual => $"{cellRef}<>{value1}",
                    Conditions.Greater => $"{cellRef}>{value1}",
                    Conditions.Less => $"{cellRef}<{value1}",
                    Conditions.GreaterOrEqual => $"{cellRef}>={value1}",
                    Conditions.LessOrEqual => $"{cellRef}<={value1}",
                    Conditions.Between => $"AND({cellRef}>={value1},{cellRef}<={value2})",
                    Conditions.NotBetween => $"OR({cellRef}<{value1},{cellRef}>{value2})",
                    Conditions.Contains => $"NOT(ISERROR(SEARCH({value1},{cellRef})))",
                    Conditions.NotContains => $"ISERROR(SEARCH({value1},{cellRef}))",
                    Conditions.BeginsWith => $"LEFT({cellRef},LEN({value1}))={value1}",
                    Conditions.EndsWith => $"RIGHT({cellRef},LEN({value1}))={value1}",
                    Conditions.ContainNulls or Conditions.ContainBlanks => $"LEN(TRIM({cellRef}))=0",
                    Conditions.ContainNonNulls or Conditions.ContainNonBlanks => $"LEN(TRIM({cellRef}))>0",
                    _ => string.Empty
                };

                if (!string.IsNullOrEmpty(formula))
                    cond.WhenIsTrue(formula);

                StyleApply(cond.Style, cf.Formating);
            } // условное форматирование таблицы

            // группировка столбцов (делаем в самом конце, т.к. добавляются строки)
            if (ti.Groups.Count > 0)
            {
                GroupsCreate(0, ti, 0, Table.Rows.Count - 1, ShowHeaders ? 2 : 1 + FirstRowIndex);
                foreach (KeyValuePair<int, int> group in Groups)
                {
                    Sheet?.Rows(group.Key, group.Value).Group(false);
                }
            }

            // выгрузка в массив байт
            MemoryStream ms = new MemoryStream();
            SaveOptions options = new SaveOptions();
            wb.SaveAs(ms, options);
            return ms.ToArray();
        }


        /// <summary>Применить формат ячеек</summary>
        /// <param name="cr">Диапазон ячеек</param>
        /// <param name="Formating">Формат ячеек</param>
        private void FormatingAplly(IXLRange? cr, FormatInfo? Formating)
        {
            if (cr == null) return;
            if (Formating == null) return;
            if (Formating.Alignment == ColumnAlignment.General) cr.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.General;
            if (Formating.Alignment == ColumnAlignment.Left) cr.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            if (Formating.Alignment == ColumnAlignment.Center) cr.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            if (Formating.Alignment == ColumnAlignment.Right) cr.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            cr.Style.Fill.BackgroundColor = XLColor.FromColor(Color.FromArgb(Formating.FillColor));
            cr.Style.Font.Bold = Formating.FontBold;
            cr.Style.Font.FontColor = XLColor.FromColor(Color.FromArgb(Formating.FontColor));
            cr.Style.Font.Italic = Formating.FontItalic;
            cr.Style.Font.FontName = Formating.FontName;
            cr.Style.Font.FontSize = Formating.FontSize;
            cr.Style.Font.Strikethrough = Formating.FontStrikethrough;
            if (Formating.FontUnderline)
                cr.Style.Font.Underline = XLFontUnderlineValues.Single;
            if (!string.IsNullOrEmpty(Formating.NumberFormat))
                cr.Style.NumberFormat.Format = Formating.NumberFormat;
            if (!string.IsNullOrEmpty(Formating.Formula))
                cr.FormulaA1 = Formating.Formula;
        }

        private void StyleApply(IXLStyle style, FormatInfo? formating)
        {
            if (formating == null) return;
            if (formating.Alignment == ColumnAlignment.General) style.Alignment.Horizontal = XLAlignmentHorizontalValues.General;
            if (formating.Alignment == ColumnAlignment.Left) style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            if (formating.Alignment == ColumnAlignment.Center) style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            if (formating.Alignment == ColumnAlignment.Right) style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            style.Fill.BackgroundColor = XLColor.FromColor(Color.FromArgb(formating.FillColor));
            style.Font.Bold = formating.FontBold;
            style.Font.FontColor = XLColor.FromColor(Color.FromArgb(formating.FontColor));
            style.Font.Italic = formating.FontItalic;
            style.Font.FontName = formating.FontName;
            style.Font.FontSize = formating.FontSize;
            style.Font.Strikethrough = formating.FontStrikethrough;
            if (formating.FontUnderline)
                style.Font.Underline = XLFontUnderlineValues.Single;
            if (!string.IsNullOrEmpty(formating.NumberFormat))
                style.NumberFormat.Format = formating.NumberFormat;
        }

        private static string FormatConditionValue(object? val)
        {
            if (val == null) return "\"\"";
            if (val is string s) return $"\"{s}\"";
            if (val is DateTime dt) return dt.ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture);
            return Convert.ToString(val, System.Globalization.CultureInfo.InvariantCulture) ?? "\"\"";
        }

        private Dictionary<int, int> Groups = new();

        /// <summary>Группировка столбцов (рекурсия)</summary>
        /// <param name="GroupColIndex">Индекс столбца в списке группировки</param>
        /// <param name="TI">Информация о таблице</param>
        /// <param name="rFrom">Строка в таблице, начало диапазона</param>
        /// <param name="rTo">Строка в таблице, конец диапазона</param>
        /// <param name="row">Строка на листе</param>
        /// <returns>Количество добавленных строк</returns>
        private Int32 GroupsCreate(Int32 GroupColIndex, TableInfo TI, Int32 rFrom, Int32 rTo, Int32 row)
        {
            string GroupCol = TI.Groups[GroupColIndex];
            Int32 colTable = Table.Columns.IndexOf(GroupCol);
            Int32 col = FirstColumnIndex;
            dynamic? GroupValue = null;
            dynamic? GroupValueLast = null;
            Int32 GroupStart = 0;
            Int32 RowsAdded = 0;
            for (Int32 r = rFrom; r <= rTo; r++)
            {
                GroupValue = Table.Rows[r][colTable];
                if (!GroupValue.Equals(GroupValueLast))
                {
                    if (GroupStart != 0)
                    {
                        Groups.Add(GroupStart, row - 1);
                    }
                    GroupValueLast = Table.Rows[r][colTable];
                    Sheet?.Row(row).InsertRowsAbove(1);
                    var value = "";
                    if (GroupValue == null) value = ""; else value = GroupValue.ToString();
                    if (Sheet != null)
                    {
                        Sheet.Cell(row, col).SetValue($"{GroupCol}: {value}");
                        //Sheet.Cell(row, col).Style.Font.Bold = true;
                        FormatingAplly(Sheet.Range(row, col, row, LastColumnIndex), ti.FormatingGroups);
                        //Sheet.Range(row, col, row, LastColumnIndex).Style.Font.Bold = true;
                        //Sheet.Range(row, col, row, LastColumnIndex).Style.Fill.BackgroundColor = XLColor.FromColor(Color.LightGray);
                        Sheet.Range(row, col, row, LastColumnIndex).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                    }
                    row = row + 1;
                    RowsAdded += 1;
                    GroupStart = row;
                    if (GroupColIndex < TI.Groups.Count - 1)
                    {
                        Int32 rTo2 = -1;
                        for (Int32 r2 = r + 1; r2 <= rTo; r2++)
                        {
                            if (GroupValue != null)
                            {
                                if (!GroupValue.Equals(Table.Rows[r2][colTable]))
                                {
                                    rTo2 = r2 - 1;
                                    break;
                                }
                            }
                        }
                        if (rTo2 == -1) rTo2 = rTo;
                        Int32 RowsAdded2 = GroupsCreate(GroupColIndex + 1, TI, r, rTo2, row);
                        row += RowsAdded2;
                        RowsAdded += RowsAdded2;
                    }
                }
                row += 1;
            }
            if (Table.Rows.Count > 0)
            {
                Groups.Add(GroupStart, row - 1);
            }
            return RowsAdded;
        }


        /// <summary>Индекс последнего столбца (с нуля)</summary>
        public Int32 LastColumnIndex
        {
            get
            {
                return FirstColumnIndex + ti.Columns.Count - 1;
            }
        }

        /// <summary>Индекс последней строки (с нуля)</summary>
        public Int32 LastRowIndex
        {
            get
            {
                Int32 H = ShowHeaders ? 1 : 0;
                return FirstRowIndex + Table.Rows.Count + H - 1;
            }
        }

        /// <summary>Диапазон ячеек столбца таблицы</summary>
        /// <param name="ColIndex">Индекс столбца в DataTable</param>
        public IXLRange? GetTableColumnRange(int ColIndex)
        {
            Int32 H = ShowHeaders ? 1 : 0;
            return Sheet?.Range(FirstRowIndex + H, ColIndex, LastRowIndex, ColIndex);
        }

        /// <summary>Диапазон ячеек данных таблицы без заголовка</summary>
        public IXLRange? GetTableDataRange()
        {
            Int32 H = ShowHeaders ? 1 : 0;
            return Sheet?.Range(FirstRowIndex + 1, FirstColumnIndex, LastRowIndex, LastColumnIndex);
        }

        /// <summary>Диапазон ячеек заголовка таблицы</summary>
        public IXLRange? GetTableHeaderRange()
        {
            if (!ShowHeaders)
                return null/* TODO Change to default(_) if this is not a reference type */;
            return Sheet?.Range(FirstRowIndex, FirstColumnIndex, FirstRowIndex, LastColumnIndex);
        }

        /// <summary>Диапазон ячеек строки таблицы</summary>
        /// <param name="ColIndex">Индекс строки в DataTable</param>
        public IXLRange? GetTableRowRange(Int32 RowIndex)
        {
            Int32 H = ShowHeaders ? 1 : 0;
            return Sheet?.Range(FirstRowIndex + H + RowIndex, FirstColumnIndex, FirstRowIndex + H + RowIndex, LastColumnIndex);
        }

        /// <summary>Диапазон ячеек всей таблицы с заголовком</summary>
        public IXLRange? GetFullTableRange()
        {
            Int32 H = ShowHeaders ? 1 : 0;
            return Sheet?.Range(FirstRowIndex, FirstColumnIndex, FirstRowIndex + Table.Rows.Count + H - 1, LastColumnIndex);
        }

        /// <summary>Подготовка таблицы. Порядок, удаление, сортировка и фильтрация столбцов</summary>
        /// <param name="CloumnPosition">Порядок столбцов. Названия столбцов через ","</param>
        /// <param name="ColumnSort">Сортировка столбцов. Названия столбцов через "," (после имени можно указать desc или asc)</param>
        /// <param name="ColumnDelete">Список столбцов для удаления через ","</param>
        /// <param name="RowFilter">Фильтрация строк. Выражение как в DataTable.RowFilter</param>
        public void TablePrepare(string CloumnPosition, string ColumnSort, string ColumnDelete, string RowFilter)
        {
            string[] Cols;
            // удаление столбцов
            if (!string.IsNullOrEmpty(ColumnDelete))
            {
                Cols = ColumnDelete.Split(",");
                foreach (string col in Cols)
                {
                    if (Table.Columns.Contains(col.Trim())) Table.Columns.Remove(col.Trim());
                }
            }

            // порядок столбцов
            if (!string.IsNullOrEmpty(CloumnPosition))
            {
                Cols = CloumnPosition.Split(",");
                Int32 ColIndex = 0;
                foreach (string col in Cols)
                {
                    if (Table.Columns.Contains(col.Trim()))
                    {
                        Table.Columns[col.Trim()]?.SetOrdinal(ColIndex);
                        ColIndex += 1;
                    }
                }
            }

            // сортировка и фильтрация столбцов
            if (!string.IsNullOrEmpty(ColumnSort) | !string.IsNullOrEmpty(RowFilter))
            {
                DataView dv = Table.DefaultView;
                if (ColumnSort != "") dv.Sort = ColumnSort;
                if (RowFilter != "") dv.RowFilter = RowFilter;
                Table = dv.ToTable();
            }
        }
    }
}
