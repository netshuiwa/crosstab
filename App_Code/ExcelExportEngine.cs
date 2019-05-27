using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

/*

public class ExcelExportEngine
{
private string _controlType = "";

private void saveTofle(MemoryStream file, string fileName)
{
    using (FileStream fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.Write))
    {
        byte[] buffer = file.ToArray();//转化为byte格式存储
        fs.Write(buffer, 0, buffer.Length);
        fs.Flush();
        buffer = null;
    }//使用using可以最后不用关闭fs 比较方便
}

#region 创建Sheet
private void CreateSheet()
{
    if (this._controlType == "crosstab")//交叉表
    {
        var style ;
        DataTable dt = new DataTable();
        IList<IQOData> qODatas = _context.AllData; // 查询结果
        Dictionary<string, int> colIndexs = _context.ColIndexs;// 每一列对应的位置
        Dictionary<string, string> fieldMap = new Dictionary<string, string>();
        foreach (string field in colIndexs.Keys)
        {
            dt.Columns.Add(field);
            var prop = _context.QOEntity.GetProperty(field);
            if (prop != null)
                fieldMap.Add(field, prop.Field);
            else
                fieldMap.Add(field, field);
        }

        foreach (IQOData qodata in qODatas)
        {
            DataRow dataRow= dt.NewRow();
            foreach (string key in colIndexs.Keys)
            {
                dataRow[key] = qodata.GetValue(fieldMap[key]).ToString();
            }
            dt.Rows.Add(dataRow);
        }
        Pivot pvt = new Pivot(dt);
        int rowDimentionCount = 0;//行维度数量
        int columnDimentionCount = 0;//列维度数量
        int valueDimentionCount = 0;//值维度数量
        string[] rowDimensions = { "Designation", "Year" };
        string[] columnDimensions = { "Company", "Department", "Name" };
        columnDimentionCount = columnDimensions.Length;
         string[] valueDimensions = { "CTC", "IsActive" };
        valueDimentionCount = valueDimensions.Length;
        bool rowGroup = true; // 行小计
        bool colGroup = true; // 列小计
        bool rowSum = true; // 行合计
        bool colSum = true; // 列合计
        DataTable dtnew = pvt.PivotData(valueDimensions, AggregateFunction.Sum, rowDimensions, columnDimensions, rowGroup, colGroup, rowSum, colSum);

        int columnCount = dtnew.Columns.Count;//列数
            int rowIndex = 0;
        if (title != "")
        {
            //创建标题行
            CreateRow(_sheet, style, rowIndex, columnCount, 20);
            _sheet.GetRow(0).GetCell(0).SetCellValue(_context.FormatSchema.titleOption.title);
            CellRangeAddress cellRangeAddress = new CellRangeAddress(rowIndex , rowIndex , 0, columnCount-1);
            _sheet.AddMergedRegion(cellRangeAddress);
            rowIndex += 1;
        }
                
        //创建副标题行
        if (_context.FormatSchema.titleOption != null && _context.FormatSchema.titleOption.subTitles != null)
            rowIndex += _context.FormatSchema.titleOption.subTitles.FindAll(st => st.type == SubTitleType.SubTitle).Count;
        int index = 0;
        for (int i = _sheet.LastRowNum + 1; i < rowIndex; i++)
        {
            style = GetStyle(_context.FormatSchema.titleOption.subTitles.FindAll(st => st.type == SubTitleType.SubTitle)[index], FontBanner.SubTitle);
            CreateRow(_sheet, style, i, columnCount, _context.FormatSchema.styleOption[FontBanner.SubTitle].rowHeight);
            _sheet.GetRow(i).GetCell(0).SetCellValue(_context.FormatSchema.titleOption.subTitles[i-1].text);
            CellRangeAddress cellRangeAddress = new CellRangeAddress(i, i, 0, columnCount - 1);
            _sheet.AddMergedRegion(cellRangeAddress);
            index++;
        }
        string[] rows = dtnew.Columns[rowDimentionCount].ColumnName.Split('.');//标题维度信息
        style= GetStyle(FontBanner.Header);
        int rowLength = rows.Length;
        for (int i = 0; i < rowLength; i++)
        {
            CreateRow(_sheet, style, i+ rowIndex, columnCount, _context.FormatSchema.styleOption[FontBanner.Header].rowHeight);
        }
        //创建标题行
        for (int j = 0; j < columnCount; j++)
        {
            for (int i = 0; i < rowLength; i++)
            {
                string[] currentvalue = dtnew.Columns[j].ColumnName.Split('.');
                if (j < rowDimentionCount)
                {
                    _sheet.GetRow(i + rowIndex).GetCell(j).SetCellValue(rowDimension[j].name);
                }
                else
                {
                    if (i == rowLength-1)
                    {
                        int currentcol = (j - rowDimentionCount) % valueDimentionCount;
                        _sheet.GetRow(i + rowIndex).GetCell(j).SetCellValue(valueDimension[currentcol].name);
                    }
                    else
                    {
                        _sheet.GetRow(i + rowIndex).GetCell(j).SetCellValue(currentvalue[i]);
                    }
                }

            }
            //合并行标题上的行标题数据
            if (j < rowDimentionCount)
            {
                CellRangeAddress cellRangeAddress = new CellRangeAddress(rowIndex, rowIndex+rowLength - 1, j, j);
                _sheet.AddMergedRegion(cellRangeAddress);
            }
        }
        int rowCount = dtnew.Rows.Count;//行数
        for (int i = 0; i < rowCount; i++)
        {
            CreateRow(_sheet, null, i + rowLength + rowIndex, columnCount, _context.FormatSchema.styleOption[FontBanner.DataArea].rowHeight);
        }
        //填充数据
        for (int j = 0; j < columnCount; j++)
        {
            for (int i = 0; i < rowCount; i++)
            {
                string currentvalue = dtnew.Rows[i][j].ToString();
                _sheet.GetRow(i + rowLength + rowIndex).GetCell(j).SetCellValue(currentvalue);
                //行标题样式
                Align align = new Align();
                style = GetStyle(FontBanner.DataArea);
                if (j < rowDimentionCount)
                {
                    currentvalue = dtnew.Rows[i][j].ToString();
                    align = rowDimension[j].align;
                }
                else
                {
                    int currentcol = (j - rowDimentionCount) % valueDimentionCount;
                    //currentvalue = valueDimension[currentcol].name;
                    align =valueDimension[currentcol].align;
                }
                switch (align)
                {
                    case Align.Center:
                        style.Alignment = HorizontalAlignment.Center;
                        break;
                    case Align.Left:
                        style.Alignment = HorizontalAlignment.Left;
                        break;
                    case Align.Right:
                        style.Alignment = HorizontalAlignment.Right;
                        break;
                    default:
                        break;
                }
                _sheet.GetRow(i + rowLength + rowIndex).GetCell(j).CellStyle = style;
            }
        }
        //合并单元格
        //合计列总计
        if (colSum)
        {
            CellRangeAddress cellRangeAddress = new CellRangeAddress(rowIndex, rowIndex+rowLength - 2, columnCount - valueDimentionCount, columnCount - 1);
            _sheet.AddMergedRegion(cellRangeAddress);
        }
        //合并行标题的列标题数据
        MergeRowWorkSheet(rowIndex, 0, rowLength, columnCount, valueDimentionCount,rowGroup);
        //合计行总计
        if (rowSum)
        {
            CellRangeAddress cellRangeAddress = new CellRangeAddress(rowIndex+rowLength + rowCount - 1, rowIndex + rowLength + rowCount - 1, 0, rowDimentionCount - 1);
            _sheet.AddMergedRegion(cellRangeAddress);
        }
        //合并列标题的行标题数据
        MergeWorkSheet(rowLength+ rowIndex, 0, rowCount, rowDimentionCount,colGroup);
    }
    else
    {
        style = GetStyle(FontBanner.Title);
        //创建标题行
        CreateRow(_sheet, style, rowIndex, _context.ColumnCount, _context.FormatSchema.styleOption[FontBanner.Title].rowHeight);
        rowIndex += 1;
        //创建副标题行
        if (_context.FormatSchema.titleOption != null && _context.FormatSchema.titleOption.subTitles != null)
            rowIndex += _context.FormatSchema.titleOption.subTitles.FindAll(st => st.type == SubTitleType.SubTitle).Count;
        int index = 0;
        for (int i = _sheet.LastRowNum + 1; i < rowIndex; i++)
        {
            style = GetStyle(_context.FormatSchema.titleOption.subTitles.FindAll(st => st.type == SubTitleType.SubTitle)[index], FontBanner.SubTitle);
            CreateRow(_sheet, style, i, _context.ColumnCount, _context.FormatSchema.styleOption[FontBanner.SubTitle].rowHeight);
            index++;
        }
        rowIndex += _context.ColumnHeaderRowCount;
        //创建列头行
        style = GetStyle(FontBanner.Header);
        for (int i = _sheet.LastRowNum + 1; i < rowIndex; i++)
        {
            CreateRow(_sheet, style, i, _context.ColumnCount, _context.FormatSchema.styleOption[FontBanner.Header].rowHeight);
        }
        //创建数据区域行
        int dataRowCount = this._context.AllData != null ? this._context.AllData.Count : 0;
        rowIndex += dataRowCount;
        for (int i = _sheet.LastRowNum + 1; i < rowIndex; i++)
        {
            CreateRow(_sheet, null, i, _context.ColumnCount, _context.FormatSchema.styleOption[FontBanner.DataArea].rowHeight);
        }
        //创建表尾行
        if (_context.FormatSchema.titleOption != null && _context.FormatSchema.titleOption.subTitles != null)
            rowIndex += _context.FormatSchema.titleOption.subTitles.FindAll(st => st.type == SubTitleType.Footer).Count;
        index = 0;
        for (int i = _sheet.LastRowNum + 1; i < rowIndex; i++)
        {
            style = GetStyle(_context.FormatSchema.titleOption.subTitles.FindAll(st => st.type == SubTitleType.Footer)[index], FontBanner.Footer);
            CreateRow(_sheet, style, _sheet.LastRowNum + 1, _context.ColumnCount, _context.FormatSchema.styleOption[FontBanner.Footer].rowHeight);
            index++;
        }
    }
}
#endregion

#region 生成标题
/// <summary>
/// 生成标题
/// </summary>
/// <param name="sheet"></param>
/// <param name="startRow"></param>
/// <param name="startColumn"></param>
private void GenerateTitle(ISheet sheet, int startRow, int startColumn)
{
    SetMergeCell(sheet, startRow, startRow, startColumn, startColumn + _context.ColumnCount - 1);
    sheet.GetRow(startRow).GetCell(startColumn).SetCellValue(_context.FormatSchema.titleOption.title);
}
#endregion

#region 生成副标题
/// <summary>
/// 生成副标题
/// </summary>
/// <param name="sheet"></param>
/// <param name="startRow"></param>
/// <param name="startColumn"></param>
private void GenerateSubTitle(ISheet sheet, int startRow, int startColumn)
{
    List<SubTitle> subTitleList = _context.FormatSchema.titleOption.subTitles.FindAll(st => st.type == SubTitleType.SubTitle);
    if (subTitleList.Count > 0)
    {
        subTitleList.Sort((a, b) => a.level.CompareTo(b.level));
        foreach (SubTitle subTitle in subTitleList)
        {
            SetMergeCell(sheet, startRow, startRow, startColumn, startColumn + _context.ColumnCount - 1);
            string subTitleValue = subTitle.text;
            sheet.GetRow(startRow).GetCell(startColumn).SetCellValue(subTitleValue);//SetCell(sheet.GetRow(startRow).GetCell(startColumn), subTitle.text, subTitle, null, FontBanner.SubTitle, 0);
            startRow++;
        }
    }
}
#endregion

#region 生成表头
/// <summary>
/// 生成表头
/// </summary>
/// <param name="sheet"></param>
/// <returns></returns>
private void GenerateTableHeader(ISheet sheet)
{
    int colIndex = 0;
    int rowIndex = 1;
    if (_context.FormatSchema.titleOption != null && _context.FormatSchema.titleOption.subTitles != null)
        rowIndex += _context.FormatSchema.titleOption.subTitles.FindAll(st => st.type == SubTitleType.SubTitle).Count;
    if (_context.ColumnHeaderRowCount == 1)
    {
        #region 无多级表头
        foreach (Column col in _context.FormatSchema.columnOption.colList)
        {
            sheet.GetRow(rowIndex).GetCell(colIndex).SetCellValue(col.name);//SetCell(sheet.GetRow(rowIndex).GetCell(colIndex), col.name, null, col, FontBanner.Header, 0);
            colIndex++;
        }
        #endregion
    }
    else
    {
        #region 多级表头
        GetTableHeader(sheet, _context.FormatSchema.columnOption.colList, rowIndex, colIndex, 1);
        #endregion
    }
}
#endregion

#region 递归生成表头数据
/// <summary>
/// 递归生成表头数据
/// </summary>
/// <param name="sheet"></param>
/// <param name="colList"></param>
/// <param name="rowIndex"></param>
/// <param name="colIndex"></param>
/// <param name="level"></param>
private void GetTableHeader(ISheet sheet, List<Column> colList, int rowIndex, int colIndex, int level)
{
    foreach (Column col in colList)
    {
        if (col.childList != null && col.childList.Count > 0)
        {
            SetMergeCell(sheet, rowIndex, rowIndex, colIndex, colIndex + ExportHelper.CalculateColumnCount(col.childList, col.childList.Count) - 1);
            sheet.GetRow(rowIndex).GetCell(colIndex).SetCellValue(col.name);
            GetTableHeader(sheet, col.childList, rowIndex + 1, colIndex, level + 1);
        }
        else
        {
            //if (col.isFixed)
            //{
            //    int frozenColCount = colIndex + 1;
            //    if (frozenColCount > _context.FrozenColumnCount) _context.FrozenColumnCount = frozenColCount;
            //}
            if (level > 1 && level < _context.ColumnHeaderRowCount)
                SetMergeCell(sheet, rowIndex, rowIndex + _context.ColumnHeaderRowCount - level, colIndex, colIndex);
            else if (level == _context.ColumnHeaderRowCount)
                SetMergeCell(sheet, rowIndex, rowIndex, colIndex, colIndex);
            else if (level == 1)
                SetMergeCell(sheet, rowIndex, rowIndex + _context.ColumnHeaderRowCount - 1, colIndex, colIndex);
        }
        sheet.GetRow(rowIndex).GetCell(colIndex).SetCellValue(col.name);// SetCell(sheet.GetRow(rowIndex).GetCell(colIndex), col.name, null, col, FontBanner.Header, 0);
        if (col.childList != null && col.childList.Count > 0)
        {
            colIndex += ExportHelper.CalculateColumnCount(col.childList, col.childList.Count);
        }
        else
        {
            colIndex++;
        }
    }
}
#endregion

#region 生成表体数据
private void GenerateTableBody(ISheet sheet,int rowIndex)
{
    #region 设置列宽、显示隐藏、 固定列、排序、对齐方式、
    _context.FormatSchema.columnOption.colList.ForEach(col =>
    {
        int colIndex = _context.ColIndexs[col.bindField];
        sheet.SetColumnHidden(colIndex, !col.visible);
        sheet.SetColumnWidth(colIndex, (int)(col.colWidth*36.6));// px转列宽单位
        ICellStyle cellStyle = sheet.GetColumnStyle(colIndex+1);//TODO 验证取得列
        switch (col.align)
        {
            case Align.Center:
                cellStyle.Alignment = HorizontalAlignment.Center;
                break;
            case Align.Left:
                cellStyle.Alignment = HorizontalAlignment.Left;
                break;
            case Align.Right:
                cellStyle.Alignment = HorizontalAlignment.Right;
                break;
            default:
                break;
        }
    });
    #endregion

    if (_context.DataRowCount == 0)
    {
        return;
    }

    IList<IQOData> qODatas = _context.AllData; // 查询结果
    Dictionary<string, int> colIndexs = _context.ColIndexs;// 每一列对应的位置
    Dictionary<string, string> fieldMap = new Dictionary<string, string>();
    foreach(string field in colIndexs.Keys)
    {
        var prop = _context.QOEntity.GetProperty(field);
        if (prop != null)
            fieldMap.Add(field, prop.Field);
        else
            fieldMap.Add(field, field);
    }

    foreach (IQOData qodata in qODatas)
    {
        foreach(string key in colIndexs.Keys)
        {
            ICell cell = sheet.GetRow(rowIndex).GetCell(colIndexs[key]);
                    
            //cell.SetCellValue((string)qodata.GetValue(key));//TODO 确认QOData里的值类型
            cell.SetCellValue(qodata.GetValue(fieldMap[key]).ToString());//TODO 确认QOData里的值类型
        }
        rowIndex++;
    }
}
#endregion

#region 生成表尾
/// <summary>
/// 生成表尾
/// </summary>
/// <param name="startRow"></param>
/// <param name="startColumn"></param>
/// <param name="sheet"></param>
private void GenerateTableFooter(ISheet sheet, int startRow, int startColumn)
{
    List<SubTitle> footerList = _context.FormatSchema.titleOption.subTitles.FindAll(st => st.type == SubTitleType.Footer);
    if (footerList.Count > 0)
    {
        footerList.Sort((a, b) => a.level.CompareTo(b.level));
        foreach (SubTitle tableFooter in footerList)
        {
            SetMergeCell(sheet, startRow, startRow, startColumn, startColumn + _context.ColumnCount - 1);
            string footerValue = tableFooter.text;
            sheet.GetRow(startRow).GetCell(startColumn).SetCellValue(footerValue);
            startRow++;
        }
    }
}
#endregion




#region 获取样式
/// <summary>
/// 获取样式
/// </summary>
/// <returns></returns>
private ICellStyle GetStyle(FontBanner fontBanner)
{
    ICellStyle style = _excelWorkbook.CreateCellStyle();
    style.Alignment = HorizontalAlignment.Center;
    style.VerticalAlignment = VerticalAlignment.Center;
    style.SetFont(_fontDictionary[fontBanner]);

    return style;
}

/// <summary>
/// 获取样式
/// </summary>
/// <returns></returns>
private ICellStyle GetStyle(SubTitle subTitle, FontBanner fontBanner)
{
    ICellStyle style = _excelWorkbook.CreateCellStyle();
    style.Alignment = HorizontalAlignment.Center;
    style.VerticalAlignment = VerticalAlignment.Center;
    switch (subTitle.align)
    {
        case Align.Left:
            style.Alignment = HorizontalAlignment.Left;
            break;
        case Align.Center:
            style.Alignment = HorizontalAlignment.Center;
            break;
        case Align.Right:
            style.Alignment = HorizontalAlignment.Right;
            break;
    }
    style.SetFont(_fontDictionary[fontBanner]);

    return style;
}
#endregion

#region 获取字体
/// <summary>
/// 获取字体
/// </summary>
/// <param name="fontBanner"></param>
/// <returns></returns>
private IFont GetFont(FontBanner fontBanner)
{
    StyleOption styleOption = this._context.FormatSchema.styleOption[fontBanner];
    IFont font = this._excelWorkbook.CreateFont();
    font.FontName = Enum.GetName(typeof(WebQuery.Api.Schema.FormatSchema.FontFamily), styleOption.fontFamily);// 设置字体
    font.FontHeightInPoints = (short)styleOption.size;// 设置大小
    font.IsBold = styleOption.fontWeight == FontWeight.bold ? true : false;// 设置加粗
    font.IsItalic = styleOption.fontStyle == FontStyle.italic ? true:false;// 设置斜体
    switch (styleOption.textDecoration) // 设置下划线
    {
        case WebQuery.Api.Schema.Spread.TextDecorationType.None:
            font.Underline = FontUnderlineType.None;
            break;
        case WebQuery.Api.Schema.Spread.TextDecorationType.Underline:
            font.Underline = FontUnderlineType.Single;
            break;
        default:
            font.Underline = FontUnderlineType.None;
            break;
    }
    // 设置字体颜色
    PaletteRecord paletteRecord = new PaletteRecord();
    HSSFPalette palette = new HSSFPalette(paletteRecord);
    System.Drawing.Color color = ExportHelper.ColorHx16ToRgb(styleOption.fontColor);
    HSSFColor xlColour = palette.FindSimilarColor(color.R, color.G, color.B);
    font.Color = xlColour.Indexed;
    return font;
}
#endregion

#region 创建行
private void CreateRow(ISheet sheet, ICellStyle style, int rowIndex, int colCount, float rowHeight)
{
    IRow row = sheet.CreateRow(rowIndex);
    row.HeightInPoints = rowHeight;//UtilConverter.MillimetersToPoints(rowHeight / 10f);
    for (int i = 0; i < colCount; i++)
    {
        row.CreateCell(i);
        if (style != null) row.GetCell(i).CellStyle = style;
    }
}
#endregion

#region 合并单元格
/// <summary>
/// 合并单元格
/// </summary>
/// <param name="sheet"></param>
/// <param name="firstRow"></param>
/// <param name="lastRow"></param>
/// <param name="firstCol"></param>
/// <param name="lastCol"></param>
private void SetMergeCell(ISheet sheet, int firstRow, int lastRow, int firstCol, int lastCol)
{
    CellRangeAddress cellRangeAddress = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
    sheet.AddMergedRegion(cellRangeAddress);
}
#endregion

    /// 合并工作表中指定行数和列数数据相同的单元格
    /// </summary>
    /// <param name="sheetIndex">工作表索引</param>
    /// <param name="beginRowIndex">开始行索引</param>
    /// <param name="beginColumnIndex">开始列索引</param>
    /// <param name="rowCount">要合并的行数</param>
    /// <param name="columnCount">要合并的列数</param>
public void MergeWorkSheet(int beginRowIndex, int beginColumnIndex, int rowCount, int columnCount,bool colGroup)
{

    //检查参数
    if (columnCount < 1 || rowCount < 1)
        return;

    for (int col = 0; col < columnCount; col++)
    {
        int mark = 0;            //标记比较数据中第一条记录位置
        int mergeCount = 1;        //相同记录数，即要合并的行数
        string text = "";

        for (int row = 0; row < rowCount; row++)
        {
            string prvName = "";
            string nextName = "";
            string prvLastColName = "";
            string nextLastColName = "";

            //最后一行不用比较
            if (row + 1 < rowCount)
            {

                if(col> beginColumnIndex)
                {
                    prvLastColName = _sheet.GetRow(beginRowIndex + row).Cells[col-1].ToString();

                    nextLastColName = _sheet.GetRow(beginRowIndex + row + 1).Cells[col-1].ToString();
                }
                else
                {
                    prvLastColName = nextLastColName;
                }
                prvName = _sheet.GetRow(beginRowIndex + row).Cells[col].ToString();

                nextName = _sheet.GetRow(beginRowIndex + row + 1).Cells[col].ToString();

                if (prvName == nextName && prvLastColName== nextLastColName)
                {
                    mergeCount++;

                    if (row == rowCount - 2)
                    {
                        CellRangeAddress cellRangeAddress = new CellRangeAddress(beginRowIndex + mark, beginRowIndex + mark + mergeCount - 1
                            , beginColumnIndex + col, beginColumnIndex + col);
                        _sheet.AddMergedRegion(cellRangeAddress);
                    }
                    else if (colGroup &&mergeCount == 1 && prvName == "小计")
                    {
                        int endCol = beginRowIndex + columnCount - col;
                        if (endCol < beginColumnIndex + col) endCol = beginRowIndex + col;
                        //合并列上的小计
                        CellRangeAddress cellRangeAddress = new CellRangeAddress(beginRowIndex + mark, beginRowIndex + mark + mergeCount - 1
                            , beginColumnIndex + col, endCol);
                        _sheet.AddMergedRegion(cellRangeAddress);
                        mergeCount = 1;
                        mark = col + 1;
                    }
                }
                else
                {
                    CellRangeAddress cellRangeAddress = new CellRangeAddress(beginRowIndex + mark, beginRowIndex + mark + mergeCount - 1, beginColumnIndex + col
                        , beginColumnIndex + col);
                    _sheet.AddMergedRegion(cellRangeAddress);
                    mergeCount = 1;
                    mark = row + 1;
                }

            }
        }
    }
}

/// 合并工作表中指定行数和列数数据相同的单元格
/// </summary>
/// <param name="beginRowIndex">开始行索引</param>
/// <param name="beginColumnIndex">开始列索引</param>
/// <param name="rowCount">要合并的行数</param>
/// <param name="columnCount">要合并的列数</param>
/// <param name="columnCount">值列数</param>
public void MergeRowWorkSheet(int beginRowIndex, int beginColumnIndex, int rowCount, int columnCount, int valueCount,bool rowGroup)
{

    //检查参数
    if (columnCount < 1 || rowCount < 1)
        return;
    for (int row = beginRowIndex; row < beginRowIndex+rowCount; row++)
    {
        int mark = 0;            //标记比较数据中第一条记录位置
        int mergeCount = 1;        //相同记录数，即要合并的行数
        string text = "";
        for (int col = 0; col < columnCount; col++)
        {
            string prvName = "";
            string prvLastRowName = "";
            string nextName = "";
            string nextLastRowName = "";

            //最后一行不用比较
            if (col + 1 < columnCount)
            {
                if(row> beginRowIndex)
                {
                    prvLastRowName = _sheet.GetRow(row-1).Cells[col].ToString();

                    nextLastRowName = _sheet.GetRow(row-1).Cells[col + 1].ToString();
                }
                else
                {
                    prvLastRowName = nextLastRowName;
                }

                prvName = _sheet.GetRow(row).Cells[col].ToString();

                nextName = _sheet.GetRow(row).Cells[col + 1].ToString();

                if (prvName == nextName && prvLastRowName == nextLastRowName)
                {
                    mergeCount++;

                    if (col == columnCount - 2)
                    {
                        CellRangeAddress cellRangeAddress = new CellRangeAddress(row, row, beginColumnIndex + mark, beginColumnIndex + mark + mergeCount - 1);
                        _sheet.AddMergedRegion(cellRangeAddress);
                    }
                    else if (rowGroup  && mergeCount == valueCount && prvName == "小计")
                    {
                        int endRow = beginRowIndex + rowCount - row - 1;
                        if (endRow < beginRowIndex + row) endRow = row;
                        //合并列上的小计
                        CellRangeAddress cellRangeAddress = new CellRangeAddress(row, endRow, beginColumnIndex + mark, beginColumnIndex + mark + mergeCount - 1);
                        _sheet.AddMergedRegion(cellRangeAddress);
                        mergeCount = 1;
                        mark = col + 1;
                    }

                }
                else if (prvName == "总计")
                {
                    mergeCount = 1;
                    mark = col + 1;
                }
                else
                {
                    CellRangeAddress cellRangeAddress = new CellRangeAddress(row, row, beginColumnIndex + mark, beginColumnIndex + mark + mergeCount - 1);
                    _sheet.AddMergedRegion(cellRangeAddress);
                    mergeCount = 1;
                    mark = col + 1;
                }

            }
        }
    }
}
}
*/