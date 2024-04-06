using Microsoft.Extensions.Hosting;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System.Collections.Generic;
using System.Drawing;

var builder = WebApplication.CreateBuilder(args);
var app = builder.Build();

test t = new test();
t.CreateATest("test.xls");
app.MapGet("/", () => "Hello World!");

app.Run();

public class test
{
    public void CreateATest(string filename)
    {
        FileStream fs = new FileStream(filename, FileMode.Create, FileAccess.Write);
        HSSFWorkbook wb = new HSSFWorkbook();
        ISheet sheet = wb.CreateSheet("NPOI");
        sheet.DisplayGridlines = false;
        
        IRow row = sheet.CreateRow(0);
        row.RowStyle = wb.CreateCellStyle();
        row.RowStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
        row.RowStyle.VerticalAlignment = VerticalAlignment.Center;
        //row.RowStyle.WrapText = true;			
        IFont font = wb.CreateFont();			
        font.IsBold = true;			
        font.Color = (short)ColorTranslator.ToWin32(Color.Red);			
        row.RowStyle.SetFont(font);

        //HSSFCellStyle headerStyle = (HSSFCellStyle)wb.CreateCellStyle();
        //headerStyle.FillForegroundColor = IndexedColors.LightBlue.Index;
        //headerStyle.FillPattern = FillPattern.SolidForeground;

        HSSFCellStyle styletop = (HSSFCellStyle)wb.CreateCellStyle();
        styletop.BorderTop = BorderStyle.Thick;
        styletop.FillForegroundColor = IndexedColors.Grey25Percent.Index;
        styletop.FillPattern = FillPattern.SolidForeground;

        HSSFCellStyle stylelefttop = (HSSFCellStyle)wb.CreateCellStyle();
        stylelefttop.BorderLeft = BorderStyle.Thick;
        stylelefttop.BorderTop = BorderStyle.Thick;
        stylelefttop.FillForegroundColor = IndexedColors.Grey25Percent.Index;
        stylelefttop.FillPattern = FillPattern.SolidForeground;

        HSSFCellStyle stylerttop = (HSSFCellStyle)wb.CreateCellStyle();
        stylerttop.BorderRight = BorderStyle.Thick;
        stylerttop.BorderTop = BorderStyle.Thick;
        stylerttop.FillForegroundColor = IndexedColors.Grey25Percent.Index;
        stylerttop.FillPattern = FillPattern.SolidForeground;

        HSSFCellStyle styleleftbottom = (HSSFCellStyle)wb.CreateCellStyle();
        styleleftbottom.BorderLeft = BorderStyle.Thick;
        styleleftbottom.BorderBottom = BorderStyle.Thick;

        HSSFCellStyle stylertbottom = (HSSFCellStyle)wb.CreateCellStyle();
        stylertbottom.BorderRight = BorderStyle.Thick;
        stylertbottom.BorderBottom = BorderStyle.Thick;

        HSSFCellStyle styleleft = (HSSFCellStyle)wb.CreateCellStyle();
        styleleft.BorderLeft = BorderStyle.Thick;


        HSSFCellStyle styleright = (HSSFCellStyle)wb.CreateCellStyle();
        styleright.BorderRight = BorderStyle.Thick;

        HSSFCellStyle stylebottom = (HSSFCellStyle)wb.CreateCellStyle();
        stylebottom.BorderBottom = BorderStyle.Thick;

        int i = 0;	
        foreach (string header in new string[] { "ID", "Name", "Age" })
        {
            ICell cell = row.CreateCell(i++);
            cell.SetCellValue(header);
            
            cell.CellStyle = styletop;
            if (cell.ColumnIndex == 0)
            {
                cell.CellStyle = stylelefttop;
            }
            if (cell.ColumnIndex == 2)
            {
                cell.CellStyle = stylerttop;
            }
            
        }
        Random rand = new Random();
        ICell? cell1 = null;
        IRow? row1 = null;
        for (i = 1; i < 5; i++)
        {
            row1 = sheet.CreateRow(i);
            for (int j = 0; j < 3; j++)
            {
                cell1 = row1.CreateCell(j);
                cell1.SetCellValue(rand.Next());
                cell1.CellStyle.FillPattern = FillPattern.NoFill;
                if (cell1!.ColumnIndex == 0)
                {
                    cell1!.CellStyle = styleleft;
                }
                if (cell1.ColumnIndex == 2)
                {
                    cell1!.CellStyle = styleright;
                }
                if (row1!.RowNum == 4 && cell1!.ColumnIndex == 1)
                {
                    cell1!.CellStyle = stylebottom;
                }
                if (row1!.RowNum == 4 && cell1!.ColumnIndex == 0)
                {
                    cell1!.CellStyle = styleleftbottom;
                }
                if (row1!.RowNum == 4 && cell1!.ColumnIndex == 2)
                {
                    cell1!.CellStyle = stylertbottom;
                }
            }
        }

        wb.Write(fs);
        fs.Close();
    }
}
