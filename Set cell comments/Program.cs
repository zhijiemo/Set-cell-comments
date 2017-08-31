using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;


namespace Set_cell_comments
{
    class Program
    {
        static void Main(string[] args)
        {
            SLDocument sl = new SLDocument();

            sl.SetCellValue(1, 1, "Commenting on comments...");
            sl.SetCellValue(2, 1, "Commenting on comments...");
            sl.SetCellValue(3, 1, "Commenting...");
            sl.SetCellValue(4, 1, "COMMenting...");
            sl.SetCellValue(5, 1, "COMMENTING!");
            sl.SetCellValue(6, 1, "*ding*");

            SLComment comm;

            // this is probably the bare minimum you need to get a comment.
            // Note that if you don't set the position, it will be automatically placed near
            // the cell it's attached to.
            comm = sl.CreateComment();
            comm.SetText("first!!111!1!!1!!");
            sl.InsertComment(2, 6, comm);

            // This simulates the "typical" Excel comment look.
            SLFont font = sl.CreateFont();
            font.SetFont("Tahoma", 9);
            font.Bold = true;//设定字体是否需要加粗
            SLRstType rst = sl.CreateRstType();//SLResType类封装了用于处理多种字符串类型的属性和方法
            rst.AppendText("Karen:\n", font);
            rst.AppendText("We have a troll! That is so immature...");//附加当前主题的小字体和默认字体大小格式的文本。
            comm = sl.CreateComment();
            comm.SetText(rst);
            sl.InsertComment(2, 10, comm);

            // You don't have to set an author explicitly. SpreadsheetLight will use this property
            // sl.DocumentProperties.Creator
            // if it's set. Otherwise, a default author name will be used. Speaking of which...

            comm = sl.CreateComment();
            comm.Author = "Ben";
            // You don't have to include the author name in the comment if you don't want to.
            // If you do, just follow the above code to set it explicitly.
            comm.SetText("Calm down Karen. We'll take care of that guy.");
            sl.InsertComment(7, 6, comm);

            comm = sl.CreateComment();
            comm.SetText("Why am I so far here? I wanna be with you guys!");
            // Positions depend on the computer's DPI.
            // Th resulting spreadsheet (if you downloaded it) is generated on a 120 DPI monitor,
            // so if you view it on 96 DPI, the positions of all the comments will likely be slightly off.
            // Positions are measured from the top-left cell.
            // This means to put the comment box at the top-left corner of row 16, column 14
            comm.SetPosition(16, 14);
            // You might want to set Visible to true. If invisible, Excel shows the comment
            // close to where it's anchored.
            sl.InsertComment(7, 10, comm);

            comm = sl.CreateComment();
            // widths and heights are measured in points. For convenience you can use
            // SLConvert.FromInchToPoint() if you're using imperial units and
            // SLConvert.FromCentimeterToPoint() if you're using metric units.
            comm.Width = 240;
            comm.Height = 180;
            comm.SetText("Can't eat anymore... I'm bloated...");
            sl.InsertComment(12, 6, comm);

            // in case you want to fit the contents snugly
            comm = sl.CreateComment();
            comm.SetText("These jeans fit me perfectly!");
            comm.AutoSize = true;
            sl.InsertComment(12, 10, comm);

            rst = sl.CreateRstType();
            font = sl.CreateFont();
            font.SetFont("Harrington", 16);
            rst.AppendText("Don't envy", font);
            font = sl.CreateFont();
            font.Bold = true;//加粗
            font.Italic = true;//斜体
            font.Strike = true;//删除线
            font.VerticalAlignment = VerticalAlignmentRunValues.Superscript;//上标
            rst.AppendText(" me because", font);
            font = sl.CreateFont();
            font.Underline = UnderlineValues.Single;
            font.SetFontThemeColor(SLThemeColorIndexValues.Accent4Color);
            rst.AppendText(" I've got style...", font);
            rst.AppendText(" and you don't.");
            comm = sl.CreateComment();
            comm.SetText(rst);
            sl.InsertComment(17, 6, comm);

            // comment text alignment
            comm = sl.CreateComment();
            comm.HorizontalTextAlignment = SLHorizontalTextAlignmentValues.Right;
            comm.VerticalTextAlignment = SLVerticalTextAlignmentValues.Center;
            comm.SetText("Stop manhandling me!");
            sl.InsertComment(17, 10, comm);

            // the top-down orientation
            comm = sl.CreateComment();
            comm.Orientation = SLCommentOrientationValues.TopDown;//文本上下朝向
            // set larger size so all the comment text can be displayed.
            comm.Width = 160;
            comm.Height = 120;
            comm.SetText("I read like those ancient Chinese texts, yah?");
            sl.InsertComment(22, 6, comm);

            // another orientation
            comm = sl.CreateComment();
            comm.Orientation = SLCommentOrientationValues.Rotated270Degrees;//旋转270°
            comm.SetText("I'm getting dizzy...");
            sl.InsertComment(22, 10, comm);

            // when you want to show comments. The default is to hide them.
            comm = sl.CreateComment();
            comm.Visible = true;
            comm.SetText("How come everyone's got the invisibility superpower?");
            sl.InsertComment(27, 6, comm);

            // when you want to style the lines surrounding the comment box
            comm = sl.CreateComment();
            comm.LineColor = System.Drawing.Color.Coral;//注释框颜色：珊瑚红
            comm.LineStyle = DocumentFormat.OpenXml.Vml.StrokeLineStyleValues.ThickBetweenThin;
            // this is in points
            comm.LineWeight = 5;
            comm.SetText("I've got fancy outlines.");
            sl.InsertComment(27, 10, comm);

            // for another fancy outline
            comm = sl.CreateComment();
            comm.SetDashStyle(SLDashStyleValues.LongDashDotDot);//线条样式是长划线
            comm.SetText("Do I look like Morse code to you?");
            sl.InsertComment(32, 6, comm);

            // for when you don't want a shadow for the comment box
            comm = sl.CreateComment();
            comm.HasShadow = false;//阴影颜色：右边和下边没有粗线
            comm.SetText("Erhmahgerd! I got no shadow! Who did this to me?");
            sl.InsertComment(32, 10, comm);

            // for when you want a different shadow colour. 
            comm = sl.CreateComment();
            comm.ShadowColor = System.Drawing.Color.HotPink;
            rst = sl.CreateRstType();
            rst.AppendText("You think ");
            rst.AppendText(" you've", new SLFont() { Italic = true });
            rst.AppendText(" got a problem? I got a pink shadow. Who has ");
            rst.AppendText("pink", new SLFont() { Italic = true, FontColor = System.Drawing.Color.HotPink });
            rst.AppendText(" shadows?");
            comm.SetText(rst);
            sl.InsertComment(37, 6, comm);

            // Next comes the background fill section. Note that the Fill property is repurposed
            // from somewhere else, and will work as intended most of the time.//接下来是背景填充部分。请注意，填充属性是从其他地方重新定义的
            // This is a limitation of the underlying VML properties, and not SpreadsheetLight.
            // For instance, the actual colour of accent colours is captured. But if you change
            // themes (and thus accent colours), the colour you used won't change automatically.

            // for when you don't want a background fill colour
            // This probably work better if you customise the shadow colour too.
            comm = sl.CreateComment();
            comm.Fill.SetNoFill();
            comm.SetText("My life is so empty...");
            sl.InsertComment(37, 10, comm);

            // for an automatic background colour. This is probably just white.
            comm = sl.CreateComment();
            comm.Fill.SetAutomaticFill();
            comm.SetText("Ooh this cup just filled up by itself! ... Can I have chocolate instead?");
            sl.InsertComment(42, 6, comm);

            // the default colour is #ffffe1
            comm = sl.CreateComment();
            // 20% transparency
            comm.Fill.SetSolidFill(System.Drawing.Color.LightSkyBlue, 20);//20%透明度
            comm.SetText("The sky's the limit!");
            sl.InsertComment(42, 10, comm);

            // linear gradients//线性渐变填充
            comm = sl.CreateComment();
            // 40% transparency on the first gradient point   第一个渐变点的透明度为40%
            comm.GradientFromTransparency = 40;
            // 80% transparency on the last gradient point
            comm.GradientToTransparency = 80;
            // 45 degrees, so gradient is from top-left to bottom-right
            comm.Fill.SetLinearGradient(SpreadsheetLight.Drawing.SLGradientPresetValues.Rainbow, 45);
            comm.SetText("I'm a unicorn! I've got rainbows coming out the wazoo!");
            sl.InsertComment(47, 6, comm);

            // path gradients  路径渐变
            comm = sl.CreateComment();
            comm.Fill.SetPathGradient(SpreadsheetLight.Drawing.SLGradientPresetValues.Ocean);
            comm.SetText("My gradients are so squarish...");
            sl.InsertComment(47, 10, comm);

            // fancier gradients
            comm = sl.CreateComment();
            comm.Fill.SetRectangularGradient(SpreadsheetLight.Drawing.SLGradientPresetValues.LateSunset, SpreadsheetLight.Drawing.SLGradientDirectionValues.CenterToBottomLeftCorner);
            // for the purposes of setting gradients, the following
            //comm.Fill.SetRadialGradient(SpreadsheetLight.Drawing.SLGradientPresetValues.LateSunset, SpreadsheetLight.Drawing.SLGradientDirectionValues.CenterToBottomLeftCorner);
            // is the same as SetRectangularGradient(). The technical explanation is that VML don't
            // support circular gradients...
            comm.SetText("Oh stop complaining... my gradients aren't that hot either.");
            sl.InsertComment(52, 6, comm);

            // pattern fills!
            comm = sl.CreateComment();
            comm.Fill.SetPatternFill(DocumentFormat.OpenXml.Drawing.PresetPatternValues.Wave, System.Drawing.Color.GhostWhite, System.Drawing.Color.LightBlue);
            comm.SetText("I'm riding the wave!");
            sl.InsertComment(52, 10, comm);

            // picture backgrounds!
            comm = sl.CreateComment();
            // left, right, top and bottom offsets. It's recommended that you just leave them as zero.
            // This particular method overload will stretch the picture. And given all-zero offsets,
            // it's effectively filling up the whole comment box.
            // The last zero is the transparency.
            comm.Fill.SetPictureFill("mandelbrot.png", 0, 0, 0, 0, 0);
            comm.SetText("I've got a fractal background!");
            sl.InsertComment(57, 6, comm);

            comm = sl.CreateComment();
            // this is one of those methods that don't quite match completely...
            // the first 2 zeroes are OffsetX and OffsetY, which aren't used, so just set them as zero.
            // 33 means 33%, so the picture will be tiled approximately 3 times (100 / 33) horizontally.
            // 50 means 50%, so it will be tiled 2 times (100 / 50) vertically.
            // You can ignore RectangleAlignmentValues and TileFlipValues (for now?)
            // The last zero is the transparency.
            comm.Fill.SetPictureFill("julia.png", 0, 0, 33, 50,
                DocumentFormat.OpenXml.Drawing.RectangleAlignmentValues.Bottom,
                DocumentFormat.OpenXml.Drawing.TileFlipValues.None,
                0);
            comm.SetText("Well, I've got Julia. *wink wink*");
            sl.InsertComment(57, 10, comm);

            sl.SaveAs("CellComments.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
