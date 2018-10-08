using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using AdvanceSoftware.VBReport8;
using AdvanceSoftware.VBReport8.Web;
using System.IO;
using System.Text;
public partial class _Default : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        WebCellReport cellReport1 = new WebCellReport();

        //テンプレートのファイルパス
        cellReport1.FileName = Server.MapPath("./Excel/Template.xlsx");

        //画像のファイルパス（方法1）
        //System.Drawing.Image image = System.Drawing.Image.FromFile(Server.MapPath("./Img/image.png"));
        //画像のファイルパス（方法2）
        FileStream fs = File.OpenRead(Server.MapPath("./Img/image.png"));
        System.Drawing.Image image = System.Drawing.Image.FromStream(fs);
        //高さと幅の設定
        
        var scale = 0.3;
        var h = image.Size.Height * scale;
        var w = image.Size.Width * scale;

        // ScaleMode プロパティでサイズの単位をピクセルに指定します。
        cellReport1.ScaleMode = ScaleMode.Pixel;
        cellReport1.Report.Start();
        cellReport1.Report.File();
        cellReport1.Page.Start("Sheet1", "1");

        //AddImage メソッドの第 1 引数に挿入する画像のフルパスを設定します。第 2、3 引数には高さ、幅のサイズを指定します。
        //cellReport1.Cell("B5").Drawing.AddImage(@"Img\image.png", 169, 150);
        //画像のフルパス、開始オフセットx、y、終了オフセットx、y、高さ、幅
        cellReport1.Cell("B5").Drawing.AddImage(image, 0, 0, 50, 50, h, w);
        cellReport1.Cell("B10").Break = true;
        cellReport1.Cell("B15").Drawing.AddImage(image, 0, 0, 50, 50, h, w);
        cellReport1.Cell("B20").Break = true;

        cellReport1.Page.End();
        cellReport1.Report.End();
        cellReport1.Report.SaveAs(Server.MapPath("./Result/Output.xlsx"));

    }
}