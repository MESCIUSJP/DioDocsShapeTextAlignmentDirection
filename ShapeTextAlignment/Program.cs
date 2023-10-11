// See https://aka.ms/new-console-template for more information
using GrapeCity.Documents.Excel.Drawing;
using GrapeCity.Documents.Excel;

Console.WriteLine("図形のテキスト配置を設定");

// 新規ワークブックの作成
var workbook = new GrapeCity.Documents.Excel.Workbook();
IWorksheet worksheet = workbook.Worksheets[0];

// 1番目の図形
IShape shape1 = worksheet.Shapes.AddShape(AutoShapeType.Rectangle, worksheet.Range["B3:E6"]);
shape1.TextFrame.TextRange.Add("1番目の図形");
//shape1.TextFrame.TextRange.Add("左揃えの配置");
// テキスト配置を左揃えに設定
shape1.TextFrame.VerticalAnchor = VerticalAnchor.AnchorTop;
shape1.TextFrame.TextRange.TextAlignment = TextAlignmentAnchor.Left;

// 2番目の図形
IShape shape2 = worksheet.Shapes.AddShape(AutoShapeType.Rectangle, worksheet.Range["B8:E11"]);
shape2.TextFrame.TextRange.Add("2番目の図形");
//shape2.TextFrame.TextRange.Add("中央揃えの配置");
// テキスト配置を中央揃えに設定
shape2.TextFrame.VerticalAnchor = VerticalAnchor.AnchorMiddle;
shape2.TextFrame.TextRange.TextAlignment = TextAlignmentAnchor.Center;

// 3番目の図形
IShape shape3 = worksheet.Shapes.AddShape(AutoShapeType.Rectangle, worksheet.Range["B13:E16"]);
shape3.TextFrame.TextRange.Add("3番目の図形");
//shape3.TextFrame.TextRange.Add("右揃えの配置");
// テキスト配置を右揃えに設定
shape3.TextFrame.VerticalAnchor = VerticalAnchor.AnchorBottom;
shape3.TextFrame.TextRange.TextAlignment = TextAlignmentAnchor.Right;

// 4番目の図形
IShape shape4 = worksheet.Shapes.AddShape(AutoShapeType.Rectangle, worksheet.Range["B18:E21"]);
shape4.TextFrame.VerticalAnchor = VerticalAnchor.AnchorMiddle;
// 最後の図形にて、3つの段落に異なる配置を設定
shape4.TextFrame.TextRange.Add("左揃え");
shape4.TextFrame.TextRange.Add("中央揃え");
shape4.TextFrame.TextRange.Add("右揃え");
shape4.TextFrame.TextRange.Paragraphs[0].TextAlignment = TextAlignmentAnchor.Left;
shape4.TextFrame.TextRange.Paragraphs[1].TextAlignment = TextAlignmentAnchor.Center;
shape4.TextFrame.TextRange.Paragraphs[2].TextAlignment = TextAlignmentAnchor.Right;

// 5番目の図形
IShape shape5 = worksheet.Shapes.AddShape(AutoShapeType.Rectangle, worksheet.Range["G3:I8"]);
shape5.TextFrame.TextRange.Add("こんにちは、DioDocsです。");
// テキストの方向を縦書き（半角文字含む）かつ右揃えに設定
shape5.TextFrame.Direction = TextDirection.StackedRtl;

// 6番目の図形
IShape shape6 = worksheet.Shapes.AddShape(AutoShapeType.Rectangle, worksheet.Range["G10:I15"]);
shape6.TextFrame.TextRange.Add("こんにちは、DioDocsです。");
// テキストの方向を縦書き（半角文字含む）かつ左揃えに設定
shape6.TextFrame.Direction = TextDirection.Stacked;

// 7番目の図形
IShape shape7 = worksheet.Shapes.AddShape(AutoShapeType.Rectangle, worksheet.Range["G17:I22"]);
shape7.TextFrame.TextRange.Add("こんにちは、DioDocsです。");
// テキストの方向を縦書きかつ右揃えに設定
shape7.TextFrame.Direction = TextDirection.Vertical;

// xlsx ファイルに保存
workbook.Save("SetShapeTextAlignment.xlsx");