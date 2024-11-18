using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;
using Sungero.Core;
using Akelon.OpenXML.Structures.Module;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Vml.Wordprocessing;
using DocumentFormat.OpenXml.Vml.Office;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System.Text;
using BorderValues = DocumentFormat.OpenXml.Wordprocessing.BorderValues;
using Aspose.BarCode.Generation;

namespace Akelon.OpenXML.Isolated.AkelonOpenXMLWrapper
{
  public class IsolatedFunctions
  {

    /// <summary>
    /// Метод для демонстрации решения.
    /// </summary>
    /// <param name="stream">Поток с версией, в которой требуется заполнить элементы управления.</param>
    /// <param name="image">Изображение в формате base64.</param>
    /// <returns>Поток с заполненной версией.</returns>
    [Public]
    public static Stream ExampleWorkWithArea(Stream stream, string image)
    {
      System.Diagnostics.Debugger.Break();
      using (WordprocessingDocument document = WordprocessingDocument.Open(stream, true))
      {
        MainDocumentPart mainPart = document.MainDocumentPart;

        #region Работа с таблицами.

        // Найти закладку с именем "InsertNewTable".
        var bookmark = mainPart.Document.Body.Descendants<BookmarkStart>()
          .Where(bm => bm.Name == "InsertNewTable")
          .FirstOrDefault();

        if (bookmark != null)
        {
          var parent = bookmark.Parent;
          Paragraph newParagraph = new Paragraph();

          // Вствить новый параграф после закладки.
          parent.InsertAfterSelf(newParagraph);

          // Создать таблицу
          Table table = CreateTable(1);
          
          // Заполнить заголовки таблицы.
          CreateHeadRow(table, new string[] { "Колонка1", "Колонка2", "Колонка3", "Колонка4", "Колонка5" });

          // Заполнить данные для таблицы.
          AddNewLastRowAndFillCells(table, new string[,] { { "1.1", "1.2", "1.3", "1.4" }, { "2.1", "2.2", "2.3", "1.4" } });

          #region Записать картинку в ячейку.
          
          // Получить ячейку в третьей строке пятой колонки.
          var cell = GetTableCell(table, 2, 4);
          
          if (cell != null)
          {
            // image - картнка в base64 из локализации.
            var bytesImage = Convert.FromBase64String(image);
            
            // Удалить дочерние элементы ячейки.
            cell.RemoveAllChildren();
            
            // Вставить изображение в ячейку.
            using (var imgStream = new MemoryStream(bytesImage))
              AddImageToCell(mainPart, cell, imgStream);
          }

          // Получить ячейку во второй строке пятой колонки.
          cell = GetTableCell(table, 1, 4);
          
          if (cell != null)
          {
            // Сгенерировать qr-код.
            var qrcode = GenerateStamp("https://akelon.com/", EncodeTypes.QR.TypeName);
            qrcode.Position = 0;
            
            // Удалить дочернии элементы ячейки.
            cell.RemoveAllChildren();
            
            AddImageToCell(mainPart, cell, qrcode);
          }
          
          #endregion

          // Вставить таблицу после нового параграфа.
          newParagraph.InsertAfterSelf(table);
        }

        #endregion

        #region Пример подсановки значений в документе.
        
        // Пример подстановки значения в элементы управления тела документа.
        ChangeParamInWordBody(mainPart, "Обычный текст", "Обычный текст");
        ChangeParamInWordBody(mainPart, "Форматированный текст 1", "Форматированный текст (цветной)", colorCode: "FF0000");
        ChangeParamInWordBody(mainPart, "Форматированный текст 2", "Форматированный текст (жирный)", bold: true);
        ChangeParamInWordBody(mainPart, "Форматированный текст 3", "Форматированный текст (цветной и жирный)", "FF0000", true);
        // Пример подстановки значения в элементы управления в верхнем колонтитуле документа.
        ChangeParamInWordHeader(mainPart, "Заголовок страницы", "Верхний колонтитул", "0000FF", true);
        // Пример подстановки значения в элементы управления в нижнем колонтитуле документа.
        ChangeParamInWordFooter(mainPart, "Подвал страницы", "Нижний колонтитул", "0000FF", true);
        
        #endregion
        
        #region Пример генерации штрих-кода и его простановки в элемент управления в нижнем колонтитуле.
        
        // Сгенерировать штрих-код типа Code128 с значение "Akelon".
        var barcode = GenerateStamp("Akelon", EncodeTypes.Code128.TypeName);
        
        // Вставить штрих-код в элемент управления с именем "img" в нижнем колонтитуле документа.
        AddBarcodeToFooter(mainPart, "ШтрихКод", barcode);
        
        #endregion
        
        #region Пример работы с подложкой документа.

        // Проверить наличие подложки в документе.
        var isAnyWatermark = AnyWatermark(mainPart);

        // При наличие подложки изменить значение на "Устарел".
        // Иначе добавить подложку с текстом "Действующий".
        if (isAnyWatermark)
          ChangeWatermarkInWord(mainPart, "Устарел");
        else
          AddWatermark(document, "Действующий");

        #endregion

        // Пример замены текста в теле документа.
        SearchAndReplace(document, "Hello world!", "Hi Everyone!");
      }
      
      return stream;
    }
    
    /// <summary>
    /// Сгенерировать штрих код.
    /// </summary>
    /// <param name="text">Значение.</param>
    /// <param name="type">Тип штрих код.</param>
    /// <returns>Поток изображения со штрих кодом.</returns>
    [Public]
    public static Stream GenerateStamp(string text, string type)
    {
      var enumType = EncodeTypes.AllEncodeTypes.FirstOrDefault(x => x.TypeName == type);
      BarcodeGenerator generator = new BarcodeGenerator(enumType, text);

      // Установить разрешение.
      generator.Parameters.Resolution = 400;
      
      // Сгенерировать штрих-код.
      var stream = new MemoryStream();
      generator.Save(stream, BarCodeImageFormat.Png);
      
      return stream;
    }
    
    /// <summary>
    /// Добавить штрих код в свойство- изображение нижнего колонтитула.
    /// </summary>
    /// <param name="mainPart">MainDocumentPart.</param>
    /// <param name="paramName">Tag свойства.</param>
    /// <param name="barcodeValue">Содержаине штрихкода.</param>
    /// <param name="height">Высота штрихкода.</param>
    /// <param name="width">Ширина штрихкода.</param>
    public static void AddBarcodeToFooter(DocumentFormat.OpenXml.Packaging.MainDocumentPart mainPart, string paramName, Stream streamImg)
    {
      if (mainPart != null)
      {
        foreach (var footerPart in mainPart.FooterParts)
        {
          DocumentFormat.OpenXml.Wordprocessing.SdtElement controlBlock = footerPart.Footer.Descendants<SdtElement>().Where(r => r.SdtProperties.GetFirstChild<Tag>().Val == paramName).FirstOrDefault();
          if (controlBlock != null)
          {
            DocumentFormat.OpenXml.Drawing.Blip blip = controlBlock.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().First();
            DocumentFormat.OpenXml.Packaging.ImagePart imagePart = footerPart.AddImagePart(DocumentFormat.OpenXml.Packaging.ImagePartType.Png);
            streamImg.Position = 0;
            imagePart.FeedData(streamImg);
            blip.Embed = footerPart.GetIdOfPart(imagePart);
          }
        }
      }
    }
    
    /// <summary>
    /// Заполнить параметр в теле документа.
    /// </summary>
    /// <param name="mainPart">MainDocumentPart</param>
    /// <param name="paramName">Tag параметра.</param>
    /// <param name="paramValue">Новое значение параметра.</param>
    /// <param name="colorCode">Цвет параметра.</param>
    /// <param name="bold">Жирность параметра.</param>
    public static void ChangeParamInWordBody(DocumentFormat.OpenXml.Packaging.MainDocumentPart mainPart, string paramName, string paramValue, string colorCode = "000000", bool bold = false)
    {
      Text textParam = null;

      if (mainPart != null)
      {
        foreach (SdtProperties propertie in mainPart.Document.Body.Descendants<SdtProperties>().Where(r => r.GetFirstChild<Tag>() != null && r.GetFirstChild<Tag>().Val.Value == paramName))
        {
          if (propertie != null)
          {
            ChangeParamInWord(propertie, textParam, paramValue, colorCode, bold);
          }
        }
      }
    }

    /// <summary>
    /// Заполнить параметр в верхнем колонтитуле.
    /// </summary>
    /// <param name="mainPart">MainDocumentPart</param>
    /// <param name="paramName">Tag параметра.</param>
    /// <param name="paramValue">Новое значение параметра.</param>
    /// <param name="colorCode">Цвет параметра.</param>
    /// <param name="bold">Жирность параметра.</param>
    public static void ChangeParamInWordHeader(DocumentFormat.OpenXml.Packaging.MainDocumentPart mainPart, string paramName, string paramValue, string colorCode = "000000", bool bold = false)
    {
      Text textParam = null;

      if (mainPart != null)
      {
        foreach (var headerPart in mainPart.HeaderParts)
        {
          foreach (SdtProperties propertie in headerPart.Header.Descendants<SdtProperties>().Where(r => r.GetFirstChild<Tag>() != null && r.GetFirstChild<Tag>().Val.Value == paramName))
          {
            if (propertie != null)
            {
              ChangeParamInWord(propertie, textParam, paramValue, colorCode, bold);
            }
          }
        }
      }
    }

    /// <summary>
    /// Заполнить параметр в нижнем колонтитуле.
    /// </summary>
    /// <param name="mainPart">MainDocumentPart</param>
    /// <param name="paramName">Tag параметра.</param>
    /// <param name="paramValue">Новое значение параметра.</param>
    /// <param name="colorCode">Цвет параметра.</param>
    /// <param name="bold">Жирность параметра.</param>
    public static void ChangeParamInWordFooter(DocumentFormat.OpenXml.Packaging.MainDocumentPart mainPart, string paramName, string paramValue, string colorCode = "000000", bool bold = false)
    {
      Text textParam = null;

      if (mainPart != null)
      {
        foreach (var footerPart in mainPart.FooterParts)
        {
          foreach (SdtProperties propertie in footerPart.Footer.Descendants<SdtProperties>().Where(r => r.GetFirstChild<Tag>() != null && r.GetFirstChild<Tag>().Val.Value == paramName))
          {
            if (propertie != null)
            {
              ChangeParamInWord(propertie, textParam, paramValue, colorCode, bold);
            }
          }
        }
      }
    }

    /// <summary>
    /// Подстановка значений в свойство документа.
    /// </summary>
    /// <param name="propertie">Свойство.</param>
    /// <param name="textParam">Tag параметра.</param>
    /// <param name="paramValue">Новое значение параметра.</param>
    /// <param name="colorCode">Цвет параметра.</param>
    /// <param name="bold">Жирность.</param>
    private static void ChangeParamInWord(DocumentFormat.OpenXml.Wordprocessing.SdtProperties propertie,
                                          DocumentFormat.OpenXml.Wordprocessing.Text textParam, string paramValue, string colorCode, bool bold)
    {

      var parent = propertie.Parent;
      var sdtContentBlock = parent.Descendants<SdtContentBlock>().FirstOrDefault();
      var runContentBlock = parent.Descendants<SdtContentRun>().FirstOrDefault();
      var cellContentBlock = parent.Descendants<SdtContentCell>().FirstOrDefault();

      if (sdtContentBlock != null)
        textParam = sdtContentBlock.Descendants<Text>().FirstOrDefault();
      else if (runContentBlock != null)
        textParam = runContentBlock.Descendants<Text>().FirstOrDefault();
      else if (cellContentBlock != null)
        textParam = cellContentBlock.Descendants<Text>().FirstOrDefault();

      if (textParam != null)
      {
        textParam.Text = paramValue;
      }
      else
      {
        textParam = new DocumentFormat.OpenXml.Wordprocessing.Text();
        textParam.Text = paramValue;
        if (sdtContentBlock != null)
          sdtContentBlock.Append(textParam);
        else if (runContentBlock != null)
          runContentBlock.Append(textParam);
        else if (cellContentBlock != null)
          cellContentBlock.Append(textParam);
      }

      if (textParam.InnerText != parent.InnerText)
      {
        var oldTextBlocks = parent.Descendants<Text>().Where(r => r.Text != paramValue || !Equals(textParam, r));
        foreach (var oldTextBlock in oldTextBlocks)
        {
          if (oldTextBlock != null)
            oldTextBlock.Remove();
        }
      }

      DocumentFormat.OpenXml.Wordprocessing.RunProperties properties = null;

      if (sdtContentBlock != null)
        properties = sdtContentBlock.Descendants<RunProperties>().FirstOrDefault();
      else if (runContentBlock != null)
        properties = runContentBlock.Descendants<RunProperties>().FirstOrDefault();
      else if (cellContentBlock != null)
        properties = cellContentBlock.Descendants<RunProperties>().FirstOrDefault();

      if (properties != null)
      {
        if (!string.IsNullOrEmpty(colorCode))
          properties.Color = new DocumentFormat.OpenXml.Wordprocessing.Color() { Val = colorCode };
        
        properties.Bold = new DocumentFormat.OpenXml.Wordprocessing.Bold() { Val = bold };
      }
    }

    /// <summary>
    /// Изменить подложку документа.
    /// </summary>
    /// <param name="mainPart">MainDocumentPart</param>
    /// <param name="value">Текст подложки.</param>
    public static void ChangeWatermarkInWord(DocumentFormat.OpenXml.Packaging.MainDocumentPart mainPart, string value)
    {
      if (mainPart != null)
      {
        foreach (var headerPart in mainPart.HeaderParts)
        {
          var hz = headerPart.Header.Descendants<DocumentFormat.OpenXml.Vml.Shape>().FirstOrDefault(r => r.Id.Value.Contains("WaterMarkObject"));
          if (hz != null)
          {
            var watermark = hz.Descendants<DocumentFormat.OpenXml.Vml.TextPath>().FirstOrDefault();
            watermark.String = value;
          }
        }
      }
    }

    /// <summary>
    /// Проверить наличие подложки в документе.
    /// </summary>
    /// <param name="mainPart">MainDocumentPart.</param>
    /// <returns>True- есть подложка, иначе false.</returns>
    public static bool AnyWatermark(DocumentFormat.OpenXml.Packaging.MainDocumentPart mainPart)
    {
      return mainPart.HeaderParts.All(x => x.Header.Descendants<DocumentFormat.OpenXml.Vml.Shape>().Any(r => r.Id.Value.Contains("WaterMarkObject")));
    }

    /// <summary>
    /// Добавить штрих код в свойство- изображение нижнего колонтитула.
    /// </summary>
    /// <param name="mainPart">MainDocumentPart.</param>
    /// <param name="paramName">Tag свойства.</param>
    /// <param name="streamImg">Поток с изображением.</param>
    public static void AddBarcodeToFooterStream(MainDocumentPart mainPart, string paramName, MemoryStream streamImg)
    {
      if (mainPart != null)
      {
        foreach (var footerPart in mainPart.FooterParts)
        {
          DocumentFormat.OpenXml.Wordprocessing.SdtElement controlBlock = footerPart.Footer.Descendants<SdtElement>().Where(r => r.SdtProperties.GetFirstChild<Tag>().Val == paramName).FirstOrDefault();
          if (controlBlock != null)
          {
            DocumentFormat.OpenXml.Drawing.Blip blip = controlBlock.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().First();
            DocumentFormat.OpenXml.Packaging.ImagePart imagePart = footerPart.AddImagePart(ImagePartType.Jpeg);
            streamImg.Position = 0;
            imagePart.FeedData(streamImg);
            blip.Embed = footerPart.GetIdOfPart(imagePart);
          }
        }
      }
    }

    /// <summary>
    /// Создать Header.
    /// </summary>
    /// <returns>Header.</returns>
    private static Header MakeHeader()
    {
      var header = new Header();
      var paragraph = new Paragraph();
      var run = new Run();
      var text = new Text();
      text.Text = "";
      run.Append(text);
      paragraph.Append(run);
      header.Append(paragraph);
      return header;
    }

    /// <summary>
    /// Добавить Подложку в документ.
    /// </summary>
    /// <param name="doc">WordprocessingDocument.</param>
    /// <param name="textWatermark">Текст подложки.</param>
    public static void AddWatermark(WordprocessingDocument doc, string textWatermark)
    {
      if (doc.MainDocumentPart.HeaderParts.Count() == 0)
      {
        doc.MainDocumentPart.DeleteParts(doc.MainDocumentPart.HeaderParts);
        var newHeaderPart = doc.MainDocumentPart.AddNewPart<HeaderPart>();
        var rId = doc.MainDocumentPart.GetIdOfPart(newHeaderPart);
        var headerRef = new HeaderReference();
        headerRef.Id = rId;
        var sectionProps = doc.MainDocumentPart.Document.Body.Elements<SectionProperties>().LastOrDefault();
        if (sectionProps == null)
        {
          sectionProps = new SectionProperties();
          doc.MainDocumentPart.Document.Body.Append(sectionProps);
        }
        sectionProps.RemoveAllChildren<HeaderReference>();
        sectionProps.Append(headerRef);

        newHeaderPart.Header = MakeHeader();
        newHeaderPart.Header.Save();
      }

      foreach (HeaderPart headerPart in doc.MainDocumentPart.HeaderParts)
      {
        var sdtBlock1 = new SdtBlock();
        var sdtProperties1 = new SdtProperties();
        var sdtId1 = new SdtId() { Val = 87908844 };
        var sdtContentDocPartObject1 = new SdtContentDocPartObject();
        var docPartGallery1 = new DocPartGallery() { Val = "Watermarks" };
        var docPartUnique1 = new DocPartUnique();
        sdtContentDocPartObject1.Append(docPartGallery1);
        sdtContentDocPartObject1.Append(docPartUnique1);
        sdtProperties1.Append(sdtId1);
        sdtProperties1.Append(sdtContentDocPartObject1);

        var sdtContentBlock1 = new SdtContentBlock();
        var paragraph2 = new Paragraph()
        {
          RsidParagraphAddition = "00656E18",
          RsidRunAdditionDefault = "00656E18"
        };
        var paragraphProperties2 = new ParagraphProperties();
        var paragraphStyleId2 = new ParagraphStyleId() { Val = "Header" };
        paragraphProperties2.Append(paragraphStyleId2);
        var run1 = new Run();
        var runProperties1 = new RunProperties();
        var noProof1 = new NoProof();
        var languages1 = new Languages() { EastAsia = "zh-TW" };
        runProperties1.Append(noProof1);
        runProperties1.Append(languages1);
        var picture1 = new Picture();
        var shapetype1 = new Shapetype()
        {
          Id = "_x0000_t136",
          CoordinateSize = "21600,21600",
          OptionalNumber = 136,
          Adjustment = "10800",
          EdgePath = "m@7,l@8,m@5,21600l@6,21600e"
        };
        var formulas1 = new Formulas();
        var formula1 = new Formula() { Equation = "sum #0 0 10800" };
        var formula2 = new Formula() { Equation = "prod #0 2 1" };
        var formula3 = new Formula() { Equation = "sum 21600 0 @1" };
        var formula4 = new Formula() { Equation = "sum 0 0 @2" };
        var formula5 = new Formula() { Equation = "sum 21600 0 @3" };
        var formula6 = new Formula() { Equation = "if @0 @3 0" };
        var formula7 = new Formula() { Equation = "if @0 21600 @1" };
        var formula8 = new Formula() { Equation = "if @0 0 @2" };
        var formula9 = new Formula() { Equation = "if @0 @4 21600" };
        var formula10 = new Formula() { Equation = "mid @5 @6" };
        var formula11 = new Formula() { Equation = "mid @8 @5" };
        var formula12 = new Formula() { Equation = "mid @7 @8" };
        var formula13 = new Formula() { Equation = "mid @6 @7" };
        var formula14 = new Formula() { Equation = "sum @6 0 @5" };

        formulas1.Append(formula1);
        formulas1.Append(formula2);
        formulas1.Append(formula3);
        formulas1.Append(formula4);
        formulas1.Append(formula5);
        formulas1.Append(formula6);
        formulas1.Append(formula7);
        formulas1.Append(formula8);
        formulas1.Append(formula9);
        formulas1.Append(formula10);
        formulas1.Append(formula11);
        formulas1.Append(formula12);
        formulas1.Append(formula13);
        formulas1.Append(formula14);
        var path1 = new DocumentFormat.OpenXml.Vml.Path()
        {
          AllowTextPath = TrueFalseValue.FromBoolean(true),
          ConnectionPointType = ConnectValues.Custom,
          ConnectionPoints = "@9,0;@10,10800;@11,21600;@12,10800",
          ConnectAngles = "270,180,90,0"
        };
        var textPath1 = new TextPath()
        {
          On = TrueFalseValue.FromBoolean(true),
          FitShape = TrueFalseValue.FromBoolean(true)
        };
        var shapeHandles1 = new ShapeHandles();

        var shapeHandle1 = new ShapeHandle()
        {
          Position = "#0,bottomRight",
          XRange = "6629,14971"
        };

        shapeHandles1.Append(shapeHandle1);

        var lock1 = new DocumentFormat.OpenXml.Vml.Office.Lock
        {
          Extension = ExtensionHandlingBehaviorValues.Edit,
          TextLock = TrueFalseValue.FromBoolean(true),
          ShapeType = TrueFalseValue.FromBoolean(true)
        };

        shapetype1.Append(formulas1);
        shapetype1.Append(path1);
        shapetype1.Append(textPath1);
        shapetype1.Append(shapeHandles1);
        shapetype1.Append(lock1);
        var shape1 = new Shape()
        {
          Id = "PowerPlusWaterMarkObject357922611",
          Style = "position:absolute;left:0;text-align:left;margin-left:0;margin-top:0;width:527.85pt;height:131.95pt;rotation:315;z-index:-251656192;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin",
          OptionalString = "_x0000_s2049",
          AllowInCell = TrueFalseValue.FromBoolean(true),
          FillColor = "silver",
          Stroked = TrueFalseValue.FromBoolean(false),
          Type = "#_x0000_t136"
        };

        var fill1 = new Fill() { Opacity = ".5" };
        TextPath textPath2 = new TextPath()
        {
          Style = "font-family:\"Calibri\";font-size:1pt",
          String = textWatermark
        };

        var textWrap1 = new TextWrap()
        {
          AnchorX = DocumentFormat.OpenXml.Vml.Wordprocessing.HorizontalAnchorValues.Margin,
          AnchorY = DocumentFormat.OpenXml.Vml.Wordprocessing.VerticalAnchorValues.Margin
        };

        shape1.Append(fill1);
        shape1.Append(textPath2);
        shape1.Append(textWrap1);
        picture1.Append(shapetype1);
        picture1.Append(shape1);
        run1.Append(runProperties1);
        run1.Append(picture1);
        paragraph2.Append(paragraphProperties2);
        paragraph2.Append(run1);
        sdtContentBlock1.Append(paragraph2);
        sdtBlock1.Append(sdtProperties1);
        sdtBlock1.Append(sdtContentBlock1);
        headerPart.Header.Append(sdtBlock1);
        headerPart.Header.Save();
      }
    }

    /// <summary>
    /// Заменить текст в теле документа.
    /// </summary>
    /// <param name="document">WordprocessingDocument</param>
    /// <param name="oldValue">Текст, который требуется заменить.</param>
    /// <param name="newValue">Новый текст.</param>
    public static void SearchAndReplace(WordprocessingDocument document, string oldValue, string newValue)
    {
      var body = document.MainDocumentPart.Document.Body;
      var paras = body.Elements<Paragraph>();

      foreach (var para in paras)
      {
        foreach (var run in para.Elements<Run>())
        {
          foreach (var text in run.Elements<Text>())
          {
            if (text.Text.Contains(oldValue))
            {
              text.Text = text.Text.Replace(oldValue, newValue);
            }
          }
        }
      }
    }

    #region Работа с таблицами

    /// <summary>
    /// Создать таблицу.
    /// </summary>
    /// <param name="borderWidth">Ширина рамки.</param>
    /// <returns>Таблица.</returns>
    public static Table CreateTable(UInt32Value borderWidth)
    {
      Table table = new Table();

      TableProperties props = new TableProperties(
        new TableBorders(
          new DocumentFormat.OpenXml.Wordprocessing.TopBorder
          {
            Val = new EnumValue<BorderValues>(BorderValues.Single),
            Size = borderWidth
          },
          new DocumentFormat.OpenXml.Wordprocessing.BottomBorder
          {
            Val = new EnumValue<BorderValues>(BorderValues.Single),
            Size = borderWidth
          },
          new DocumentFormat.OpenXml.Wordprocessing.LeftBorder
          {
            Val = new EnumValue<BorderValues>(BorderValues.Single),
            Size = borderWidth
          },
          new DocumentFormat.OpenXml.Wordprocessing.RightBorder
          {
            Val = new EnumValue<BorderValues>(BorderValues.Single),
            Size = borderWidth
          },
          new InsideHorizontalBorder
          {
            Val = new EnumValue<BorderValues>(BorderValues.Single),
            Size = borderWidth
          },
          new InsideVerticalBorder
          {
            Val = new EnumValue<BorderValues>(BorderValues.Single),
            Size = borderWidth
          }));

      table.AppendChild<TableProperties>(props);

      return table;
    }

    /// <summary>
    /// Сформировать первую строку для таблицы с заголовками.
    /// </summary>
    /// <param name="table">Таблица.</param>
    /// <param name="data">Массив строк с наименование колонок.</param>
    public static void CreateHeadRow(Table table, string[] data)
    {
      // Создать и добавить строку.
      TableRow row = new TableRow();
      table.Append(row);

      // Создать и заполнить ячейки.
      foreach (var item in data)
      {
        ParagraphProperties paragraphProperties = new ParagraphProperties();
        SpacingBetweenLines spacing = new SpacingBetweenLines() { After = "0" };
        paragraphProperties.SpacingBetweenLines = spacing;
        paragraphProperties.Append(new Justification() { Val = JustificationValues.Center });
        Paragraph paragraph = new Paragraph(paragraphProperties);
        paragraph.Append(new Run(new Text(item)));
        TableCellVerticalAlignment tcVA = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };
        TableCellProperties tcp = new TableCellProperties();
        tcp.Append(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "9600" });
        tcp.Append(tcVA);
        var tc = new TableCell();
        tc.Append(tcp);
        tc.AppendChild<Paragraph>(paragraph);
        row.Append(tc);
      }
    }

    /// <summary>
    /// Получить ячейку.
    /// </summary>
    /// <param name="body">Тело документа.</param>
    /// <param name="tableNumber">Порядковый номер таблицы(начинается с 0).</param>
    /// <param name="rowNumber">Порядковый номер строки(начинается с 0).</param>
    /// <param name="cellNumber">Порядковый номер ячейки в строке(начинается с 0).</param>
    /// <returns>Ячейка таблицы.</returns>
    public static TableCell GetTableCell(Body body, int tableNumber, int rowNumber, int cellNumber)
    {
      Table table = body.Elements<Table>().ElementAt<Table>(tableNumber);
      TableRow tableRow = table.Elements<TableRow>().ElementAt(rowNumber);
      TableCell tableCell = tableRow.Elements<TableCell>().ElementAt(cellNumber);
      return tableCell;
    }

    /// <summary>
    /// Получить ячейку.
    /// </summary>
    /// <param name="table">Таблица.</param>
    /// <param name="rowNumber">Порядковый номер строки(начинается с 0).</param>
    /// <param name="cellNumber">Порядковый номер ячейки в строке(начинается с 0).</param>
    /// <returns>Ячейка таблицы.</returns>
    public static TableCell GetTableCell(Table table, int rowNumber, int cellNumber)
    {
      TableRow tableRow = table.Elements<TableRow>().ElementAt(rowNumber);
      TableCell tableCell = tableRow.Elements<TableCell>().ElementAt(cellNumber);
      return tableCell;
    }

    /// <summary>
    /// Запиcать текст ячейки с новым пораграфом.
    /// </summary>
    /// <param name="cellText">Новый тест.</param>
    /// <param name="tableCell">Ячейка.</param>
    /// <param name="justificationValues">Выравнивание текста перечесление JustificationValues или null.</param>
    /// <param name="isBold">Жирный.</param>
    /// <param name="isItalic">Курсив.</param>
    public static void WriteNewParagraphTextToCell(string cellText, TableCell tableCell, JustificationValues? justificationValues, bool isBold, bool isItalic)
    {
      Paragraph paragraph = tableCell.AppendChild(new Paragraph());

      // Добавить свойства параграфа и выровнить содержимое.
      if (justificationValues != null)
      {
        var paragraphProp = new ParagraphProperties();
        paragraphProp.Append(new Justification() { Val = justificationValues });
        paragraph.Append(paragraphProp);
      }

      //Создать новый элемент Run.
      Run run = paragraph.AppendChild<Run>(new Run());
      var runPropertyes = run.AppendChild(new RunProperties());
      if (isBold)
        runPropertyes.Append(new Bold());
      if (isItalic)
        runPropertyes.Append(new Italic());
      runPropertyes.Append(new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" });
      runPropertyes.Append(new FontSize() { Val = "24" });

      // Вставить текст в новый жлемент Run.
      var text = run.AppendChild<Text>(new Text());
      text.Text = cellText;
    }

    /// <summary>
    /// Добавить в таблицу новую строку с нужным индексом путем копирования строки таблицы.
    /// (Добавить строку в середину таблицы).
    /// </summary>
    /// <param name="table">Таблица, в которую нужно добавить ячейки.</param>
    /// <param name="copyRowIndex">Индекс ячейки, котрую нужно скопировать</param>
    /// <param name="rowPlaseIndex">Индекс, куда нужно вставить ячейки</param>
    /// <param name="insertCount">Количество ячеек, которые нужно добавить</param>
    public static void AddTableRow(Table table, int copyRowIndex, int rowPlaseIndex, int insertCount)
    {
      TableRow tableRow = table.Elements<TableRow>().ElementAt(copyRowIndex);

      var tableRowList = table.ChildElements.Where(w => w is TableRow).ToList();
      table.RemoveAllChildren<TableRow>();
      int rowIndex = 0;
      foreach (var row in tableRowList)
      {
        if (rowPlaseIndex == rowIndex)
        {
          for (int i = 0; i < insertCount; i++)
          {
            TableRow copyRow = (TableRow)tableRow.CloneNode(true);
            table.AppendChild(copyRow);
          }
        }
        table.AppendChild(row);
        rowIndex++;
      }
    }

    /// <summary>
    /// Удалить все строки талицы, кроме первой.
    /// </summary>
    /// <param name="table">Таблица.</param>
    public static void RemoveRowTable(Table table)
    {
      int count = table.Descendants<TableRow>().Count();

      for (int i = count - 1; i >= 1; i--)
        table.Descendants<TableRow>().ElementAt(i).Remove();
    }

    /// <summary>
    /// Добавить строку в конец таблицы и заполнить в ней ячейки.
    /// </summary>
    /// <param name="table">Таблица.</param>
    /// <param name="arrRow">Двумерный массив c данными для заполнения.</param>
    public static void AddNewLastRowAndFillCells(Table table, string[,] arrRow)
    {
      TableRow theRow = table.Elements<TableRow>().Last();

      for (int i = 0; i < arrRow.GetLength(0); i++)
      {
        TableRow rowCopy = (TableRow)theRow.CloneNode(true);

        for (int j = 0; j < arrRow.GetLength(1); j++)
        {
          if (rowCopy.Descendants<TableCell>().ElementAt(j).HasChildren)
            rowCopy.Descendants<TableCell>().ElementAt(j).RemoveAllChildren<Paragraph>();
          rowCopy.Descendants<TableCell>().ElementAt(j).Append(new Paragraph(new Run(new Text(arrRow[i, j]))));
        }

        table.AppendChild(rowCopy);
      }
    }

    /// <summary>
    /// Получить таблицу по количеству столбцов и имеющемуся значению в таблице.
    /// </summary>
    /// <param name="mainPart">Тело документа.</param>
    /// <param name="containce">Значение, коорое содержит таблица.</param>
    /// <param name="countColumn">Количество столбцов.</param>
    /// <returns>Таблица.</returns>
    public static Table GetTable(MainDocumentPart mainPart, string containce, int countColumn)
    {
      return mainPart.Document.Body.Descendants<Table>().FirstOrDefault(x => x.GetFirstChild<TableRow>().Descendants<TableCell>().Count() == countColumn && x.InnerText.Contains(containce));
    }

    /// <summary>
    /// Вставить картинку в ячейку.
    /// </summary>
    /// <param name="mainPart">MainDocumentPart документа.</param>
    /// <param name="cell">Ячейка.</param>
    /// <param name="imageStream">Поток изображения.</param>
    public static void AddImageToCell(MainDocumentPart mainPart, TableCell cell, Stream imageStream)
    {
      using (var memoryStream = new MemoryStream())
      {
        imageStream.CopyTo(memoryStream);
        DocumentFormat.OpenXml.Packaging.ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);
        memoryStream.Position = 0;
        imagePart.FeedData(memoryStream);
        var relationshipId = mainPart.GetIdOfPart(imagePart);
        
        var element =
          new Drawing(
            new DW.Inline(
              new DW.Extent() { Cx = 990000L, Cy = 792000L },
              new DW.EffectExtent()
              {
                LeftEdge = 0L,
                TopEdge = 0L,
                RightEdge = 0L,
                BottomEdge = 0L
              },
              new DW.DocProperties()
              {
                Id = (UInt32Value)1U,
                Name = "Picture 1"
              },
              new DW.NonVisualGraphicFrameDrawingProperties(
                new A.GraphicFrameLocks() { NoChangeAspect = true }),
              new A.Graphic(
                new A.GraphicData(
                  new PIC.Picture(
                    new PIC.NonVisualPictureProperties(
                      new PIC.NonVisualDrawingProperties()
                      {
                        Id = (UInt32Value)0U,
                        Name = "New Bitmap Image.jpg"
                      },
                      new PIC.NonVisualPictureDrawingProperties()),
                    new PIC.BlipFill(
                      new A.Blip(
                        new A.BlipExtensionList(
                          new A.BlipExtension()
                          {
                            Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                          })
                       )
                      {
                        Embed = relationshipId,
                        CompressionState =
                          A.BlipCompressionValues.Print
                      },
                      new A.Stretch(
                        new A.FillRectangle())),
                    new PIC.ShapeProperties(
                      new A.Transform2D(
                        new A.Offset() { X = 0L, Y = 0L },
                        new A.Extents() { Cx = 990000L, Cy = 792000L }),
                      new A.PresetGeometry(
                        new A.AdjustValueList()
                       )
                      { Preset = A.ShapeTypeValues.Rectangle }))
                 )
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
             )
            {
              DistanceFromTop = (UInt32Value)0U,
              DistanceFromBottom = (UInt32Value)0U,
              DistanceFromLeft = (UInt32Value)0U,
              DistanceFromRight = (UInt32Value)0U
            });

        cell.Append(new Paragraph(new Run(element)));
      }
    }
    
    #endregion
  }
}