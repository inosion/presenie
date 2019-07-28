package inosion.pptx


import org.apache.poi.xslf.usermodel._ // XMLSlideShow
import org.apache.poi.sl.usermodel._
import java.io.File
import java.io._

import scala.collection.JavaConverters._
import com.fasterxml.jackson.databind.JsonNode


import org.openxmlformats.schemas.drawingml.x2006.main.CTTableRow

import io.gatling.jsonpath.JsonPath

import scala.util._
import scala.collection.mutable
import org.apache.poi.sl.usermodel.PictureData.PictureType
import java.awt.geom.Rectangle2D
import com.typesafe.scalalogging._
//import org.openxmlformats.schemas.presentationml.x2006.main.


object ShapeGroup extends Enumeration  {
  type ShapeGroup = Value
  val Image = Value
}


// Value holder for data context
case class DataContext(rootJsonNode: JsonNode, contextJsonNode: Option[JsonNode] = None)


object PPTXMerger extends StrictLogging {

    type TextHolder = TextShape[XSLFShape, XSLFTextParagraph]

    type Foo = XSLFSheet

    import scala.reflect.ClassTag

    def f[T](v: T)(implicit ev: ClassTag[T]) = ev.toString

    val JsonPathReg                   = raw"[\$$@]|[\$$@]\.[0-9A-Za-z_:\-\?\.\[\]\*]+"
    val matchRegexpTemplate           = raw".*(\{\{\s*(" + JsonPathReg + raw")\s*\}\}).*"
    val matchContextControl           = (raw"(\{\s*context\s*=\s*(" + JsonPathReg + raw")\s*\})").r
    val matchGroupShapeControl        = (raw"(\{\s*context\s*=\s*(" + JsonPathReg + raw")\s*dir\s*=\s*(\d+)\s+gap=(\d+)\s*\})").r

    // new methods for iteration here

    /**
     * Iterates over the source template slide.
     */
    def render2(fc: FilesContext): Either[Error, Unit] = {

        // we need to track which SlideLayouts we ported across
        val visitedLayouts: mutable.Seq[XSLFSlideLayout] = mutable.Seq()
        logger.debug(s"! Regexp for templates = ${matchRegexpTemplate}")
        logger.debug(s"! Regexp for controls = ${matchContextControl}")
        logger.debug(s"! Regexp for group shape controls = ${matchGroupShapeControl}")

        val pptTemplate: XMLSlideShow = new XMLSlideShow(new FileInputStream(fc.src.getAbsolutePath()))
        val pptNew: XMLSlideShow      = new XMLSlideShow(new FileInputStream(fc.src.getAbsolutePath()))

        // To retain the Master Slides, fonts etc... we start with the original Slide and
        // to retain the master slide layouts.
        for (idx <- 0 until pptNew.getSlides.size()) {
            logger.debug(s"√ Removing slide ${prettyPrintSlide(pptNew.getSlides().get(0))}")
            pptNew.removeSlide(0)
        }

        // get the data
        val jsonData = JsonYamlTools.readFileToJson(fc.data)

        // iterate over the old slide
        for (srcSlide <- pptTemplate.getSlides().asScala) {
            val sc = SlidesContext(srcSlide = srcSlide, destPptx = pptNew, destSlides = List())
            processSlide(fc, sc, dc = DataContext(rootJsonNode = jsonData), visitedLayouts)
        }

        // save out the resulting new pptx
        try {
            pptNew.write(new FileOutputStream(fc.dst))
            Right(())
        } catch {
            case e: Exception => Left(new Error(s"failed writing the new pptx [${fc.dstAbsPath}]", e))
        }

    }

    def processSlide(fc: FilesContext, sc: SlidesContext, dc: DataContext, visitedLayouts: mutable.Seq[XSLFSlideLayout]) {

        // We will copy this slide from source to new ppt, then fill it's data
        val newSlide = PPTXTools.createSlide(sc.destPptx, sc.srcSlide, visitedLayouts)
        val newSlidesContext = sc.copy(destSlides = sc.destSlides :+ newSlide)

        System.err.println(s":: New slide from [${fc.srcRelPath}(${prettyPrintSlide(sc.srcSlide)})] to [${fc.dstRelPath}(${prettyPrintSlide(sc.srcSlide)})] ")

        findControlJsonPath(sc.srcSlide.getShapes().asScala) match {
            case Some(PageControlData(shape, contextMatch, jsonPath)) => {
                logger.debug(s"√ Slide [${prettyPrintSlide(sc.srcSlide)}] has a jsonPath context -> ${jsonPath}")
                JsonPath.query(jsonPath, dc.rootJsonNode).map{ i =>
                    for (jsonNode <- i) {
                        logger.debug(s"√ node = ${jsonNode.toString()}")
                        processSlide(fc, newSlidesContext, dc.copy(contextJsonNode = Some(jsonNode)), visitedLayouts)
                    }
                }
                // now we have templated it out, let's remove it
                sc.destPptx.removeSlide(newSlide.getSlideNumber() - 1)
                newSlide.removeShape(shape)
            }
            case None => processAllShapes(newSlide, dc)
        }

    }

    def processAllShapes(slide: XSLFSlide,  dc: DataContext) {

        for (shape <- slide.getShapes().asScala) {

            shape match {
                case textShape : TextHolder =>
                    if (hasTemplate(textShape)) {
                      logger.debug(s"√ ${f(shape)} -> ${textShape.getText()}")
                      changeText(textShape, dc)
                      logger.debug(s"√ now == ${f(shape)} -> ${textShape.getText()}")

                    } else {
                      logger.debug(s"✖ '${textShape.getText()}' did not match")
                    }

                case group : XSLFGroupShape => processGroupShape(group)
                case table : XSLFTable  if (hasControl(table.getRows().get(0).getCells().get(0)))    => {
                    logger.debug(s"√ we have a table - ${table.getRows().get(0).getCells().get(0).getText()}")
                    iterateTable(table, dc)
                }
                case _ => logger.debug(s"✖ ${f(shape)} is not TextHolder")
            }

        }
    }

    def processGroupShape(groupShape: XSLFGroupShape) {
        logger.debug(s"⸮ Inspecting XSLFGroupShape[${groupShape.getShapeName()}]...")
        findControlJsonPath(groupShape.getShapes().asScala) match {
            case None => logger.debug(s"✖ XSLFGroupShape[${groupShape.getShapeName()}] - no control, ignoring")
            case Some(GroupShapeControlData(shape, controlMatch, jsonPath, direction, gap)) => {
                logger.debug(s"√ Found the XSLFGroupShape[${groupShape.getShapeName()}] with control fields")
                // from here we get a Exception in thread "main" java.util.ConcurrentModificationException
                // val newGroupShape: XSLFGroupShape = groupShape.getSheet().createGroup()
                // newGroupShape.setAnchor(groupShape.getAnchor())

                for (shape <- groupShape.getShapes().asScala) {
                    shape match {
                        case s: XSLFAutoShape => logger.debug("not impl yet") // newGroupShape.createAutoShape()
                        case t: XSLFTextBox   => logger.debug("groupshape - not yet") // newGroupShape.createTextBox().setText(t.getText())
                        case _ => logger.error(s"The group ${prettyPrintShape(groupShape)} has a shape ${prettyPrintShape(shape)} that did not match")
                    }

                }
                // nned to remove it on the outer /// groupShape.removeShape(shape)
            }
        }
    }

    def iterateTable(table: XSLFTable, dc: DataContext) {
        val firstCell = table.getRows().get(0).getCells().get(0)
        val firstCellText = firstCell.getText()
        // val (a, b, tableContextJsonPath) = for (m <- matchContextControl.findFirstMatchIn(firstCellText)) yield m.group
        val mm = for (m <- matchContextControl.findFirstMatchIn(firstCellText)) yield m
        val tableContextJsonPath = mm.get.group(2)
        val controlString = mm.get.group(0)

        val (jsonNode, _tableContextJsonPath) = nodeAndQuery(tableContextJsonPath, dc)

        JsonPath.query(_tableContextJsonPath, jsonNode).map{ iter =>
            for ((jsonNode, i) <- iter.zipWithIndex) {
                logger.debug(s"√ Table node = ${jsonNode.toString()}")
                RowCloner.cloneRow(table, 1) // including cells
                for ((cell, ci) <- table.getRows().get(table.getRows().size() - 1).getCells().asScala.zipWithIndex) {
                    // cell.setText(r)
                    cell.setStrokeStyle(cell.getStrokeStyle())
                       // get the current style and size etc
                       // src cell
                    val srcCell = table.getRows().get(0).getCells().get(ci)
                    val newText = replaceText(dc.copy(contextJsonNode = Some(jsonNode)), cell.getText())
                    applyNewText(newText, cell, srcCell)
                }
            }
        }

        // now remove the template row
        table.removeRow(1)
        // remove the context string
        //table.getRows().get(0).getCells().get(0).setText(firstCellText.replace(controlString, ""))


        applyNewText(firstCellText.replace(controlString, ""), firstCell, firstCell)



    }

    /**
     * Replaces the text in the "textShape", cloning the style out of the "first" paragraph textrun
     */
    def applyNewText(newText: String, textHolder: XSLFTextShape, original: XSLFTextShape) {
        logger.debug(s"√ applyNewText :: [${textHolder.getText()}] newText = [${newText}]")
        val color = original.getTextParagraphs().get(0).getTextRuns().get(0).getFontColor()
        val font = original.getTextParagraphs().get(0).getTextRuns().get(0).getFontFamily()
        val size = original.getTextParagraphs().get(0).getTextRuns().get(0).getFontSize()

        textHolder.setText("") // effectively removes all paragraphs and textruns

        val tr:XSLFTextRun = textHolder.appendText(newText,true)
        tr.setFontColor(color)
        tr.setFontFamily(font)
        tr.setFontSize(size)


    }

    def changeText(textShape: TextHolder, dc: DataContext) {
        val text = textShape.getText()
        val matchRegexpTemplate.r(replacingText, jsonQuery) = text

        logger.debug(s"√ found = [${replacingText}] jsonpath = [${jsonQuery}]")

        val (theJsonNode, theJsonPath) = nodeAndQuery(jsonQuery, dc)

        val dataText = JsonPath.query(theJsonPath, theJsonNode) match {
            case Left(error) => throw new java.lang.Error(error.reason)
            case Right(i)    => try {
                Some(i.next().asText())
            } catch {
                case e: java.util.NoSuchElementException => logger.error(s"The JSONPath expression ==> ${theJsonPath} <== did not resolve to any data. Ignoring"); None
                case e: Exception => throw e
            }
        }

        dataText.map { txt =>
            val newText = text.replace(replacingText, txt)

            applyNewText(newText, textShape.asInstanceOf[XSLFTextShape], textShape.asInstanceOf[XSLFTextShape])

            /*
            // get the current style and size etc
            val color = textShape.getTextParagraphs().get(0).getTextRuns().get(0).getFontColor()
            val font = textShape.getTextParagraphs().get(0).getTextRuns().get(0).getFontFamily()
            val size = textShape.getTextParagraphs().get(0).getTextRuns().get(0).getFontSize()
            logger.debug(s"√ dataText = [${txt}] newText = [${newText}]")

            textShape.setText("")
            val tr:XSLFTextRun = textShape.appendText(newText,true).asInstanceOf[XSLFTextRun]
            tr.setFontColor(color)
            tr.setFontFamily(font)
            tr.setFontSize(size)
            */
        }
    }

    def replaceText(dc: DataContext, text: String) :String = {
        // fullTemplateText is {{someJsonPath}}
        // jsonQuery is what was found
        val matchRegexpTemplate.r(fullTemplateText, jsonQuery) = text

        val (jsonNode, jsonPath) = nodeAndQuery(jsonQuery, dc)

        val dataText = JsonPath.query(jsonPath, jsonNode) match {
            case Left(error) => throw new java.lang.Error(error.reason)
            case Right(i)    => try {
                Some(i.next().asText())
            } catch {
                case e: java.util.NoSuchElementException => System.err.println(s"The JSONPath expression ==> ${jsonQuery} <== did not resolve to any data. Ignoring"); None
                case e: Exception => throw e
            }
        }

        dataText.map( text.replace(fullTemplateText, _) ).getOrElse("")

    }

    /**
     * Given a JSONPath query ($. or @.) we will determine
     * if it is to use the context node [@.], or the rootnode [$.]
     * for the lookup
     *
     * ! This is a hack because the io.gatling.JsonPath does not support a context object.
     */
    def nodeAndQuery(jsonQuery: String, dc: DataContext): (JsonNode, String) = {
        jsonQuery(0) match {
            case '$' => (dc.rootJsonNode,        jsonQuery)
            case '@' => dc.contextJsonNode match {
                            case Some(jn) => (jn, "$" + jsonQuery.stripPrefix("@"))
                            case None     => {
                                val newJsonPath = "$" + jsonQuery.stripPrefix("@")
                                logger.warn(s"! jsonPath starts is ${jsonQuery} but the context object is empty. Using root object instead (eg: ${newJsonPath})")
                                (dc.rootJsonNode, "$" + jsonQuery.stripPrefix("@"))
                            }
                        }
        }
    }

    def addImage(ppt: XMLSlideShow, slide: XSLFSlide, imagePath: String, imageShapeName: String, shape: XSLFShape, pictureType: PictureType) {

        val picIS: FileInputStream = new FileInputStream(new File(imagePath))
        // https://stackoverflow.com/questions/4905393/scala-inputstream-to-arraybyte commons-io still the best
        val picture: Array[Byte]       = org.apache.commons.io.IOUtils.toByteArray(picIS)

        val anchor: Rectangle2D = shape.getAnchor()
        slide.removeShape(shape)

        val pd: XSLFPictureData    = ppt.addPicture(picture, pictureType)
        val pics: XSLFPictureShape = slide.createPicture(pd)
        pics.setAnchor(anchor)

    }

    trait ControlData {
        // the shape that holds the matching control string (for removal by caller)
        def shape: XSLFShape
        // the string that matched
        def controlText: String
        // the jsonPath
        def jsonPath: String
    }
    case class PageControlData(shape: XSLFShape, controlText: String, jsonPath: String) extends ControlData
    case class GroupShapeControlData(shape: XSLFShape, controlText: String, jsonPath: String, direction: Int, gap: Int) extends ControlData
    case class ImageControlData(shape: XSLFShape, controlText: String, jsonPath: String) extends ControlData


    def findControlJsonPath(shapes: scala.collection.mutable.Seq[XSLFShape]): Option[ControlData] = {

        for (shape <- shapes) {
            if (shape.isInstanceOf[TextHolder]) {
                val textShape = shape.asInstanceOf[TextHolder]
                logger.debug(s"⸮ inspecting - Shape[${shape.getShapeName()}] `${textShape.getText()}`")
                textShape.getText() match {
                    case matchContextControl(controlText, jsonPath) => { // page control
                        logger.debug(s"√ Match - Shape[${shape.getShapeName()}] control=`${controlText}` jp=`${jsonPath}`")
                        return Some(PageControlData(shape, controlText, jsonPath))
                    }
                    case matchGroupShapeControl(controlText, jsonPath, direction, gap) => {
                        logger.debug(s"√ Match - Shape[${shape.getShapeName()}] control=`${controlText}` jp=`${jsonPath}` dir=${direction} gap=${gap}")
                        return Some(GroupShapeControlData(shape, controlText, jsonPath, direction.toInt, gap.toInt))
                    }
                    case _ => {
                        logger.debug(s"✖ shape:${shape.getShapeName()} did not have a controlData")

                    }
                }
            }
        }
        None
    }

    def getShape(slide: XSLFSlide, shapeName: String): Option[XSLFShape] = {
        for (shape <- slide.getSlideLayout().getShapes().asScala) {
            shape.getShapeName().toLowerCase() match {
              case shapeName => return Some(shape)
              case _ =>
            }
        }
        return None;
    }

    def hasTemplate(shape: TextHolder): Boolean =
      shape.getText().matches(matchRegexpTemplate)

    def hasControl(shape: TextHolder): Boolean = matchContextControl.findFirstIn(shape.getText()).isDefined


    def prettyPrintSlide(slide: XSLFSlide): String = {
        List(Option(slide.getSlideName()), Some(s"#${slide.getSlideNumber()}")).flatten.mkString("/")
    }

    def prettyPrintShape(shape: XSLFShape): String = shape.getShapeName()


}

