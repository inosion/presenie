package inosion.presenie.pptx


import org.apache.poi.xslf.usermodel._ // XMLSlideShow
import org.apache.poi.sl.usermodel._
import java.io.File
import java.io._

// import scala.collection.JavaConverters._
import scala.jdk.CollectionConverters._


import org.openxmlformats.schemas.drawingml.x2006.main.CTTableRow

import com.filippodeluca.jsonpath.parser.JsonPathParser
import com.filippodeluca.jsonpath.circe.CirceSolver
import io.circe.Json

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

object PPTXMerger extends StrictLogging {
    type TextHolder = TextShape[XSLFShape, XSLFTextParagraph]

    import scala.reflect.ClassTag

    def f[T](v: T)(implicit ev: ClassTag[T]) = ev.toString

    val JsonPathReg                   = raw"[\$$@]|[\$$@]\.[0-9A-Za-z_:\-\?\.\[\]\*]+"
    val matchRegexpTemplate           = raw".*(\{\{\s*(" + JsonPathReg + raw")\s*\}\}).*"
    val matchContextControl           = (raw"(\{\s*context\s*=\s*(" + JsonPathReg + raw")\s*\})").r
    val matchGroupShapeControl        = (raw"(\{\s*context\s*=\s*(" + JsonPathReg + raw")\s*dir\s*=\s*(\d+)\s+gap=(\d+)\s*\})").r

    def render(data: File, template: File, outFile: File): Either[Error, Unit] = {

        // we need to track which SlideLayouts we ported across
        val visitedLayouts: mutable.Seq[XSLFSlideLayout] = mutable.Seq()
        logger.debug(s"! Regexp for templates = ${matchRegexpTemplate}")
        logger.debug(s"! Regexp for controls = ${matchContextControl}")
        logger.debug(s"! Regexp for group shape controls = ${matchGroupShapeControl}")

        val pptTemplate: XMLSlideShow = new XMLSlideShow(new FileInputStream(template.getAbsolutePath()))
        val pptNew: XMLSlideShow      = new XMLSlideShow(new FileInputStream(template.getAbsolutePath()))

        // fastest way to clone the sheet I can see is copy it, then remove all slides
        // to retain the master slide layouts.
        for (idx <- 0 until pptNew.getSlides.size()) {
            logger.debug(s"√ Removing slide ${prettyPrintSlide(pptNew.getSlides().get(0))}")
            pptNew.removeSlide(0)
        }

        val jsonData = JsonYamlTools.readFileToJson(data) match {
            case Right(a) => a
            case Left(e) => throw e
        }


        for (srcSlide <- pptTemplate.getSlides().asScala) {
            processSlide(template.getPath(), outFile.getPath, srcSlide, pptNew, jsonData, None, visitedLayouts)
        }

        pptNew.write(new FileOutputStream(outFile))

        Right(())

    }

    def prettyPrintSlide(slide: XSLFSlide): String = {
        List(Option(slide.getSlideName()), Some(s"#${slide.getSlideNumber()}")).flatten.mkString("/")
    }

    def prettyPrintShape(shape: XSLFShape): String = shape.getShapeName()


    def processSlide(srcPath: String, destPath: String, srcSlide: XSLFSlide, pptNew: XMLSlideShow, rootJson: Json, contextJson: Option[Json], visitedLayouts: mutable.Seq[XSLFSlideLayout]) : Unit = {

        // We will copy this slide from source to new ppt, then fill it's data
        val newSlide = PPTXTools.createSlide(pptNew, srcSlide, visitedLayouts)
        System.err.println(s":: ------ New Slide from src [${srcPath}(${srcSlide.getSlideNumber()})] to [${destPath}(${newSlide.getSlideNumber()})] ")

        // on the new slide, locate all control fields, any that are page based, we will process
        findControlJsonPath(newSlide.getShapes().asScala) match {

            // page based control fields
            case Some(PageControlData(shape, contextMatch, jsonPath)) => {
                logger.debug(s"√ Slide [${prettyPrintSlide(newSlide)}] has a PageControl - jsonPath context -> ${jsonPath}")

                // remove controlShape
                newSlide.removeShape(shape)

                val result = JsonPathParser.parse(jsonPath).map { jp =>
                    CirceSolver.solve(jp, rootJson)
                }

                result.map{ iter =>
                    for ((jsonNode, i) <- iter.zipWithIndex) {
                        logger.debug(s"√ node = ${jsonNode.toString()}")
                        processSlide(srcPath, destPath, newSlide, pptNew, rootJson, Some(jsonNode), visitedLayouts)
                    }
                }

                // now we have templated it out, let's remove it
                pptNew.removeSlide(newSlide.getSlideNumber() - 1)
            }
            // unsupported control fields
            case Some(_) => logger.error(s"! Slide [${prettyPrintSlide(newSlide)}] has a control but it is not a PageControlData")

            // no page control fields found, process all shapes
            case None => processAllShapes(newSlide, rootJson, contextJson)
        }

    }

    def processAllShapes(slide: XSLFSlide,  rootJson: Json, contextJson: Option[Json]) : Unit = {

        for (shape <- slide.getShapes().asScala) {

            shape match {
                case textShape : TextHolder =>
                    if (hasTemplate(textShape)) {
                      logger.debug(s"√ ${f(shape)} is a templated shape")
                      changeText(textShape, rootJson, contextJson)
                    } else {
                      logger.debug(s"✖ '${textShape.getText()}' did not match")
                    }

                case group : XSLFGroupShape => processGroupShape(group)
                case table : XSLFTable  if (hasControl(table.getRows().get(0).getCells().get(0)))    => {
                    logger.debug(s"√ we have a table - ${table.getRows().get(0).getCells().get(0).getText()}")
                    iterateTable(table, rootJson, contextJson)
                }
                case _ => logger.debug(s"✖ ${f(shape)} is not TextHolder")
            }

        }
    }

    def processGroupShape(groupShape: XSLFGroupShape) : Unit = {
        logger.debug(s"⸮ Inspecting XSLFGroupShape[${groupShape.getShapeName()}]...")
        findControlJsonPath(groupShape.getShapes().asScala) match {
            case None => logger.debug(s"✖ XSLFGroupShape[${groupShape.getShapeName()}] - no control, ignoring")
            case Some(GroupShapeControlData(shape, controlMatch, jsonPath, direction, gap)) => {
                logger.debug(s"√ Found the XSLFGroupShape[${groupShape.getShapeName()}] with control fields")
                val newGroupShape: XSLFGroupShape = groupShape.getSheet().createGroup()
                newGroupShape.setAnchor(groupShape.getAnchor())

                for (shape <- groupShape.getShapes().asScala) {
                    shape match {
                        case s: XSLFAutoShape => logger.debug("not impl yet") // newGroupShape.createAutoShape()
                        case t: XSLFTextBox   => newGroupShape.createTextBox().setText(t.getText())
                        case _ => logger.error(s"The group ${prettyPrintShape(groupShape)} has a shape ${prettyPrintShape(shape)} that did not match")
                    }

                }
                // nned to remove it on the outer /// groupShape.removeShape(shape)
            }
        }
    }

    def iterateTable(table: XSLFTable, rootJson: Json, contextJson: Option[Json]) : Unit = {
        val firstCellText = table.getRows().get(0).getCells().get(0).getText()
        // val (a, b, tableContextJsonPath) = for (m <- matchContextControl.findFirstMatchIn(firstCellText)) yield m.group
        val mm = for (m <- matchContextControl.findFirstMatchIn(firstCellText)) yield m
        val tableContextJsonPath = mm.get.group(2)
        val controlString = mm.get.group(0)

        val (jsonNode, _tableContextJsonPath) = nodeAndQuery(tableContextJsonPath, rootJson, contextJson)


        val jpResult = JsonPathParser.parse(_tableContextJsonPath).map { jp =>
            CirceSolver.solve(jp, jsonNode)
        }

        jpResult.map{ iter =>
            for ((jsonNode, i) <- iter.zipWithIndex) {
                logger.debug(s"√ Table node = ${jsonNode.toString()}")
                RowCloner.cloneRow(table, 1) // including cells
                for ((cell, ci) <- table.getRows().get(table.getRows().size() - 1).getCells().asScala.zipWithIndex) {
                    changeText(cell, rootJson, Some(jsonNode))
                    cell.setStrokeStyle(cell.getStrokeStyle())
                }
            }
        }

        // now remove the template row
        table.removeRow(1)
        // remove the context string
        table.getRows().get(0).getCells().get(0).setText(firstCellText.replace(controlString, ""))


    }



    def changeText(textShape: TextHolder, rootJsonNode: Json, contextJsonNode: Option[Json]) : Unit = {
        val text = textShape.getText()
        val matchRegexpTemplate.r(replacingText, jsonQuery) = text

        logger.debug(s"√ found = [$replacingText] jsonpath = [$jsonQuery]")

        val (jsonNode, jsonPath) = nodeAndQuery(jsonQuery, rootJsonNode, contextJsonNode)

        val matchedText = JsonPathParser.parse(jsonPath).map { jp =>
            CirceSolver.solve(jp, jsonNode)
        }

        matchedText.map { iter =>
            for ((jsonNode, i) <- iter.zipWithIndex) {
                logger.debug(s"√ node = ${jsonNode.toString()}")
                val newText = text.replace(replacingText, {jsonNode.asString.getOrElse(jsonNode.toString())})
                logger.debug(s"√ dataText = [${jsonNode.toString()}] newText = [$newText]")
                textShape.setText(newText)
            }
        }
    }

    /**
     * Given a JSONPath query ($. or @.) we will determine
     * if it is to use the context node [@.], or the rootnode [$.]
     * for the lookup
     *
     * ! This is a hack because the io.gatling.JsonPath does not support a context object.
     */
    def nodeAndQuery(jsonQuery: String, rootJson: Json, contextJson: Option[Json]): (Json, String) = {
        jsonQuery(0) match {
            case '$' => (rootJson,        jsonQuery)
            case '@' => contextJson match {
                            case Some(jn) => (jn, "$" + jsonQuery.stripPrefix("@"))
                            case None     => {
                                val newJsonPath = "$" + jsonQuery.stripPrefix("@")
                                logger.warn(s"! jsonPath starts is ${jsonQuery} but the context object is empty. Using root object instead (eg: ${newJsonPath})")
                                (rootJson, "$" + jsonQuery.stripPrefix("@"))
                            }
                        }
        }
    }

    def addImage(ppt: XMLSlideShow, slide: XSLFSlide, imagePath: String, imageShapeName: String, shape: XSLFShape, pictureType: PictureType) : Unit = {

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


    /**
     * Given a list of shapes, we will find the first shape that has a control string
     * and return the control data
     * ControlData is a looping object on the slide
     */
    def findControlJsonPath(shapes: scala.collection.mutable.Seq[XSLFShape]): Option[ControlData] = {

        for (shape <- shapes) {
            if (shape.isInstanceOf[TextHolder]) {
                val textShape = shape.asInstanceOf[TextHolder]
                logger.debug(s"⸮ inspecting - Shape[${shape.getShapeName()}] `${textShape.getText()}`")
                textShape.getText() match {
                    case matchContextControl(controlText, jsonPath) => { // page control
                        logger.debug(s"√ Match - PageControl - Shape[${shape.getShapeName()}] control=`${controlText}` jp=`${jsonPath}`")
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

}

