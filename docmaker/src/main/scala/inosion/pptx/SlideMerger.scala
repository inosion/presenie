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
import java.awt.Color

object SlideMergerX extends StrictLogging {

    type JsonNodeStack = List[JsonNode]

    def prettyPrintSlide(slide: XSLFSlide): String = {
        List(Option(slide.getSlideName()), Some(s"#${slide.getSlideNumber()}")).flatten.mkString("/")
    }

    def prettyPrintShape(shape: XSLFShape): String = shape.getShapeName()


    import scala.reflect.ClassTag

    def f[T](v: T)(implicit ev: ClassTag[T]) = ev.toString

    val jpre                          = raw"[\$$@]|[\$$@]\.[0-9A-Za-z_:\-\?\.\[\]\*]+"
    val matchRegexpTemplate           = raw".*(\{\{\s*(" + jpre + raw")\s*\}\}).*"
    val matchContextControl           = (raw"(\{\s*context\s*=\s*(" + jpre + raw")\s*\})").r
    val matchGroupShapeControl        = (raw"(\{\s*context\s*=\s*(" + jpre + raw")\s*dir\s*=\s*(\d+)\s+gap=(\d+)\s*\})").r

    def render2(fc: FilesContext): Either[Error, Unit] = {

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
            templateSlide(srcSlide, fc, sc, jn = List(jsonData))
        }

           // save out the resulting new pptx
        try {
            pptNew.write(new FileOutputStream(fc.dst))
            Right(())
        } catch {
            case e: Exception => Left(new Error(s"failed writing the new pptx [${fc.dstAbsPath}]", e))
        }
    }

    /**
     * Curried function
     * Takes the srcslide, and optionally control data object.
     * Clones the slide (importSource)
     * if controlData shape (that determines repeating a slide) , remove it from this slide
     * return the new slide
     */
    def cloneSlide(pptx: XMLSlideShow) (srcSlide: XSLFShapeContainer, pcd: Option[ControlData]): XSLFShapeContainer = {

        val newSlide: XSLFSlide = pptx.createSlide();
        PPTXTools.copySlideContent(srcSlide.asInstanceOf[XSLFSlide], newSlide)
        newSlide.importContent(srcSlide.asInstanceOf[XSLFSlide])
        newSlide
    }

    def cloneGroupShape(slide: XSLFSlide) (srcSlide: XSLFShapeContainer, pcd: Option[ControlData]): XSLFShapeContainer = {

        slide.createGroup()
    }

    type AddToParentFN = (XSLFShapeContainer, Option[ControlData]) => XSLFShapeContainer

    def notControlShapes(ctrlShape: XSLFTextShape, shapesIter: scala.collection.Iterable[XSLFShape]) =
        shapesIter.filterNot(s => {
            // s match {
            //     case x: XSLFTextShape => logger.debug(s"[${x.getShapeName()}]${x.getText()} --> [${ctrlShape.getShapeName()}]${ctrlShape.getText()}");
            //     case _ =>
            // };

            s.getShapeName().equals(ctrlShape.getShapeName())
          }
        )

    def templateSlide(srcSlide: XSLFSlide, fc: FilesContext, sc: SlidesContext, jn: JsonNodeStack) {

        // if there is a context, load and push on stack
        findControlJsonPath(srcSlide.getShapes().asScala) match {
            // it is a slide, with page control
            case pcd @ Some(PageControlData(ctrlShape, controlText, jsonPath)) => {
                JsonPath.query(jsonPath, jn.head).map { i =>
                    for (jsonNode <- i) {
                        logger.debug(s"√ (slide w/- JsonNode = ${jsonNode.toString().subSequence(0,20)}...")
                        val newSlide = cloneSlide(sc.destPptx) (srcSlide, None).asInstanceOf[XSLFSlide]

                        //templateShapes(newSlide, srcSlide, newSlide, fc,sc,jn :+ jsonNode)
                    }
                }

            }
            case None => cloneSlide(sc.destPptx) (srcSlide, None)
        }
    }

    /**
     * We are always iterating over shapes from source (as we don't modify those)
     * And "applying" transformations to the "sibling" shape in the destination
     * The shapes could be
     * - all shapes in a group shape
     * - all shapes on a slide
     */
    def templateShapes(destSlide: XSLFSlide
                    , sourceShapeContainer: XSLFShapeContainer
                    , destShapeContainer: XSLFShapeContainer
                    , fc: FilesContext
                    , sc: SlidesContext
                    , jn: JsonNodeStack) {

        for (shape <- sourceShapeContainer.getShapes().asScala) {

            // if shape container (processShapes)
            logger.debug(s"on Container: ${shape.getParent()} => ${prettyPrintShape(shape)}")

            shape match {
                case textShape : XSLFTextShape =>
                    if (hasTemplate(textShape)) {
                      logger.debug(s"√ ${f(shape)} -> ${textShape.getText()}")

                      val destTextShape = destShapeContainer.getShapes().asScala.filter { x =>
                        logger.debug(s"∞ comparing ${x.getShapeName()} with ${shape.getShapeName()}")
                        x.getShapeName().equalsIgnoreCase(shape.getShapeName())
                      }.head.asInstanceOf[XSLFTextShape]

                      //applyNewText(templateOutNewText(textShape.getText(), jn), destTextShape, textShape)
                      logger.debug(s"√ now == ${f(shape)} -> ${textShape.getText()}")

                    } else {
                      logger.debug(s"✖ '${textShape.getText()}' did not match")
                    }

                case srcGroupShape : XSLFGroupShape =>
                    templateShapes(destSlide, srcGroupShape, null, fc ,sc ,jn)
                case table : XSLFTable  if (hasControl(table.getRows().get(0).getCells().get(0)))    => {
                     logger.debug(s"√ we have a table - ${table.getRows().get(0).getCells().get(0).getText()}")
                     iterateTable(table, jn)
                }
                case _ => logger.debug(s"✖ ${f(shape)} is not TextHolder")
            }


        }

        // remove the context-holding-shape
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

    def hasControl(shape: XSLFTextShape): Boolean = matchContextControl.findFirstIn(shape.getText()).isDefined

    def hasTemplate(shape: XSLFTextShape): Boolean = shape.getText().matches(matchRegexpTemplate)
        /**
     * Given a JSONPath query ($. or @.) we will determine
     * if it is to use the context node [@.], or the rootnode [$.]
     * for the lookup
     *
     * ! This is a hack because the io.gatling.JsonPath does not support a context object.
     */
    def nodeAndQuery(jsonQuery: String, jn: List[JsonNode]): (JsonNode, String) = {
        jsonQuery(0) match {
            case '$' =>
                (jn.head, jsonQuery)
            case '@' => {
                val newJsonPath = "$" + jsonQuery.stripPrefix("@")
                if (jn.size == 1) logger.warn(s"! jsonPath starts is ${jsonQuery} but there is no context JsonNode object (just root). Query changed (eg: ${newJsonPath})")
                (jn.last, newJsonPath)
            }
        }
    }


    def findControlJsonPath(shapes: scala.collection.Iterable[XSLFShape]): Option[ControlData] = {

        for (shape <- shapes) {
            if (shape.isInstanceOf[XSLFTextShape]) {
                val textShape = shape.asInstanceOf[XSLFTextShape]
                logger.debug(s"⸮ inspecting - Shape[${shape.getShapeName()}] `${textShape.getText()}`")
                textShape.getText() match {
                    case matchContextControl(controlText, jsonPath) => { // page control
                        logger.debug(s"√ Match - Shape[${shape.getShapeName()}] control=`${controlText}` jp=`${jsonPath}`")
                        return Some(PageControlData(textShape, controlText, jsonPath))
                    }
                    case matchGroupShapeControl(controlText, jsonPath, direction, gap) => {
                        logger.debug(s"√ Match - Shape[${shape.getShapeName()}] control=`${controlText}` jp=`${jsonPath}` dir=${direction} gap=${gap}")
                        return Some(GroupShapeControlData(textShape, controlText, jsonPath, direction.toInt, gap.toInt))
                    }
                    case _ => {
                        logger.debug(s"✖ shape:${shape.getShapeName()} did not have a controlData")

                    }
                }
            }
        }
        None
    }

    def templateOutNewText(templatableText: String, jn: List[JsonNode]) : String = {
        val matchRegexpTemplate.r(replacingText, jsonQuery) = templatableText

        logger.debug(s"√ found = [${replacingText}] jsonpath = [${jsonQuery}]")

        val (theJsonNode, theJsonPath) = nodeAndQuery(jsonQuery, jn)

        val dataText = JsonPath.query(theJsonPath, theJsonNode) match {
            case Left(error) => throw new java.lang.Error(error.reason)
            case Right(i)    => try {
                i.toSeq match {
                    case s if s.length > 1 => Some(s.map( j => s"• ${j.asText()}").mkString("\n"))
                    case s => Some(s.head.asText())
                }
            } catch {
                case e: java.util.NoSuchElementException => logger.error(s"The JSONPath expression ==> ${theJsonPath} <== did not resolve to any data. Ignoring"); None
                case e: Exception => throw e
            }
        }

        dataText.map { txt => templatableText.replace(replacingText, txt) }.getOrElse("")
    }

    /**
     * Replaces the text in the "textShape", cloning the style out of the "first" paragraph textrun
     */
    def applyNewText(newText: String, textShape: XSLFTextShape, original: XSLFTextShape) {
        logger.debug(s"√ applyNewText :: [${textShape.getText().take(10)}] newText = [${newText.take(10)}]")
        val color = original.getTextParagraphs().get(0).getTextRuns().get(0).getFontColor()
        val font = original.getTextParagraphs().get(0).getTextRuns().get(0).getFontFamily()
        val size = original.getTextParagraphs().get(0).getTextRuns().get(0).getFontSize()

        textShape.clearText()

        // ignore the val tr:XSLFTextRun = textShape.appendText(newText,true)
        textShape.appendText(newText,true)

        for (pr <- textShape.getTextParagraphs().asScala) {
            for (tr <- pr.getTextRuns().asScala) {
                tr.setFontColor(color)
                tr.setFontFamily(font)
                tr.setFontSize(size)
            }
        }
    }

    def iterateTable(table: XSLFTable, jn: List[JsonNode]) {
        val firstCell = table.getRows().get(0).getCells().get(0)
        val firstCellText = firstCell.getText()
        val mm = for (m <- matchContextControl.findFirstMatchIn(firstCellText)) yield m
        val tableContextJsonPath = mm.get.group(2)
        val controlString = mm.get.group(0)

        val (jsonNode, _tableContextJsonPath) = nodeAndQuery(tableContextJsonPath, jn)

        JsonPath.query(_tableContextJsonPath, jsonNode).map{ iter =>
            for ((jsonNode, i) <- iter.zipWithIndex) {
                logger.debug(s"√ Table node = ${jsonNode.toString().take(10)}")
                RowCloner.cloneRow(table, 1) // including cells
                for ((cell, ci) <- table.getRows().get(table.getRows().size() - 1).getCells().asScala.zipWithIndex) {
                    cell.setStrokeStyle(cell.getStrokeStyle())
                    val srcCell = table.getRows().get(0).getCells().get(ci)
                    tableCellApplyBorders(cell,srcCell)
                    applyNewText(templateOutNewText(cell.getText(), jn :+ jsonNode), cell, srcCell)
                }
            }
        }

        // now remove the template row
        table.removeRow(1)
        // remove the context string

        applyNewText(firstCellText.replace(controlString, ""), firstCell, firstCell)
    }

    def tableCellApplyBorders(cell: XSLFTableCell, srcCell: XSLFTableCell) {
        for (beType <- List(TableCell.BorderEdge.bottom, TableCell.BorderEdge.top, TableCell.BorderEdge.right, TableCell.BorderEdge.left)) {
            if(srcCell.getBorderCap(beType) != null)      cell.setBorderCap(     beType,    cell.getBorderCap(beType))
            if(srcCell.getBorderColor(beType) != null)    cell.setBorderColor(   beType,    srcCell.getBorderColor(beType))
            if(srcCell.getBorderCompound(beType) != null) cell.setBorderCompound(beType,    cell.getBorderCompound(beType))
            if(srcCell.getBorderDash(beType) != null)     cell.setBorderDash(    beType,    cell.getBorderDash(beType))
            if(srcCell.getBorderStyle(beType) != null)    cell.setBorderStyle(   beType,    srcCell.getBorderStyle(beType))
            if(srcCell.getBorderWidth(beType) != null)    cell.setBorderWidth(   beType,    srcCell.getBorderWidth(beType))
            cell.setBorderWidth(   beType,    1)
            cell.setBorderColor(   beType,    Color.DARK_GRAY)
          }
    }
}