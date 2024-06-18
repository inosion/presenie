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


object SlideMerger extends StrictLogging {

    type JsonNodeStack = List[JsonNode]

    val GroupShapeDefaultDir = 90
    val GroupShapeDefaultGap = 0.1

    def prettyPrintSlide(slide: XSLFSlide): String = {
        List(Option(slide.getSlideName()), Some(s"#${slide.getSlideNumber()}")).flatten.mkString("/")
    }

    def prettyPrintShape(shape: XSLFShape): String = shape.getShapeName()

    import scala.reflect.ClassTag

    def f[T](v: T)(implicit ev: ClassTag[T]) = ev.toString

    val jpre                          = raw"[\$$@]|[\$$@]\.[0-9A-Za-z_:\-\?\.\[\]\*]+"
    val matchRegexpTemplate           = raw".*(\{\{\s*(" + jpre + raw")\s*\}\}).*"
    val matchContextControl           = (raw"(\{\s*context\s*=\s*(" + jpre + raw")\s*\})").r
    val matchGroupShapeControl        = (raw"(\{\s*context\s*=\s*(" + jpre + raw")\s*(dir\s*=\s*(\d+))?\s*(gap=(\d+))?\s*\})").r
    val matchImageControl             = (raw"(\{\s*image\s+path\s*=\s*(" + jpre + raw")\s*\})").r


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
            produceSlides(srcSlide, srcSlide.getShapes().asScala, fc, sc, jn = List(jsonData), cd = None )
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
    def cloneSlide(pptx: XMLSlideShow) (srcSlide: XSLFShapeContainer, pcd: Option[ControlData]): XSLFSlide = {

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

    /**
     * We are always iterating over shapes from source (as we don't modify those)
     * And "applying" transformations to the "sibling" shape in the destination
     * The shapes could be
     * - all shapes in a group shape
     * - all shapes on a slide
     */
    def templateShapes(sourceContainer: XSLFShapeContainer
                    , fc: FilesContext
                    , sc: SlidesContext
                    , jn: JsonNodeStack
                    , cd: Option[ControlData]
                    , destContainer: XSLFShapeContainer) {

        /* this section creates a new "dest" container */
        logger.debug(s"templateShapes: on ${f(sourceContainer)}")

        // mark the contextShape for removal (later)

        // remove the control data
        if (cd.isDefined) {
            val remove = destContainer.getShapes().asScala.filter(_.getShapeName().equalsIgnoreCase(cd.get.shape.getShapeName())).head
            destContainer.removeShape(remove)
        }


        /*
         * This is quite tricky. We iterate over the sourceContainer (the source slide) and
         * when we find a shape we like, we see if it needs to be templated.
         * But, we need to locate it, in the destination slide. I assume there is no "correlation"
         * from source to destination shape... only that the names and the text will be the same
         * .. that won't hold when we begin "cloning the shapes.. but that is a recursive iteration
         * task
         */
        logger.debug(s"iter of shapes on [${sourceContainer}] => templating shapes on [${destContainer}]")
        for (srcShape <- sourceContainer.getShapes().asScala) {

            // if shape container (processShapes)
            logger.debug(s"=> ${prettyPrintShape(srcShape)}")

            srcShape match {
                case srcTextShape : XSLFTextShape if (hasTemplate(srcTextShape)) => templateTextShape(destContainer, srcTextShape, jn)

                case srcTextShape : XSLFTextShape => logger.debug(s"✖ ${prettyPrintShape(srcTextShape)} (text)-> '${srcTextShape.getText()}' had no template '{{...}}' to change")

                case srcGroupShape : XSLFGroupShape => {
                     logger.debug(s"GroupShape:: size = ${srcGroupShape.getShapes.size()}")
                     // we need to find the matching groupShape
                     logger.debug(s"srcGroupShape: => ${srcGroupShape.getShapeName}")
                     //destContainer.getShapes().asScala.filter(_.isInstanceOf[XSLFGroupShape]).foreach(s => logger.debug(s"${s.getShapeName()}"))

                     val destGroupShape = destContainer.getShapes().asScala.filter { x =>
                        x.isInstanceOf[XSLFGroupShape] &&
                        x.getShapeName().equalsIgnoreCase(srcGroupShape.getShapeName())
                    }.head.asInstanceOf[XSLFGroupShape]
                     processGroupShape(srcGroupShape, fc, sc, jn, cd , destGroupShape)

                }

                case table : XSLFTable  if (hasControl(table.getRows().get(0).getCells().get(0)))    => {

                     val destTable = destContainer.getShapes().asScala.filter { t =>
                        t.isInstanceOf[XSLFTable] &&
                        t.getShapeName().equals(table.getShapeName())
                     }.head.asInstanceOf[XSLFTable]
                     logger.debug(s"√ we have a table - ${destTable.getRows().get(0).getCells().get(0).getText()}")
                     iterateTable(destTable, jn)
                }
                case _ => logger.debug(s"✖ ${f(srcShape)} is not TextHolder")
            }
        }

        // remove the context-holding-shape
    }

    /**
     * We have the container in the dest slide pack, that "has" this textShape we need to change the
     * text on. We don't "actually have it" because we are not iterating over it.
     * Iterating and modifying the collection at the same time is a no-no.
     */
    def templateTextShape(container: XSLFShapeContainer, srcTextShape: XSLFTextShape, jn: JsonNodeStack) {

        val destTextShape = container.getShapes().asScala.filter { x =>
            logger.debug(s"∞ comparing ${x.getShapeName()} with ${srcTextShape.getShapeName()}")
            x.isInstanceOf[XSLFTextShape] &&
            // x.getShapeName().equalsIgnoreCase(srcTextShape.getShapeName()) &&
            x.asInstanceOf[XSLFTextShape].getText().equals(srcTextShape.getText())
        }.head.asInstanceOf[XSLFTextShape]

        applyNewText(templateOutNewText(srcTextShape.getText(), jn), destTextShape, srcTextShape)

    }

    def processGroupShape(srcGroupShape: XSLFGroupShape, fc: FilesContext
                          , sc: SlidesContext
                          , jn: JsonNodeStack
                          , cd: Option[ControlData], destGroupShape: XSLFGroupShape) {

        // if there is a context, load and push on stack
        findControlJsonPath(destGroupShape.getShapes().asScala) match {

            case gcd @ Some(GroupShapeControlData(controlShape, control, jsonPath, direction, gap)) => {

                logger.debug(s"¡ It is a group shape w/- control :: ${prettyPrintShape(controlShape)} --> ${jsonPath},${direction},${gap}")

                // in will need to come back in here

                // cloneShapeToContainer(destContainer, groupShape)
                val (jsonNode, _groupShapeContextJsonPath) = nodeAndQuery(jsonPath, jn)

                val sourceAnchor = srcGroupShape.getAnchor()

                JsonPath.query(_groupShapeContextJsonPath, jsonNode).map { jsNodes =>
                    for ((groupJn, idx) <- jsNodes.zipWithIndex) {
                        logger.debug(s"√ groupShape = ${idx}/${groupJn.toString().subSequence(0, 20)}")

                        val clonedGroupShape = ShapeImporter.addShape(shape = srcGroupShape, srcSheet = sc.srcSlide, destSheet = sc.destSlides.last, sc = sc).asInstanceOf[XSLFGroupShape]

                        val anchor = clonedGroupShape.getAnchor()
                        anchor.setRect(srcGroupShape.getAnchor.getX + (idx*(gap + srcGroupShape.getAnchor.getWidth)), srcGroupShape.getAnchor.getY, srcGroupShape.getAnchor.getWidth, srcGroupShape.getAnchor.getHeight)
                        clonedGroupShape.setAnchor(anchor)

                        templateShapes(
                            srcGroupShape
                            , fc
                            , sc
                            , jn :+ groupJn
                            , gcd
                            , clonedGroupShape
                        )
                    }
                }
                val parent = destGroupShape.getParent
                parent.removeShape(destGroupShape)

            }

            case Some(ImageControlData(shape,controlText,jsonPath)) => logger.debug(s"TODO we found an image - ${jsonPath}")

            // no control data.. so just a normal GroupShape so just process it's inner shapes
            case None => {
                logger.debug(s"¡ normal groupShape (no Control Data) [${srcGroupShape}")
                templateShapes(
                             srcGroupShape
                             ,fc
                             ,sc
                             ,jn
                             ,None
                             ,destGroupShape
                             )
            }

        }

    }

    def cloneShapeToContainer(container: XSLFShapeContainer, srcShape: XSLFShape) {

        //ShapeImporter.addShape(shape = srcShape, )
        val newShape = srcShape match {
            case t: XSLFTextBox => cloneTextBox(container, t)
            case _ => logger.error("•••• cloning that shape is not supported")
        }

    }

    def cloneTextBox(container: XSLFShapeContainer, srcText: XSLFTextBox) {
        //val t: XSLFTextBox = container.createTextBox()
       // t.setBorderCap(srcShape.getBorderCap())
       // t.setBorderColor(srcShape.getBorderColor())

    }


    /**
     * We are always iterating over shapes from source (as we don't modify those)
     * And "applying" transformations to the "sibling" shape in the destination
     * The shapes could be
     * - all shapes in a group shape
     * - all shapes on a slide
     */
    def produceSlides(srcSlide: XSLFSlide
                    , srcIter: scala.collection.Iterable[XSLFShape]
                    , fc: FilesContext
                    , sc: SlidesContext
                    , jn: JsonNodeStack
                    , cd: Option[ControlData]) {

        /* this section creates a new "dest" container */
        logger.debug(s"produceSlides: on ${f(srcSlide)}")

        // if there is a context, load and push on stack
        findControlJsonPath(srcIter) match {
            // it is a slide, with page control
            case pcd @ Some(PageControlData(ctrlShape, controlText, jsonPath)) => {
                logger.debug(s"¡ It's a slide ${pcd}")
                JsonPath.query(jsonPath, jn.head).map { i =>
                    for (jsonNode <- i) {
                        logger.debug(s"√ node = ${jsonNode.toString().subSequence(0,20)}")
                        val newSlide = cloneSlide(sc.destPptx)(srcSlide, None)
                        templateShapes(
                            srcSlide
                             //,notControlShapes(ctrlShape, srcIter)
                             ,fc
                             ,sc.copy(srcSlide = srcSlide, destSlides = sc.destSlides :+ newSlide)
                             ,jn :+ jsonNode
                             ,pcd
                             ,newSlide
                             )
                    }
                }

            }

            // no control data.. so just a normal slide
            case None => {
                logger.debug(s"¡ no ctl data on this container [${srcSlide}")
                val newSlide = cloneSlide(sc.destPptx)(srcSlide, None)
                templateShapes(
                             srcSlide
                             ,fc
                             ,sc.copy(destSlides = sc.destSlides :+ newSlide)
                             ,jn
                             ,None
                             ,newSlide
                             )
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
     * ! This is a manipulation on the JsonPath standard because the io.gatling.JsonPath does not support a context object.
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


    /**
      * Finder method that determines, if on this container (iterator)
      * there is "context" meta data, then return what that JSON Path is,
      * and the type of control logic we think it is.
      * @param shapes
      * @return Option[ControlData]
      */
    def findControlJsonPath(shapes: scala.collection.Iterable[XSLFShape]): Option[ControlData] = {

        for (shape <- shapes) {
            if (shape.isInstanceOf[XSLFTextShape]) {
                val textShape = shape.asInstanceOf[XSLFTextShape]
                logger.debug(s"⸮ inspecting - Shape[${shape.getShapeName()}] `${textShape.getText()}`")
                textShape.getText() match {
                    case matchContextControl(controlText, jsonPath)
                        if shape.getParent.isInstanceOf[XSLFSlide] => { // page control
                        logger.debug(s"√ Match - Shape[${shape.getShapeName()}] control=`${controlText}` jp=`${jsonPath}`")
                        return Some(PageControlData(textShape, controlText, jsonPath))
                    }

                    case matchGroupShapeControl(controlText, jsonPath, null, null, null, null)
                        if shape.getParent.isInstanceOf[XSLFGroupShape] =>
                        return Some(GroupShapeControlData(textShape, controlText, jsonPath, GroupShapeDefaultDir, GroupShapeDefaultGap))

                    case matchGroupShapeControl(controlText, jsonPath, _, direction, _, gap)
                        if shape.getParent.isInstanceOf[XSLFGroupShape] =>
                        return Some(GroupShapeControlData(textShape, controlText, jsonPath, direction.toInt, gap.toInt))

                    case matchImageControl(controlText, jsonPath)
                        if shape.getParent.isInstanceOf[XSLFGroupShape] &&
                          shape.getParent.getShapes.size == 2 => { // we expect in this groupshape to be the text control, and the image
                        logger.debug(s"√ Match - Image path-to-image=`${jsonPath}`")
                        return Some(ImageControlData(textShape, controlText, jsonPath))
                    }

                    // we shouldn't really match here .. as we have the 'if' guards above
                    case matchContextControl(controlText, jsonPath) =>
                        logger.error(s"✖ shape:${shape.getShapeName()} has controlData, but not in a shape we know (${shape.getParent().isInstanceOf[XSLFGroupShape]})")
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

    /**
     * Template out the table
     */
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