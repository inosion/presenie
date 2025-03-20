package inosion.presenie.pptx

import org.slf4j.MDC

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
    val matchRegexpTemplate           = raw".*(\{\{\s*(?<jsonpath>" + JsonPathReg + raw")\s*\}\}).*"
    val matchContextControl           = (raw"(\{\s*context\s*=\s*(?<jsonpath>" + JsonPathReg + raw")\s*\})").r
    val _grpShape = raw"(?s)(\{\s*context\s*=\s*(?<jsonpath>" + JsonPathReg + raw")\s*dir\s*=\s*(?<dir>\d+)(\s+gap=(?<gap>\d+))?\s*\})"
    val matchGroupShapeControl        = _grpShape.r
    type SlideIndexMap = mutable.ListBuffer[Int]

    def render(data: File, template: File, outFile: File): Either[Error, Unit] = {

        // we need to track which SlideLayouts we ported across
        val visitedLayouts: mutable.Seq[XSLFSlideLayout] = mutable.Seq()

        logger.debug(s"! Regexp for templates = ${matchRegexpTemplate}")
        logger.debug(s"! Regexp for controls = ${matchContextControl}")
        logger.debug(s"! Regexp for group shape controls = ${matchGroupShapeControl}")

        // clone the slide 
        PPTXTools.clonePptx(template.toPath(), outFile.toPath())

        val pptTemplate: XMLSlideShow = new XMLSlideShow(new FileInputStream(template.getAbsolutePath()))
        val pptNew: XMLSlideShow      = new XMLSlideShow(new FileInputStream(outFile.getAbsolutePath()))

        val jsonData = JsonYamlTools.readFileToJson(data) match {
            case Right(a) => a
            case Left(e) => throw e
        }

        // because we are looping over slides, we can't "modify" the slide in place
        // so we loop over the source template, as Read-Only, and use the destination
        // slide to clone slides, and data from / to

        // track which slide index corresponds to which slide in the new ppt
        val slideIndexMap: SlideIndexMap = (0 to pptTemplate.getSlides().size() - 1).toList.to(mutable.ListBuffer)
        

        logger.debug(s"! Slide Count == ${pptTemplate.getSlides().size()}")
        for (slideIndex <- (0 to pptTemplate.getSlides().size() - 1)) {
              logger.debug(s"in loop slide count idx=${slideIndex} pptTemplate.size()=${pptTemplate.getSlides().size()}")
              logger.debug(s"! Slide index map = ${slideIndexMap.mkString(",")} Getting --> ${slideIndexMap(slideIndex)}")
              val sourceSlide = pptNew.getSlides().get(slideIndexMap(slideIndex))
              logger.debug(s"render: [${slideIndex} from ${slideIndexMap.mkString(",")}] ${prettyPrintSlide(sourceSlide)}")
              processSlide(slideIndex, template.getPath(), outFile.getPath, sourceSlide, pptNew, jsonData, None, slideIndexMap)
        }

        pptNew.write(new FileOutputStream(outFile))

        Right(())

    }

    def prettyPrintSlide(slide: XSLFSlide): String = {
        List(Option(slide.getSlideName()), Some(s"#${slide.getSlideNumber()}")).flatten.mkString("/")
    }

    def prettyPrintShape(shape: XSLFShape): String = shape.getShapeName()

    def prettyPrintAnchor(anchor: Rectangle2D): String = {
        // rounded to 3 decimal places output 
        s"${anchor.getWidth()}%.3fx${anchor.getHeight()}%.3f @ x=${anchor.getX()}%.3f y=${anchor.getY()}%.3f"        
    }    

    def moveIndexesDown(slideIndexMap: SlideIndexMap, start: Int) : Unit = {
        logger.debug(s"Slide index map b4  = ${slideIndexMap.mkString(",")}")
        for (i <- start to slideIndexMap.length - 1) {
            slideIndexMap(i) = slideIndexMap(i) + 1
        }
        logger.debug(s"Slide index map aft = ${slideIndexMap.mkString(",")}")
    }

    def moveIndexesUp(slideIndexMap: SlideIndexMap, start: Int) : Unit = {
        logger.debug(s"Slide index map b4  = ${slideIndexMap.mkString(",")}")
        for (i <- start to slideIndexMap.length - 1) {
            slideIndexMap(i) = slideIndexMap(i) - 1
        }
        logger.debug(s"Slide index map aft = ${slideIndexMap.mkString(",")}")
    }

    def processSlide(srcIdx: Int, srcPath: String, destPath: String, sourceSlide: XSLFSlide, pptNew: XMLSlideShow, rootJson: Json, contextJson: Option[Json], slideIndexMap: SlideIndexMap) : Unit = {
        withDepthLogging {

            // on the new slide, locate all control fields, any that are page based, we will process
            findControlJsonPath(sourceSlide.getShapes().asScala) match {

                // page based control fields
                case Some(PageControlData(shape, contextMatch, jsonPath)) => {
                    logger.debug(s"√ Slide [${prettyPrintSlide(sourceSlide)}] has a PageControl - jsonPath context -> ${jsonPath}")

                    // remove controlShape
                    sourceSlide.removeShape(shape)

                    val dataContextOnPage = JsonPathParser.parse(jsonPath).map { jp =>
                        CirceSolver.solve(jp, rootJson)
                    }

                    dataContextOnPage.map{ iter =>
                        for ((jsonNode, i) <- iter.zipWithIndex) {
                            logger.debug(s"√ node = ${jsonNode.toString()}")

                            // clone the slide
                            val newSlide: XSLFSlide = PPTXTools.createSlide(pptNew, sourceSlide, srcIdx + 1)
                            logger.debug(s"√ Cloned slide [${prettyPrintSlide(sourceSlide)}] to [${prettyPrintSlide(newSlide)}]")
                            moveIndexesDown(slideIndexMap, srcIdx + 1)
                            processSlide(srcIdx, srcPath, destPath, newSlide, pptNew, rootJson, Some(jsonNode), slideIndexMap)
                        }
                    }

                    // now we have templated it out, let's remove the original page
                    pptNew.removeSlide(sourceSlide.getSlideNumber() - 1)
                    moveIndexesUp(slideIndexMap, srcIdx + 1)
                }
                // unsupported control fields
                case Some(_) => logger.error(s"! Slide [${prettyPrintSlide(sourceSlide)}] has a control but it is not a PageControlData")

                // no page control fields found, process all shapes
                case None => processAllShapes(sourceSlide, rootJson, contextJson)
            }
        }

    }

    def processAllShapes(slide: XSLFSlide,  rootJson: Json, contextJson: Option[Json]) : Unit = {

        val shapeCount = slide.getShapes().size()
        for (shapeIndex <- (0 to shapeCount -1)) {

            val shape = slide.getShapes().get(shapeIndex)
            shape match {
                case textShape : TextHolder =>
                    if (hasTemplate(textShape)) {
                      logger.debug(s"√ ${f(shape)} is a templated shape")
                      changeText(textShape, rootJson, contextJson)
                    } else {
                      logger.debug(s"✖ '${textShape.getText()}' did not match")
                    }

                case group : XSLFGroupShape => { 
                    processGroupShape(slide, group, rootJson, contextJson)
                    // remote the group shape
                    // slide.removeShape(group)
                }
                case table : XSLFTable  if (hasControl(table.getRows().get(0).getCells().get(0)))    => {
                    logger.debug(s"√ we have a table - ${table.getRows().get(0).getCells().get(0).getText()}")
                    iterateTable(table, rootJson, contextJson)
                }
                case _ => logger.debug(s"✖ ${f(shape)} is not TextHolder")
            }

        }
    }

    def calcNewAnchor(currentAnchor: Rectangle2D, direction: Int, gap: Int, iteration: Int) : Rectangle2D = {
        val newD = gap + ((currentAnchor.getWidth() * Math.abs(Math.cos(Math.toRadians(direction)))) + (currentAnchor.getHeight() * Math.abs(Math.sin(Math.toRadians(direction)))) ) / 2

        // the gap is the distance from the edge of the shape, on the vector of the direction
        val newAnchorX = currentAnchor.getX() + (newD * Math.cos(Math.toRadians(direction))) 
        val newAnchorY = currentAnchor.getY() + (newD * Math.sin(Math.toRadians(direction)))

        new Rectangle2D.Double(newAnchorX, newAnchorY, currentAnchor.getWidth(), currentAnchor.getHeight())
    }
    def calcNewAnchorFixedDown(currentAnchor: Rectangle2D, direction: Int, gap: Int, iteration: Int) : Rectangle2D = {

        // the gap is the distance from the edge of the shape, on the vector of the direction
        val newAnchorX = currentAnchor.getX()
        val newAnchorY = currentAnchor.getY() + ((currentAnchor.getHeight() + gap) * iteration)

        new Rectangle2D.Double(newAnchorX, newAnchorY, currentAnchor.getWidth(), currentAnchor.getHeight())
    }


    def cloneShape(slide: XSLFSlide, sourceShape: XSLFShape, direction: Int, gap: Int, iteration: Int, translate: Boolean, rootJson: Json, contextJson: Option[Json], parentShape: Option[XSLFGroupShape]): XSLFShape = {

        logger.debug(s"⸮ Cloning shape ${prettyPrintShape(sourceShape)} in direction ${direction} gap ${gap}")
        // direction is degrees 0 to 360
        // gap is how far in that direction to move the new shape
        val currentAnchor = sourceShape.getAnchor()

        val newAnchor = if (translate) {
            calcNewAnchorFixedDown(currentAnchor, direction, gap, iteration)
        } else {
            currentAnchor.clone().asInstanceOf[Rectangle2D]
        }
        
        logger.debug(s"√ currentAnchor = ${prettyPrintAnchor(currentAnchor)} newAnchor = ${prettyPrintAnchor(newAnchor)}")

        sourceShape match { 
            case g: XSLFGroupShape => {
                val newGroupShape = slide.createGroup();
                newGroupShape.setInteriorAnchor(g.getInteriorAnchor())
                newGroupShape.setAnchor(newAnchor)
                for (shape <- g.getShapes().asScala) {
                    // 0,0,0, False no direction, no gap, always 0 iteration, no translation
                     val newShape = cloneShape(slide, shape, 0, 0, 0, false, rootJson, contextJson, Some(newGroupShape))
                }
                return newGroupShape
            }
            case t: XSLFTextBox => {
                val newShape = parentShape match { 
                    case Some(g) => g.createTextBox()
                    case None => slide.createTextBox()
                }
                newShape.setAnchor(newAnchor)
                newShape.setText(t.getText())
                return newShape
            }
            case a: XSLFAutoShape => {
                logger.debug(s"√ Cloning AutoShape ${a.getShapeName()}")
                val newShape = parentShape match { 
                    case Some(g) => g.createAutoShape()
                    case None => slide.createAutoShape()
                }
                newShape.setShapeType(a.getShapeType());
                newShape.setAnchor(newAnchor)
                newShape.setText(a.getText())
                newShape.setFillColor(a.getFillColor())
                newShape.setLineColor(a.getLineColor())
                newShape.setLineWidth(a.getLineWidth())
                newShape.setFlipHorizontal(a.getFlipHorizontal())
                newShape.setFlipVertical(a.getFlipVertical())
                newShape.setRotation(a.getRotation())
                changeText(newShape, rootJson, contextJson)
                return newShape
            }
            
            case _ => {
                logger.error(s"! Shape ${prettyPrintShape(sourceShape)} is currently not supported for cloning");
                return null
            }
        }

    }

    def processGroupShape(slide: XSLFSlide, groupShape: XSLFGroupShape, rootJson: Json, contextJson: Option[Json]) : Unit = {
        logger.debug(s"⸮ Inspecting XSLFGroupShape[${groupShape.getShapeName()}]...")
        // loop through all 
        findControlJsonPath(groupShape.getShapes().asScala) match {
            case None => logger.debug(s"✖ XSLFGroupShape[${groupShape.getShapeName()}] - no control, ignoring")

            // we found the control field inside this group object
            case Some(GroupShapeControlData(shape, controlMatch, jsonPath, direction, gap)) => {
                logger.debug(s"√ Found the XSLFGroupShape[${groupShape.getShapeName()}] with control fields")
                // remove the control shape, from the groupshape.
                groupShape.removeShape(shape)

                val (jsonNode, newJp) = nodeAndQuery(jsonPath, rootJson, contextJson)

                val dataContextOnGroupShape = JsonPathParser.parse(newJp).map { jp =>
                    CirceSolver.solve(jp, jsonNode)
                }

                dataContextOnGroupShape.map{ iter =>
                    for ((jn, i) <- iter.zipWithIndex) {
                        logger.debug(s"√ node = ${jn.toString()}")

                        // clone the shape
                        val theGap = if (i == 0) 0 else gap
                        val translate = true
                        cloneShape(slide, groupShape, direction, theGap, i, translate, rootJson, Some(jn), None)
                    }
                }

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

        // if there is no text, returm
        if (text == null || text.isEmpty()) {
            return
        }
        
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
                val text = textShape.getText()
                logger.debug(s"⸮ inspecting - Shape[${shape.getShapeName()}] `${textShape.getText()}`")

                // Direct regex match test
                val groupShapeMatch = matchGroupShapeControl.findFirstMatchIn(text)
                if (groupShapeMatch.isDefined) {
                    val jsonPath = groupShapeMatch.get.group("jsonpath")
                    val direction = groupShapeMatch.get.group("dir")
                    val gap = if (groupShapeMatch.get.group("gap") != null) groupShapeMatch.get.group("gap").toInt else 0
                    val controlText = groupShapeMatch.get.group(0)
                    logger.debug(s"√ Match - Shape[${shape.getShapeName()}] control=`${controlText}` jp=`${jsonPath}` dir=${direction} gap=${gap}")
                    return Some(GroupShapeControlData(shape, controlText, jsonPath, direction.toInt, gap))
                } else {
                    logger.debug(s"✖ Direct regex match not found for group shape control")
                }
                                

                val contextControlMatch = matchContextControl.findFirstMatchIn(text)
                if (contextControlMatch.isDefined) {
                    val jsonPath = contextControlMatch.get.group("jsonpath")
                    val controlText = contextControlMatch.get.group(0)
                    logger.debug(s"√ Match - Shape[${shape.getShapeName()}] control=`${controlText}` jp=`${jsonPath}`")
                    return Some(PageControlData(shape, controlText, jsonPath))
                } else {
                    logger.debug(s"✖ Direct regex match not found for context control")
                }

                logger.debug(s"✖ shape:${shape.getShapeName()} ${text} did not have a controlData")
        
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


    def withDepthLogging[T](block: => T): T = {
        val depth = Option(MDC.get("depth")).map(_.toInt).getOrElse(0)
        MDC.put("depth", (depth + 1).toString)
        try {
            block
        } finally {
            MDC.put("depth", depth.toString)
        }
    }

}

