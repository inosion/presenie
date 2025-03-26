package inosion.presenie.pptx

import org.slf4j.MDC

import org.apache.poi.xslf.usermodel._ // XMLSlideShow
import org.apache.poi.sl.usermodel._
import java.io.File
import java.io._
import java.awt.Color

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

    val JsonPathReg                   = raw"[\$$@]|[\$$@]\.[0-9A-Za-z_:\-\?\.\[\]\*@=']+"
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

    def prettyPrintShape(shape: Shape[XSLFShape,XSLFTextParagraph]): String = shape.getShapeName()

    def prettyPrintAnchor(anchor: Rectangle2D): String = {
        // rounded to 3 decimal places output 
        s"${anchor.getWidth()%.3f}x${anchor.getHeight()%.3f} @ x=${anchor.getX()%.3f} y=${anchor.getY()%.3f}"        
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
                            logger.debug(s" Applying data, under PageControl idx=${i} - ${jsonNode.toString()}")

                            // clone the slide
                            val newSlide: XSLFSlide = PPTXTools.createSlide(pptNew, sourceSlide, srcIdx + 1)
                            logger.debug(s"+ Cloned slide [${prettyPrintSlide(sourceSlide)}] to [${prettyPrintSlide(newSlide)}]")
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
            logger.debug(s"⸮ Inspecting Shape[${shape.getShapeName()}]... ")
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
                    logger.debug(s"√ Table with ControlContext found - ${table.getRows().get(0).getCells().get(0).getText()}")
                    changeTextInTable(table, rootJson, contextJson, tableIteration = true)
                }
                case table : XSLFTable  => {
                    logger.debug(s"√ we have a table - but no control ... rowCount=${table.getRows().size()}")
                    changeTextInTable(table, rootJson, contextJson, tableIteration = false)
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
                logger.debug(s"√ Cloning AutoShape ${a.getShapeName()} {${a.getText()}}")
                val newShape = parentShape match { 
                    case Some(g) => g.createAutoShape()
                    case None => slide.createAutoShape()
                }
                // don't set the type if it is null

                if (a.getShapeType() != null) {
                    newShape.setShapeType(a.getShapeType())
                }
                newShape.setAnchor(newAnchor)
                newShape.setText(a.getText())
                newShape.setFillColor(a.getFillColor())
                newShape.setLineColor(a.getLineColor())
                newShape.setLineWidth(a.getLineWidth())
                newShape.setFlipHorizontal(a.getFlipHorizontal())
                newShape.setFlipVertical(a.getFlipVertical())
                newShape.setRotation(a.getRotation())
                changeText(newShape, rootJson, contextJson, Some(a))
                return newShape
            }
            case p: XSLFPictureShape => {
                logger.debug(s"√ Cloning PictureShape ${p.getShapeName()}")
                val newShape = parentShape match { 
                    case Some(g) => g.createPicture(p.getPictureData())
                    case None => slide.createPicture(p.getPictureData())
                }
                newShape.setAnchor(newAnchor)
                return newShape
            }

            case t: XSLFTable => {
                logger.debug(s"√ Cloning Table ${t.getShapeName()}")
                val newShape = parentShape match { 
                    case Some(g) => g.createTable()
                    case None => slide.createTable()
                }
                newShape.setAnchor(newAnchor)
                for (row <- t.getRows().asScala) {
                    val newRow = newShape.addRow()
                    for (cell <- row.getCells().asScala) {
                        val newCell = newRow.addCell()
                        newCell.setText(cell.getText())
                        newCell.setStrokeStyle(cell.getStrokeStyle())
                        newCell.setFillColor(cell.getFillColor())
                        newCell.setLineColor(cell.getLineColor())
                        newCell.setLineWidth(cell.getLineWidth())
                        newCell.setFlipHorizontal(cell.getFlipHorizontal())
                        newCell.setFlipVertical(cell.getFlipVertical())
                        newCell.setRotation(cell.getRotation())
                        newCell.setAnchor(cell.getAnchor())
                        // TODO - we should determine what type of table iteration to do here
                        changeTextInTable(t, rootJson, contextJson, tableIteration = false)
                        
                    }
                }
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

                logger.debug(s" -> Running the loop on jsonNode = ${jsonNode.toString()}, newJp = ${newJp}")
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

    def changeTextInTable(table: XSLFTable, rootJson: Json, contextJson: Option[Json], tableIteration: Boolean) : Unit = {
        if (!tableIteration) {
            logger.debug(s"No Control String on the table // just change text in table")
            // for every row and every cell, we will change the text
            for (row <- table.getRows().asScala) {
                for (cell <- row.getCells().asScala) {
                    changeText(cell, rootJson, contextJson)
                }
            }
        } else { 
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
                        // cell.setStrokeStyle(cell.getStrokeStyle())
                    }
                }
            }

            // now remove the template row
            table.removeRow(1)
            // remove the context string
            table.getRows().get(0).getCells().get(0).setText(firstCellText.replace(controlString, ""))
        }
    }

    def rawChangeTextPreserveStyling(textShape: TextHolder, newText: String) : Unit = { 
        val textShapeParagraphs = textShape.getTextParagraphs()
        val textShapeParagraphTextRuns = textShapeParagraphs.get(0).getTextRuns()
        logger.debug(s"Reapplying style text for : Shape[${prettyPrintShape(textShape)}] `${textShape.getText()}` to `${newText}`")
        for ((r, i) <- textShapeParagraphTextRuns.asScala.zipWithIndex) {
            logger.debug(s"  TextRun[${i}] -> `${r.getRawText()}`")
            r.setText(newText);
        }
    }

    case class TextAndStyle(isBold: Boolean, colorRgb: Int, fontSize: Double, fontFamily: String = "Arial", isItalic: Boolean = false, isUnderlined: Boolean = false, isStrikethrough: Boolean = false, isSubscript: Boolean = false, isSuperscript: Boolean = false, characterSpacing: Double = 0.0)

    def rawChangeTextPreserveStyling2(textShape: TextHolder, newText: String, templateText: String, originalShape: Option[TextHolder] = None) : Unit = {

        logger.debug(s"Reapplying style text for : Shape[${prettyPrintShape(textShape)}] `${textShape.getText()}` to `${newText}`")
        // get the original style
        val textRun = originalShape match { 
            case Some(os) => os.getTextParagraphs().get(0).getTextRuns().get(0)
            case None => textShape.getTextParagraphs().get(0).getTextRuns().get(0)
        }
            

        // get Color:
        val solidPaint = textRun.getFontColor.asInstanceOf[PaintStyle.SolidPaint]
        val color = solidPaint.getSolidColor.getColor.getRGB

        // save text & styles:
        val textAndStyle = TextAndStyle(textRun.isBold, color, textRun.getFontSize, textRun.getFontFamily, textRun.isItalic, textRun.isUnderlined, textRun.isStrikethrough, textRun.isSubscript, textRun.isSuperscript, textRun.getCharacterSpacing)

        val currentText = textShape.getText()
        val aa_newText = currentText.replace(templateText, newText)

        textShape.setText(aa_newText)

        // reapply all the style information
        textShape.getTextParagraphs().get(0).getTextRuns().get(0).setFontColor(new Color(textAndStyle.colorRgb))
        textShape.getTextParagraphs().get(0).getTextRuns().get(0).setBold(textAndStyle.isBold)
        textShape.getTextParagraphs().get(0).getTextRuns().get(0).setFontSize(textAndStyle.fontSize)

        textShape.getTextParagraphs().get(0).getTextRuns().get(0).setFontFamily(textAndStyle.fontFamily)
        textShape.getTextParagraphs().get(0).getTextRuns().get(0).setItalic(textAndStyle.isItalic)
        textShape.getTextParagraphs().get(0).getTextRuns().get(0).setUnderlined(textAndStyle.isUnderlined)
        textShape.getTextParagraphs().get(0).getTextRuns().get(0).setStrikethrough(textAndStyle.isStrikethrough)
        textShape.getTextParagraphs().get(0).getTextRuns().get(0).setSubscript(textAndStyle.isSubscript)
        textShape.getTextParagraphs().get(0).getTextRuns().get(0).setSuperscript(textAndStyle.isSuperscript)
        textShape.getTextParagraphs().get(0).getTextRuns().get(0).setCharacterSpacing(textAndStyle.characterSpacing)
        
    }

    def debugTextHolder(textHolder: TextHolder) : Unit = { 
        logger.debug(s"TextHolder[${textHolder.getShapeName()}] text = ${textHolder.getText()}")
        for (p <- textHolder.getTextParagraphs().asScala) {
            logger.debug(s"  Paragraph[${p.getText()}]")
            for (r <- p.getTextRuns().asScala) {
                logger.debug(s"    Run[${r.getRawText()}]")
            }
        }
    }

    def changeText(textShape: TextHolder, rootJsonNode: Json, contextJsonNode: Option[Json], originalShape: Option[TextHolder] = None) : Unit = {
        // the originalShape has style we need to clone. 
        val text = textShape.getText()

        // if there is no text, returm
        if (text == null || text.isEmpty()) {
            logger.debug(s"Shape[${prettyPrintShape(textShape)}] is empty.. no changes")
            return
        } 

        try {
             
            val jsonQuery = matchRegexpTemplate.r.findFirstMatchIn(text) match { 
                case Some(m) => m.group("jsonpath")
                case None => {
                    logger.debug(s"(None on Regex) Ignoring text --> Shape[${prettyPrintShape(textShape)}] `${textShape.getText()}`")
                    if (originalShape.isDefined) {
                        rawChangeTextPreserveStyling2(textShape, text, text, originalShape)
                    }
                    return
                }
            }
                
        } catch {
            case e: scala.MatchError => {
                logger.debug(s"(MatchError) no jsonPath in text --> Shape[${prettyPrintShape(textShape)}] `${textShape.getText()}`")
                return
            }
        }

        try {
            val matchRegexpTemplate.r(templateText, jsonQuery) = text

            logger.debug(s"(Matched) found RegEx [$templateText] jsonPath=[$jsonQuery]")

            val (jsonNode, jsonPath) = nodeAndQuery(jsonQuery, rootJsonNode, contextJsonNode)

            val matchedText = JsonPathParser.parse(jsonPath).map { jp =>
                CirceSolver.solve(jp, jsonNode)
            }

            // join the matcheed entries via a comma

            val newText = matchedText match {
                case Right(iter) =>
                    iter.map { jsonNode =>
                        jsonNode.asString.getOrElse(jsonNode.toString())
                    }.mkString(", ")
                case Left(error) =>
                    logger.error(s"Error parsing JSONPath: $error")
                    ""
            }
            rawChangeTextPreserveStyling2(textShape, newText, templateText, originalShape)
                        
        } catch {
            case e: scala.MatchError => {
                logger.debug(s"(MatchError) Ignoring text --> Shape[${prettyPrintShape(textShape)}] `${textShape.getText()}`")
                return
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
            case '$' => { 
                logger.debug(s"! jsonPath - using root object")
                (rootJson,        jsonQuery)
            }
            case '@' => contextJson match {
                            case Some(jn) => { 
                                logger.debug(s"! haveContext - query is @ relative")
                                (jn, "$" + jsonQuery.stripPrefix("@"))
                            }
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
                    logger.debug(s"√ GroupControl Match - Shape[${shape.getShapeName()}] control=`${controlText}` jp=`${jsonPath}` dir=${direction} gap=${gap}")
                    return Some(GroupShapeControlData(shape, controlText, jsonPath, direction.toInt, gap))
                } else {
                    logger.debug(s"✖ Direct regex match not found for group shape control")
                }
                                

                val contextControlMatch = matchContextControl.findFirstMatchIn(text)
                if (contextControlMatch.isDefined) {
                    val jsonPath = contextControlMatch.get.group("jsonpath")
                    val controlText = contextControlMatch.get.group(0)
                    logger.debug(s"√ PageControl Match - Shape[${shape.getShapeName()}] control=`${controlText}` jp=`${jsonPath}`")
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

