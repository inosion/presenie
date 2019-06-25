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

//import org.openxmlformats.schemas.presentationml.x2006.main.

object PPTXMerger { 
    type TextHolder = TextShape[XSLFShape, XSLFTextParagraph]

    import scala.reflect.ClassTag

    def f[T](v: T)(implicit ev: ClassTag[T]) = ev.toString

    val JsonPathReg                   = raw"[\$$@]|[\$$@]\.[0-9A-Za-z\?\.\[\]\*]+"
    val matchRegexpTemplate           = raw".*(\{\{\s*(" + JsonPathReg + raw")\s*\}\}).*"
    val matchContextControl             = (raw"(\{\s*context\s*=\s*(" + JsonPathReg + raw")\s*\})").r
    val matchGroupShapeControl        = (raw"(\{\s*context\s*=\s*(" + JsonPathReg + raw")\s*\})").r    

    def render(config: File, data: File, template: File, outFile: File): Either[Error, Unit] = {

        println(matchRegexpTemplate)

        val pptTemplate: XMLSlideShow = new XMLSlideShow(new FileInputStream(template.getAbsolutePath()))
        val pptNew: XMLSlideShow      = new XMLSlideShow(new FileInputStream(template.getAbsolutePath()))
        
        // fastest way to clone the sheet I can see is copy it, then remove all slides
        // to retain the master slide layouts.
        for (idx <- 0 until pptNew.getSlides.size()) { 
            pptNew.removeSlide(0)
            System.out.println(s"removed $idx")
        }

        val jsonData = JsonYamlTools.readFileToJson(data)

        for (srcSlide <- pptTemplate.getSlides().asScala) {
            processSlide(template.getPath(), outFile.getPath, srcSlide, pptNew, jsonData, None)
        }

        pptNew.write(new FileOutputStream(outFile))

        Right(())

    }

    def processSlide(srcPath: String, destPath: String, srcSlide: XSLFSlide, pptNew: XMLSlideShow, rootJsonNode: JsonNode, contextJsonNode: Option[JsonNode]) {


        // We will copy this slide from source to new ppt, then fill it's data
        pptNew.createSlide().importContent(srcSlide)
        val newSlide = pptNew.getSlides().get(pptNew.getSlides().size() - 1)
        System.err.println(s":: Cloning from src [${srcPath}(${srcSlide.getSlideNumber()})] to [${destPath}(${newSlide.getSlideNumber()})] ")


        findSlideIterator(newSlide) match {
            case Some(jsonPath) => {
                System.out.println(s"! found a jsonPath context ${jsonPath}")
                JsonPath.query(jsonPath, rootJsonNode).map{ i => 
                    for (jsonNode <- i) {
                        System.out.println(s"! node = ${jsonNode.toString()}")

                        processSlide(srcPath, destPath, newSlide, pptNew, rootJsonNode, Some(jsonNode))
                    }
                }
                // now we have templated it out, let's remove it
                pptNew.removeSlide(newSlide.getSlideNumber() - 1)
            }
            case None => processAllShapes(newSlide, rootJsonNode, contextJsonNode)
        }

    }

    def processAllShapes(slide: XSLFSlide,  rootJsonNode: JsonNode, contextJsonNode: Option[JsonNode]) {

        for (shape <- slide.getShapes().asScala) {

            shape match { 
                case textShape : TextHolder =>
                    if (hasTemplate(textShape)) {
                      System.out.println(s":: ${f(shape)} is a templated shape")
                      changeText(textShape, rootJsonNode, contextJsonNode)
                    } else { 
                      System.out.println(s"!! '${textShape.getText()}' did not match")
                    }
                
                case group : XSLFGroupShape => System.out.println(s":: Group Shape")
                case table : XSLFTable  if (hasControl(table.getRows().get(0).getCells().get(0)))    => {
                    System.out.println(s":: we have a table - ${table.getRows().get(0).getCells().get(0).getText()}")
                    iterateTable(table, rootJsonNode, contextJsonNode)
                }
                case _ => System.out.println(s"\n:: ${f(shape)} is not TextHolder")
            }
            
        }
    }

    def iterateTable(table: XSLFTable, rootJsonNode: JsonNode, contextJsonNode: Option[JsonNode]) {
        val firstCellText = table.getRows().get(0).getCells().get(0).getText()
        // val (a, b, tableContextJsonPath) = for (m <- matchContextControl.findFirstMatchIn(firstCellText)) yield m.group
        val mm = for (m <- matchContextControl.findFirstMatchIn(firstCellText)) yield m
        val tableContextJsonPath = mm.get.group(2)
        val controlString = mm.get.group(0)

        val (jsonNode, _tableContextJsonPath) = nodeAndQuery(tableContextJsonPath, rootJsonNode, contextJsonNode)

        JsonPath.query(_tableContextJsonPath, jsonNode).map{ iter => 
            for ((jsonNode, i) <- iter.zipWithIndex) {
                System.out.println(s"! Table node = ${jsonNode.toString()}")
                RowCloner.cloneRow(table, 1) // including cells
                for ((cell, ci) <- table.getRows().get(table.getRows().size() - 1).getCells().asScala.zipWithIndex) {
                    cell.setText(replaceText(rootJsonNode, jsonNode, cell.getText()))
                    cell.setStrokeStyle(cell.getStrokeStyle())
                }
            }
        }

        // now remove the template row
        table.removeRow(1)
        // remove the context string
        table.getRows().get(0).getCells().get(0).setText(firstCellText.replace(controlString, ""))

    
    }
    
    def changeText(textShape: TextHolder, rootJsonNode: JsonNode, contextJsonNode: Option[JsonNode]) {
        val text = textShape.getText()
        val matchRegexpTemplate.r(replacingText, jsonQuery) = text

        System.out.println(s"found = [${replacingText}] jsonpath = [${jsonQuery}]")

        val (theJsonNode, theJsonPath) = nodeAndQuery(jsonQuery, rootJsonNode, contextJsonNode)

        val dataText = JsonPath.query(theJsonPath, theJsonNode) match { 
            case Left(error) => throw new java.lang.Error(error.reason)
            case Right(i)    => try { 
                Some(i.next().asText())
            } catch { 
                case e: java.util.NoSuchElementException => System.err.println(s"The JSONPath expression ==> ${theJsonPath} <== did not resolve to any data. Ignoring"); None
                case e: Exception => throw e
            }
        }

        dataText.map { txt =>
            val newText = text.replace(replacingText, txt)
            System.out.println(s"dataText = [${txt}] newText = [${newText}]")
            textShape.setText(newText)
        }
    }

    def replaceText(rootJsonNode: JsonNode, contextJsonNode: JsonNode, text: String) :String = {
        // fullTemplateText is {{someJsonPath}}
        // jsonQuery is what was found
        val matchRegexpTemplate.r(fullTemplateText, jsonQuery) = text

        val (jsonNode, jsonPath) = nodeAndQuery(jsonQuery, rootJsonNode, Some(contextJsonNode))

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
    def nodeAndQuery(jsonQuery: String, rootJsonNode: JsonNode, contextJsonNode: Option[JsonNode]): (JsonNode, String) = {
        jsonQuery(0) match {
            case '$' => (rootJsonNode,        jsonQuery)
            case '@' => contextJsonNode match {
                            case Some(jn) => (jn, "$" + jsonQuery.stripPrefix("@"))
                            case None     => {
                                System.err.println("Not Context object. Using root instead")
                                (rootJsonNode, "$" + jsonQuery.stripPrefix("@"))
                            }  
                        }
        }
    }

    // has side affect of removing the control 
    def findSlideIterator(currentSlide: XSLFSlide): Option[String] = {

        for (shape <- currentSlide.getShapes().asScala) {
            if (shape.isInstanceOf[TextHolder]) {
                val textShape = shape.asInstanceOf[TextHolder]
                if (hasControl(textShape)) {
                    val jsonPath = for (m <- matchContextControl.findFirstMatchIn(textShape.getText())) yield m.group(2)
                    currentSlide.removeShape(shape)
                    return jsonPath
                }
            } 
        }
        None

    }    

    def hasTemplate(shape: TextHolder): Boolean = 
      shape.getText().matches(matchRegexpTemplate)

    def hasControl(shape: TextHolder): Boolean = matchContextControl.findFirstIn(shape.getText()).isDefined

}

