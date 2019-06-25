package inosion.pptx

import org.apache.poi.xslf.usermodel._ // XMLSlideShow
import org.apache.poi.sl.usermodel._
import java.io.File
import java.io.FileInputStream

import com.fasterxml.jackson.databind.{ JsonNode, ObjectMapper }

import scala.collection.JavaConverters._
import scala.util.Try

object PPTXTools { 

    def listSlideLayouts(template: File) = { 
        val ppt: XMLSlideShow = new XMLSlideShow(new FileInputStream(template.getAbsolutePath()))
        for((master, i) <- ppt.getSlideMasters().asScala.zipWithIndex) {
          System.out.println(s":: Master [${i} ${master.getXmlObject().getCSld().getName()}]" )
          for(layout <- master.getSlideLayouts()) {
              System.out.println(s"Name: ${layout.getName} - Type: ${layout.getType()}")
          }
        }
    }
}

object JsonYamlTools {

    val mapper = new ObjectMapper
    def parseJson(s: String) = mapper.readValue(s, classOf[JsonNode])
    def readFileToJson(data: File): JsonNode = {

        val filecontents = scala.io.Source.fromFile(data).getLines.mkString
        parseJson(filecontents)
    }
}