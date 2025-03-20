package inosion.presenie.pptx

import org.apache.poi.xslf.usermodel._ // XMLSlideShow
import org.apache.poi.sl.usermodel._
import java.io.File
import java.io.FileInputStream
import org.apache.poi.xslf.usermodel.MasterSlideTooling

import io.circe.Json

import io.circe._
import io.circe.parser._
import scala.io.Source

import scala.collection.JavaConverters._
import scala.util.Try
import scala.collection.mutable
import java.io.FileOutputStream

object PPTXTools {

    def listSlideLayouts(template: File) = {
        System.out.println(s":: Slide Layouts for ${template.getAbsolutePath()}" )

        val ppt: XMLSlideShow = new XMLSlideShow(new FileInputStream(template.getAbsolutePath()))
        for((master, i) <- ppt.getSlideMasters().asScala.zipWithIndex) {
          System.out.println(s"  :: Master [${i} ${master.getXmlObject().getCSld().getName()}]" )
          for(layout <- master.getSlideLayouts()) {
              System.out.println(s"    Name: ${layout.getName} - Type: ${layout.getType()}")
          }
        }
    }

    // Thanks ! from https://bz.apache.org/bugzilla/attachment.cgi?id=36089&action=edit
    def createSlide(ppt: XMLSlideShow, srcSlide: XSLFSlide, position: Int): XSLFSlide = {
        val slide: XSLFSlide = ppt.createSlide();
        slide.getSlideLayout().importContent(srcSlide.getSlideLayout());
        slide.importContent(srcSlide);
        ppt.setSlideOrder(slide, position);
        slide
    }

    def clonePptx(srcFile:  java.nio.file.Path, destFile:  java.nio.file.Path) : Unit = { 
        import java.nio.file.StandardCopyOption.REPLACE_EXISTING

        // implicit def toPath (filename: String) = java.nio.file.Paths.get(filename)

        java.nio.file.Files.copy(srcFile, destFile, REPLACE_EXISTING)        
    }

}


object JsonYamlTools {
  def parseJson(s: String): Either[ParsingFailure, Json] = parse(s)

  def readFileToJson(data: File): Either[ParsingFailure, Json] = {
    val fileContents = Source.fromFile(data).getLines().mkString
    parseJson(fileContents)
  }
}