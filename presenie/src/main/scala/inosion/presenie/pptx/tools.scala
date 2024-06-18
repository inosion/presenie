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
    def createSlide(ppt: XMLSlideShow, srcSlide:XSLFSlide, visitedLOs: mutable.Seq[XSLFSlideLayout] ): XSLFSlide = {
        val slide: XSLFSlide = ppt.createSlide();
        val srcLayout: XSLFSlideLayout = srcSlide.getSlideLayout()
        if (!visitedLOs.contains(srcLayout)) {
            visitedLOs :+ srcLayout
            val dstLayout: XSLFSlideLayout = slide.getSlideLayout();
            dstLayout.importContent(srcLayout);
        }
        slide.importContent(srcSlide);
        slide
    }

    def clonePptSlides(srcFile: File, destFile: File) : Unit = {
        val pptSrc: XMLSlideShow = new XMLSlideShow(new FileInputStream(srcFile.getAbsolutePath()))
        val pptDest: XMLSlideShow      = new XMLSlideShow()
        pptDest.setPageSize(pptSrc.getPageSize())

        val visitedLayouts: mutable.Seq[XSLFSlideLayout] = mutable.Seq()

        for (s <- pptSrc.getSlides().asScala) {
            //val slide = createSlide(pptDest, s, visitedLayouts)
            // val slide = pptDest.createSlide(s.getSlideLayout())
            val slide = pptDest.createSlide()
            copySlideContent(s, slide)
            slide.importContent(s)
            // not supported for XML -- slide.setFollowMasterBackground(s.getFollowMasterBackground())
            // not supported for XML --  slide.setFollowMasterColourScheme(s.getFollowMasterColourScheme())
            slide.getTheme()
        }


        /* not sure how to clone a master sheet */
        for (ms <- pptSrc.getSlideMasters().asScala) {
            val newms:XSLFSlideMaster = MasterSlideTooling.cloneMasterSlide(ms)
//            newms.importContent(ms)
            pptDest.getSlideMasters().add(newms)
        }

        for (p <- List(pptSrc, pptDest)) {
            System.out.println(s":: Available slide layouts [${p}]")
            //getting the list of all slide masters
            for(m: XSLFSlideMaster <- p.getSlideMasters().asScala) {

                //getting the list of the layouts in each slide master
                for(l: XSLFSlideLayout <- m.getSlideLayouts()) {

                    //getting the list of available slides
                    System.out.println("- " + l.getType());
                }
            }

        }

        pptDest.write(new FileOutputStream(destFile))

    }


    def copySlideContent(srcSlide: XSLFSlide, destSlide: XSLFSlide) = {

        val destSlideLayout: XSLFSlideLayout  = destSlide.getSlideLayout();
        val destSlideMaster: XSLFSlideMaster  = destSlide.getSlideMaster();

        val srcSlideLayout: XSLFSlideLayout = srcSlide.getSlideLayout();
        val srcSlideMaster: XSLFSlideMaster = srcSlide.getSlideMaster();

        try {
                // copy source layout to the new layout
                destSlideLayout.importContent(srcSlideLayout);
                // copy source master to the new master
                destSlideMaster.importContent(srcSlideMaster);
        } catch {
            case e: Exception => e.printStackTrace()
        }

    }

}


object JsonYamlTools {
  def parseJson(s: String): Either[ParsingFailure, Json] = parse(s)

  def readFileToJson(data: File): Either[ParsingFailure, Json] = {
    val fileContents = Source.fromFile(data).getLines().mkString
    parseJson(fileContents)
  }
}