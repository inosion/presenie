package inosion.pptx

import org.docx4j.openpackaging.packages.{OpcPackage, PresentationMLPackage}
import java.io.File
import java.io.FileInputStream

import com.fasterxml.jackson.databind.{JsonNode, ObjectMapper}

import scala.collection.JavaConverters._
import org.docx4j.dml.{CTRegularTextRun, CTTextField, CTTextLineBreak, CTTextParagraph}
import org.docx4j.mce.AlternateContent
import org.docx4j.openpackaging.parts.PresentationML.SlidePart
import org.jvnet.jaxb2_commons.ppp.Child
import org.pptx4j.pml.{CTEmbeddedFontListEntry, CTGraphicalObjectFrame, CTRel, CxnSp, GroupShape, Pic, Presentation, Shape}

object PPTXTools {

  /*
    def listSlideLayouts(template: File) = {
        val ppt: XMLSlideShow = new XMLSlideShow(new FileInputStream(template.getAbsolutePath()))
        for((master, i) <- ppt.getSlideMasters().asScala.zipWithIndex) {
          System.out.println(s":: Master [${i} ${master.getXmlObject().getCSld().getName()}]" )
          for(layout <- master.getSlideLayouts()) {
              System.out.println(s"Name: ${layout.getName} - Type: ${layout.getType()}")
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
   */
  def listFonts(pres: Presentation) {
    if (pres.getEmbeddedFontLst != null) {
      val embeddedFontList = pres.getEmbeddedFontLst
      for (fe: CTEmbeddedFontListEntry <- embeddedFontList.getEmbeddedFont.asScala) {
        System.out.println(s"Embedded Font: ${fe.getFont().getTypeface()}")
      }
    } else {
      System.out.println("no embedded font list")
    }
  }

  def clonePptSlides(srcFile: File, destFile: File) {
    val pptSrc = OpcPackage
      .load(new FileInputStream(srcFile.getAbsolutePath()))
      .asInstanceOf[PresentationMLPackage]
    val pptDest = PresentationMLPackage.createPackage()
    pptDest.getMainPresentationPart.getContents
      .setSldSz(pptSrc.getMainPresentationPart.getContents.getSldSz)

    // clone the fonts

    listFonts(pptSrc.getMainPresentationPart.getContents)

    for (s <- pptSrc.getMainPresentationPart.getSlideParts.asScala) {
      val slide = pptDest.getMainPresentationPart.addSlide(s)
    }

    System.out.println(
      s"Masters == ${pptDest.getMainPresentationPart.getContents.getSldMasterIdLst.getSldMasterId.asScala}"
    )

    /*
        /* not sure how to clone a master sheet */
        for (ms <- pptSrc.getSlideMasters().asScala) {
            val newms:XSLFSlideMaster = MasterSlideTooling.cloneMasteSlide(ms)
            try {
                newms.importContent(ms)
            } catch {
                case e: NullPointerException => {
                    System.out.println("may be missing a font")
                    e.printStackTrace()
                }
            }

            pptDest.getSlideMasters().add(newms)
        }
     */

    System.out.println(
      s"Masters == ${pptDest.getMainPresentationPart.getContents.getSldMasterIdLst.getSldMasterId.asScala}"
    )
    pptDest.save(destFile)

  }

  /*

    def copySlideContent(srcSlide: XSLFSlide, destSlide: XSLFSlide) = {

        val destSlideLayout: XSLFSlideLayout  = destSlide.getSlideLayout()
        val destSlideMaster: XSLFSlideMaster  = destSlide.getSlideMaster()

        val srcSlideLayout: XSLFSlideLayout = srcSlide.getSlideLayout()
        val srcSlideMaster: XSLFSlideMaster = srcSlide.getSlideMaster()

        try {
                // copy source layout to the new layout
                destSlideLayout.importContent(srcSlideLayout);
                // copy source master to the new master
                destSlideMaster.importContent(srcSlideMaster);
        } catch {
            case e: Exception => e.printStackTrace()
        }

    }


   */

  /**
    * Pretty print a Shape (for debugging)
    * @param thing
    * @return
    */
  def ppShape(thing: Child): String = {

    val s = thing match {
      case s: Shape if s.getTxBody != null => s.getNvSpPr.getCNvPr.getName + "/‘" + s.getTxBody().getP.asScala.foldLeft("")(_ + toText(_)) + "’"
        // // System.out.println(indent + tp.getClass.getName + "\n\n" + XmlUtils.marshaltoString(tp, true, true, org.pptx4j.jaxb.Context.jcPML, "http://schemas.openxmlformats.org/presentationml/2006/main", "txBody", classOf[CTTextParagraph]))
      case s: Shape => s.getNvSpPr.getCNvPr.getName + "/<empty>"
      case g: GroupShape =>
        s"GroupShape(${g.getSpOrGrpSpOrGraphicFrame.size()})"
      case o: CTGraphicalObjectFrame => "ObjectFrame"
      case x: CxnSp                  => "?CxnSp"
      case p: Pic                    => "?Pic"
      case r: CTRel                  => "?CTRel"
      case a: AlternateContent       => "AltContent"
      case _                         => "UNKNOWN Type"
    }

    s"${thing.getClass.getName}[${s}]"

  }

  def textRunToText()

  def toText(para: CTTextParagraph): String = {

    para.getEGTextRun.size() match {
      case 1 => para.getEGTextRun.get(0) match {
        case t: CTRegularTextRun => t.getT
        case l: CTTextLineBreak => "\n"
        case f: CTTextField => s"« ${f.getT} »"
      }
      case 0 => ""
      case _ => toText(para.getEGTextRun.asScala.head) + toText(para.getEGTextRun.asScala.tail)
    }
  }

  def doTraversal(slide: SlidePart): Unit = {

    import org.docx4j.TraversalUtil
    import org.docx4j.XmlUtils
    import org.docx4j.dml.CTTextBody
    import org.docx4j.dml.CTTextParagraph

    new TraversalUtil(
      slide.getJaxbElement.getCSld.getSpTree.getSpOrGrpSpOrGraphicFrame,
      new org.docx4j.TraversalUtil.Callback() {
        var indent = "" //                      @Override
        def apply(o: Any): java.util.List[AnyRef] = {
          val text = ""
          try
          // System.out.println(indent + o.getClass.getName + "\n\n" + XmlUtils.marshaltoString(o, true, org.pptx4j.jaxb.Context.jcPML))
            System.out.println(indent + ppShape(o.asInstanceOf[Child]))
          catch {
            case me: RuntimeException =>
              System.out.println(indent + "[error] " + o.getClass.getName + (me.getMessage))
          }

          null
        }

        def shouldTraverse(o: Any) = true // Depth first
        def walkJAXBElements(parent: Any): Unit = {
          indent += "    "
          val children = getChildren(parent).asScala
          if (children != null) {
            for (o <- children) { // if its wrapped in javax.xml.bind.JAXBElement, get its
              // value
              this.apply(XmlUtils.unwrap(o))
              if (this.shouldTraverse(o)) walkJAXBElements(o)
            }
          }
          indent = indent.substring(0, indent.length - 4)
        }

        def getChildren(o: Any): java.util.List[AnyRef] =
          TraversalUtil.getChildrenImpl(o)
      }
    )

  }


  def slideInfo(filename: File) {
    val presentationMLPackage =
      OpcPackage.load(filename).asInstanceOf[PresentationMLPackage]

    for ((slidePart, i) <- presentationMLPackage.getMainPresentationPart.getSlideParts.asScala.zipWithIndex) {
      val slideLayoutPart = slidePart.getSlideLayoutPart
      System.out.println(slidePart.getPartName.getName)
      val layoutName = slideLayoutPart.getJaxbElement.getCSld.getName
      System.out.println(
        ".. uses layout: " + slideLayoutPart.getPartName.getName + " with layout name cSld/@name='" + layoutName + "'"
      )
      System.out.println(
        "   .. which uses master: " + slideLayoutPart.getSlideMasterPart.getPartName.getName
      )
      doTraversal(slidePart)

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
