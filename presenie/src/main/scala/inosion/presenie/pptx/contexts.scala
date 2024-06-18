
package inosion.presenie.pptx

import org.apache.poi.xslf.usermodel._ // XMLSlideShow & friends
import java.io.File
import io.circe.Json

/**
  * A Control object on the slide, that holds a subcontext for the control
  */
trait ControlData {
    // the shape that holds the matching control string (for removal by caller)
    def shape: XSLFShape
    // the string that matched
    def controlText: String
    // the jsonPath
    def jsonPath: String
}
// not sure if these need to be XSLFTextShape or XSLFShape
case class PageControlData(shape: XSLFShape, controlText: String, jsonPath: String) extends ControlData
case class GroupShapeControlData(shape: XSLFShape, controlText: String, jsonPath: String, direction: Double, gap: Double) extends ControlData
case class ImageControlData(shape: XSLFShape, controlText: String, jsonPath: String) extends ControlData

case class DataContext(rootJsonNode: Json, contextJsonNode: Option[Json] = None)


/**
  * Context for the files being processed
  *
  * @param src
  * @param dst
  * @param data
  */
case class FilesContext(src: File, dst: File, data: File) {
    def srcRelPath = src.getPath()
    def srcAbsPath = src.getAbsolutePath()
    def dstRelPath = dst.getPath()
    def dstAbsPath = dst.getAbsolutePath()
    def dataRelPath = data.getPath()
    def dataAbsPath = data.getAbsolutePath()
}

case class SlidesContext(srcSlide: XSLFSlide, destPptx: XMLSlideShow, destSlides: Seq[XSLFSlide])