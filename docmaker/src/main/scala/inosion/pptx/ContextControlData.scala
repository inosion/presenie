package inosion.pptx

import java.io.File

/*
trait ControlData {
    // the shape that holds the matching control string (for removal by caller)
    def shape: XSLFShape
    // the string that matched
    def controlText: String
    // the jsonPath
    def jsonPath: String
}
case class PageControlData(shape: XSLFTextShape, controlText: String, jsonPath: String) extends ControlData
case class GroupShapeControlData(shape: XSLFTextShape, controlText: String, jsonPath: String, direction: Double, gap: Double) extends ControlData
case class ImageControlData(shape: XSLFTextShape, controlText: String, jsonPath: String) extends ControlData

*/

case class FilesContext(src: File, dst: File, data: File) {
    def srcRelPath = src.getPath()
    def srcAbsPath = src.getAbsolutePath()
    def dstRelPath = dst.getPath()
    def dstAbsPath = dst.getAbsolutePath()
    def dataRelPath = data.getPath()
    def dataAbsPath = data.getAbsolutePath()
}

/*
case class SlidesContext(srcSlide: XSLFSlide, destPptx: XMLSlideShow, destSlides: Seq[XSLFSlide])
*/