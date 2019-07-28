package main.scala.inosion.pptx;

object MasterSlideTooling {

    def cloneMasteSlide(slideMaster: XSLFSlideMaster): XSLFSlideMaster = {
        val part = slideMaster.getPackagePart()
        new XSLFSlideMaster(part)
    }
}
