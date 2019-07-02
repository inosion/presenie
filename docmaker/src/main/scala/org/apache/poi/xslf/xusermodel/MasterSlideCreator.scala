
package org.apache.poi.xslf.usermodel

object MasterSlideTooling {

    def cloneMasteSlide(slideMaster: XSLFSlideMaster): XSLFSlideMaster = {
        val part = slideMaster.getPackagePart()
        new XSLFSlideMaster(part)
    }
}
