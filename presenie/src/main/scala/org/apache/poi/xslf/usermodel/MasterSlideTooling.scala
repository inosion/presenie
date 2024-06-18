
package org.apache.poi.xslf.usermodel

object MasterSlideTooling {

    def cloneMasterSlide(slideMaster: XSLFSlideMaster): XSLFSlideMaster = {
        val part = slideMaster.getPackagePart()
        new XSLFSlideMaster(part)
    }
}
