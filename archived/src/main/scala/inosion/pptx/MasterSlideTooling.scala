
package org.apache.poi.xslf.usermodel

import org.apache.poi.xslf.usermodel._

object MasterSlideTooling {

    def cloneMasteSlide(slideMaster: XSLFSlideMaster): XSLFSlideMaster = {
        val part = slideMaster.getPackagePart()
        new XSLFSlideMaster(part)
    }
}
