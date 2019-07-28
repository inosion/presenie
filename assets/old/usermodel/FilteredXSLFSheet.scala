package org.apache.poi.xslf.usermodel

import org.apache.poi.sl.usermodel._
import org.openxmlformats.schemas.presentationml.x2006.main._
import java.util.{Iterator => JIterator}
import java.util.{List => JList}
import java.awt.Graphics2D
import java.awt.geom.Rectangle2D
import scala.collection.JavaConverters._


/**
 * This class is a "hack" to be able to use POI, to import a shape
 * from a different sheet.
 * It relies on the fact that XSLFSheet.appendContent() https://github.com/apache/poi/blob/trunk/src/ooxml/java/org/apache/poi/xslf/usermodel/XSLFSheet.java#L465
 * uses the underlying XML to add slide objects.
 *
 * We filter out ALL other shapes except the one we want
 */

 abstract class FilteredXSLFSheet(sheet: XSLFSheet, shape: XSLFShape) extends XSLFSheet {

       /**
         * We only return the shape
         * @return
         */
       override def getShapes(): JList[XSLFShape] = {
         bufferAsJavaList(sheet.getShapes.asScala.filter( x => x.hashCode().equals(shape.hashCode())))
       }

  /*
       override protected def getSpTree()                           = sheet.getSpTree()
       {
         val root = sheet.getXmlObject();
         val sp = root.selectPath(
          s"declare namespace p='http://schemas.openxmlformats.org/presentationml/2006/main' .//${shape.getCNvPr.getId/p:spTree");
         if(sp.length == 0) {
          throw new IllegalStateException("CTGroupShape was not found");
         }
         _spTree = (CTGroupShape)sp[0];
        }
        return _spTree;
       }
       }

   */


       // public delegated methods
        override def getSlideShow()                                  = sheet.getSlideShow()
        override def getXmlObject()                                  = sheet.getXmlObject()
        override def createAutoShape()                               = sheet.createAutoShape()
        override def createFreeform()                                = sheet.createFreeform()
        override def createTextBox()                                 = sheet.createTextBox()
        override def createConnector()                               = sheet.createConnector()
        override def createGroup()                                   = sheet.createGroup
        override def createPicture(pictureData: PictureData)         = sheet.createPicture(pictureData)
        override def createTable()                                   = sheet.createTable()
        override def createTable(numRows: Int, numCols: Int)         = sheet.createTable(numRows, numCols)
        override def createOleShape(pictureData: PictureData)        = sheet.createOleShape(pictureData)
        override def iterator(): JIterator[XSLFShape]                = sheet.iterator()
        override def addShape(shape: XSLFShape)                      = sheet.addShape(shape)
        override def removeShape(xShape: XSLFShape)                  = sheet.removeShape(xShape)
        override def clear()                                         = sheet.clear()
        override def importContent(src: XSLFSheet)                   = sheet.importContent(src)
        override def appendContent(src: XSLFSheet)                   = sheet.appendContent(src)
        override def getTheme()                                      = sheet.getTheme()
        override def getPlaceholder( ph: Placeholder)                = sheet.getPlaceholder(ph)
        override def getPlaceholder(idx: Int)                        = sheet.getPlaceholder(idx)
        override def getPlaceholders()                               = sheet.getPlaceholders()
        override def getFollowMasterGraphics()                       = sheet.getFollowMasterGraphics()
        override def getBackground()                                 = sheet.getBackground()
        override def draw(graphics: Graphics2D)                      = sheet.draw(graphics)
        override def getPlaceholderDetails(placeholder: Placeholder) = sheet.getPlaceholderDetails(placeholder)
        override def addChart(chart: XSLFChart)                      = sheet.addChart(chart)
        override def addChart(chart: XSLFChart, rect2D: Rectangle2D) = sheet.addChart(chart, rect2D)

        override protected def allocateShapeId()                     = sheet.allocateShapeId()
        override protected def registerShapeId(shapeId: Int)         = sheet.registerShapeId(shapeId)
        override protected def deregisterShapeId(shapeId: Int)       = sheet.deregisterShapeId(shapeId)
        override protected def getRootElementName()                  = sheet.getRootElementName()
        override protected def getTextShapeByType(t: Placeholder)    = sheet.getTextShapeByType(t)
 }

 object FilteredXSLFSheet {
    protected def buildShapes(spTree: CTGroupShape, parent: XSLFShapeContainer) = XSLFSheet.buildShapes(spTree, parent)
 }