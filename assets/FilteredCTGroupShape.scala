package org.openxmlformats.schemas.presentationml.x2006.main
import java.io.{File, InputStream, OutputStream, Reader, Writer}
import java.util

import javax.xml.namespace.QName
import javax.xml.stream.XMLStreamReader
import org.apache.xmlbeans.{QNameSet, SchemaType, XmlCursor, XmlDocumentProperties, XmlObject, XmlOptions}
import org.apache.xmlbeans.xml.stream.XMLInputStream
import org.openxmlformats.schemas.drawingml.x2006.main.CTGroupShapeProperties
import org.w3c.dom.Node
import org.xml.sax.ContentHandler
import org.xml.sax.ext.LexicalHandler

class FilteredCTGroupShape(ctgroupshape: CTGroupShape) extends CTGroupShape {

  override def getNvGrpSpPr: CTGroupShapeNonVisual = ctgroupshape.getNvGrpSpPr

  override def setNvGrpSpPr(ctGroupShapeNonVisual: CTGroupShapeNonVisual): Unit = ctgroupshape.setNvGrpSpPr(ctGroupShapeNonVisual)

  override def addNewNvGrpSpPr(): CTGroupShapeNonVisual = ctgroupshape.addNewNvGrpSpPr()

  override def getGrpSpPr: CTGroupShapeProperties = ctgroupshape.getGrpSpPr

  override def setGrpSpPr(ctGroupShapeProperties: CTGroupShapeProperties): Unit = ctgroupshape.setGrpSpPr(ctGroupShapeProperties)

  override def addNewGrpSpPr(): CTGroupShapeProperties = ctgroupshape.addNewGrpSpPr()

  override def getSpList: util.List[CTShape] = ctgroupshape.getSpList

  override def getSpArray: Array[CTShape] = ctgroupshape.getSpArray()

  override def getSpArray(i: Int): CTShape = ctgroupshape.getSpArray(i)

  override def sizeOfSpArray(): Int = ctgroupshape.sizeOfSpArray()

  override def setSpArray(ctShapes: Array[CTShape]): Unit = ctgroupshape.setSpArray(ctShapes)

  override def setSpArray(i: Int, ctShape: CTShape): Unit = ???

  override def insertNewSp(i: Int): CTShape = ???

  override def addNewSp(): CTShape = ???

  override def removeSp(i: Int): Unit = ???

  override def getGrpSpList: util.List[CTGroupShape] = ???

  override def getGrpSpArray: Array[CTGroupShape] = ???

  override def getGrpSpArray(i: Int): CTGroupShape = ???

  override def sizeOfGrpSpArray(): Int = ???

  override def setGrpSpArray(ctGroupShapes: Array[CTGroupShape]): Unit = ???

  override def setGrpSpArray(i: Int, ctGroupShape: CTGroupShape): Unit = ???

  override def insertNewGrpSp(i: Int): CTGroupShape = ???

  override def addNewGrpSp(): CTGroupShape = ???

  override def removeGrpSp(i: Int): Unit = ???

  override def getGraphicFrameList: util.List[CTGraphicalObjectFrame] = ???

  override def getGraphicFrameArray: Array[CTGraphicalObjectFrame] = ???

  override def getGraphicFrameArray(i: Int): CTGraphicalObjectFrame = ???

  override def sizeOfGraphicFrameArray(): Int = ???

  override def setGraphicFrameArray(ctGraphicalObjectFrames: Array[CTGraphicalObjectFrame]): Unit = ???

  override def setGraphicFrameArray(i: Int, ctGraphicalObjectFrame: CTGraphicalObjectFrame): Unit = ???

  override def insertNewGraphicFrame(i: Int): CTGraphicalObjectFrame = ???

  override def addNewGraphicFrame(): CTGraphicalObjectFrame = ???

  override def removeGraphicFrame(i: Int): Unit = ???

  override def getCxnSpList: util.List[CTConnector] = ???

  override def getCxnSpArray: Array[CTConnector] = ???

  override def getCxnSpArray(i: Int): CTConnector = ???

  override def sizeOfCxnSpArray(): Int = ???

  override def setCxnSpArray(ctConnectors: Array[CTConnector]): Unit = ???

  override def setCxnSpArray(i: Int, ctConnector: CTConnector): Unit = ???

  override def insertNewCxnSp(i: Int): CTConnector = ???

  override def addNewCxnSp(): CTConnector = ???

  override def removeCxnSp(i: Int): Unit = ???

  override def getPicList: util.List[CTPicture] = ???

  override def getPicArray: Array[CTPicture] = ???

  override def getPicArray(i: Int): CTPicture = ???

  override def sizeOfPicArray(): Int = ???

  override def setPicArray(ctPictures: Array[CTPicture]): Unit = ???

  override def setPicArray(i: Int, ctPicture: CTPicture): Unit = ???

  override def insertNewPic(i: Int): CTPicture = ???

  override def addNewPic(): CTPicture = ???

  override def removePic(i: Int): Unit = ???

  override def getExtLst: Nothing = ???

  override def isSetExtLst: Boolean = ???

  override def setExtLst(ctExtensionListModify: Nothing): Unit = ???

  override def addNewExtLst(): Nothing = ???

  override def unsetExtLst(): Unit = ???

  override def schemaType(): SchemaType = ???

  override def validate(): Boolean = ???

  override def validate(options: XmlOptions): Boolean = ???

  override def selectPath(path: String): Array[XmlObject] = ???

  override def selectPath(path: String, options: XmlOptions): Array[XmlObject] = ???

  override def execQuery(query: String): Array[XmlObject] = ???

  override def execQuery(query: String, options: XmlOptions): Array[XmlObject] = ???

  override def changeType(newType: SchemaType): XmlObject = ???

  override def substitute(newName: QName, newType: SchemaType): XmlObject = ???

  override def isNil: Boolean = ???

  override def setNil(): Unit = ???

  override def isImmutable: Boolean = ???

  override def set(srcObj: XmlObject): XmlObject = ???

  override def copy(): XmlObject = ???

  override def copy(options: XmlOptions): XmlObject = ???

  override def valueEquals(obj: XmlObject): Boolean = ???

  override def valueHashCode(): Int = ???

  override def compareTo(obj: Any): Int = ???

  override def compareValue(obj: XmlObject): Int = ???

  override def selectChildren(elementName: QName): Array[XmlObject] = ???

  override def selectChildren(elementUri: String, elementLocalName: String): Array[XmlObject] = ???

  override def selectChildren(elementNameSet: QNameSet): Array[XmlObject] = ???

  override def selectAttribute(attributeName: QName): XmlObject = ???

  override def selectAttribute(attributeUri: String, attributeLocalName: String): XmlObject = ???

  override def selectAttributes(attributeNameSet: QNameSet): Array[XmlObject] = ???

  override def monitor(): AnyRef = ???

  override def documentProperties(): XmlDocumentProperties = ???

  override def newCursor(): XmlCursor = ???

  override def newXMLInputStream(): XMLInputStream = ???

  override def newXMLStreamReader(): XMLStreamReader = ???

  override def xmlText(): String = ???

  override def newInputStream(): InputStream = ???

  override def newReader(): Reader = ???

  override def newDomNode(): Node = ???

  override def getDomNode: Node = ???

  override def save(ch: ContentHandler, lh: LexicalHandler): Unit = ???

  override def save(file: File): Unit = ???

  override def save(os: OutputStream): Unit = ???

  override def save(w: Writer): Unit = ???

  override def newXMLInputStream(options: XmlOptions): XMLInputStream = ???

  override def newXMLStreamReader(options: XmlOptions): XMLStreamReader = ???

  override def xmlText(options: XmlOptions): String = ???

  override def newInputStream(options: XmlOptions): InputStream = ???

  override def newReader(options: XmlOptions): Reader = ???

  override def newDomNode(options: XmlOptions): Node = ???

  override def save(ch: ContentHandler, lh: LexicalHandler, options: XmlOptions): Unit = ???

  override def save(file: File, options: XmlOptions): Unit = ???

  override def save(os: OutputStream, options: XmlOptions): Unit = ???

  override def save(w: Writer, options: XmlOptions): Unit = ???

  override def dump(): Unit = ???
}