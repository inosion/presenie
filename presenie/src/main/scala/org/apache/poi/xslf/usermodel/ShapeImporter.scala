package org.apache.poi.xslf.usermodel

import com.typesafe.scalalogging._
import inosion.presenie.pptx.SlidesContext
import org.openxmlformats.schemas.presentationml.x2006.main.CTShape
import org.openxmlformats.schemas.presentationml.x2006.main.CTGroupShape
import org.openxmlformats.schemas.presentationml.x2006.main.CTConnector
import org.openxmlformats.schemas.presentationml.x2006.main.CTPicture
import org.openxmlformats.schemas.presentationml.x2006.main.CTGraphicalObjectFrame

import scala.collection.JavaConverters._

// credits - https://gist.github.com/jorgeortiz85/908035

class PrivateFieldSetter(x: AnyRef, fieldName: String) {
  def apply(arg: Any): Any = {
    def _parents: Stream[Class[_]] =
      Stream(x.getClass) #::: _parents.map(_.getSuperclass)
    val parents = _parents.takeWhile(_ != null).toList
    val fields = parents.flatMap(_.getDeclaredFields)
    val field = fields
      .find(_.getName == fieldName)
      .getOrElse(
        throw new IllegalArgumentException(
          "Field " + fieldName + " not found"
        )
      )
    field.setAccessible(true)
    field.set(x, arg)
  }
}

class PrivateFieldExposer(x: AnyRef) {
  def apply(field: scala.Symbol): PrivateFieldSetter =
    new PrivateFieldSetter(x, field.name)
}

class PrivateMethodCaller(x: AnyRef, methodName: String) {
  def apply(_args: Any*): Any = {
    val args = _args.map(_.asInstanceOf[AnyRef])
    def _parents: Stream[Class[_]] =
      Stream(x.getClass) #::: _parents.map(_.getSuperclass)
    val parents = _parents.takeWhile(_ != null).toList
    val methods = parents.flatMap(_.getDeclaredMethods)
    val method = methods
      .find(_.getName == methodName)
      .getOrElse(
        throw new IllegalArgumentException(
          "Method " + methodName + " not found"
        )
      )
    method.setAccessible(true)
    method.invoke(x, args: _*)
  }
}

class PrivateMethodExposer(x: AnyRef) {
  def apply(method: scala.Symbol): PrivateMethodCaller =
    new PrivateMethodCaller(x, method.name)
}


object ShapeImporter extends StrictLogging {

  private def m(x: AnyRef): PrivateMethodExposer = new PrivateMethodExposer(x)
  private def f(x: AnyRef): PrivateFieldExposer = new PrivateFieldExposer(x)

  def addShapex(shape: XSLFShape, srcSheet: XSLFSheet, destSheet: XSLFSheet, sc: SlidesContext):XSLFShape = {

    val newShape = shape match {
      case _ : XSLFGroupShape => destSheet.createGroup()
      case _ : XSLFConnectorShape => destSheet.createConnector()
      case _ : XSLFAutoShape => { val x = destSheet.createAutoShape(); destSheet.getSpTree().addNewSp().set(shape.getXmlObject.copy()); x }
      case _ : XSLFTable => destSheet.createTable()
      case _ : XSLFTextBox => { val k = destSheet.createTextBox() ; destSheet.getSpTree().addNewSp().set(shape.getXmlObject.copy()) ; k}
      case _ : XSLFFreeformShape => { val k = destSheet.createFreeform() ; destSheet.getSpTree().addNewSp().set(shape.getXmlObject.copy()) ; k}
      case o : XSLFObjectShape => destSheet.createOleShape(o.getPictureData())
      case p : XSLFPictureShape => destSheet.createPicture(p.getPictureData())
      case _ => throw new UnsupportedOperationException(s"[${shape.getClass.getCanonicalName}] shape is not supported")
    }

    m(newShape)('copy)(shape)

    newShape

  }

  def addShape2(shape: XSLFShape, srcSheet: XSLFSheet, destSheet: XSLFSheet, sc: SlidesContext):XSLFShape = {

    val offset = destSheet.getShapes.size()

    logger.debug(s"before = ${destSheet.getShapes.size()}")

    destSheet.importContent(srcSheet)

    logger.debug(s"after = ${destSheet.getShapes.size()}")

    //val shapesToRemove = for (x <- offset until destSheet.getShapes.size; sh <- destSheet.getShapes().get(x); if (! sh.getShapeName().equals(shape.getShapeName)) ) yield shape

    val shape2Keep = destSheet.getShapes.asScala.filter(_.getShapeName.equals(shape.getShapeName))
    val shapes2Ditch = destSheet.getShapes.asScala.filter(!_.getShapeName.equals(shape.getShapeName))

    for (z <- shapes2Ditch) {
      destSheet.removeShape(z)
    }
    shape2Keep.last
  }

  import scala.reflect.ClassTag
  def matches[A, B: ClassTag](a: A, b: B) = a match {
    case _: B => true
    case _ => false
  }

  def addShape(shape: XSLFShape, srcSheet: XSLFSheet, destSheet: XSLFSheet, sc: SlidesContext):XSLFShape = {

    val dummySlide = sc.destPptx.createSlide()
    dummySlide.importContent(srcSheet)
    val removalShapes = dummySlide.getShapes.asScala.filter(x => ! (matches(x,shape) && x.getShapeName.equals(shape.getShapeName)) ).toSeq

    logger.debug(s"Size ${dummySlide.getShapes.size()}")

    for (x <- removalShapes) {
      logger.debug(s"removing ${x.getShapeName}")
      dummySlide.removeShape(x)
    }
    logger.debug(s"Size ${dummySlide.getShapes.size()}")

    destSheet.appendContent(dummySlide)

    val last = destSheet.getShapes.size() - 1
    destSheet.getShapes.get(last)
  }



  def addShapeX(shape: XSLFShape, srcSheet: XSLFSheet, destSheet: XSLFSheet, sc: SlidesContext):XSLFShape = {

    // Step 1 = logic from XSLFSheet.appendContent
    val spTree: CTGroupShape = destSheet.getSpTree()
    val currentShapeHashCodes = destSheet.getShapes().asScala.map(_.hashCode())
    val offset = destSheet.getShapes().size()

    logger.debug(s"∞∞∞∞∞ we want to clone [${shape.getShapeId()}] = destSheet has ${currentShapeHashCodes.size}")

    val xShape = shape.getXmlObject()

    xShape match {
      case _: CTShape                => spTree.addNewSp().set(xShape.copy())
      case _: CTGroupShape           => spTree.addNewGrpSp().set(xShape.copy())
      case _: CTConnector            => spTree.addNewCxnSp().set(xShape.copy())
      case _: CTPicture              => spTree.addNewPic().set(xShape.copy())
      case _: CTGraphicalObjectFrame => spTree.addNewGraphicFrame().set(xShape.copy())
    }

    // Step 2 = XSLFSheet.appendContent
    //          --> calls wipeAndReinitialize()
    //          --> calls initDrawingAndShapes()()

    // these three lines will rebuild the full shape tree
    f(destSheet)('_shapes)(null)
    f(destSheet)('_drawing)(null)
    m(destSheet)('initDrawingAndShapes)()

    // The new shape is the indexed offset from the last (meaning that the array of shapes remains in it's order
    // brittle yes - correct - yes
    val newShape = destSheet.getShapes().get(offset)



    // now we copy in the contents on the shape that has been added
    // Ref: update the shape according to its own additional copy rules
    m(newShape)('copy)(shape)

    f(destSheet)('_placeholders)(null)
    m(destSheet)('initPlaceholders)()
    m(destSheet)('commit)()

    newShape
  }

}
