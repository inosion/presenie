// import org.apache.poi.xslf.usermodel.{XMLSlideShow, XSLFSlide}
// import org.scalatest.flatspec.AnyFlatSpec
// import org.scalatest.matchers.should.Matchers

// class ToolsTest extends AnyFlatSpec with Matchers {

//   "copySlideContent" should "copy content from source slide to destination slide" in {
//     val tools = new PPTXTools() // Assuming Tools is a class containing copySlideContent method

//     val srcPpt = new XMLSlideShow(getClass.getResourceAsStream("/srcSlide.pptx"))
//     val destPpt = new XMLSlideShow(getClass.getResourceAsStream("/destSlide.pptx"))

//     val srcSlide: XSLFSlide = srcPpt.getSlides.get(0)
//     val destSlide: XSLFSlide = destPpt.createSlide()

//     tools.copySlideContent(srcSlide, destSlide)

//     // Here you should add assertions to check if the content has been copied correctly.
//     // This depends on what you consider as "content" and how you can access it.
//     // For example, if you consider text as content, you can check if the text in the destination slide
//     // is the same as the text in the source slide.
//     // destSlide.getText should be (srcSlide.getText)
//   }
// }