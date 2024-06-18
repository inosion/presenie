package inosion.presenie

import java.io.File
import org.rogach.scallop._


object Presenie extends App {

  val version = "1.0.0"

  class Conf(arguments: Seq[String]) extends ScallopConf(arguments) {


      trait TemplateArg { _: ScallopConf  =>
        val template = opt[File](required = true, descr = "The template to merge data with")
        mainOptions = Seq(template)
      }
      version(s"Presenie ${Presenie.version} (c) 2024 Inosion")
      val verbose =  opt[Boolean]()

      val inspect =     new Subcommand("inspect") with TemplateArg {}
      addSubcommand(inspect)

      val merge =     new Subcommand("merge") with TemplateArg {
        val data =     opt[File](required = true, descr = "The data file to merge with the template")
        val config =   opt[File](required = false, descr = "Config file to drive merging")
        val outFile =  opt[File](required = true, descr = "Output filename")
      }
      addSubcommand(merge)

      val cloner =     new Subcommand("clone") with TemplateArg {
        val outFile =  opt[File](required = true, descr = "Output filename")
      }
      addSubcommand(cloner)

      verify()

  }

  val conf = new Conf(args)

  conf.subcommand match {
    /*
     * A debug method of sorts
     */
    case Some(conf.inspect) => {
      pptx.PPTXTools.listSlideLayouts(conf.inspect.template())
    }
    /*
     * Do the merge
     */
    case Some(conf.merge) => {
      val files = pptx.FilesContext(src = conf.merge.template.apply(), dst = conf.merge.outFile.apply(), data = conf.merge.data.apply())
      pptx.PPTXMerger.render(files)
      Console.println(s"Wrote ${conf.merge.outFile.apply().getAbsolutePath()}")
    }
    case Some(conf.cloner) => {
      Console.println(s"Cloned ${conf.merge.template.apply().getAbsolutePath()} --> ${conf.merge.outFile.apply().getAbsolutePath()}")
      pptx.PPTXTools.clonePptSlides(conf.cloner.template(), conf.cloner.outFile())
    }
    case _ => conf.printHelp()
  }

}
