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

      val list =     new Subcommand("list") with TemplateArg {}
      addSubcommand(list)

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
    case Some(conf.list) => pptx.PPTXTools.listSlideLayouts(conf.list.template())
    case Some(conf.merge) => {
      pptx.PPTXMerger.render(conf.merge.data(), conf.merge.template(), conf.merge.outFile())
      Console.println(s"Wrote ${conf.merge.outFile.apply().getAbsolutePath()}")
    }
    case Some(conf.cloner) => {
      Console.println(s"Cloned ${conf.merge.template.apply().getAbsolutePath()} --> ${conf.merge.outFile.apply().getAbsolutePath()}")
      pptx.PPTXTools.clonePptx(conf.cloner.template().toPath(), conf.cloner.outFile().toPath())
    }
    case _ => conf.printHelp()
  }

}
