package inosion

import java.io.File
import org.rogach.scallop._


object DocMaker extends App {

  val version = "1.0.0"
  
  class Conf(arguments: Seq[String]) extends ScallopConf(arguments) {    


      trait TemplateArg { _: ScallopConf  => 
        val template = opt[File](required = true, descr = "The template to merge data with")
        mainOptions = Seq(template)
      }
      version(s"docmaker ${DocMaker.version} (c) 2019 Inosion")
      val verbose =  opt[Boolean]()

      val list =     new Subcommand("list") with TemplateArg {
      }
      addSubcommand(list)

      val merge =     new Subcommand("merge") with TemplateArg {
        val data =     opt[File](required = true, descr = "The data file to merge with the template")
        val config =   opt[File](required = true, descr = "Config file to drive merging")
        val outFile =  opt[File](required = true, descr = "Output filename")
      }

      addSubcommand(merge)

      verify()

  }

  val conf = new Conf(args)

  conf.subcommand match { 
    case Some(conf.list) => pptx.PPTXTools.listSlideLayouts(conf.list.template())
    case Some(conf.merge) => { 
      pptx.PPTXMerger.render(conf.merge.config(), conf.merge.data(), conf.merge.template(), conf.merge.outFile())
      Console.println(s"Wrote ${conf.merge.outFile.apply().getAbsolutePath()}")
    }
    case _ => conf.printHelp()
  }

}
