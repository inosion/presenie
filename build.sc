// build.sc
import mill._, scalalib._
import coursier.maven.MavenRepository

import $ivy.`com.lihaoyi::mill-contrib-bloop:0.4.1`

object CustomZincWorkerModule extends ZincWorkerModule {
  def repositories() = super.repositories ++ Seq(
    MavenRepository("http://artifactory.ai.cba/artifactory/maven")
  )
}

object docmaker extends ScalaModule {

  override def resources = T.sources{ millSourcePath / 'src / 'main / 'resources }

  def zincWorker = CustomZincWorkerModule

  def scalaVersion = "2.12.8"

  def ivyDeps = Agg(


    ivy"org.docx4j:docx4j-JAXB-ReferenceImpl:8.1.1",


    // ivy"org.docx4j:docx4j-JAXB-MOXy:8.1.1",
    // ivy"org.docx4j:docx4j-JAXB-Internal:8.1.1",


    // ivy"javax.xml.bind:jaxb-api:2.2.11",

    // ivy"com.sun.xml.bind:jaxb-core:2.2.11",

    // ivy"com.sun.xml.bind:jaxb-impl:2.2.11",

    // ivy"javax.activation:activation:1.1.1",

    ivy"org.rogach::scallop:3.3.1",
    ivy"commons-io:commons-io:2.6",
    ivy"ch.qos.logback:logback-classic:1.2.3",
    ivy"com.typesafe.scala-logging::scala-logging:3.9.2",
    ivy"io.gatling::jsonpath:0.7.0",
    ivy"com.fasterxml.jackson.core:jackson-databind:2.9.9",
    ivy"com.fasterxml.jackson.dataformat:jackson-dataformat-yaml:2.9.2",
    ivy"org.slf4j:slf4j-api:1.7.5",
  )
    /* circe may be a library we will test
    ivy"io.circe::circe-parser:0.11.1",
    ivy"io.circe::circe-yaml:0.11.0-M1",
    ivy"io.circe::circe-optics:0.11.0",
    */

}
