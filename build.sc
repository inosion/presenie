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


  def zincWorker = CustomZincWorkerModule

  def scalaVersion = "2.12.8"

  def ivyDeps = Agg(
    ivy"org.apache.poi:poi:4.1.0",
    ivy"org.apache.poi:poi-ooxml:4.1.0",
    ivy"org.rogach::scallop:3.3.1",
    ivy"commons-io:commons-io:2.6",
    ivy"ch.qos.logback:logback-classic:1.2.3",
    ivy"com.typesafe.scala-logging::scala-logging:3.9.2",
    ivy"io.gatling::jsonpath:0.7.0",
    ivy"com.fasterxml.jackson.core:jackson-databind:2.9.9",
    ivy"com.fasterxml.jackson.dataformat:jackson-dataformat-yaml:2.9.2",
  )
    /* circe may be a library we will test
    ivy"io.circe::circe-parser:0.11.1",
    ivy"io.circe::circe-yaml:0.11.0-M1",
    ivy"io.circe::circe-optics:0.11.0",
    */

}
