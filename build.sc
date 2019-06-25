// build.sc
import mill._, scalalib._

object docmaker extends ScalaModule {
  def scalaVersion = "2.12.8"
  def ivyDeps = Agg(
    ivy"org.apache.poi:poi:4.1.0",
    ivy"org.apache.poi:poi-ooxml:4.1.0",
    ivy"org.rogach::scallop:3.3.1",

    /*
    ivy"io.circe::circe-parser:0.11.1",
    ivy"io.circe::circe-yaml:0.11.0-M1",
    ivy"io.circe::circe-optics:0.11.0",
    */
    ivy"io.gatling::jsonpath:0.7.0",
    ivy"com.fasterxml.jackson.core:jackson-databind:2.9.9",
    ivy"com.fasterxml.jackson.dataformat:jackson-dataformat-yaml:2.9.2",
  )
}
