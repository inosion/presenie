plugins {
    scala
    application
}

repositories {
    mavenCentral()
}

dependencies {
    implementation("org.scala-lang:scala-library:2.13.11")

    implementation("org.apache.poi:poi:5.2.5")
    implementation("org.apache.poi:poi-ooxml:5.2.5")

    // for the CLI
    implementation("org.rogach:scallop_2.13:5.1.0")

    implementation("commons-io:commons-io:2.6")
    implementation("ch.qos.logback:logback-classic:1.5.6")
    implementation("org.slf4j:slf4j-api:1.7.32")

    // add log4j2 core
     implementation("org.apache.logging.log4j:log4j-core:2.23.1")
     implementation("org.apache.logging.log4j:log4j-api:2.23.1")
     implementation("com.typesafe.scala-logging:scala-logging_2.13:3.9.5")
    implementation("com.filippodeluca:jsonpath-parser_2.13:0.0.28")
    implementation("com.filippodeluca:jsonpath-circe_2.13:0.0.28")

    testImplementation("junit:junit:4.13.2")
    testImplementation("org.scalatest:scalatest_2.13:3.2.9")
}

application {
    mainClass = "inosion.presenie.Presenie"
}

tasks.jar {
    manifest.attributes["Main-Class"] = "inosion.presenie.Presenie"
    val dependencies = configurations
        .runtimeClasspath
        .get()
        .map(::zipTree) // OR .map { zipTree(it) }
    from(dependencies)
    duplicatesStrategy = DuplicatesStrategy.EXCLUDE
}