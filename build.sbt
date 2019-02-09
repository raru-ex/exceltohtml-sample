name := "exceltohtml-sample"
organization := "net.syrup16g"

version := "0.1"

scalaVersion := "2.12.8"

libraryDependencies ++= Seq(
  "org.apache.poi" % "poi" % "4.0.1",
  "org.apache.poi" % "poi-scratchpad" % "4.0.1",
  "org.apache.poi" % "poi-ooxml" % "4.0.1",
  "org.apache.poi" % "poi-ooxml-schemas" % "4.0.1",
  "org.dom4j" % "dom4j" % "2.1.1",
)