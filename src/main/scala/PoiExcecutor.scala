package net.syrup16g.sample.excel

import java.io._
import java.time.LocalDateTime
import java.time.format.DateTimeFormatter

import javax.xml.parsers.DocumentBuilderFactory
import javax.xml.transform.{OutputKeys, TransformerFactory}
import javax.xml.transform.dom.DOMSource
import javax.xml.transform.stream.StreamResult
import org.apache.poi.hssf.converter.ExcelToHtmlConverter
import org.apache.poi.hssf.usermodel.HSSFWorkbook

object PoiExecutor {

  def main(args: Array[String]): Unit = {
    val xlsFilePath = "files/input/sample1.xls"
    val outputFile = "files/output/sample2html-" + getFormatedCurrentDateTimeString() + ".html"
    val workbook = loadExcel(xlsFilePath)

    val excelToHtmlConverter = createExcelToHtmlConverter(workbook)
    excelToHtmlConverter.processWorkbook(workbook)

    val htmlDocument = excelToHtmlConverter.getDocument
    val transformer = TransformerFactory.newInstance().newTransformer()
    transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8")
    transformer.setOutputProperty(OutputKeys.INDENT, "yes")
    transformer.setOutputProperty(OutputKeys.METHOD, "html")

    val outStream = new ByteArrayOutputStream()
    val domSource = new DOMSource(htmlDocument)
    val streamResult = new StreamResult(outStream)
    transformer.transform(domSource, streamResult)
    outStream.close()

    val printStream = new PrintStream(new FileOutputStream(new File(outputFile)))
    printStream.print(new String(outStream.toByteArray))
    printStream.close()
  }


  private def loadExcel(path: String): HSSFWorkbook = {
    // pathからexcelを読み込む
    val input = new FileInputStream(path)
    return new HSSFWorkbook(input)
  }

  private def createExcelToHtmlConverter(workbook: HSSFWorkbook): ExcelToHtmlConverter =
    new ExcelToHtmlConverter(DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument())

  private def getFormatedCurrentDateTimeString(): String = {
    return LocalDateTime.now.format(DateTimeFormatter.ofPattern("yyyy-MM-dd_HH-mm-ss"))
  }
}
