package org.apache.poi.xslf.usermodel


import org.apache.poi.sl.usermodel._
import org.openxmlformats.schemas.drawingml.x2006.main.CTTableRow

import inosion.presenie.pptx.PPTXMerger
import inosion.presenie.pptx.PPTXTools
import scala.collection.JavaConverters._

object RowCloner {

    def cloneRow(table: XSLFTable, rowId: Int) : Unit = {
        val oldRow = table.getRows.get(rowId)

        try {
             // val ctrow = CTTableRow.Factory.parse(oldRow.getXmlObject().newInputStream())
             val row = table.addRow()
             val rowSize = oldRow.getCells().size()

             for (x <- 0 until rowSize) {
                val cell = row.addCell()
                val oldCell = oldRow.getCells().get(x)
                cell.setText(oldCell.getText())
                
                cell.setVerticalAlignment(oldCell.getVerticalAlignment())
                cell.setLeftInset(oldCell.getLeftInset())
                cell.setRightInset(oldCell.getRightInset())
                cell.setTopInset(oldCell.getTopInset())
                cell.setBottomInset(oldCell.getBottomInset())

                // get the font styling from the old cell first paragraph first run
                val oldRun = oldCell.getTextParagraphs().get(0).getTextRuns().get(0)
                PPTXTools.textParagraphCloneStyle(oldCell.getTextParagraphs().get(0), cell.getTextParagraphs().get(0))
                cell.getTextParagraphs().get(0).getTextRuns().get(0).setFontColor(oldRun.getFontColor())
                cell.getTextParagraphs().get(0).getTextRuns().get(0).setFontFamily(oldRun.getFontFamily())
                cell.getTextParagraphs().get(0).getTextRuns().get(0).setFontSize(oldRun.getFontSize())
                cell.getTextParagraphs().get(0).getTextRuns().get(0).setUnderlined(oldRun.isUnderlined())
                cell.getTextParagraphs().get(0).getTextRuns().get(0).setStrikethrough(oldRun.isStrikethrough())
                cell.getTextParagraphs().get(0).getTextRuns().get(0).setSuperscript(oldRun.isSuperscript())
                cell.getTextParagraphs().get(0).getTextRuns().get(0).setSubscript(oldRun.isSubscript())
                cell.getTextParagraphs().get(0).getTextRuns().get(0).setItalic(oldRun.isItalic())
                cell.getTextParagraphs().get(0).getTextRuns().get(0).setBold(oldRun.isBold())

                 
                 // reapply the same font styling
                 // PPTXMerger.rawChangeTextPreserveStyling(cell, oldCell.getText())
                 
                 cell.setFillColor(oldCell.getFillColor())

                 for (beType <- List(TableCell.BorderEdge.bottom, TableCell.BorderEdge.top, TableCell.BorderEdge.right, TableCell.BorderEdge.left)) {
                   //  cell.setBorderCap(     beType,    oldCell.getBorderCap(beType))
                   if(oldCell.getBorderColor(beType) != null) cell.setBorderColor(   beType,    oldCell.getBorderColor(beType))
                   // cell.setBorderCompound(beType,    oldCell.getBorderCompound(beType))
                   //   cell.setBorderDash(    beType,    oldCell.getBorderDash(beType))
                   if(oldCell.getBorderStyle(beType) != null) cell.setBorderStyle(   beType,    oldCell.getBorderStyle(beType))
                   if(oldCell.getBorderWidth(beType) != null) cell.setBorderWidth(   beType,    oldCell.getBorderWidth(beType))
                 }

                 cell.getTextParagraphs().get(0).getTextRuns().get(0).setFontSize(cell.getTextParagraphs().get(0).getTextRuns().get(0).getFontSize())
             }

        } catch {
            //case e: XmlException => e.printStackTrace()
            //case e: IOException  => e.printStackTrace()
            case e: Exception  => {
                e.printStackTrace()
            }
        }

    }

    def cloneRow2(table: XSLFTable, rowId: Int) : Unit = {

        val oldRow = table.getRows.get(rowId)


        val newRow: XSLFTableRow = table.addRow();
        newRow.setHeight(oldRow.getHeight());
        for (x <- 0 until oldRow.getCells().size()) {

          val newCell: XSLFTableCell = newRow.addCell()
          val oldCell = oldRow.getCells().get(x)

          // clone all paragraphs in this cell
          for (p <- oldCell.getTextParagraphs().asScala) {
            val para: XSLFTextParagraph = newCell.addNewTextParagraph();
            para.setTextAlign(p.getTextAlign())

            for (t <- p.getTextRuns().asScala) {
                val r1 = para.addNewTextRun();
                r1.setText(t.getRawText());
                r1.setFontColor(t.getFontColor());
                r1.setFontSize(t.getFontSize());
                r1.setFontFamily(t.getFontFamily())
            }
          }
        }
    }
}