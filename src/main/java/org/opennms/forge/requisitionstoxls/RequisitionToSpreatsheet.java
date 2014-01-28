package org.opennms.forge.requisitionstoxls;

import java.io.File;
import java.io.IOException;
import java.util.List;
import javax.xml.bind.JAXBContext;
import javax.xml.bind.JAXBException;
import javax.xml.bind.Unmarshaller;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Colour;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import org.opennms.netmgt.provision.persist.requisition.Requisition;
import org.opennms.netmgt.provision.persist.requisition.RequisitionCategory;
import org.opennms.netmgt.provision.persist.requisition.RequisitionCollection;
import org.opennms.netmgt.provision.persist.requisition.RequisitionInterface;
import org.opennms.netmgt.provision.persist.requisition.RequisitionMonitoredService;
import org.opennms.netmgt.provision.persist.requisition.RequisitionNode;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class RequisitionToSpreatsheet {

    private static final Logger LOGGER = LoggerFactory.getLogger(RequisitionToSpreatsheet.class);

    private Requisition requisition;
    private final String HEADER = "Node_Lable\tIP_Management\tIfType\tAsset_Description\tSvc_Forced\tCat_Location\tCat_OperatingSystem\tCat_Environment\tCat_General";
    private final String ASSET_DESCRIPTION = "description";

    private final File requisitionFile = new File("/tmp/Requisition.xml");
    private final File requisitionsFile = new File("/tmp/svorcmonitor.xml");
    private final File spreatSheetFile = new File("/tmp/svorcmonitor.xls");

    public void runRequisition(File inputFile, File outputFile) {
        requisition = readRequisitionFromFile(requisitionFile);
        if (requisition != null) {
            requisition2SpreatSheet(requisition, spreatSheetFile);
            LOGGER.info("Wrote requisition {} into file {}", requisition.getForeignSource(), outputFile.getAbsolutePath());
        } else {
            LOGGER.error("InputFile dose not contain requisition stoping process");
        }
    }

    public void runRequisitions(File inputFile, File outputFile) {
        List<Requisition> requisitions = readRequisitonsFromFile(inputFile);
        if (requisitions != null) {
            requisitions2SpreatSheet(requisitions, outputFile);
            LOGGER.info("Wrote all {} requisitions into file {}", requisitions.size(), outputFile.getAbsolutePath());
        } else {
            LOGGER.error("InputFile dose not contain requisitions stoping process");
        }
    }

    public void requisitions2SpreatSheet(List<Requisition> requisitions, File outFile) {
        for (Requisition requisition : requisitions) {
            requisition2SpreatSheet(requisition, outFile);
        }
    }

    public List<Requisition> readRequisitonsFromFile(File requisitionsFile) {
        RequisitionCollection requisitions = null;
        try {
            JAXBContext jaxbContext = JAXBContext.newInstance(RequisitionCollection.class);
            Unmarshaller jaxbUnmarshaller = jaxbContext.createUnmarshaller();
            requisitions = (RequisitionCollection) jaxbUnmarshaller.unmarshal(requisitionsFile);
        } catch (JAXBException ex) {
            LOGGER.error("Reading requisitions from inputFile failed", ex);
        }
        return requisitions;
    }

    public Requisition readRequisitionFromFile(File requFile) {
        Requisition requisition = null;
        try {
            JAXBContext jaxbContext = JAXBContext.newInstance(Requisition.class);
            Unmarshaller jaxbUnmarshaller = jaxbContext.createUnmarshaller();
            requisition = (Requisition) jaxbUnmarshaller.unmarshal(requFile);
        } catch (JAXBException ex) {
            LOGGER.error("Reading requisition from inputFile failed", ex);
        }
        return requisition;
    }

    public void requisition2SpreatSheet(Requisition requisition, File outFile) {

        try {
            WritableWorkbook workbook = null;
            if (outFile.exists()) {
                Workbook existingWorkbook = Workbook.getWorkbook(outFile);
                System.out.println("Adding requisition " + requisition.getForeignSource() + " to file " + outFile.getAbsolutePath());
                workbook = Workbook.createWorkbook(outFile, existingWorkbook);

            } else {
                System.out.println("Creating requisition " + requisition.getForeignSource() + " in new file " + outFile.getAbsolutePath());
                workbook = Workbook.createWorkbook(outFile);
            }
            WritableSheet sheet = workbook.createSheet(requisition.getForeignSource(), workbook.getNumberOfSheets());

            //Write Header into Sheet
            Integer rowIndex = 0;
            Integer cellIndex = 0;
            String[] headerRow = HEADER.split("\t");
            for (String headerCell : headerRow) {

                Label lable = new Label(cellIndex, rowIndex, headerCell);
                sheet.addCell(lable);
                cellIndex++;
            }
            rowIndex++;
            cellIndex = 0;

            for (RequisitionNode node : requisition.getNodes()) {
                Boolean nodeAdded = false;
                Label label = null;
                for (RequisitionInterface reqInterface : node.getInterfaces()) {

                    if (!nodeAdded) {
                        label = new Label(cellIndex, rowIndex, node.getNodeLabel());
                        sheet.addCell(label);
                        nodeAdded = true;
                    } else {
                        WritableCellFormat additionalInterfaceRow = new WritableCellFormat();
                        additionalInterfaceRow.setAlignment(Alignment.RIGHT);
                        additionalInterfaceRow.setBackground(Colour.AQUA);
                        label = new Label(cellIndex, rowIndex, node.getNodeLabel(), additionalInterfaceRow);
                        sheet.addCell(label);
                    }
                    cellIndex++;

                    label = new Label(cellIndex, rowIndex, reqInterface.getIpAddr());
                    sheet.addCell(label);
                    cellIndex++;
                    label = new Label(cellIndex, rowIndex, reqInterface.getSnmpPrimary().getCode());
                    sheet.addCell(label);
                    cellIndex++;
                    if (node.getAsset(ASSET_DESCRIPTION) != null) {
                        label = new Label(cellIndex, rowIndex, node.getAsset(ASSET_DESCRIPTION).getValue());
                        sheet.addCell(label);
                    }
                    cellIndex++;
                    label = new Label(cellIndex, rowIndex, getForcedServices(reqInterface));
                    sheet.addCell(label);
                    cellIndex++;

                    for (RequisitionCategory category : node.getCategories()) {
                        label = new Label(cellIndex, rowIndex, category.getName());
                        sheet.addCell(label);
                        cellIndex++;
                    }

                    cellIndex = 0;
                    rowIndex++;
                }
                cellIndex = 0;
            }

            workbook.write();
            workbook.close();

        } catch (WriteException writeException) {
            LOGGER.error("Writing to xls caused a problem", writeException);
        } catch (BiffException biffException) {
            LOGGER.error("Working with the xls data caused a problem", biffException);
        } catch (IOException io) {
            LOGGER.error("Writing the output file {} caused a problem", outFile.getAbsolutePath(), io);
        }
        LOGGER.debug("Wrote output for requisition {} into file {}", requisition.getForeignSource(), outFile.getAbsolutePath());
    }

    private String getForcedServices(RequisitionInterface reqInterface) {
        StringBuilder sb = new StringBuilder();
        for (RequisitionMonitoredService service : reqInterface.getMonitoredServices()) {
            sb.append(service.getServiceName());
            sb.append(", ");
        }
        if (sb.toString().endsWith(", ")) {
            return sb.toString().substring(0, sb.toString().length() - 2);
        }
        return sb.toString();
    }
}
