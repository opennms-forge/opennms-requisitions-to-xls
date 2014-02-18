package org.opennms.forge.requisitionstoxls;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Set;
import java.util.TreeSet;
import javax.xml.bind.JAXBContext;
import javax.xml.bind.JAXBException;
import javax.xml.bind.Unmarshaller;
import jxl.Cell;
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
import org.opennms.netmgt.provision.persist.requisition.RequisitionAsset;
import org.opennms.netmgt.provision.persist.requisition.RequisitionCategory;
import org.opennms.netmgt.provision.persist.requisition.RequisitionCollection;
import org.opennms.netmgt.provision.persist.requisition.RequisitionInterface;
import org.opennms.netmgt.provision.persist.requisition.RequisitionMonitoredService;
import org.opennms.netmgt.provision.persist.requisition.RequisitionNode;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class RequisitionToSpreatsheet {

    private static final Logger LOGGER = LoggerFactory.getLogger(RequisitionToSpreatsheet.class);
    private List<String> headers = new ArrayList<>();
    private final String ASSET_PREFIX = "Asset_";
    private final String CATEGORY_PREFIX = "cat_";
    private final String SERVICE_PREFIX = "svc_";

    public void runRequisition(File inputFile, File outputFile) {
        Requisition requisition = readRequisitionFromFile(inputFile);
        if (requisition != null) {
            headers.add("Node_Label");
            headers.add("IP_Interface");
            headers.add("IfType_SNMP");
            headers.addAll(buildHeader(requisition));
            requisition2SpreatSheet(requisition, outputFile);
            LOGGER.info("Wrote requisition {} into file {}", requisition.getForeignSource(), outputFile.getAbsolutePath());
        } else {
            LOGGER.error("InputFile dose not contain requisition stoping process");
        }
    }

    public void runRequisitions(File inputFile, File outputFile) {
        List<Requisition> requisitions = readRequisitonsFromFile(inputFile);
        if (requisitions != null) {
            headers.add("Node_Label");
            headers.add("IP_Interface");
            headers.add("IfType_SNMP");
            this.headers.addAll(buildHeader(requisitions));
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
            for (String header : headers) {
                Label label = new Label(cellIndex, rowIndex, header);
                sheet.addCell(label);
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

                        for (Cell headerCell : sheet.getRow(0)) {
                            if (headerCell.getContents().startsWith(ASSET_PREFIX)) {
                                RequisitionAsset asset = node.getAsset(headerCell.getContents().substring(ASSET_PREFIX.length()));
                                if (asset != null) {
                                    label = new Label(headerCell.getColumn(), rowIndex, asset.getValue());
                                    sheet.addCell(label);
                                }
                            }
                            if (headerCell.getContents().startsWith(CATEGORY_PREFIX)) {
                                RequisitionCategory category = node.getCategory(headerCell.getContents().substring(CATEGORY_PREFIX.length()));
                                if (category != null) {
                                    label = new Label(headerCell.getColumn(), rowIndex, category.getName());
                                    sheet.addCell(label);
                                }
                            }
                        }
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

                    for (Cell headerCell : sheet.getRow(0)) {
                        if (headerCell.getContents().startsWith(SERVICE_PREFIX)) {

                            for (RequisitionMonitoredService service : reqInterface.getMonitoredServices()) {
                                if (service.getServiceName().equals(headerCell.getContents().substring(SERVICE_PREFIX.length()))) {
                                    label = new Label(headerCell.getColumn(), rowIndex, service.getServiceName());
                                    sheet.addCell(label);
                                }
                            }
                        }
                    }

                    rowIndex++;
                    cellIndex = 0;
                }
                cellIndex = 0;
            }

            workbook.write();
            workbook.close();

        } catch (WriteException writeException) {
            LOGGER.error("Writing to xls caused a pheaderNamesroblem", writeException);
        } catch (BiffException biffException) {
            LOGGER.error("Working with the xls data caused a problem", biffException);
        } catch (IOException io) {
            LOGGER.error("Writing the output file {} caused a problem", outFile.getAbsolutePath(), io);
        }

        LOGGER.debug(
                "Wrote output for requisition {} into file {}", requisition.getForeignSource(), outFile.getAbsolutePath());
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

    private Set<String> addFieldsToHeader(RequisitionNode node) {
        Set<String> headerNames = new TreeSet<>();
        for (RequisitionCategory category : node.getCategories()) {
            headerNames.add(CATEGORY_PREFIX + category.getName());
        }
        for (RequisitionAsset asset : node.getAssets()) {
            headerNames.add(ASSET_PREFIX + asset.getName());
        }
        for (RequisitionInterface reqInterface : node.getInterfaces()) {
            for (RequisitionMonitoredService service : reqInterface.getMonitoredServices()) {
                headerNames.add(SERVICE_PREFIX + service.getServiceName());
            }
        }
        return headerNames;
    }

    private Set<String> buildHeader(List<Requisition> requisitions) {
        Set<String> headerNames = new TreeSet<>();
        for (Requisition requisition : requisitions) {
            for (RequisitionNode node : requisition.getNodes()) {
                headerNames.addAll(addFieldsToHeader(node));
            }
        }
        return headerNames;
    }

    private Set<String> buildHeader(Requisition requisition) {
        Set<String> headerNames = new TreeSet<>();
        for (RequisitionNode node : requisition.getNodes()) {
            headerNames.addAll(addFieldsToHeader(node));
        }
        return headerNames;
    }
}
