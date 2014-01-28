package org.opennms.forge.requisitionstoxls;

import java.io.File;
import javax.xml.bind.JAXBContext;
import javax.xml.bind.JAXBException;
import javax.xml.bind.Unmarshaller;
import org.opennms.netmgt.provision.persist.requisition.Requisition;
import org.opennms.netmgt.provision.persist.requisition.RequisitionCollection;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class Starter {

    private static final Logger LOGGER = LoggerFactory.getLogger(Starter.class);

    private final static String OUTPUT_FILE_PARAMETER = "output-xls";
    private final static String INPUT_FILE_PARAMETER = "input-xml";
    private static File outFile;
    private static File inFile;

    public static void main(String[] args) {
        LOGGER.info("Hallo Welt");

        if (System.getProperty(OUTPUT_FILE_PARAMETER, null) != null && System.getProperty(INPUT_FILE_PARAMETER, null) != null) {
            inFile = new File(System.getProperty(INPUT_FILE_PARAMETER));
            outFile = new File(System.getProperty(OUTPUT_FILE_PARAMETER));
            LOGGER.debug("Input  File :: {}", inFile.getAbsolutePath());
            LOGGER.debug("Output File :: {}", outFile.getAbsolutePath());
            if (inFile.exists() && inFile.canRead()) {

                LOGGER.debug("Can read Input File");

                RequisitionToSpreatsheet requisitionToSpreatsheet = new RequisitionToSpreatsheet();
                try {
                    if (isInputManyRequisitions(inFile)) {
                        LOGGER.debug("running multiple requisitions");
                        requisitionToSpreatsheet.runRequisitions(inFile, outFile);
                    } else {
                        LOGGER.debug("running single requisition");
                        requisitionToSpreatsheet.runRequisition(inFile, outFile);
                    }
                } catch (Exception ex) {
                    LOGGER.error("Input file dose not contain one or many Requisitions", ex);
                }
                LOGGER.info("Thanks for computing with OpenNMS");
            }
        } else {
            LOGGER.info("Please provide the following parameters: -D" + OUTPUT_FILE_PARAMETER + " -D" + INPUT_FILE_PARAMETER);
        }
    }

    private static boolean isInputManyRequisitions(File inFile) throws Exception {
        try {
            JAXBContext jaxbContext = JAXBContext.newInstance(RequisitionCollection.class);
            Unmarshaller jaxbUnmarshaller = jaxbContext.createUnmarshaller();
            RequisitionCollection requisitions = (RequisitionCollection) jaxbUnmarshaller.unmarshal(inFile);
            LOGGER.debug("Unmarshalled {} requisitions", requisitions.size());
            return true;
        } catch (ClassCastException | JAXBException ex) {
            LOGGER.debug("Input file dose not contain multiple requisitions");
            LOGGER.trace("Input file dose not contain multiple requisitions", ex);
            try {
                JAXBContext jaxbContext = JAXBContext.newInstance(Requisition.class);
                Unmarshaller jaxbUnmarshaller = jaxbContext.createUnmarshaller();
                Requisition requisition = (Requisition) jaxbUnmarshaller.unmarshal(inFile);
                LOGGER.debug("Unmarshalled requisition {}", requisition.getForeignSource());
                return false;
            } catch (ClassCastException | JAXBException ex1) {
                LOGGER.debug("Input file dose not contain a single requisition");
                LOGGER.trace("Input file dose not contain a single requisition", ex1);
            }
        }
        throw new Exception("Input file dose not contain one or many Requisitions");
    }
}