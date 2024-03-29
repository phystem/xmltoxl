package xml2xl;

import com.jcabi.xml.XML;
import com.jcabi.xml.XMLDocument;
import net.sf.practicalxml.DomUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.w3c.dom.Element;
import org.w3c.dom.Node;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Stream;

public class X2XConverter {

    private String parentNodeName;

    List<HashMap<String, String>> values = new ArrayList<>();

    HashMap<String, String> currentXmlMap;

    HashSet<String> columnNames = new LinkedHashSet<>();

    int invalidXmls = 0;

    public X2XConverter(String parentNodeName) {
        this.parentNodeName = parentNodeName;
    }

    private void writeToExcel(String outputFile) {
        if (!values.isEmpty()) {

            System.out.println("No of columns - " + columnNames.size());
            System.out.println("No of Rows - " + values.size());
            if (invalidXmls > 0)
                System.out.println("Couldn't Process xmls - " + invalidXmls);

            File excelFile = new File(outputFile + ".xlsx");
            System.out.println("Storing in Excel File - " + excelFile.getAbsolutePath());
            try {
                SXSSFWorkbook workbook = new SXSSFWorkbook();
                workbook.setCompressTempFiles(true);

                int sheetSize = workbook.getNumberOfSheets();
                Sheet sheet = workbook.createSheet("Sheet" + sheetSize);
                Row headerRow = sheet.createRow(0);
                columnNames.forEach(columnName -> {
                    headerRow.createCell(headerRow.getPhysicalNumberOfCells()).setCellValue(columnName);
                });
                values.forEach(excelValue -> {
                    Row dataRow = sheet.createRow(sheet.getPhysicalNumberOfRows());
                    columnNames.forEach(columnName -> {
                        Cell dataCell = dataRow.createCell(dataRow.getPhysicalNumberOfCells());
                        if (excelValue.containsKey(columnName)) {
                            dataCell.setCellValue(excelValue.get(columnName));
                        } else {
                            dataCell.setBlank();
                        }
                    });
                });
                workbook.write(new FileOutputStream(excelFile));
                workbook.dispose();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private void convert(List<String> files, String outputFile) {
        System.out.println("Processing Files");
        files.forEach(file -> {
            System.out.println("Processing File - " + file);
            parseAndProcessXml(file);
        });
        writeToExcel(outputFile);
    }

    private void writeXmlToXl(String xmlText) {
        xmlText = xmlText.replaceAll("\\P{Print}", "");
        xmlText = xmlText.substring(0, xmlText.lastIndexOf(">") + 1);
        currentXmlMap = new LinkedHashMap<>();
        values.add(currentXmlMap);
        try {
            XML xml = new XMLDocument(xmlText).registerNs("ebm", getNameSpace(xmlText));
            List<XML> nodes = xml.nodes("//" + parentNodeName);
            if (!nodes.isEmpty()) {
                processNode(nodes.get(0).node());
            } else {
                System.err.println("Couldn't found nodes matching //" + parentNodeName);
            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }


    private String getNameSpace(String xmlContent) {
        Pattern pattern = Pattern.compile("<ebm:.* xmlns:ebm=\"(.*)\">");
        Matcher matcher = pattern.matcher(xmlContent);
        if (matcher.find()) {
            return matcher.group(1);
        }
        return "";
    }

    private void parseAndProcessXml(String file) {
        StringBuilder xmlString = new StringBuilder();
        try {
            String fileContent = new String(Files.readAllBytes(Paths.get(file)));
            Stream.of(fileContent.split("\n"))
                    .forEach(line -> {
                        if (line.contains("<?xml version")) {
                            if (line.startsWith("</")) {
                                xmlString.append(line.substring(0, line.indexOf(">") + 1));
                            }
                            if (!xmlString.toString().isEmpty()) {
                                writeXmlToXl(xmlString.toString());
                                xmlString.setLength(0);
                            }
                        } else {
                            xmlString.append(line);
                        }
                    });
            writeXmlToXl(xmlString.toString());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void storeDetails(Node childNode) {
        String columnName = DomUtil.getAbsolutePath((Element) childNode);
        columnName = columnName
                .substring(columnName.indexOf("/", 1) + 1)
                .replace("/", "_")
                .replace("[", "")
                .replace("]", "");
        String columnValue = childNode.getTextContent();
        columnNames.add(columnName);
        currentXmlMap.put(columnName, columnValue);
    }

    private void processNode(Node nodeToProcess) {
        DomUtil.getChildren(nodeToProcess).forEach(element -> {
            if (DomUtil.hasElementChildren(element)) {
                processNode(element);
            } else {
                storeDetails(element);
            }
        });
    }

    public static void main(String[] args) {
        if (args != null && args.length >= 3) {
            String xmlNodeName = args[0];
            String excelFileName = args[1];
            List<String> files = new ArrayList<>();
            if (!args[2].endsWith(".xml")) {
                File folder = new File(args[2]);
                Stream.of(folder.listFiles((dir, name) -> name.endsWith(".xml"))).map(File::getAbsolutePath).forEach(files::add);
            } else {
                for (int i = 2; i < args.length; i++) {
                    files.add(args[i]);
                }
            }
            new X2XConverter(xmlNodeName).convert(files, excelFileName);
        } else {
            System.out.println("Usage --");
            System.out.println("java -jar xml2xl-1.0-SNAPSHOT.jar xmlNodeName OutputExcelFileName input1.xml");
            System.out.println("java -jar xml2xl-1.0-SNAPSHOT.jar xmlNodeName OutputExcelFileName input1.xml input2.xml");
            System.out.println("java -jar xml2xl-1.0-SNAPSHOT.jar xmlNodeName OutputExcelFileName inputXmlFolderLocation");
            System.out.println("java -jar xml2xl-1.0-SNAPSHOT.jar ebm:item OutputExcelFileName input1.xml");
        }
    }

}
