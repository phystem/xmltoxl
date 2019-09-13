package xml2csv;

import com.jcabi.xml.XML;
import com.jcabi.xml.XMLDocument;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Node;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Stream;

public class X2CConverter {


    List<HashMap<String, String>> values = new ArrayList<>();

    HashMap<String, String> currentXmlMap;

    HashSet<String> columnNames = new LinkedHashSet<>();

    private void writeToExcel(String outputFile) {
        File excelFile = new File(outputFile + ".xlsx");
        System.out.println("Storing in Excel File - " + excelFile.getAbsolutePath());
        try {
            Workbook workbook;
            if (excelFile.exists())
                workbook = new XSSFWorkbook(new FileInputStream(excelFile));
            else
                workbook = new XSSFWorkbook();
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
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    private void convert(List<String> files, String outputFile) {
        System.out.println("Processing Files");
        files.forEach(file -> {
            System.out.println("Processing File - " + file);
            parseXmlsFromFile(file).forEach(xmlText -> {
                currentXmlMap = new LinkedHashMap<>();
                values.add(currentXmlMap);
                XML xml = new XMLDocument(xmlText).registerNs("ebm", getNameSpace(xmlText));
                Node itemNode = xml.nodes("//ebm:item").get(0).node();
                processNode(itemNode, getSimpleNodeName(itemNode));
            });
        });
        writeToExcel(outputFile);
    }

    private String getNameSpace(String xmlContent) {
        Pattern pattern = Pattern.compile("<ebm:PublishItem xmlns:ebm=\"(.*)\">");
        Matcher matcher = pattern.matcher(xmlContent);
        if (matcher.find()) {
            return matcher.group(1);
        }
        return "";
    }

    private List<String> parseXmlsFromFile(String file) {
        List<String> xmls = new ArrayList<>();
        StringBuilder xmlString = new StringBuilder();
        try {
            String fileContent = new String(Files.readAllBytes(Paths.get(file)));
            Stream.of(fileContent.split("\n"))
                    .forEach(line -> {
                        if (line.contains("<?xml version")) {
                            if (!xmlString.toString().isEmpty()) {
                                xmls.add(xmlString.toString());
                                xmlString.setLength(0);
                            }
                        } else {
                            xmlString.append(line);
                        }
                    });
            xmls.add(xmlString.toString());
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("Found " + xmls.size() + " xmls in file");
        return xmls;
    }

    private void storeDetails(Node childNode, String parentName) {
        String columnName = parentName + "_" + getSimpleNodeName(childNode);
        columnName = columnName.startsWith("_") ? columnName.substring(1) : columnName;
        String uniqColName = columnName;
        int counter = 1;
        while (currentXmlMap.containsKey(uniqColName)) {
            uniqColName = columnName + counter++;
        }
        String columnValue = childNode.getTextContent();
        columnNames.add(uniqColName);
        currentXmlMap.put(uniqColName, columnValue);
    }

    private String getSimpleNodeName(Node node) {
        return node.getNodeName().replace(node.getPrefix(), "").replace(":", "");
    }

    private void processNode(Node nodeToProcess, String parentName) {
        if (nodeToProcess.getNodeType() == 3) {
            return;
        }
        if (nodeToProcess.getChildNodes().getLength() > 1) {
            for (int i = 0; i < nodeToProcess.getChildNodes().getLength(); i++) {
                String processNodeName = getSimpleNodeName(nodeToProcess);
                processNode(nodeToProcess.getChildNodes().item(i), parentName.equals(processNodeName) ? "" : parentName + "_" + processNodeName);
            }
        } else {
            storeDetails(nodeToProcess, parentName);
        }
    }

    public static void main(String[] args) {
        if (args != null && args.length >= 2) {
            String excelFileName = args[0];
            List<String> files = new ArrayList<>();
            if (!args[1].endsWith(".xml")) {
                File folder = new File(args[1]);
                Stream.of(folder.listFiles((dir, name) -> name.endsWith(".xml"))).map(File::getAbsolutePath).forEach(files::add);
            } else {
                for (int i = 1; i < args.length; i++) {
                    files.add(args[i]);
                }
            }
            new X2CConverter().convert(files, excelFileName);
        } else {
            System.out.println("Usage --");
            System.out.println("java -jar xml2csv-1.0-SNAPSHOT.jar OutputExcelFileName input1.xml");
            System.out.println("java -jar xml2csv-1.0-SNAPSHOT.jar OutputExcelFileName input1.xml input2.xml");
            System.out.println("java -jar xml2csv-1.0-SNAPSHOT.jar OutputExcelFileName inputXmlFolderLocation");
        }
    }

}
