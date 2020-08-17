package sample;
/*
Licensed under the Apache License, Version 2.0 (the "License");
        you may not use this file except in compliance with the License.
        You may obtain a copy of the License at

        http://www.apache.org/licenses/LICENSE-2.0

        Unless required by applicable law or agreed to in writing, software
        distributed under the License is distributed on an "AS IS" BASIS,
        WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
        See the License for the specific language governing permissions and
        limitations under the License.

        Last Edited by Stephen Peters on June 4th 2020

        Created by Stephen Peters to evaluate the SAPPhIRE Causal Model as a framework for this conceptual design concurrent engineering tool.
*/

import java.awt.*;
import java.io.*;
import java.io.File;
import java.io.IOException;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import javafx.fxml.FXML;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import javafx.scene.layout.GridPane;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;


import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;

import java.lang.reflect.Array;
import java.net.URI;
import java.util.Scanner;

import javafx.event.ActionEvent;
import javafx.scene.control.Button;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.scene.Parent;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.stage.Stage;
import javafx.scene.Node;

import java.math.RoundingMode;
import java.text.DecimalFormat;
import java.util.*;

public class Controller

{
    public DecimalFormat DF = new DecimalFormat("######.###");

    @FXML TextArea ControlText;

    @FXML public TextArea MainThermalTextArea;

    @FXML public TextArea MainPowerTextArea;

    @FXML public TextArea MainCommunicationsTextArea;

    @FXML public TextArea ThermalInputsText;

    @FXML public TextArea MainStructureTextArea;

    @FXML public TextArea OPTextArea;

    @FXML public TextField UserPathField;

    @FXML public static String UserPath = "C://Users//steph//OneDrive//Desktop"; //--module-path C:\Program Files\Java\jdk-13.0.2\lib --add-modules javafx.controls,javafx.fxml

    /* This is for the UPEI Desktop
    public String SpudNik1Assembly = "C://Users//sepeters.ADDS//Desktop//Conceptual Design Files//Subset of Parts//Current SpudNik-1 Assembly.SLDASM";
    public String FileLocationThermal = "C://Users//sepeters.ADDS//Desktop//Conceptual Design Files//Thermal Study v_03.xlsx";
    public String BOM = "C://Users//sepeters.ADDS//Desktop//Conceptual Design Files//SpudNik-1 Bill of Materials Test.xlsx";
    public String Requirements = "C://Users//sepeters.ADDS//Desktop//Conceptual Design Files//CCP UPEI SpudNik-1 Master Requirements List.xlsx";
    public String FileLocationFlyWheel = "C://Users//sepeters.ADDS//Desktop//Conceptual Design Files//Command and Control//CC Analysis//Flywheel analysis.xlsx";
    public String FileLocationFlyWheelCAD = "C://Users//sepeters.ADDS//Desktop//Conceptual Design Files//Command and Control//CC CAD//Flywheel design.SLDPRT";
    public String FileLocationMagnetorquers = "C://Users//sepeters.ADDS//Desktop//Conceptual Design Files//Command and Control//CC Analysis//IMU & Magnetometer Options.xlsx";
    public String MagnetorquerCAD = "C://Users//sepeters.ADDS//Desktop//Conceptual Design Files//Command and Control//CC CAD//Magnetorquer.SLDPRT";
    public String FileLocationReactionWheel = "C://Users//sepeters.ADDS//Desktop//Conceptual Design Files//Command and Control//CC Analysis//Reaction_Wheel_Master_V003.xlsx";
    public String ReactionWheelCAD = "C://Users//sepeters.ADDS//Desktop//Conceptual Design Files//Command and Control//CC CAD//Reaction wheel.sldprt";
    public String FileLocationPowerBudget = "C://Users//sepeters.ADDS//Desktop//Conceptual Design Files//Power//Power Analysis//Power Budget.xlsx";
    public String ControlSimulationFileLocation = "C://Users//sepeters.ADDS//Desktop//Conceptual Design Files//Command and Control//CC Analysis//CubeSat Simulink Simulation//CubeSat_Simulation_V2.slx";
    public String SolarPanelsCAD = "C://Users//sepeters.ADDS//Desktop//Conceptual Design Files//Subset of Parts//Power Components//16 cell array.SLDPRT";
    public String BatteryCAD = "C://Users//sepeters.ADDS//Desktop//Conceptual Design Files//Subset of Parts//Power Components//ClydeSpace 20 Wh battery part.SLDPRT";
    */

    public void ReadUserPath(ActionEvent event)
    {
        UserPath = UserPathField.getText();

        System.out.println(""+ UserPath);

        UserPath = "C://Users//steph//OneDrive//Desktop";
    }

    public String MSIPathtoConceptualDesignFiles= UserPath + "C://Users//steph//OneDrive//Desktop";
    // This is for the MSI
    public String SpudNik1Assembly = UserPath + "//Conceptual Design Files//Subset of Parts//Current SpudNik-1 Assembly.SLDASM";
    public String FileLocationThermal = UserPath + "//Conceptual Design Files//Thermal Study v_03.xlsx";
    public String FileLocationPointingBudget = UserPath + "//Conceptual Design Files//SpudNik-1 Pointing Budget.xlsx";
    public String BOM = UserPath + "//Conceptual Design Files//SpudNik-1 Bill of Materials V4.0.xlsx";
    public String Requirements = UserPath + "//Conceptual Design Files//CCP UPEI SpudNik-1 Master Requirements List.xlsx";

    public String FileLocationFlyWheel = UserPath + "//Conceptual Design Files//Command and Control//CC Analysis//Flywheel analysis.xlsx";
    public String FileLocationFlyWheelCAD = UserPath + "//Conceptual Design Files//Command and Control//CC CAD//Flywheel design.SLDPRT";
    public String FileLocationMagnetorquers = UserPath + "//Conceptual Design Files//Command and Control//CC Analysis//IMU & Magnetometer Options.xlsx";
    public String MagnetorquerCAD = UserPath + "//Conceptual Design Files//Command and Control//CC CAD//Magnetorquer.SLDPRT";
    public String FileLocationReactionWheel = UserPath + "//Conceptual Design Files//Command and Control//CC Analysis//Reaction_Wheel_Master_V003.xlsx";
    public String ReactionWheelCAD = UserPath + "//Conceptual Design Files//Command and Control//CC CAD//Reaction wheel.sldprt";
    public String ControlSimulationFileLocation = UserPath + "//Conceptual Design Files//Command and Control//CC Analysis//CubeSat Simulink Simulation//CubeSat_Simulation_V2.slx";
    public String FileLocationMagnetorquerOptimization = UserPath + "//Conceptual Design Files//Command and Control//CC Analysis//Magnetorquer//Single_Magnetorquer_Model_with_Feedback_without_COM.slx";
    public String FileLocationMotorSpeed = UserPath + "//Conceptual Design Files//Command and Control//CC Analysis//Motor Speed Torque Testing//Speed_vs_Duty_Cycle_Test.ino";

    public String FileLocationPowerBudget = UserPath + "//Conceptual Design Files//Power//Power Analysis//Power Budget.xlsx";
    public String SolarPanelsCAD = UserPath + "//Conceptual Design Files//Subset of Parts//Power Components//16 cell array.SLDPRT";
    public String BatteryCAD = UserPath + "//Conceptual Design Files//Subset of Parts//Power Components//ClydeSpace 20 Wh battery part.SLDPRT";
    public String FileLocationPowerAnalysis = UserPath + "//Mission Design Tools//CubeSatToolbox//CubeSatToolbox//CubeSat//Power//SolarCellPower.m";
    public String FileLocationBatteryDS = UserPath + "//Conceptual Design Files//Power//Power Reference Datasheets//Circuit Design Resources//Batteries Clyde Space.pdf";
    //public String FileLocationPanelsDS = UserPath + "//Conceptual Design Files//Power//Power Analysis//Power Budget.xlsx";

    public String FileLocationCommunicationsAnalysis = UserPath + "//Conceptual Design Files//NEW SpudNik-1 S band transmitter Link budget uplink UHF Downlink UHF IARU_Link_Model_Rev2.5.5.xlsx";
    public String CommunicationsAntennaGain = UserPath + "//Conceptual Design Files//Antenna Gain.xlsx";
    public String CommunicationsAtmPolLosses = UserPath + "//Conceptual Design Files//Atm and Pol losses.xlsx";
    public String CommunicationsAntennaPattern = UserPath + "//Conceptual Design Files//Antenna Patterns.xlsx";
    public String CommunicationsAntennaPointingLosses = UserPath + "//Conceptual Design Files//Antenna Pointing Losses.xlsx";
    public String CommunicationsSBandAnt = UserPath + "//Conceptual Design Files//HiSPICO S-Band Patch antenna.pdf";
    public String CommunicationsSBandTrans = UserPath + "//Conceptual Design Files//HiSPICO S-Band Datasheet 06.2019.pdf";
    public String CommunicationsUHFAnt = UserPath + "//Conceptual Design Files//UHF Antenna ICD.docx";
    public String CommunicationsUHFTrans = UserPath + "//Conceptual Design Files//EnduroSat UHF Transceiver II Data sheet.pdf";
    public String CommunicationsUHFTransCAD = UserPath + "//Conceptual Design Files//Subset of Parts//Communication//Tranciever Module V2.00.sldasm";
    public String CommunicationsSBandAntCAD = UserPath + "//Conceptual Design Files//Subset of Parts//Communication//HipSiCo sband Antenna.SLDPRT";
    public String CommunicationsSBandTransCAD = UserPath + "//Conceptual Design Files//Subset of Parts//Communication//HiPSICO S-Band Transmitter.sldasm";
    public String CommunicationsUHFAntCAD = UserPath + "//Conceptual Design Files//Subset of Parts//Communication//Communications system assem//Antenna Release//Total Assembly.sldasm";

    public String FileLocationMassBudget = UserPath + "//Conceptual Design Files//SpudNik-1 Mass Budget.xlsx";
    public String RailsCAD = UserPath + "//Conceptual Design Files//Subset of Parts//Structure_Payload//CubeSat Assembly V2.00//Rails V3.00.SLDPRT";
    public String FileLocationAL6061 = UserPath + "//Conceptual Design Files//Aluminum alloy 6061 prop.xlsx";

    public String FileLocationRailDrawing = UserPath + "//Conceptual Design Files//Structure and Payload//CAD//CAD Drawings//Drawing files//Rail - A4.SLDDRW";
    public String FileLocationPayloadCAD = UserPath + "//Conceptual Design Files//Subset of Parts//Structure_Payload//Payload V3.0//CubeSat Optical Payload COP-00-00-000.sldasm";
    public String FileLocationImageSensorCAD = UserPath + "//Conceptual Design Files//Subset of Parts//Structure_Payload//Payload V3.0//Python 5000 sensor.sldasm";
    public String FileLocationCharacteristics = UserPath + "//Conceptual Design Files//Imaging Characteristics V2.xlsx";
    public String FileLocationImageSensoroRgan = UserPath + "//Conceptual Design Files//Power//Power Reference Datasheets//NOIP1SN5000A-D.pdf";
    public String FileLocationPayloadAlignment = UserPath + "//Conceptual Design Files//Payload Alignment.png";
    public String FileLocationSpotSize = UserPath + "//Conceptual Design Files//Spot Size Image.png";
    //-------------------------COLOUR CODES:----------------------------------------//
    /*
   Thermal Red: #FF6262
   ADCS Blue: #4EFF3F
   Power Green: #3F8CFF
   Communications: #BA55D3
   OP: #e3ee16
   Structure: #bbadad
   */

    //------------------------ Scene Switchers -------------------------------------------------------//
    public void ThermalScene(ActionEvent event) throws Exception {

        Parent Scene2parent = FXMLLoader.load(getClass().getResource("MainThermalScene.fxml"));
        Scene Scene2 = new Scene(Scene2parent);
        Stage window = (Stage)((Node)event.getSource()).getScene().getWindow();
        window.setScene(Scene2);

    }
    public void BacktoMainMenu(ActionEvent event) throws Exception{

        Parent Scene2parent = FXMLLoader.load(getClass().getResource("MainPageScene.fxml"));
        Scene Scene2 = new Scene(Scene2parent);
        Stage window = (Stage)((Node)event.getSource()).getScene().getWindow();
        window.setScene(Scene2);

    }
    public void ThermalInputs(ActionEvent event) throws Exception{

        Parent Scene2parent = FXMLLoader.load(getClass().getResource("MainThermalScene.fxml"));
        Scene Scene2 = new Scene(Scene2parent);

        Stage window = (Stage)((Node)event.getSource()).getScene().getWindow();

        window.setScene(Scene2);
    }
    public void ControlScene(ActionEvent event) throws Exception{

        Parent Scene2parent = FXMLLoader.load(getClass().getResource("MainControlScene.fxml"));
        Scene Scene2 = new Scene(Scene2parent);

        Stage window = (Stage)((Node)event.getSource()).getScene().getWindow();

        window.setScene(Scene2);
    }
    public void PowerScene(ActionEvent event) throws Exception {

        Parent Scene2parent = FXMLLoader.load(getClass().getResource("MainPowerScene.fxml"));
        Scene Scene2 = new Scene(Scene2parent);

        Stage window = (Stage) ((Node) event.getSource()).getScene().getWindow();

        window.setScene(Scene2);
    }
    public void CommunicationsScene(ActionEvent event) throws Exception{

        Parent Scene2parent = FXMLLoader.load(getClass().getResource("MainCommunicationsScene.fxml"));
        Scene Scene2 = new Scene(Scene2parent);

        Stage window = (Stage)((Node)event.getSource()).getScene().getWindow();

        window.setScene(Scene2);
    }
    public void StructureScene(ActionEvent event) throws Exception{

        Parent Scene2parent = FXMLLoader.load(getClass().getResource("MainStructureScene.fxml"));
        Scene Scene2 = new Scene(Scene2parent);

        Stage window = (Stage)((Node)event.getSource()).getScene().getWindow();

        window.setScene(Scene2);
    }
    public void OpticalPayloadScene(ActionEvent event) throws Exception{

        Parent Scene2parent = FXMLLoader.load(getClass().getResource("OP.fxml"));
        Scene Scene2 = new Scene(Scene2parent);

        Stage window = (Stage)((Node)event.getSource()).getScene().getWindow();

        window.setScene(Scene2);
    }
    //---------------------------------- Common to All -----------------------------------//
    public void OrbitalParameters(ActionEvent event) throws Exception{
        MainThermalTextArea.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + FileLocationThermal));
        Sheet sheet1 = workbook1.getSheetAt(0);
        FileInputStream file = new FileInputStream("" + FileLocationThermal); //Inputs File

        MainThermalTextArea.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("B10:H22");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        System.out.println(firstRow + ("\n") + lastRow +("\n") + firstColumn + ("\n") +lastColumn);


        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) {
            InputsArray[ArrayVar]="";
            // System.out.println(ArrayVar);
            MainThermalTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.20s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.20s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            MainThermalTextArea.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
    }
    public void CADAssembly(ActionEvent event){
        try {
            Desktop desktop = null;
            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }

            desktop.open(new File("" + SpudNik1Assembly));
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    }
    public void BOM(ActionEvent event){
        try {
            Desktop desktop = null;
            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }

            desktop.open(new File("" + BOM));
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    }
    public void Requirements(ActionEvent event){
        try {
            Desktop desktop = null;
            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }

            desktop.open(new File("" + Requirements));
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    }
    //------------------------ THERMAL -------------------------------------------------------//
    public void ThermalPhysicalPhenomena(ActionEvent event) throws Exception{
        MainThermalTextArea.setText(" The Thermal Subsystem seeks to provide a preliminary estimate of the satellite thermal environment based on worst case hot and cold " +
                "scenarios. " +
                "The 'Incident Flux' Input provides data \n regarding intensity and amount of incident radiation, and is used by the 'Energy Exposed' Physical Effect calculations " +
                "regarding external heat input. These electromagnetic waves" +
                " travel\n through the vacuum of space and provide thermal energy to the satellite, total heat transfer also depends on the surface absorptivity and emissivity.\n\n" +
                " The 'Power Input' provides details on the rate of energy consumption" +
                ", and is used by the 'Internal Energy Generated' Physical Effect to estimate heat given off due to resistance within \n power consuming components. " +
                "The 'Orbital Parameters' Input provides satellite orbital period and position data, affecting time within direct sunlight" +
                " and distances.\n\n The 'Geometric Envelope' oRgan provides a surface area value for calculating " +
                "incident radiation. The 'Material Properties' oRgan lists thermal conductivities for the preliminary\n thermal interface materials, with a higher " +
                "conductivity allowing for better conductive heat transfer. ");

    }
    public void ThermalMinMaxTemperature (ActionEvent event) throws Exception{
        MainThermalTextArea.setText(" Current temperature estimates are between -15 and 47 Degrees Celcius");

    }
    public void EnergyExposedHot(ActionEvent event) throws Exception {
        MainThermalTextArea.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + FileLocationThermal));
        Sheet sheet1 = workbook1.getSheetAt(2);
        FileInputStream file = new FileInputStream("" + FileLocationThermal); //Inputs File

        XSSFWorkbook workbook = new XSSFWorkbook(file); //Apache POI code. XSSF means the newer Excel versions, HSSF were the older ones

        MainThermalTextArea.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("I14:O21");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) {

            InputsArray[ArrayVar]=""; //Gets rid of null being printed.
            MainThermalTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.20s", NumericCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.20s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            MainThermalTextArea.appendText(InputsArray[ArrayVar] + "              ");
            ArrayVar=ArrayVar+1;
        }
    }
    public void ExposedAreaHot(ActionEvent event) throws Exception {
        MainThermalTextArea.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + FileLocationThermal));
        Sheet sheet1 = workbook1.getSheetAt(2);
        FileInputStream file = new FileInputStream("" + FileLocationThermal); //Inputs File

        XSSFWorkbook workbook = new XSSFWorkbook(file); //Apache POI code. XSSF means the newer Excel versions, HSSF were the older ones
        XSSFSheet sheet = workbook.getSheetAt(0);
        Cell cell = null;
        BufferedReader reader = new BufferedReader(new InputStreamReader(System.in)); // Needed for reading user input

        MainThermalTextArea.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("B45:G64");  //B45:G64 I14:O20

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        System.out.println(firstRow + ("\n") + lastRow +("\n") + firstColumn + ("\n") +lastColumn);


        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) {
            InputsArray[ArrayVar]="";
            // System.out.println(ArrayVar);
            MainThermalTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.100s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.100s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.100s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            MainThermalTextArea.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
    }
    public void InternalEnergyGenerateHot(ActionEvent event) throws Exception{
        MainThermalTextArea.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + FileLocationThermal));
        Sheet sheet1 = workbook1.getSheetAt(2);

        MainThermalTextArea.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("I31:P34");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        System.out.println(firstRow + ("\n") + lastRow +("\n") + firstColumn + ("\n") +lastColumn);

        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) { //Outputting properly formatted Excel Data
            InputsArray[ArrayVar]="";

            MainThermalTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {

                Row row = sheet1.getRow(r);

                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.20s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.20s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            MainThermalTextArea.appendText(InputsArray[ArrayVar] + "              "); //Adding to rows
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
    }
    public void EnergyExposedCold(ActionEvent event) throws Exception{
        MainThermalTextArea.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + FileLocationThermal));
        Sheet sheet1 = workbook1.getSheetAt(2);
        FileInputStream file = new FileInputStream("" + FileLocationThermal); //Inputs File

        XSSFWorkbook workbook = new XSSFWorkbook(file); //Apache POI code. XSSF means the newer Excel versions, HSSF were the older ones
        XSSFSheet sheet = workbook.getSheetAt(0);
        Cell cell = null;
        BufferedReader reader = new BufferedReader(new InputStreamReader(System.in)); // Needed for reading user input

        MainThermalTextArea.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("I5:O11");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        System.out.println(firstRow + ("\n") + lastRow +("\n") + firstColumn + ("\n") +lastColumn);


        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) {
            InputsArray[ArrayVar]="";
            // System.out.println(ArrayVar);
            MainThermalTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.20s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.20s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            MainThermalTextArea.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
    }
    public void ExposedAreaCold(ActionEvent event) throws Exception {
        MainThermalTextArea.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + FileLocationThermal));
        Sheet sheet1 = workbook1.getSheetAt(2);
        FileInputStream file = new FileInputStream("" + FileLocationThermal); //Inputs File

        XSSFWorkbook workbook = new XSSFWorkbook(file); //Apache POI code. XSSF means the newer Excel versions, HSSF were the older ones
        XSSFSheet sheet = workbook.getSheetAt(0);
        Cell cell = null;
        BufferedReader reader = new BufferedReader(new InputStreamReader(System.in)); // Needed for reading user input

        MainThermalTextArea.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("B3:G23");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        System.out.println(firstRow + ("\n") + lastRow +("\n") + firstColumn + ("\n") +lastColumn);


        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) {
            InputsArray[ArrayVar]="";
            // System.out.println(ArrayVar);
            MainThermalTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.20s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.20s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            MainThermalTextArea.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
    }
    public void InternalEnergyGeneratedCold(ActionEvent event) throws Exception{
        MainThermalTextArea.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + FileLocationThermal));
        Sheet sheet1 = workbook1.getSheetAt(2);
        FileInputStream file = new FileInputStream("" + FileLocationThermal); //Inputs File

        XSSFWorkbook workbook = new XSSFWorkbook(file); //Apache POI code. XSSF means the newer Excel versions, HSSF were the older ones
        XSSFSheet sheet = workbook.getSheetAt(0);
        Cell cell = null;
        BufferedReader reader = new BufferedReader(new InputStreamReader(System.in)); // Needed for reading user input

        MainThermalTextArea.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("I25:P28");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        System.out.println(firstRow + ("\n") + lastRow +("\n") + firstColumn + ("\n") +lastColumn);


        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) {
            InputsArray[ArrayVar]="";
            // System.out.println(ArrayVar);
            MainThermalTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.20s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.20s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            MainThermalTextArea.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
    }
    public void ThermalExcel(ActionEvent event) throws Exception{

        try {
            Desktop desktop = null;
            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }

            desktop.open(new File("" + FileLocationThermal));
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    }
    public void ThermalMaterialProperties(ActionEvent event) throws Exception {
        MainThermalTextArea.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + FileLocationThermal));
        Sheet sheet1 = workbook1.getSheetAt(4);
        FileInputStream file = new FileInputStream("" + FileLocationThermal); //Inputs File

        XSSFWorkbook workbook = new XSSFWorkbook(file); //Apache POI code. XSSF means the newer Excel versions, HSSF were the older ones
        XSSFSheet sheet = workbook.getSheetAt(0);
        Cell cell = null;
        BufferedReader reader = new BufferedReader(new InputStreamReader(System.in)); // Needed for reading user input

        MainThermalTextArea.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("B2:K5");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        System.out.println(firstRow + ("\n") + lastRow +("\n") + firstColumn + ("\n") +lastColumn);


        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) {
            InputsArray[ArrayVar]="";
            // System.out.println(ArrayVar);
            MainThermalTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.20s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.20s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            MainThermalTextArea.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
    }
    public void ThermalPowerInput(ActionEvent event) throws Exception{
        MainThermalTextArea.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + FileLocationPowerBudget));
        Sheet sheet1 = workbook1.getSheetAt(0);
        FileInputStream file = new FileInputStream("" + FileLocationPowerBudget); //Inputs File

        MainThermalTextArea.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("B2:H2");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        System.out.println(firstRow + ("\n") + lastRow +("\n") + firstColumn + ("\n") +lastColumn);


        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) { //This is to print the title row
            InputsArray[ArrayVar]="";
            // System.out.println(ArrayVar);
            MainThermalTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.5s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.5s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            MainThermalTextArea.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
        cellRangeAddress = CellRangeAddress.valueOf("B16:H17");

        firstRow = cellRangeAddress.getFirstRow();
        lastRow = cellRangeAddress.getLastRow();
        firstColumn = cellRangeAddress.getFirstColumn();
        lastColumn = cellRangeAddress.getLastColumn();
        for (int r = firstRow; r <= lastRow; r++) {
            InputsArray[ArrayVar]="";
            // System.out.println(ArrayVar);
            MainThermalTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.5s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.5s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            MainThermalTextArea.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
    }
    public void IncidentFlux(ActionEvent event) throws Exception{
        MainThermalTextArea.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + FileLocationThermal));
        Sheet sheet1 = workbook1.getSheetAt(1);
        FileInputStream file = new FileInputStream("" + FileLocationThermal); //Inputs File

        MainThermalTextArea.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("B2:G21");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        System.out.println(firstRow + ("\n") + lastRow +("\n") + firstColumn + ("\n") +lastColumn);


        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) {
            InputsArray[ArrayVar]="";
            // System.out.println(ArrayVar);
            MainThermalTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.5s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.5s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            MainThermalTextArea.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
    }
    public void ThermalRequirements(ActionEvent event) throws Exception {
        MainThermalTextArea.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + Requirements));
        Sheet sheet1 = workbook1.getSheetAt(0);
        FileInputStream file = new FileInputStream("" + Requirements); //Inputs File

        XSSFWorkbook workbook = new XSSFWorkbook(file); //Apache POI code. XSSF means the newer Excel versions, HSSF were the older ones
        XSSFSheet sheet = workbook.getSheetAt(0);
        Cell cell = null;

        MainThermalTextArea.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("C1:E1");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) {

            InputsArray[ArrayVar]=""; //Gets rid of null being printed.
            MainThermalTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.20s", NumericCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.20s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-30.200s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            MainThermalTextArea.appendText(InputsArray[ArrayVar] + "              ");
            ArrayVar=ArrayVar+1;
        }
        cellRangeAddress = CellRangeAddress.valueOf("C109:E113");

        firstRow = cellRangeAddress.getFirstRow();
        lastRow = cellRangeAddress.getLastRow();
        firstColumn = cellRangeAddress.getFirstColumn();
        lastColumn = cellRangeAddress.getLastColumn();

        InputsArray = new String[100];
        ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) {

            InputsArray[ArrayVar]=""; //Gets rid of null being printed.
            MainThermalTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.20s", NumericCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.20s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-30.200s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            MainThermalTextArea.appendText(InputsArray[ArrayVar] + "              ");
            ArrayVar=ArrayVar+1;
        }
    }
    public void ThermalGeoEnvelope(ActionEvent event) throws Exception {
        MainThermalTextArea.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + FileLocationThermal));
        Sheet sheet1 = workbook1.getSheetAt(2);
        FileInputStream file = new FileInputStream("" + FileLocationThermal); //Inputs File

        XSSFWorkbook workbook = new XSSFWorkbook(file); //Apache POI code. XSSF means the newer Excel versions, HSSF were the older ones

        MainThermalTextArea.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("B28:E44");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) {

            InputsArray[ArrayVar]=""; //Gets rid of null being printed.
            MainThermalTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.20s", NumericCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.20s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.100s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            MainThermalTextArea.appendText(InputsArray[ArrayVar] + "              ");
            ArrayVar=ArrayVar+1;
        }
        cellRangeAddress = CellRangeAddress.valueOf("C109:H113");

        firstRow = cellRangeAddress.getFirstRow();
        lastRow = cellRangeAddress.getLastRow();
        firstColumn = cellRangeAddress.getFirstColumn();
        lastColumn = cellRangeAddress.getLastColumn();

        InputsArray = new String[100];
        ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) {

            InputsArray[ArrayVar]=""; //Gets rid of null being printed.
            MainThermalTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-100.100s", NumericCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-100.100s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-100.70s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            MainThermalTextArea.appendText(InputsArray[ArrayVar] + "              ");
            ArrayVar=ArrayVar+1;
        }
    }
    //------------------------ CONTROLS -------------------------------------------------------//
    public void ControlPhysicalPhenomena(ActionEvent event) throws Exception{
        ControlText.setText("The Physical Effects are simulations of the CubeSat's control scheme during orbit. The 'Orbital Parameters' input provides the orbital velocities, period etc." +
                "The 'Power Input', from the 'Power' Subsystem, influences the choice\n of components (example: the motor) based on their power consumption. The Flywheels, Reaction Wheels and Magnetourquers oRgans affect the mass, interia" +
                "and applied torque values of the simulation.\n\n When spun, the reaction wheel induces a torque in the CubeSat, causing it to rotate in the opposite direction. The magnetorquers can also induce torques by interacting with" +
                "the earth's magnetic field when they're energized.\n The simulation provides an initial feasibility study of the control scheme, and provides pointing error data to the 'Communications' Subsystem.");
    }
    public void FlyWheels(ActionEvent event) throws Exception{
        ControlText.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + FileLocationFlyWheel));
        Sheet sheet1 = workbook1.getSheetAt(0);
        FileInputStream file = new FileInputStream("" + FileLocationFlyWheel); //Inputs File

        XSSFWorkbook workbook = new XSSFWorkbook(file); //Apache POI code. XSSF means the newer Excel versions, HSSF were the older ones
        ControlText.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("G1:M8");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) {
            InputsArray[ArrayVar]="";
            // System.out.println(ArrayVar);
            ControlText.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.20s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.20s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            ControlText.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
    }
    public void Magnetorquer(ActionEvent event) throws Exception {
        ControlText.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + FileLocationMagnetorquers));
        Sheet sheet1 = workbook1.getSheetAt(0);
        FileInputStream file = new FileInputStream("" + FileLocationMagnetorquers); //Inputs File

        XSSFWorkbook workbook = new XSSFWorkbook(file); //Apache POI code. XSSF means the newer Excel versions, HSSF were the older ones
        ControlText.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("A1:E10");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        String[] InputsArray = new String[100];
        int ArrayVar = 0;

        for (int r = firstRow; r <= lastRow; r++) {
            InputsArray[ArrayVar] = "";
            // System.out.println(ArrayVar);
            ControlText.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-25.25s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar] = InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-25.25s", StringCell);
                                InputsArray[ArrayVar] = InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-25.25s", content);

                        InputsArray[ArrayVar] = InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            ControlText.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar = ArrayVar + 1;
        }
    }
    public void ReactionWheels(ActionEvent event) throws Exception{
        ControlText.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + FileLocationReactionWheel));
        Sheet sheet1 = workbook1.getSheetAt(0);
        FileInputStream file = new FileInputStream("" + FileLocationReactionWheel); //Inputs File

        XSSFWorkbook workbook = new XSSFWorkbook(file); //Apache POI code. XSSF means the newer Excel versions, HSSF were the older ones
        ControlText.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("A1:C23");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) {
            InputsArray[ArrayVar]="";
            // System.out.println(ArrayVar);
            ControlText.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.20s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.20s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            ControlText.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
    }
    public void ControlPowerInput(ActionEvent event) throws Exception{
        ControlText.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + FileLocationPowerBudget));
        Sheet sheet1 = workbook1.getSheetAt(0);
        FileInputStream file = new FileInputStream("" + FileLocationPowerBudget); //Inputs File

        ControlText.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("B2:H2");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        System.out.println(firstRow + ("\n") + lastRow +("\n") + firstColumn + ("\n") +lastColumn);


        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) { //This is to print the title row
            InputsArray[ArrayVar]="";
            // System.out.println(ArrayVar);
            ControlText.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.5s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.5s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            ControlText.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
        cellRangeAddress = CellRangeAddress.valueOf("B8:H13");

        firstRow = cellRangeAddress.getFirstRow();
        lastRow = cellRangeAddress.getLastRow();
        firstColumn = cellRangeAddress.getFirstColumn();
        lastColumn = cellRangeAddress.getLastColumn();
        for (int r = firstRow; r <= lastRow; r++) {
            InputsArray[ArrayVar]="";
            // System.out.println(ArrayVar);
            ControlText.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.5s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.5s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            ControlText.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
    }
    public void FlyWheelCAD(ActionEvent event) throws Exception{
        try {
            Desktop desktop = null;
            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }

            desktop.open(new File("" + FileLocationFlyWheelCAD));
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    }
    public void MagnetorquerCAD(ActionEvent event) throws Exception{
        try {
            Desktop desktop = null;
            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }

            desktop.open(new File("" + MagnetorquerCAD));
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    }
    public void ReactionWheelCAD(ActionEvent event) throws Exception{
        try {
            Desktop desktop = null;
            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }

            desktop.open(new File("" + ReactionWheelCAD));
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    }
    public void ControlSimulation(ActionEvent event) throws Exception{
        try {
            Desktop desktop = null;
            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }

            desktop.open(new File("" + ControlSimulationFileLocation));
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    }
    public void MagnetorquerOptimization(ActionEvent event) throws Exception{
        try {
            Desktop desktop = null;
            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }

            desktop.open(new File("" + FileLocationMagnetorquerOptimization));
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    }
    public void MotorSpeed(ActionEvent event) throws Exception{
        try {
            Desktop desktop = null;
            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }

            desktop.open(new File("" + FileLocationMotorSpeed));
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    }
    public void ControlOrbitalParameters(ActionEvent event) throws Exception{
        ControlText.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + FileLocationThermal));
        Sheet sheet1 = workbook1.getSheetAt(0);
        FileInputStream file = new FileInputStream("" + FileLocationThermal); //Inputs File

        ControlText.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("B10:H22");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        System.out.println(firstRow + ("\n") + lastRow +("\n") + firstColumn + ("\n") +lastColumn);


        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) {
            InputsArray[ArrayVar]="";
            // System.out.println(ArrayVar);
            ControlText.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.20s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.20s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            ControlText.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
    }
    public void ControlRequirements(ActionEvent event) throws Exception {
        ControlText.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + Requirements));
        Sheet sheet1 = workbook1.getSheetAt(0);
        FileInputStream file = new FileInputStream("" + Requirements); //Inputs File

        XSSFWorkbook workbook = new XSSFWorkbook(file); //Apache POI code. XSSF means the newer Excel versions, HSSF were the older ones
        XSSFSheet sheet = workbook.getSheetAt(0);
        Cell cell = null;

        ControlText.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("C1:E1");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) {

            InputsArray[ArrayVar]=""; //Gets rid of null being printed.
            ControlText.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.20s", NumericCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.20s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-30.200s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            ControlText.appendText(InputsArray[ArrayVar] + "              ");
            ArrayVar=ArrayVar+1;
        }
        cellRangeAddress = CellRangeAddress.valueOf("C30:E54");

        firstRow = cellRangeAddress.getFirstRow();
        lastRow = cellRangeAddress.getLastRow();
        firstColumn = cellRangeAddress.getFirstColumn();
        lastColumn = cellRangeAddress.getLastColumn();

        InputsArray = new String[100];
        ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) {

            InputsArray[ArrayVar]=""; //Gets rid of null being printed.
            ControlText.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.20s", NumericCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.20s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-30.200s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            ControlText.appendText(InputsArray[ArrayVar] + "              ");
            ArrayVar=ArrayVar+1;
        }
    }
    public void PointingBudget(ActionEvent event) throws Exception{
        ControlText.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + FileLocationPointingBudget));
        Sheet sheet1 = workbook1.getSheetAt(0);
        FileInputStream file = new FileInputStream("" + FileLocationPointingBudget); //Inputs File

        XSSFWorkbook workbook = new XSSFWorkbook(file); //Apache POI code. XSSF means the newer Excel versions, HSSF were the older ones
        ControlText.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("A1:D13");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) {
            InputsArray[ArrayVar]="";
            // System.out.println(ArrayVar);
            ControlText.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.20s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.20s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            ControlText.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
    }
    public void ControlMassAllocated(ActionEvent event) throws Exception{
        ControlText.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + FileLocationMassBudget));
        Sheet sheet1 = workbook1.getSheetAt(0);
        FileInputStream file = new FileInputStream("" + FileLocationMassBudget); //Inputs File

        ControlText.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("E5:J5");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        System.out.println(firstRow + ("\n") + lastRow +("\n") + firstColumn + ("\n") +lastColumn);


        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) { //This is to print the title row
            InputsArray[ArrayVar]="";
            // System.out.println(ArrayVar);
            ControlText.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.5s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.5s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            ControlText.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
        cellRangeAddress = CellRangeAddress.valueOf("E8:J8");

        firstRow = cellRangeAddress.getFirstRow();
        lastRow = cellRangeAddress.getLastRow();
        firstColumn = cellRangeAddress.getFirstColumn();
        lastColumn = cellRangeAddress.getLastColumn();
        for (int r = firstRow; r <= lastRow; r++) {
            InputsArray[ArrayVar]="";
            // System.out.println(ArrayVar);
            ControlText.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.5s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.5s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            ControlText.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
    }
    //---------------------POWER-------------------------------------------------------------//
    public void PowerPhysicalPhenomena(ActionEvent event) throws Exception{
        MainPowerTextArea.setText(" The Power Subsystem seeks to provide a power budget output to each subsystem to control the selection of electrical energy consumption, based on operational modes\n" +
                "The 'Orbital Parameters' Input includes data regarding the orbital path and velocity of the satellite, which influences radiation exposure and intensity. The 'Solar Flux' Input\n" +
                "contains information specific to the radiation sources, including the Sub and the Earth, critical for calculating power generated. The oRgans dictate both the amount of \n" +
                "electrical storage capcity, and power generation efficiency. The physical effects produce an estimate of orbital power values by taking into account the solar cell efficiency, along with the inputs. ");

    }
    public void PowerTemperature(ActionEvent event) throws Exception{
        MainPowerTextArea.setText(" Current temperature estimates are between -15 and 47 Degrees Celcius");

    }
    public void PowerRequirements(ActionEvent event) throws Exception {
        MainPowerTextArea.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + Requirements));
        Sheet sheet1 = workbook1.getSheetAt(0);
        FileInputStream file = new FileInputStream("" + Requirements); //Inputs File

        XSSFWorkbook workbook = new XSSFWorkbook(file); //Apache POI code. XSSF means the newer Excel versions, HSSF were the older ones
        XSSFSheet sheet = workbook.getSheetAt(0);
        Cell cell = null;

        MainPowerTextArea.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("C1:E1");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) {

            InputsArray[ArrayVar]=""; //Gets rid of null being printed.
            MainPowerTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.20s", NumericCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.20s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-30.200s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            MainPowerTextArea.appendText(InputsArray[ArrayVar] + "              ");
            ArrayVar=ArrayVar+1;
        }
        cellRangeAddress = CellRangeAddress.valueOf("C30:E54");

        firstRow = cellRangeAddress.getFirstRow();
        lastRow = cellRangeAddress.getLastRow();
        firstColumn = cellRangeAddress.getFirstColumn();
        lastColumn = cellRangeAddress.getLastColumn();

        InputsArray = new String[100];
        ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) {

            InputsArray[ArrayVar]=""; //Gets rid of null being printed.
            MainPowerTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.20s", NumericCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.20s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-30.200s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            MainPowerTextArea.appendText(InputsArray[ArrayVar] + "              ");
            ArrayVar=ArrayVar+1;
        }
    }
    public void PowerIncidentFlux(ActionEvent event) throws Exception{
        MainPowerTextArea.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + FileLocationThermal));
        Sheet sheet1 = workbook1.getSheetAt(1);
        FileInputStream file = new FileInputStream("" + FileLocationThermal); //Inputs File

        MainPowerTextArea.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("B2:G21");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        System.out.println(firstRow + ("\n") + lastRow +("\n") + firstColumn + ("\n") +lastColumn);


        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) {
            InputsArray[ArrayVar]="";
            // System.out.println(ArrayVar);
            MainPowerTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.5s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.5s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            MainPowerTextArea.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
    }
    public void SolarPanelCAD(ActionEvent event) throws Exception{
        try {
            Desktop desktop = null;
            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }

            desktop.open(new File("" + SolarPanelsCAD));
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    }
    public void BatteryCAD(ActionEvent event) throws Exception{
        try {
            Desktop desktop = null;
            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }

            desktop.open(new File("" + BatteryCAD));
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    }
    public void PowerAnalysis(ActionEvent event) throws Exception{
        try {
            Desktop desktop = null;
            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }

            desktop.open(new File("" + FileLocationPowerAnalysis));
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    }
    public void PowerOrbitalParameters(ActionEvent event) throws Exception{
        MainPowerTextArea.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + FileLocationThermal));
        Sheet sheet1 = workbook1.getSheetAt(0);
        FileInputStream file = new FileInputStream("" + FileLocationThermal); //Inputs File

        MainPowerTextArea.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("B10:H22");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        System.out.println(firstRow + ("\n") + lastRow +("\n") + firstColumn + ("\n") +lastColumn);


        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) {
            InputsArray[ArrayVar]="";
            // System.out.println(ArrayVar);
            MainPowerTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.20s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.20s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            MainPowerTextArea.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
    }
    public void PowerBudget(ActionEvent event) throws Exception{
        MainPowerTextArea.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + FileLocationThermal));
        Sheet sheet1 = workbook1.getSheetAt(3);
        FileInputStream file = new FileInputStream("" + FileLocationThermal); //Inputs File

        MainPowerTextArea.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("B2:H2");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        System.out.println(firstRow + ("\n") + lastRow +("\n") + firstColumn + ("\n") +lastColumn);


        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) { //This is to print the title row
            InputsArray[ArrayVar]="";
            // System.out.println(ArrayVar);
            MainPowerTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.5s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.5s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            MainPowerTextArea.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
        cellRangeAddress = CellRangeAddress.valueOf("B4:I16");

        firstRow = cellRangeAddress.getFirstRow();
        lastRow = cellRangeAddress.getLastRow();
        firstColumn = cellRangeAddress.getFirstColumn();
        lastColumn = cellRangeAddress.getLastColumn();
        for (int r = firstRow; r <= lastRow; r++) {
            InputsArray[ArrayVar]="";
            // System.out.println(ArrayVar);
            MainPowerTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.5s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.5s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            MainPowerTextArea.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
    }
    public void BatteryDS(ActionEvent event) throws Exception{
        try {
            Desktop desktop = null;
            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }

            desktop.open(new File("" + FileLocationBatteryDS));
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    }
    public void PowerMassAllocated(ActionEvent event) throws Exception{
        MainPowerTextArea.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + FileLocationMassBudget));
        Sheet sheet1 = workbook1.getSheetAt(0);
        FileInputStream file = new FileInputStream("" + FileLocationMassBudget); //Inputs File

        MainPowerTextArea.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("E5:J5");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        System.out.println(firstRow + ("\n") + lastRow +("\n") + firstColumn + ("\n") +lastColumn);


        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) { //This is to print the title row
            InputsArray[ArrayVar]="";
            // System.out.println(ArrayVar);
            MainPowerTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.5s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.5s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            MainPowerTextArea.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
        cellRangeAddress = CellRangeAddress.valueOf("E10:J10");

        firstRow = cellRangeAddress.getFirstRow();
        lastRow = cellRangeAddress.getLastRow();
        firstColumn = cellRangeAddress.getFirstColumn();
        lastColumn = cellRangeAddress.getLastColumn();
        for (int r = firstRow; r <= lastRow; r++) {
            InputsArray[ArrayVar]="";
            // System.out.println(ArrayVar);
            MainPowerTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.5s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.5s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            MainPowerTextArea.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
    }
    /*public void PanelsDS(ActionEvent event) throws Exception{
        try {
            Desktop desktop = null;
            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }

            desktop.open(new File("" + BatteryCAD));
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    } */
    //---------------------------- Communications ------------------------------------//
    public void CommunicationsPhysicalPhenomena(ActionEvent event) throws Exception{
        MainCommunicationsTextArea.setText("The objective of the physical effects are to produce a 'Downlink' and 'Uplink' budget to provide an initial assessment of data transfer capabilities. The losses calculated are due to atmospheric gas molecules \n and misaligned polarization fields." +
                "The antenna gain is akin to the power of the transmission, and depends on a wide variety of factors including efficiency, geometry and frequency. Finally, the antenna \npatterns are calculated primarily via the gain, antenna type and the polarization field." +
                "\n\nThe 'Pointing Error' Input is provided by the 'Controls' Subsystem and is needed for the pointing loss calculation. The 'Orbital Parameters' input provides satellite velocity and LOS information." +
                "The 'Image\n Data' input is provided by the 'Payload' Subsystem, and represents the required data file to be downlinked. The 'Power Input', provided by the 'Power' Subsystem, influences " +
                "the component selections, in \nconjunction with the 'Allocated Mass', provided by the 'Structure' Subsystem. The oRgans are all data sheets, which contain essential information for" +
                " each component.");

    }
    public void CommunicationsPowerInput(ActionEvent event) throws Exception{
        MainCommunicationsTextArea.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + FileLocationPowerBudget));
        Sheet sheet1 = workbook1.getSheetAt(0);
        FileInputStream file = new FileInputStream("" + FileLocationPowerBudget); //Inputs File

        MainCommunicationsTextArea.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("B2:H2");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) { //This is to print the title row
            InputsArray[ArrayVar]="";

            MainCommunicationsTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.5s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.5s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            MainCommunicationsTextArea.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
        cellRangeAddress = CellRangeAddress.valueOf("B6:H6");

        firstRow = cellRangeAddress.getFirstRow();
        lastRow = cellRangeAddress.getLastRow();
        firstColumn = cellRangeAddress.getFirstColumn();
        lastColumn = cellRangeAddress.getLastColumn();
        for (int r = firstRow; r <= lastRow; r++) {
            InputsArray[ArrayVar]="";
            // System.out.println(ArrayVar);
            MainCommunicationsTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.5s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.5s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            MainCommunicationsTextArea.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
    }
    public void CommunicationsDownlinkBudget(ActionEvent event) throws Exception{
        MainCommunicationsTextArea.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + FileLocationCommunicationsAnalysis));
        Sheet sheet1 = workbook1.getSheetAt(12);
        FileInputStream file = new FileInputStream("" + FileLocationCommunicationsAnalysis); //Inputs File

        XSSFWorkbook workbook = new XSSFWorkbook(file); //Apache POI code. XSSF means the newer Excel versions, HSSF were the older ones
        MainCommunicationsTextArea.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("A2:D41");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) {
            InputsArray[ArrayVar]="";
            // System.out.println(ArrayVar);
            MainCommunicationsTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.20s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.20s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            MainCommunicationsTextArea.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
    }
    public void CommunicationsUplinkBudget(ActionEvent event) throws Exception{
        MainCommunicationsTextArea.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + FileLocationCommunicationsAnalysis));
        Sheet sheet1 = workbook1.getSheetAt(11);
        FileInputStream file = new FileInputStream("" + FileLocationCommunicationsAnalysis); //Inputs File

        XSSFWorkbook workbook = new XSSFWorkbook(file); //Apache POI code. XSSF means the newer Excel versions, HSSF were the older ones
        MainCommunicationsTextArea.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("A2:D41");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) {
            InputsArray[ArrayVar]="";
            // System.out.println(ArrayVar);
            MainCommunicationsTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.20s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.20s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            MainCommunicationsTextArea.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
    }
    public void CommunicationsOrbitalParameters(ActionEvent event) throws Exception{
        MainCommunicationsTextArea.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + FileLocationThermal));
        Sheet sheet1 = workbook1.getSheetAt(0);
        FileInputStream file = new FileInputStream("" + FileLocationThermal); //Inputs File

        MainCommunicationsTextArea.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("B10:H22");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        System.out.println(firstRow + ("\n") + lastRow +("\n") + firstColumn + ("\n") +lastColumn);


        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) {
            InputsArray[ArrayVar]="";
            // System.out.println(ArrayVar);
            MainCommunicationsTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.20s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.20s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            MainCommunicationsTextArea.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
    }
    public void CommunicationsAntennaGain(ActionEvent event) throws Exception{
        try {
            Desktop desktop = null;
            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }

            desktop.open(new File("" + CommunicationsAntennaGain));
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    }
    public void CommunicationsAtmospherePolarizationLosses(ActionEvent event) throws Exception{
        try {
            Desktop desktop = null;
            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }

            desktop.open(new File("" + CommunicationsAtmPolLosses));
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    }
    public void CommunicationsAntennaPatterns(ActionEvent event) throws Exception{
        try {
            Desktop desktop = null;
            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }

            desktop.open(new File("" + CommunicationsAntennaPattern));
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    }
    public void CommunicationsAntennaPointingLosses(ActionEvent event) throws Exception{
        try {
            Desktop desktop = null;
            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }

            desktop.open(new File("" + CommunicationsAntennaPointingLosses));
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    }
    public void CommunicationsSBandAnt(ActionEvent event) throws Exception{
        try {
            Desktop desktop = null;
            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }

            desktop.open(new File("" + CommunicationsSBandAnt));
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    }
    public void CommunicationsSBandTrans(ActionEvent event) throws Exception{
        try {
            Desktop desktop = null;
            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }

            desktop.open(new File("" + CommunicationsSBandTrans));
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    }
    public void CommunicationsUHFAnt(ActionEvent event) throws Exception{
        try {
            Desktop desktop = null;
            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }

            desktop.open(new File("" + CommunicationsUHFAnt));
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    }
    public void CommunicationsUHFTrans(ActionEvent event) throws Exception{
        try {
            Desktop desktop = null;
            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }

            desktop.open(new File("" + CommunicationsUHFTrans));
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    }
    public void CommunicationsSBandAntCAD(ActionEvent event) throws Exception{
        try {
            Desktop desktop = null;
            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }

            desktop.open(new File("" + CommunicationsSBandAntCAD));
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    }
    public void CommunicationsSBandTransCAD(ActionEvent event) throws Exception{
        try {
            Desktop desktop = null;
            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }

            desktop.open(new File("" + CommunicationsSBandTransCAD));
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    }
    public void CommunicationsUHFTransCAD(ActionEvent event) throws Exception{
        try {
            Desktop desktop = null;
            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }

            desktop.open(new File("" + CommunicationsUHFTransCAD));
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    }
    public void CommunicationsUHFAntCAD(ActionEvent event) throws Exception{
        try {
            Desktop desktop = null;
            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }

            desktop.open(new File("" + CommunicationsUHFAntCAD));
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    } //FileLocationPointingBudget  MainCommunicationsTextArea
    public void CommunicationsPointingError(ActionEvent event) throws Exception{
        MainCommunicationsTextArea.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + FileLocationThermal));
        Sheet sheet1 = workbook1.getSheetAt(5);
        FileInputStream file = new FileInputStream("" + FileLocationThermal); //Inputs File

        XSSFWorkbook workbook = new XSSFWorkbook(file); //Apache POI code. XSSF means the newer Excel versions, HSSF were the older ones
        XSSFSheet sheet = workbook.getSheetAt(5);
        Cell cell = null;
        BufferedReader reader = new BufferedReader(new InputStreamReader(System.in)); // Needed for reading user input

        MainCommunicationsTextArea.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("A11:E12");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        System.out.println(firstRow + ("\n") + lastRow +("\n") + firstColumn + ("\n") +lastColumn);


        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) {
            InputsArray[ArrayVar]="";
            // System.out.println(ArrayVar);
            MainCommunicationsTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.20s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.20s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            MainCommunicationsTextArea.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }

    }
    public void CommunicationsMassAllocated(ActionEvent event) throws Exception{
        MainCommunicationsTextArea.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + FileLocationMassBudget));
        Sheet sheet1 = workbook1.getSheetAt(0);
        FileInputStream file = new FileInputStream("" + FileLocationMassBudget); //Inputs File

        MainCommunicationsTextArea.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("E5:J5");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        System.out.println(firstRow + ("\n") + lastRow +("\n") + firstColumn + ("\n") +lastColumn);


        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) { //This is to print the title row
            InputsArray[ArrayVar]="";
            // System.out.println(ArrayVar);
            MainCommunicationsTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.5s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.5s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            MainCommunicationsTextArea.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
        cellRangeAddress = CellRangeAddress.valueOf("E9:J9");

        firstRow = cellRangeAddress.getFirstRow();
        lastRow = cellRangeAddress.getLastRow();
        firstColumn = cellRangeAddress.getFirstColumn();
        lastColumn = cellRangeAddress.getLastColumn();
        for (int r = firstRow; r <= lastRow; r++) {
            InputsArray[ArrayVar]="";
            // System.out.println(ArrayVar);
            MainCommunicationsTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.5s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.5s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            MainCommunicationsTextArea.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
    }
    public void CommunicationsImageFileSize(ActionEvent event) throws Exception {
        MainCommunicationsTextArea.setText("After compression file size: 144kB");

    }
    //---------------------------- Structure ------------------------------------//
    public void StructurePhysicalPhenomena(ActionEvent event) throws Exception{
        MainStructureTextArea.setText("A load and vibration simulation is required to satisfy structural integrity requirements. The 'Z-axis Compressive Force' Input is the required force to be " +
                "withstood by each individual rail.\n The 'Launch Vibration' Input represents the various frequencies that need to be withstood by the assembly during launch. " +
                "The oRgans include AL 6061 mechanical\n and electrical properties, as well as the specific geometries of the Rails themselves. The output of this subsystem," +
                " based on these simulations, is a preliminary mass budget.");
    }
    public void CompressiveForce(ActionEvent event) throws Exception {
        MainStructureTextArea.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + Requirements));
        Sheet sheet1 = workbook1.getSheetAt(0);
        FileInputStream file = new FileInputStream("" + Requirements); //Inputs File

        XSSFWorkbook workbook = new XSSFWorkbook(file); //Apache POI code. XSSF means the newer Excel versions, HSSF were the older ones
        XSSFSheet sheet = workbook.getSheetAt(0);
        Cell cell = null;

        MainStructureTextArea.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("B99:E99");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) {

            InputsArray[ArrayVar]=""; //Gets rid of null being printed.
            MainStructureTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {

                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {

                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.20s", NumericCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.20s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-30.200s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            MainStructureTextArea.appendText(InputsArray[ArrayVar] + "              ");
            ArrayVar=ArrayVar+1;
        }

    }
    public void VibrationInput(ActionEvent event) throws Exception {
        MainStructureTextArea.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + Requirements));
        Sheet sheet1 = workbook1.getSheetAt(0);
        FileInputStream file = new FileInputStream("" + Requirements); //Inputs File

        XSSFWorkbook workbook = new XSSFWorkbook(file); //Apache POI code. XSSF means the newer Excel versions, HSSF were the older ones
        XSSFSheet sheet = workbook.getSheetAt(0);
        Cell cell = null;

        MainStructureTextArea.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("C12:I12");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) {

            InputsArray[ArrayVar]=""; //Gets rid of null being printed.
            MainStructureTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.20s", NumericCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.20s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-30.200s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            MainStructureTextArea.appendText(InputsArray[ArrayVar] + "              ");
            ArrayVar=ArrayVar+1;
        }

    }
    public void RailsCAD(ActionEvent event) throws Exception{
        try {
            Desktop desktop = null;
            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }

            desktop.open(new File("" + RailsCAD));
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    }
    public void MassBudget(ActionEvent event) throws Exception {
        MainStructureTextArea.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + FileLocationMassBudget));
        Sheet sheet1 = workbook1.getSheetAt(0);
        FileInputStream file = new FileInputStream("" + FileLocationMassBudget); //Inputs File

        XSSFWorkbook workbook = new XSSFWorkbook(file); //Apache POI code. XSSF means the newer Excel versions, HSSF were the older ones
        MainStructureTextArea.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("A1:L21");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) {
            InputsArray[ArrayVar]="";
            // System.out.println(ArrayVar);
            MainStructureTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.20s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.20s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            MainStructureTextArea.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
    }
    public void AL6061oRgan(ActionEvent event) throws Exception {
        MainStructureTextArea.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + FileLocationAL6061));
        Sheet sheet1 = workbook1.getSheetAt(2);
        FileInputStream file = new FileInputStream("" + FileLocationAL6061); //Inputs File

        XSSFWorkbook workbook = new XSSFWorkbook(file); //Apache POI code. XSSF means the newer Excel versions, HSSF were the older ones

        MainStructureTextArea.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("B4:O5");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) {

            InputsArray[ArrayVar]=""; //Gets rid of null being printed.
            MainStructureTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.20s", NumericCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.20s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            MainStructureTextArea.appendText(InputsArray[ArrayVar] + "              ");
            ArrayVar=ArrayVar+1;
        }
    }
    //-------------------- Payload -------------------------------------//
    public void PayloadPhysicalPhenomena(ActionEvent event) throws Exception{
        OPTextArea.setText("The Optical Payload subsystem's objective is to design for 5-7 metre resolution, and to deliver a preliminary file size transfer requirement to the Communications subsystem. " +
                "The 'File Size'\n physical effect calculates the approximate image file size based on the pixel density, image quality and compression algorithms. The 'Imaging Characteristics' PE produces " +
                "a layout concept\n for the payload to attain the required resolution based on the geoemtries of the mirrors and lenses, as well as the focal length. The 'Power Input' is used for " +
                "determining a suitable image \nsensor. The 'Image Sensor', in conjunction with its oRgan, provides information on pixel density, image size and quality. 'Tolerances' are " +
                "used for creating\n the geometric layout of the payload.");
    }
    public void OPPowerInput(ActionEvent event) throws Exception{
        OPTextArea.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + FileLocationPowerBudget));
        Sheet sheet1 = workbook1.getSheetAt(0);
        FileInputStream file = new FileInputStream("" + FileLocationPowerBudget); //Inputs File

        OPTextArea.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("B2:H2");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        System.out.println(firstRow + ("\n") + lastRow +("\n") + firstColumn + ("\n") +lastColumn);


        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) { //This is to print the title row
            InputsArray[ArrayVar]="";
            // System.out.println(ArrayVar);
            OPTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.5s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.5s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            OPTextArea.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
        cellRangeAddress = CellRangeAddress.valueOf("B4:H4");

        firstRow = cellRangeAddress.getFirstRow();
        lastRow = cellRangeAddress.getLastRow();
        firstColumn = cellRangeAddress.getFirstColumn();
        lastColumn = cellRangeAddress.getLastColumn();
        for (int r = firstRow; r <= lastRow; r++) {
            InputsArray[ArrayVar]="";
            // System.out.println(ArrayVar);
            OPTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.5s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.5s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            OPTextArea.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
    }
    public void PayloadCAD(ActionEvent event) throws Exception{
        try {
            Desktop desktop = null;
            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }

            desktop.open(new File("" + FileLocationPayloadCAD));
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    }
    public void ImagingCharacteristics(ActionEvent event) throws Exception{
        try {
            Desktop desktop = null;
            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }

            desktop.open(new File("" + FileLocationCharacteristics));
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    }
    public void RailDrawing(ActionEvent event) throws Exception{
        try {
            Desktop desktop = null;
            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }

            desktop.open(new File("" + FileLocationRailDrawing));
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    }
    public void ImageSensorCAD(ActionEvent event) throws Exception{
        try {
            Desktop desktop = null;
            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }

            desktop.open(new File("" + FileLocationImageSensorCAD));
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    }
    public void PayloadMassAllocated(ActionEvent event) throws Exception{
        OPTextArea.setText(""); //This clears the text area

        DataFormatter formatter = new DataFormatter();
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("" + FileLocationMassBudget));
        Sheet sheet1 = workbook1.getSheetAt(0);
        FileInputStream file = new FileInputStream("" + FileLocationMassBudget); //Inputs File

        OPTextArea.setStyle("-fx-font-family: monospace, 20px");

        CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf("E5:J5");

        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        System.out.println(firstRow + ("\n") + lastRow +("\n") + firstColumn + ("\n") +lastColumn);


        String[] InputsArray = new String[100];
        int ArrayVar=0;

        for (int r = firstRow; r <= lastRow; r++) { //This is to print the title row
            InputsArray[ArrayVar]="";
            // System.out.println(ArrayVar);
            OPTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.5s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.5s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            OPTextArea.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
        cellRangeAddress = CellRangeAddress.valueOf("E7:J7");

        firstRow = cellRangeAddress.getFirstRow();
        lastRow = cellRangeAddress.getLastRow();
        firstColumn = cellRangeAddress.getFirstColumn();
        lastColumn = cellRangeAddress.getLastColumn();
        for (int r = firstRow; r <= lastRow; r++) {
            InputsArray[ArrayVar]="";
            // System.out.println(ArrayVar);
            OPTextArea.appendText("\n");

            for (int c = firstColumn; c < lastColumn; c++) {
                //Gets rid of null being printed.
                Row row = sheet1.getRow(r);
                //Row CoSrow = sheet1.getRow(SMfirstRow);
                if (row == null) {
                    System.out.printf(new CellAddress(r, c) + " is in an empty row.");
                } else {
                    Cell cell2 = row.getCell(c);
                    if (cell2 == null) {
                        System.out.printf("Empty Cell ");
                    } else if (cell2.getCellType() == CellType.FORMULA) {

                        switch (cell2.getCachedFormulaResultType()) {
                            case NUMERIC:
                                String NumericCell = DF.format(cell2.getNumericCellValue());
                                String NumericFormatted = String.format("%-20.5s", NumericCell); //%-20.20s
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericFormatted;
                                break;
                            case STRING:
                                String StringCell = cell2.getRichStringCellValue().getString();
                                StringCell = String.format("%-20.5s", StringCell);
                                InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + StringCell;
                                break;
                        }
                    } else {

                        String content = formatter.formatCellValue(cell2);
                        String NumericCell = String.format("%-20.20s", content);

                        InputsArray[ArrayVar]=InputsArray[ArrayVar] + "     " + NumericCell;
                    }

                }
            }
            OPTextArea.appendText(InputsArray[ArrayVar] + "              ");
            System.out.println(InputsArray[ArrayVar]);
            ArrayVar=ArrayVar+1;
        }
    }
    public void ImageSensoroRgan(ActionEvent event) throws Exception{
        try {
            Desktop desktop = null;
            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }

            desktop.open(new File("" + FileLocationImageSensoroRgan));
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    }
    public void FileSizeCalculator(ActionEvent event) throws Exception{
        try {
            Desktop desktop = null;
            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }

            java.awt.Desktop.getDesktop().browse(URI.create("http://jan.ucc.nau.edu/lrm22/pixels2bytes/calculator.htm"));
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    }
    public void PayloadAlignment(ActionEvent event) throws Exception{
        try {
            Desktop desktop = null;
            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }

            desktop.open(new File("" + FileLocationPayloadAlignment));
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    }
    public void LightInput(ActionEvent event) throws Exception {
        OPTextArea.setText("Parallel, polychromatic light rays enter the payload aperture.");

    }
    public void ImageFileSize(ActionEvent event) throws Exception {
        OPTextArea.setText("After compression file size: 144kB");

    }
//---------------------------------------- MISC CODING STUFF --------------------------------//

    /*
    public void ButtonPress1(ActionEvent event){



        //This is all the relevant code regarding buttons

        System.out.println("This can print out the button's Name!");

        Button tempButton = ((Button)event.getSource()); //Get the Button that is being pressed

        String currentButtonPress = tempButton.getText(); //Get the text of the button that is being pressed

        System.out.println("Button Press event "+ currentButtonPress) ;

        //ControlOutput.setText("Hello World!");

        outputtext.setText("I MADE THIS TEXT FROM THE CONTROLLER CLASS!!!!!! 11111111");

        outputtext2.setText("I can set this one independantly 2");


    }
    public void TextFieldTest(ActionEvent event){

        //This is all the relevant code regarding Reading user inputs

        System.out.println("Text Field Test");

        TextField TextField1 = ((TextField)event.getSource());//Get the Button that is being pressed

        String TextFieldContents = TextField1.getText();//Get the text of the button that is being pressed

        System.out.println(" This is What is inside Text Field???"+ TextFieldContents) ;

        //Added from outputtest


    }
    public void TextoutputTest(ActionEvent event){

        //This is all the relevant code regarding outputting text to user

        TextArea TextArea1 = ((TextArea)event.getSource());

        //MainThermalTextArea.setText("" + test);

        String TextAreaContents = TextArea1.getText(); //Reads input

        //TextArea1.setText(TextAreaContents.getText());





        System.out.println(" This is What is inside Text Field???"+ TextArea1) ;
    }
*/
}

