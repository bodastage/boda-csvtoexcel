package com.bodastage.csvtoexcel;

import com.opencsv.CSVReader;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import javax.xml.stream.XMLStreamException;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/*
 * Converts csv files to excel format
 * @version 1.0.0
 */

/**
 *
 * @author Bodastage Solutions
 */
public class BodaCSVToExcel {
    
    /**
     * Input csv file or directory containing a bunch of csv files
     */
    private String dataSource = "";
    
    /**
     * Output directory.
     *
     * @since 1.0.0
     */
    private String outputFilename = "CSV_to_Excel.xlsx";
    
    
    /**
     * The base file name of the file being parsed.
     * 
     * @since 1.0.0
     */
    private String baseFileName = "";
    
    /**
     * The file being  parsed.
     * 
     * @since 1.0.0
     */
    private String dataFile;
    
    
    private XSSFWorkbook workbook = new XSSFWorkbook();
    
    /**
     * Parser start time. 
     * 
     * @version 1.0.0
     */
    final long startTime = System.currentTimeMillis();
    
    public static void main(String[] args){
    
        try{
            //show help
            if(args.length != 2 || (args.length == 1 && args[0] == "-h")){
                showHelp();
                System.exit(1);
            }
            //Get bulk CM XML file to parse.
            String filename = args[0];
            String outputFilename = args[1];
           
            BodaCSVToExcel converter = new BodaCSVToExcel();
            converter.setDataSource(filename);
            converter.setOutputFilename(outputFilename);
            converter.createWorkBook();
            converter.printExecutionTime();
        }catch(Exception e){
            System.out.println(e.getMessage());
            System.exit(1);
        }

    }
    
    /**
     * Show parser help.
     * 
     * @since 1.0.0
     * @version 1.0.0
     */
    static public void showHelp(){
        System.out.println("boda-csvtoexcel 1.0.0 Copyright (c) 2017 Bodastage(http://www.bodastage.com)");
        System.out.println("Creates an excel workbook from csv files.");
        System.out.println("Usage: java -jar boda-csvtoexcel.jar <inputDirectory> outputFile.xlsx");
    }
    
    public void createWorkBook() throws UnsupportedEncodingException, IOException{
        processFileOrDirectory();
    }
    /**
     * Set the data source i.e. the input file or directory
     * 
     * @param dataSource 
     */
   public void setDataSource(String dataSource){
       this.dataSource = dataSource;
   }
   
public void processFileOrDirectory()
            throws FileNotFoundException, UnsupportedEncodingException, IOException {
        //this.dataFILe;
        Path file = Paths.get(this.dataSource);
        boolean isRegularExecutableFile = Files.isRegularFile(file)
                & Files.isReadable(file);

        boolean isReadableDirectory = Files.isDirectory(file)
                & Files.isReadable(file);

        if (isRegularExecutableFile) {
            this.setFileName(this.dataSource);
            baseFileName =  getFileBasename(this.dataFile);
            System.out.print("Adding  " + this.baseFileName + " to MS Excel workbook...");
            this.addFileToWorkBook(this.dataSource);
            System.out.println("Done");
        }

        if (isReadableDirectory) {

            File directory = new File(this.dataSource);

            //get all the files from a directory
            File[] fList = directory.listFiles();

            for (File f : fList) {
                this.setFileName(f.getAbsolutePath());
                try {
                   
                    baseFileName =  getFileBasename(this.dataFile);
                    System.out.print("Adding " + this.baseFileName + " to MS Excel workbook...");
                    this.addFileToWorkBook(f.getAbsolutePath());
                    System.out.println("Done");
                   
                } catch (Exception e) {
                    System.out.println(e.getMessage());
                    System.out.println("Skipping file: " + this.baseFileName + "\n");
                }
            }

        }
        
        try {
            FileOutputStream outputStream = new FileOutputStream(this.outputFilename);
            workbook.write(outputStream);
            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        

    }

    /**
     * Set name of file to parser.
     * 
     * @since 1.0.0
     * @version 1.0.0
     * @param filename 
     */
    private void setFileName(String filename ){
        this.dataFile = filename;
    }
    
    /**
     * @since 1.0.0
     * @param filename 
     */
    public void addFileToWorkBook(String filename) throws FileNotFoundException, IOException{
        XSSFSheet sheet = this.workbook.createSheet(getFilenameMinusExtension(filename));
        CSVReader reader = new CSVReader(new FileReader(filename));
     String [] nextLine;
	 Integer rowNum = 0;
     while ((nextLine = reader.readNext()) != null) {
		Row row = sheet.createRow(rowNum++);
		int colNum = 0;
		for (Object field : nextLine) {
			Cell cell = row.createCell(colNum++);
			if (field instanceof String) {
				cell.setCellValue((String) field);
			} else if (field instanceof Integer) {
				cell.setCellValue((Integer) field);
			}
		}
    }
    }
    /**
     * Set the output directory.
     * 
     * @since 1.0.0
     * @version 1.0.0
     * @param directoryName 
     */
    public void setOutputFilename(String filename ){
        if(!filename.endsWith(".xlsx")){
            filename = filename + ".xlsx";
        }
        this.outputFilename = filename;
    }
    
    
/**
     * Print program's execution time.
     * 
     * @since 1.0.0
     */
    public void printExecutionTime(){
        float runningTime = System.currentTimeMillis() - startTime;
        
        String s = "Processing completed. ";
        s = s + "Total time:";
        
        //Get hours
        if( runningTime > 1000*60*60 ){
            int hrs = (int) Math.floor(runningTime/(1000*60*60));
            s = s + hrs + " hours ";
            runningTime = runningTime - (hrs*1000*60*60);
        }
        
        //Get minutes
        if(runningTime > 1000*60){
            int mins = (int) Math.floor(runningTime/(1000*60));
            s = s + mins + " minutes ";
            runningTime = runningTime - (mins*1000*60);
        }
        
        //Get seconds
        if(runningTime > 1000){
            int secs = (int) Math.floor(runningTime/(1000));
            s = s + secs + " seconds ";
            runningTime = runningTime - (secs/1000);
        }
        
        //Get milliseconds
        if(runningTime > 0 ){
            int msecs = (int) Math.floor(runningTime/(1000));
            s = s + msecs + " milliseconds ";
            runningTime = runningTime - (msecs/1000);
        }

        
        System.out.println(s);
    }

    /**
     * Get file base name.
     * 
     * @since 1.0.0
     */
     public String getFileBasename(String filename){
        try{
            return new File(filename).getName();
        }catch(Exception e ){
            return filename;
        }
    }
     
    public String getFilenameMinusExtension(String filename){
        String fName = getFileBasename(filename);
        return fName.replaceAll("(?i)\\.csv$", "");
    }
}
