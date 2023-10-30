package gui;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import javafx.application.Application;
import javafx.beans.property.ReadOnlyObjectWrapper;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.event.ActionEvent;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.cell.TextFieldTableCell;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class MainApp extends Application
{
	private static int rowCnt;
	private static int sheetNum = 0;
	private static int[] currColCnt = {13,8,4};
	private static String[][] colNames = 
	{
		{"Student ID", "Name", "Surname", "Gender", "Primary Email", "Secondary Email", 
		"Department", "Study Program", "Year Enrolled", "Year Graduated", "Student Phone Number", "Country", "City"},
		{"Company ID", "Name", "Profile", "Contact Person Name", "Contact Person Position", "Contact Phone Number", "Employed Students", "Internship Students"},
		{"Student ID", "Company ID", "Position", "Sector"}
	};
    @Override
    public void start(Stage stg0) throws Exception
    {        
        ObservableList<ObservableList<String>> data = FXCollections.observableArrayList();        
        List<List<String>> studData = readExcelFile("dbms.xls",0);
        List<List<String>> compData = readExcelFile("dbms.xls",1);
        List<List<String>> empData = readExcelFile("dbms.xls",2);
                
        TableView<ObservableList<String>> tableView = populateCells(studData, data, currColCnt[0], 0);
        
        final Button btn2 = new Button("Insert");
        btn2.setOnAction
        (
        	(ActionEvent t) -> 
        	{
        		try 
        		{
					insertRow(sheetNum,currColCnt[sheetNum]);
				}
        		catch (Exception e) 
        		{
					e.printStackTrace();
				}
        	}
        );
        
        final Button btn3 = new Button("Delete");
        btn3.setOnAction
        (
            (ActionEvent t) -> 
            {
            	try 
            	{
            		int selRowIndex = tableView.getSelectionModel().getSelectedIndex();
    				deleteRow(sheetNum,selRowIndex);
    				fixIDs(sheetNum);
    			} 
            	catch (Exception e) 
            	{
    				e.printStackTrace();
    			}
            }
         );
        
        final Button btn4 = new Button("Update/Goto StudentDB");
        btn4.setOnAction
        (
            (ActionEvent t) -> 
            {
                try 
                {
                	sheetNum = 0;
                	tableView.getItems().clear();
                	List<List<String>> newData = readExcelFile("dbms.xls",0);
                	populateCells(newData, data, currColCnt[sheetNum], 0);
                } 
                catch (Exception e) 
                {
        			e.printStackTrace();
        		}
           }
        );
        
        final Button btn5 = new Button("Update/Goto CompanyDB");
        btn5.setOnAction
        (
            (ActionEvent t) -> 
            {
                try 
                {
                	sheetNum = 1;
                	tableView.getItems().clear();
                	List<List<String>> newData = readExcelFile("dbms.xls",1);
                	populateCells(newData, data, currColCnt[sheetNum], 1);
                } 
                catch (Exception e) 
                {
        			e.printStackTrace();
        		}
           }
        );
        
        final Button btn6 = new Button("Update/Goto EmployeeDB");
        btn6.setOnAction
        (
            (ActionEvent t) -> 
            {
                try 
                {
                	sheetNum = 2;
                	//tableView.getItems().clear();
                	//populateCells(compData, data, currColCnt[sheetNum], 2);
                } 
                catch (Exception e) 
                {
        			e.printStackTrace();
        		}
           }
        );
        
        final Button btn7 = new Button("Save Changes");
        btn7.setOnAction
        (
            (ActionEvent t) -> 
            {
                try 
                {
                	saveChanges(tableView);
                } 
                catch (Exception e) 
                {
        			e.printStackTrace();
        		}
           }
        );
        
        tableView.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY);
        VBox.setVgrow(tableView, Priority.ALWAYS);
        
        tableView.setEditable(true);        
        
        HBox bottom = new HBox();
        bottom.getChildren().addAll(btn2,btn3,btn4,btn5,btn6,btn7);
        VBox top = new VBox(tableView,bottom);
        
        Scene scene = new Scene(top,1280,720);
		stg0.setTitle("Database Management System (BETA)");
        stg0.setScene(scene);
        stg0.show();
    }

    public static TableView<ObservableList<String>> populateCells(List<List<String>> excelData, ObservableList<ObservableList<String>> data, int colCnt, int sheetNum)
    {
    	TableView<ObservableList<String>> tableView = new TableView<ObservableList<String>>();
    	tableView.getItems().clear();
        for (int i = 0; i < excelData.size(); i++)
            data.add(FXCollections.observableArrayList(excelData.get(i)));
        
        tableView.setItems(data);

    	 for (int i = 0; i < excelData.get(0).size(); i++) 
         {
    		 int curCol = i;
    		 if (curCol >= colCnt) break;
    		 
             final TableColumn<ObservableList<String>, String> column = new TableColumn<>(colNames[sheetNum][curCol]);
             column.setCellValueFactory
             (
             	param -> new ReadOnlyObjectWrapper<>(param.getValue().get(curCol))
             );
             column.setCellFactory(TextFieldTableCell.forTableColumn());
             column.setOnEditCommit(e -> e.getRowValue());

             tableView.getColumns().add(column);
         }
    	 
    	 return tableView;
    }
    
    public static List<List<String>> readExcelFile(String fileName, int sheetNum) throws Exception
    {
    	ArrayList<List<String>> excelInfo = new ArrayList<List<String>>();
    	rowCnt = 0;

        InputStream inputStream = new FileInputStream(fileName);
        DataFormatter formatter = new DataFormatter();

        Workbook workbook = WorkbookFactory.create(inputStream);
        Sheet sheet = workbook.getSheetAt(sheetNum);

        for (Row row : sheet)
        {
            ArrayList<String> tempList = new ArrayList<String>();

            for (Cell cell : row) 
            {
                String text = formatter.formatCellValue(cell);
                tempList.add(text.length() == 0 ? "" : text);
            }
            excelInfo.add(tempList);
                
            if (rowCnt==0) continue;
            if (rowCnt == 200) break;

            rowCnt++;
        }
        return excelInfo;
     }
    
    public static void insertRow(int sheetNum, int colCnt) throws Exception
    {
    	InputStream input = new FileInputStream("dbms.xls");
        Workbook workbook = WorkbookFactory.create(input);
    	Sheet sheet = workbook.getSheetAt(sheetNum);
    	int rows = sheet.getLastRowNum();
    	
    	Row row = sheet.createRow(rows+1);
    	row.createCell(0).setCellValue(rows+2);
    	for (int i=1; i<colCnt; i++)
    		row.createCell(i).setCellValue("Placeholder");
        
        FileOutputStream output = new FileOutputStream("dbms.xls");
        workbook.write(output);
        output.close();
    }
    
    public static void deleteRow(int sheetNum, int rowIndex) throws Exception
    {
    	InputStream input = new FileInputStream("dbms.xls");
        Workbook workbook = WorkbookFactory.create(input);
    	Sheet sheet = workbook.getSheetAt(sheetNum);
        int rows = sheet.getLastRowNum();
        
        if (rowIndex >= 0 && rowIndex < rows)
            sheet.shiftRows(rowIndex + 1, rows, -1);
        if (rowIndex == rows) 
        {
            Row delRow = sheet.getRow(rowIndex);
            if (delRow != null) 
                sheet.removeRow(delRow);
        }
        
        FileOutputStream output = new FileOutputStream("dbms.xls");
        workbook.write(output);
        output.close();
    }
    
    public static void saveChanges(TableView<ObservableList<String>> tableView) throws Exception
    {
    	InputStream input = new FileInputStream("dbms.xls");
        Workbook workbook = WorkbookFactory.create(input);
    	Sheet sheet = workbook.getSheetAt(sheetNum);
    	Row row = null;
        
        for(int i=0;i<tableView.getItems().size();i++)
        {
            row = sheet.createRow(i);          
            for(int j=0; j<tableView.getColumns().size();j++)
            {                
                if(tableView.getColumns().get(j).getCellData(i) != null)
                    row.createCell(j).setCellValue(tableView.getColumns().get(j).getCellData(i).toString());
                else
                    row.createCell(j).setCellValue("");
            }
        }
        
        FileOutputStream output = new FileOutputStream("dbms-modified.xls");
        workbook.write(output);
        output.close();
    }
    
    public static void fixIDs(int sheetNum) throws Exception
    {
    	InputStream input = new FileInputStream("dbms.xls");
        Workbook workbook = WorkbookFactory.create(input);
    	Sheet sheet = workbook.getSheetAt(sheetNum);
    	int ID = 1;
    
        for (Row row : sheet)
        {
            for (Cell cell : row) 
                cell.setCellValue(ID);   
            if (ID == 201) break;
            ID++;
        }
        
        FileOutputStream output = new FileOutputStream("dbms.xls");
        workbook.write(output);
        output.close();
    }
}
