import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Scanner;

import org.apache.poi.xdgf.usermodel.shape.exceptions.StopVisitingThisBranch;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class Hello {

	static int COL;
	
	public static void main(String []args) throws RowsExceededException, WriteException {
		
                Scanner sc = new Scanner(System.in);
                 System.out.println("Enter the number Columns:");
                 int col=sc.nextInt();
		         COL=col;
                  
                 System.out.println("Enter the every Column name:");
                 String colname[]=new String[col];
                 
                 for(int i=0;i<col;i++) {
                	 colname[i]=sc.next();
                 }
                 System.out.println("Enter the number of Rows:");
                 int row=sc.nextInt();
		
		
		try {
			WritableWorkbook writableWorkbook=Workbook.createWorkbook(new File("Workbook.xls"));
		    WritableSheet writableSheet =writableWorkbook.createSheet("1", 0);
		   
		    
		    for(int i=0;i<col;i++) {
		     writableSheet.addCell(new Label(i,0,colname[i]));     	
		    }
		    
		    
		    ArrayList<Student> students = createList(row,col);
            int rownum=0;
            for (Student student : students) {
				rownum++;	
				String var[] =student.getVarX();
				
			   	  for(int i=0;i<col;i++) {
				  writableSheet.addCell(new Label(i,rownum,var[i]));	
				}
			     	
            }
		    
		  
		      writableWorkbook.write();
		    writableWorkbook.close();
		    
		    System.out.println("Your Excel file is Created");} 
		   catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		}	
	
	
		
	static ArrayList<Student> createList(int row,int col){
            
		System.out.println("Excel sheet is created\nEnter data one by one :\n");
		
		Scanner sc = new Scanner(System.in);
		
		ArrayList<Student> students = new ArrayList<>();
		
		for(int i=0;i<row;i++) {
			String var[]=new String[col];
			
			for(int j=0;j<col;j++) {
			var[j]=sc.next();	
			}
			
		students.add(new Student(var));
		}
		
	return students;	
	}
	
	
      	
	
	
	
}

  


class Student{
       
        	            
     int size=Hello.COL;
      	  
     String var[] =new String[size];
     String entry;
     public Student(String var[]) {
		this.var=var;
	}
          
    public String[] getVarX() {
        return var;	
       }
         
}
