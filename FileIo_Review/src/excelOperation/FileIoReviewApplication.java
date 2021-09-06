package excelOperation;

import java.io.IOException;
import java.util.List;
import java.util.Scanner;

public class FileIoReviewApplication {

	static Scanner sc = new Scanner(System.in);
	public static void main(String[] args) throws IOException {
		boolean flag = true;
		do {
		int choice = display();
		
		ReadExcel read = new ReadExcel();
		WriteExcel update = new WriteExcel();
		
		switch(choice) {
		
		case 1: 
		//List<Excelmodel> record=
		read.readData(); 
		break;
		
		case 2:
		System.out.println("Enter name");
		String name1 = sc.next();
		read.search(name1);
		break;
		
		case 3: 
		System.out.println("Enter name");
		String name = sc.next();
		System.out.println("Enter update city");
		String city  = sc.next();
		int n = update.writeData(name, city);
		
		if(n!=0)
		{
			System.out.println("City updated Succeefully");
		}
		break;
		
		case 4:  int p =  update.addString();
		if(p!=0)
		{
			System.out.println("City updated Succeefully");
		}
		break;
		
		case 6:  flag = false;
		
		}
		}while(flag);
		
		
//		ReadExcel read = new ReadExcel();
//		//List<Excelmodel> record=
//		read.readData();
//		
	//System.out.println(p.getCity());

	}
	
	public static int display() {
		System.out.println("-------------------------");
		System.out.println("1.Read excel and write to notepad ");
		System.out.println("2.Search name and print Citi on console (take name as user input and find corresponding City )");
		System.out.println("3.Search name and update City (take name as user input and update corresponding City )");
		System.out.println("4.Add “City” to all Cities   (ex : Bangalore -> Bangalore City)");
		System.out.println("6. Close the application");
		System.out.println("-------------------------");

		Scanner sc = new Scanner(System.in);
		int n = sc.nextInt();

		return n;
	}

}
