package basepackage;

public class Runner {
	public static void main(String[] args) {
		System.out.println("Application started!!");
		ReportService service = new ReportService();
		service.fetchData();
	}
}
