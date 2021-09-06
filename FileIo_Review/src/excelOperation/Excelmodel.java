package excelOperation;

public class Excelmodel {
	
	private int id;
	private String name;
	private String city;
	
	public Excelmodel(int id, String name, String city) {
		super();
		this.id = id;
		this.name = name;
		this.city = city;
	}

	public Excelmodel() {
		// TODO Auto-generated constructor stub
		super();
	}

	public int getId() {
		return id;
	}

	public String getName() {
		return name;
	}

	public String getCity() {
		return city;
	}

	public void setId(int id) {
		this.id = id;
	}

	public void setName(String name) {
		this.name = name;
	}

	public void setCity(String city) {
		this.city = city;
	}

}
