
public class Student {

	String epwnimo;
	String onoma;
	int apousies;
	int merikesapousies;
	int ApousiesApoArxh;
	
	
	
	public Student(String epwnimo, String onoma, int apousies, int merikesapousies,int ApousiesApoArxh) {
		super();
		this.epwnimo = epwnimo;
		this.onoma = onoma;
		this.apousies = apousies;
		this.merikesapousies = merikesapousies;
		this.ApousiesApoArxh=ApousiesApoArxh;
	}

	public int getApousiesApoArxh() {
		return ApousiesApoArxh;
	}

	public void setApousiesApoArxh(int apousiesApoArxh) {
		ApousiesApoArxh = apousiesApoArxh;
	}

	public int getMerikesapousies() {
		return merikesapousies;
	}

	public void setMerikesapousies(int merikesapousies) {
		this.merikesapousies = merikesapousies;
	}

	public String getEpwnimo() {
		return epwnimo;
	}
	public void setEpwnimo(String epwnimo) {
		this.epwnimo = epwnimo;
	}
	public String getOnoma() {
		return onoma;
	}
	public void setOnoma(String onoma) {
		this.onoma = onoma;
	}
	public int getApousies() {
		return apousies;
	}
	public void setApousies(int apousies) {
		this.apousies = apousies;
	}
	
	
	
}
