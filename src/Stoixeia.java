import java.io.File;
import java.util.Date;

public class Stoixeia {

	Date arxh;
	Date telos;
	String path;
	File f;
	String fname;
	String fname2;
	File f2;
	String path2;
	
	public Stoixeia() {
		super();
	}


	


	public Stoixeia(Date arxh, Date telos, String path, File f, String fname, String fname2, File f2, String path2) {
		super();
		this.arxh = arxh;
		this.telos = telos;
		this.path = path;
		this.f = f;
		this.fname = fname;
		this.fname2 = fname2;
		this.f2 = f2;
		this.path2 = path2;
	}





	public String getFname2() {
		return fname2;
	}





	public void setFname2(String fname2) {
		this.fname2 = fname2;
	}





	public File getF2() {
		return f2;
	}





	public void setF2(File f2) {
		this.f2 = f2;
	}





	public String getPath2() {
		return path2;
	}





	public void setPath2(String path2) {
		this.path2 = path2;
	}





	public String getFname() {
		return fname;
	}

	public void setFname(String fname) {
		this.fname = fname;
	}

	public File getF() {
		return f;
	}

	public void setF(File f) {
		this.f = f;
	}

	public Date getArxh() {
		return arxh;
	}
	public void setArxh(Date arxh) {
		this.arxh = arxh;
	}
	public Date getTelos() {
		return telos;
	}
	public void setTelos(Date telos) {
		this.telos = telos;
	}
	public String getPath() {
		return path;
	}
	public void setPath(String path) {
		this.path = path;
	}
	
	
	
	
}
