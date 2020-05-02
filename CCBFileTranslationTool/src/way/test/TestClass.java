package way.test;

public class TestClass {

	public static void main(String[] args) {
		String a = "								<string>endSpin</string>";
		
//		System.out.println(a.indexOf("\\"));
//		System.out.println(a.substring(a.indexOf("\\")+1));
		System.out.println(new String (new String(a.getBytes()).replace("endSpin", "what").getBytes()));
//		System.out.println(a.contains("endSpain"));
//		System.out.println(a.replace("endSpin", "what"));
		

	}

}
