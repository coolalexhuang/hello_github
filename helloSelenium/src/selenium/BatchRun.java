package selenium;
import junit.framework.Test;
import junit.framework.TestSuite;

public class BatchRun {

	public static Test suite() {
		TestSuite suite = new TestSuite();
		
		suite.addTestSuite(Test1.class);
		return suite;
	}

	public static void main(String[] args) {
		junit.textui.TestRunner.run(suite());
	}
}
