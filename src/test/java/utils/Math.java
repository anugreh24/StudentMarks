package utils;

import java.io.IOException;

public class Math extends Subject {
	
	
	public static void main(String[] args) throws IOException {
		
		double[] math_marks = null;
		
		for (int i=0;i<5;i++) {
				math_marks[i] = al.get(i).maths_score;
				
		}
		
		System.out.print(math_marks);
	}
}
