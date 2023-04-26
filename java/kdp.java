import controller.CrossWordGenerator8x5by11;
import controller.wordFileGenerator6by9;
import controller.wordFileGenerator8x5by11;

import java.io.FileNotFoundException;
import java.io.IOException;

public class kdp {
    public static void main(String[] args) throws FileNotFoundException {
        wordFileGenerator6by9 a = new wordFileGenerator6by9();
        wordFileGenerator8x5by11 b = new wordFileGenerator8x5by11();
        CrossWordGenerator8x5by11 c = new CrossWordGenerator8x5by11();

        String donePath = "/done5/";
        try {
//            a.Start();
//            b.Start();
              c.Start(donePath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
