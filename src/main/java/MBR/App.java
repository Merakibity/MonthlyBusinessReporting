package MBR;

import java.io.IOException;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args ) throws IOException
    {
        Mbr m = new Mbr();
       Mbr.createAndStartService();
        m.createDriver();
        int l = m.ketData();
        for (int i = 1; i <= l; i++) {
            m.getValues(i);
            try { 
                m.mbrgen(1);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
            try { 
                m.mbrgen(2);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
    }
}
}
