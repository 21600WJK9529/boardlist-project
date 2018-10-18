
/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package projectfinal;

import com.healthmarketscience.jackcess.Database;
import com.healthmarketscience.jackcess.DatabaseBuilder;
import java.io.File;
import javax.swing.JOptionPane;
import projectfinal.LoadingScreen.HomeScreen;
import projectfinal.LoadingScreen.LoadingScreen;

/**
 *
 * @author Matthew
 */
public class ProjectFinal {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws InterruptedException {
        // TODO code application logic here
        ProjectFinal f = new ProjectFinal();
        f.run();
    }
    
    public void run() throws InterruptedException  {
        LoadingScreen ls = new LoadingScreen();
        ls.setVisible(true);
        //ls.setSize(600,400);
        for(int i = 0; i<=100;i++){
            Thread.sleep(1);
            ls.progressCount.setText(String.valueOf(i+"%"));
            ls.progressBar.setValue(i);
            
            File tempFile = new File("publisher.mdb");
            
            if(tempFile.exists()){
                
            }
            else{
                try {
            
                 DatabaseBuilder.create(Database.FileFormat.V2010, new File("publisher.mdb"));
        } catch (Exception e) {
                    JOptionPane.showMessageDialog(null, "Error creating Database file");
        }
                
            }
            
            if(i == 100){
                ls.dispose();
                HomeScreen home = new HomeScreen();
                home.setVisible(true);
                
            }
        }
    }
    
}
