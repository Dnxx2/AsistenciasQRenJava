/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ConexionSQL;
import java.sql.Connection; 
import java.sql.DriverManager;
import javax.swing.JOptionPane;
/**
 *
 * @author 28082
 */
public class conexionsql {
    Connection conectar = null; 
    
    public Connection conexion(){
        try{
            Class.forName("com.mysql.cj.jdbc.Driver");
            conectar = (Connection)DriverManager.getConnection("jdbc:mysql://localhost/loginai","root","daniel2020");
            //JOptionPane.showMessageDialog(null,"Conexion Exitosa!!");
                        
        }catch(Exception e){
            JOptionPane.showMessageDialog(null, "Conexion Fallida :( " +e.getMessage());
        }

        
      return conectar;  
    }
    
}
