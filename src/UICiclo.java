import com.jose.model.CicloContable;
import com.jose.model.FileWorkbook;

import javax.swing.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

/**
 * Created by agua on 14/07/15.
 */
public class UICiclo extends JFrame {

    private JPanel UICiclo;
    private JButton mCreateCicloButton;
    private JLabel mStatusLabel;

    public static void main(String[] args) {
        JFrame frame = new JFrame("UICiclo");
        frame.setContentPane(new UICiclo().UICiclo);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setResizable(false);
        frame.pack();
        frame.setVisible(true);

    }

    public UICiclo() {
        FileWorkbook fileWorkbook = new FileWorkbook();

        mCreateCicloButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                int statusCode = fileWorkbook.saveWorkBook();

                if (statusCode == fileWorkbook.FILE_CREATED) {
                    mStatusLabel.setText("Libro Excel creado correctamente");
                } else if (statusCode == fileWorkbook.FILE_UPDATED) {
                    mStatusLabel.setText("Libro Excel actualizado correctamente");
                } else if (statusCode == fileWorkbook.FILE_ERROR) {
                    mStatusLabel.setText("Hubo un error al leer/escribir el archivo");
                }
            }
        });
    }

}
