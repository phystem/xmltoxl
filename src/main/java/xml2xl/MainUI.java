package xml2xl;

import javax.swing.*;

public class MainUI extends JFrame {

    public MainUI() {
        setTitle("XML2XL - InProgress");
        initComponents();
    }

    private void initComponents() {


    }

    public static void main(String[] args) {
        MainUI mainUI = new MainUI();
        mainUI.setSize(500, 500);
        mainUI.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        mainUI.setLocationRelativeTo(null);
        mainUI.setVisible(true);
    }
}
