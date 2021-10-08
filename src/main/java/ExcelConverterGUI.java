import javax.swing.*;

public class ExcelConverterGUI extends  JFrame{
    private JButton selectTxtFileButton;
    private JButton selectExcelFileButton;
    private JButton convertTxtExcelButton;
    private JButton splitExcelButton;
    private JPanel mainPanel;

    public ExcelConverterGUI(String title) {
        super(title);

        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        this.setContentPane(mainPanel);
        this.pack();

    }
}
