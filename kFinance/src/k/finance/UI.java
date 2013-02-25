package k.finance;

import java.awt.EventQueue;

import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;

import java.awt.Font;
import javax.swing.JTextField;
import javax.swing.JButton;
import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;
import java.io.File;

public class UI {

	private JFrame frmOracleerpby;
	private JTextField textField;
	private final JFileChooser fc = new JFileChooser();

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					UI window = new UI();
					window.frmOracleerpby.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public UI() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frmOracleerpby = new JFrame();
		frmOracleerpby.setTitle("OracleERP\u9884\u8BA1\u8D1F\u503A\u6570\u636E\u5904\u7406\u5DE5\u5177By\u607A\u54E5");
		frmOracleerpby.setBounds(100, 100, 580, 202);
		frmOracleerpby.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frmOracleerpby.getContentPane().setLayout(null);
		
		JLabel lbloracle = new JLabel("\u8BF7\u9009\u62E9\u8981\u5904\u7406\u7684\u6587\u4EF6\uFF08\u4ECEOracle\u5BFC\u51FA\u540E\u9700\u8981\u53E6\u5B58\u4E00\u4EFD\uFF09");
		lbloracle.setFont(new Font("Dialog", Font.PLAIN, 14));
		lbloracle.setBounds(34, 21, 336, 20);
		frmOracleerpby.getContentPane().add(lbloracle);
		
		textField = new JTextField();
		textField.setBounds(34, 53, 394, 31);
		frmOracleerpby.getContentPane().add(textField);
		textField.setColumns(10);
		
		fc.setDialogTitle("请选择OracleERP导出的xls文件");
		
		JButton button = new JButton("\u9009\u62E9\u6587\u4EF6");
		button.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				int result = fc.showOpenDialog(null);
				if (result == JFileChooser.APPROVE_OPTION) {
			          File selectedFile = fc.getSelectedFile();
			          textField.setText(selectedFile.getAbsolutePath());
			    }
			}
		});
		button.setBounds(440, 53, 86, 31);
		frmOracleerpby.getContentPane().add(button);
		
		JButton btnNewButton = new JButton("\u5904\u7406\u5E76\u5BFC\u51FAExcel");
		btnNewButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				String inFile = textField.getText();
				GetContentFromXls.get(inFile);
				JOptionPane.showMessageDialog(null,"处理完毕，请进入D盘查找新创建的文件！","message",JOptionPane.INFORMATION_MESSAGE);
			}
		});
		btnNewButton.setBounds(34, 98, 492, 38);
		frmOracleerpby.getContentPane().add(btnNewButton);
	}
}
