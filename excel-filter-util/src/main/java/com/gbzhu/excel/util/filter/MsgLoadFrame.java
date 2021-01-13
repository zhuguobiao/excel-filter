package com.gbzhu.excel.util.filter;

import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowEvent;
import java.io.File;
import java.io.IOException;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.JTextField;

import org.apache.log4j.Appender;
import org.apache.log4j.FileAppender;
import org.apache.log4j.Layout;
import org.apache.log4j.Logger;
import org.apache.log4j.PatternLayout;


public class MsgLoadFrame extends JFrame implements ActionListener {
    /**
     * @author Roy_70
     * @date
     */
    private static final long serialVersionUID = -1189035634361220261L;
    private static Logger logger = Logger.getLogger(MsgLoadFrame.class);
    JFrame mainframe;
    JPanel panel;
    //创建相关的Label标签
    JLabel inFilePathLabel = new JLabel("导入文件(Excel):");
    JLabel outFilePathLabel = new JLabel("导出文件路径:");
    JLabel outLogPathLabel = new JLabel("过程日志路径:");
    JLabel excelColIndexLabel = new JLabel("指定分组的列:");

    //创建相关的文本域
    JTextField inFilePathText = new JTextField(20);
    JTextField outFilePathText = new JTextField(20);
    JTextField outLogPathText = new JTextField(20);
    JTextField excelColIndexTextField = new JTextField(20);

    //创建滚动条以及输出文本域
    JScrollPane jscrollPane;
    JTextArea outTextArea = new JTextArea();
    //创建按钮
    JButton inFilePathButton = new JButton("...");
    JButton outFilePathButton = new JButton("...");
    JButton outLogPathButton = new JButton("...");
    JButton startButton = new JButton("开始");

    public void show() {
        mainframe = new JFrame("Excel分组导出-1.0");
        // Setting the width and height of frame
        mainframe.setSize(575, 580);
        mainframe.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        mainframe.setResizable(false);//固定窗体大小

        Toolkit kit = Toolkit.getDefaultToolkit(); // 定义工具包
        Dimension screenSize = kit.getScreenSize(); // 获取屏幕的尺寸
        int screenWidth = screenSize.width / 2; // 获取屏幕的宽
        int screenHeight = screenSize.height / 2; // 获取屏幕的高
        int height = mainframe.getHeight(); //获取窗口高度
        int width = mainframe.getWidth(); //获取窗口宽度
        mainframe.setLocation(screenWidth - width / 2, screenHeight - height / 2);//将窗口设置到屏幕的中部
        //窗体居中，c是Component类的父窗口
        //mainframe.setLocationRelativeTo(c);
        Image myimage = kit.getImage("resourse/hxlogo.gif"); //由tool获取图像
        mainframe.setIconImage(myimage);
        initPanel();//初始化面板
        mainframe.add(panel);
        mainframe.setVisible(true);
    }

    /* 创建面板，这个类似于 HTML 的 div 标签
     * 我们可以创建多个面板并在 JFrame 中指定位置
     * 面板中我们可以添加文本字段，按钮及其他组件。
     */
    public void initPanel() {
        this.panel = new JPanel();
        panel.setLayout(null);
        //this.panel = new JPanel(new GridLayout(3,2)); //创建3行3列的容器
        /* 这个方法定义了组件的位置。
         * setBounds(x, y, width, height)
         * x 和 y 指定左上角的新位置，由 width 和 height 指定新的大小。
         */
        inFilePathLabel.setBounds(10, 20, 120, 25);
        inFilePathText.setBounds(120, 20, 400, 25);
        inFilePathButton.setBounds(520, 20, 30, 25);
        this.panel.add(inFilePathLabel);
        this.panel.add(inFilePathText);
        this.panel.add(inFilePathButton);

        outFilePathLabel.setBounds(10, 50, 120, 25);
        outFilePathText.setBounds(120, 50, 400, 25);
        outFilePathButton.setBounds(520, 50, 30, 25);
        this.panel.add(outFilePathLabel);
        this.panel.add(outFilePathText);
        this.panel.add(outFilePathButton);

        outLogPathLabel.setBounds(10, 80, 120, 25);
        outLogPathText.setBounds(120, 80, 400, 25);
        outLogPathButton.setBounds(520, 80, 30, 25);
        this.panel.add(outLogPathLabel);
        this.panel.add(outLogPathText);
        this.panel.add(outLogPathButton);

        excelColIndexLabel.setBounds(10, 110, 120, 25);
        excelColIndexTextField.setBounds(120, 110, 400, 25);
        this.panel.add(excelColIndexLabel);
        this.panel.add(excelColIndexTextField);

        startButton.setBounds(10, 150, 80, 25);
        this.panel.add(startButton);

        outTextArea.setEditable(false);
        outTextArea.setFont(new Font("标楷体", Font.BOLD, 16));
        jscrollPane = new JScrollPane(outTextArea);
        jscrollPane.setBounds(10, 200, 550, 330);
        this.panel.add(jscrollPane);
        //增加动作监听
        inFilePathButton.addActionListener(this);
        outFilePathButton.addActionListener(this);
        outLogPathButton.addActionListener(this);
        startButton.addActionListener(this);
    }

    /**
     * 单击动作触发方法
     *
     * @param event
     */
    @Override
    public void actionPerformed(ActionEvent event) {
        System.out.println(event.getActionCommand());
        if (event.getSource() == startButton) {
            //确认对话框弹出
            int result = JOptionPane.showConfirmDialog(null, "是否开始转换?", "确认", 0);//YES_NO_OPTION
            if (result == 1) {//是：0，否：1，取消：2
                return;
            }
            System.out.println(inFilePathText.getText());
            if (inFilePathText.getText().equals("") || outFilePathText.getText().equals("")
                    || outLogPathText.getText().equals("")) {
                JOptionPane.showMessageDialog(null, "路径不能为空", "提示", 2);//弹出提示对话框，warning
                return;
            }
            String excelColIndex = excelColIndexTextField.getText();
            if (excelColIndex.equals("")) {
                JOptionPane.showMessageDialog(null, "指定分组的列不能为空", "提示", 2);//弹出提示对话框，warning
                return;
            }

            outTextArea.setText("");
            String outlogpath = outLogPathText.getText();
            //设置log4j日志输出格式以及路径
            Layout layout = new PatternLayout("%-d{yyyy-MM-dd HH:mm:ss}  [ %C{1}--%M:%L行 ] - [ %p ]  %m%n");
            Appender appender = null;
            try {
                appender = new FileAppender(layout, outlogpath + "\\log.log");
            } catch (IOException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
            logger.addAppender(appender);
            System.out.println(outlogpath + "\\log.log");
            logger.debug("数据转换开始");

            File dir = new File( outFilePathText.getText());
            if (!dir.exists()) {// 判断目录是否存在
                dir.mkdir();
            }
            new ExcelUtil(inFilePathText.getText(), outFilePathText.getText(), excelColIndex, logger).filterExcel();

            logger.debug("数据转换结束");
            outTextArea.setText("此处放输出信息(非日志)");

            result = JOptionPane.showConfirmDialog(null, "是否打开转换结果文件夹?", "确认", 0);//YES_NO_OPTION
            if (result == 0) {//是：0，否：1，取消：2
                try {
                    Desktop.getDesktop().open(new File(outFilePathText.getText()));
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }

            result = JOptionPane.showConfirmDialog(null, "是否打开日志文件?", "确认", 0);//YES_NO_OPTION
            if (result == 0) {//是：0，否：1，取消：2
                try {
                    @SuppressWarnings("unused")
                    Process process = Runtime.getRuntime().exec("cmd.exe  /c notepad " + outlogpath + "\\log.log");//调用cmd方法使用记事本打开文件
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }

        }


        //判断三个选择按钮并对应操作
        if (event.getSource() == inFilePathButton) {
            File file = openChoseWindow(JFileChooser.FILES_ONLY);
            if (file == null)
                return;
            inFilePathText.setText(file.getAbsolutePath());
            outFilePathText.setText(file.getParent() + "\\out");
            outLogPathText.setText(file.getParent() + "\\log");
        }
        if (event.getSource() == outFilePathButton) {
            File file = openChoseWindow(JFileChooser.DIRECTORIES_ONLY);
            if (file == null)
                return;
            outFilePathText.setText(file.getAbsolutePath() + "\\out");
        }
        if (event.getSource() == outLogPathButton) {
            File file = openChoseWindow(JFileChooser.DIRECTORIES_ONLY);
            if (file == null)
                return;
            outLogPathText.setText(file.getAbsolutePath() + "\\log");
        }

    }

    /**
     * 打开选择文件窗口并返回文件
     *
     * @param type
     * @return
     */
    public File openChoseWindow(int type) {
        JFileChooser jfc = new JFileChooser();
        jfc.setFileSelectionMode(type);//选择的文件类型(文件夹or文件)
        jfc.showDialog(new JLabel(), "选择");
        File file = jfc.getSelectedFile();
        return file;
    }

    public void windowClosed(WindowEvent arg0) {
        System.exit(0);
    }

    public void windowClosing(WindowEvent arg0) {
        System.exit(0);
    }

}