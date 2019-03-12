import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

public class runfrom {

    private static String OUT_FILE = "";
    private static String INPUT_PATH = "";

    private static String DEFAULT_PATH = "";

    private static List<String> fileList = new ArrayList<>();
    private static List<Object[]> dataList = new ArrayList<>();
    private static JFrame jFrame = null;
    private static JTextArea textArea = null;
    private static JRadioButton jrb1=null;
    private  static  JComboBox jComboBox=null;

    public static void main(String arg[]) {

        // 创建 JFrame 实例
        jFrame = new JFrame("数据合并");
        // Setting the width and height of frame
        jFrame.setSize(700, 400);
        jFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        /* 创建面板，这个类似于 HTML 的 div 标签
          我们可以创建多个面板并在 JFrame 中指定位置
          面板中我们可以添加文本字段，按钮及其他组件。
         */
        JPanel panel = new JPanel();
        // 添加面板
        jFrame.add(panel);
       /*
         调用用户定义的方法并添加组件到面板
         */
        placeComponents(panel);

        // 设置界面可见
        jFrame.setVisible(true);


    }

    private static void placeComponents(JPanel panel) {

        /* 布局部分我们这边不多做介绍
         * 这边设置布局为 null
         */
        panel.setLayout(null);


        /*
         * 创建文本域用于用户输入
         */
        JTextField userText = new JTextField(20);
        userText.setBounds(90, 40, 400, 50);
        panel.add(userText);


        JButton jb = new JButton("输入文件夹");
        jb.setBounds(500, 45, 115, 40);
        fileBrowser(jb, userText);

        panel.add(jb, userText);

        /*
         *这个类似用于输入的文本域
         * 但是输入的信息会以点号代替，用于包含密码的安全性
         */

        JTextField userText2 = new JTextField(20);
        userText2.setBounds(90, 120, 400, 50);
        panel.add(userText2);

        JButton jb2 = new JButton("输出文件夹");
        jb2.setBounds(500, 123, 115, 40);
        fileBrowser(jb2, userText2);
        panel.add(jb2);
        // 创建登录按钮
        JButton loginButton = new JButton("数据生成");
        loginButton.setBounds(240, 200, 160, 50);
        panel.add(loginButton);
        jrb1=new JRadioButton("是否分成两个文件");
        jrb1.setBounds(490,170,200,50);
        panel.add(jrb1);
        JLabel jLabel = new JLabel("Logo长度");
        Integer []integer = {0,1,2,3,4,5,6,7,8,9,10};
        jComboBox = new JComboBox(integer);
        jComboBox.setBounds(550,220,70,30);
        jComboBox.setSelectedIndex(5);
        jLabel.setBounds(492,210,200,50);
        panel.add(jLabel);
        panel.add(jComboBox);
        confirm(loginButton, userText, userText2);


    }

    /*private static JTextArea PopTextArea() {

        TextAreaFrame frame = new TextAreaFrame();
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setVisible(true);
        return frame.getTextArea();

    }*/

    private static void fileBrowser(JButton jButton, final JTextField jTextField) {
        jButton.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {

                JFileChooser chooser = new JFileChooser();
                chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
                FileNameExtensionFilter FILTER = new FileNameExtensionFilter(
                        ".xls&.xlsx", "xls", "xlsx");
                if (!DEFAULT_PATH.equals("")) {
                    chooser.setCurrentDirectory(new File(DEFAULT_PATH));
                }
                chooser.setCurrentDirectory(new File(""));
                chooser.setFileFilter(FILTER);

                int returnVal = chooser.showOpenDialog(new JPanel());

                if (returnVal == JFileChooser.APPROVE_OPTION) {
                   /* System.out.println("你打开的文件是: " +
                            chooser.getSelectedFile().getAbsolutePath());*/
                    DEFAULT_PATH = chooser.getSelectedFile().getAbsolutePath();
                    jTextField.setText(DEFAULT_PATH);
                }
            }
        });

    }


    private static void confirm(JButton jButton, final JTextField jTextField1, final JTextField jTextField2) {

        jButton.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {

                INPUT_PATH = stringUtil(jTextField1.getText());

                OUT_FILE = stringUtil(jTextField2.getText());

                if (!INPUT_PATH.equals("") && !OUT_FILE.equals("")) {
                 /*   new Runnable() {
                        @Override
                        public void run() {
                            textArea = PopTextArea();
                        }
                    }.run();*/
                    run(jTextField1.getText());

                }


            }
        });


    }

    private static String stringUtil(String string) {

        StringBuffer stringBuffer = new StringBuffer();
        String[] split = string.split("\\\\");

        for (int i = 0; i < split.length; i++) {
            if (i != split.length - 1) {
                stringBuffer.append(split[i] + "\\\\");
            } else {
                stringBuffer.append(split[i]);
            }
        }
        return stringBuffer.toString();

    }

    private static void getInputFileList(String path) {

        File file = new File(path);
        if (file.exists()) {
            File[] files = file.listFiles();

            for (File file2 : files) {
                if (file2.isDirectory()) {
                    getInputFileList(file2.getAbsolutePath());
                } else {
                    fileList.add(file2.getAbsolutePath());
                }
            }

        } else {
            System.out.println("文件夹不存在!");
        }
    }

    private static void read(String file) {
        try {
            String ext = file.substring(file.lastIndexOf("."));
            InputStream fis = new FileInputStream(file);
            Workbook wb = null;
            if (".xls".equals(ext)) {
                wb = new HSSFWorkbook(fis);
            } else if (".xlsx".equals(ext)) {
                wb = new XSSFWorkbook(fis);
            }
            Sheet sheet = wb.getSheetAt(0);
            int rows = sheet.getPhysicalNumberOfRows();
            for (int i = 0; i < rows; i++) {
                if (dataList.size() > 0 && i == 0) {
                    System.out.println(jComboBox.getSelectedIndex());
                    i=jComboBox.getSelectedIndex()+1;
                }else if(dataList.size()<=0 && i==0){
                    i+=jComboBox.getSelectedIndex();
                }
                Row row = sheet.getRow(i);

                Object[] objects = new Object[row.getPhysicalNumberOfCells()];
                int index = 0;
                for (Cell cell : row) {
                    if (cell.getCellType().equals(CellType.NUMERIC)) {
                        objects[index] = cell.getNumericCellValue();
                        if (index == 2) {
                            SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");
                            objects[index] = sdf.format(cell.getDateCellValue());
                        }
                    }
                    if (cell.getCellType().equals(CellType.STRING)) {
                        objects[index] = cell.getStringCellValue();
                    }
                    index++;
                }
                dataList.add(objects);
            }
            wb.close();
            fis.close();
            System.out.println("数据量：" + dataList.size());
           /* textArea.append("数据量：" + dataList.size() + "\r\n");*/
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void write(String file, String defaultFileName) {
        try {

            FileOutputStream fos = new FileOutputStream(file + "\\\\" + defaultFileName + ".xlsx");

            SXSSFWorkbook wb = new SXSSFWorkbook(10000);
            Sheet sheet = wb.createSheet("sheet");
            //在sheet里创建行
            for (int j = 0; j < dataList.size(); j++) {
//                System.out.println(j);
                Row row = sheet.createRow(j);
                Object[] item = dataList.get(j);
                //在sheet里创建列
                for (int i = 0; i < item.length; i++) {
                    Cell cell = row.createCell(i);
                    if (item[i] instanceof String) {
                        cell.setCellValue(String.valueOf(item[i]));
                    } else {

                        if (item[i] != null) {
                            cell.setCellValue((double) item[i]);
                        }
                    }
                }
            }
            wb.write(fos);
            wb.dispose();
            fos.close();
            if(!jrb1.isSelected()) {
                fileList.clear();
            }
            System.out.println("写入完成！");
           /* textArea.append("写入完成!");*/
            /*  PopDialog(jFrame,"提示框","写入完成",true);*/
            JOptionPane.showMessageDialog(jFrame, "写入完成", "提示框", JOptionPane.INFORMATION_MESSAGE);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    private static void run(String filename) {


        getInputFileList(INPUT_PATH);
        System.out.println("总文件数：" + fileList.size());
       /* textArea.append("总文件数：" + fileList.size());*/

        //第一个输出文件
        for (int m = 0; m < Math.ceil(jrb1.isSelected()==true?fileList.size()/2:fileList.size()); m++) {
            String s = fileList.get(m);
            System.out.println(s);
           /* textArea.append(s);*/
            // 读
            read(fileList.get(m));
            if(jrb1.isSelected()) {
                fileList.remove(m);
            }
        }
        //写
        if(jrb1.isSelected()) {
            write(OUT_FILE, "_demo_1");
        }else{
            write(OUT_FILE, "_demo_");
        }
        Object[] header = dataList.get(0);
        dataList.clear();
        dataList.add(header);
        if(jrb1.isSelected()) {
               //第二个输出文件
                for (String path : fileList) {
                    System.out.println(path);
                    // 读
                    read(path);
                }
                fileList.clear();
                //写
                write(OUT_FILE ,"_demo_2");
                dataList.clear();
        }
    }

    //自动生成文件名字

    private static String autonGenerateFileName(String FilePath) {
        //D:\hehe\CostaRica-EXP-201803

        //EXP_COSTARICA_2018

        if (FilePath.contains("EXP")) {
            return str2Str(FilePath, "EXP");

        } else if (FilePath.contains("IMP")) {
            return str2Str(FilePath, "IMP");
        } else {
            return "";
        }


    }

    //字符串处理封装
    private static String str2Str(String str, String mode) {
        int i = str.indexOf(mode);
        String substring = str.substring(0, i - 1);
        String[] split = substring.split("\\\\");
        String start = split[split.length - 1].toUpperCase();
        String[] split1 = str.split("-");
        String end = split1[split1.length - 1];
        return start + "_" + mode + "_" + end;

    }


}



