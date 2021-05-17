package cn.xxx.xmind2excel.ui;

import cn.xxx.xmind2excel.util.FileExtension;
import cn.xxx.xmind2excel.biz.XMind2Excel;
import org.jb2011.lnf.beautyeye.ch3_button.BEButtonUI;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.swing.*;
import javax.swing.border.Border;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.File;
import java.util.concurrent.atomic.AtomicReference;

/**
 * @author xiongchenghui
 * @date 2020-08-12
 * &Desc 工具UI
 */
public class SwingClient extends JFrame{
    private static Logger logger = LoggerFactory.getLogger(SwingClient.class);

    private static SwingClient instance = null;
    private JProgressBar progressBar;
    private AtomicReference<JFileChooser> fileChooser;

    private SwingClient(){ }

    public static SwingClient getInstance(){
        if(null == instance){
            synchronized (SwingClient.class){
                if(null == instance){
                    instance =  new SwingClient();
                }
            }
        }
        return instance;
    }

    /***
     * &History: 2020-08-13 createBy xiongchenghui
     * @param
     * @return void
     * &Desc: 初始化窗体
     */
    public void initUI(){
        /** 主窗体部分 **/
        this.setTitle("XMind转换Excel工具");
        this.setResizable(false);
        // 设置为关闭窗口同事关闭项目
        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        this.setBounds(100,100,700,540);

        /** JPanel面板设置 **/
        JPanel panel = new JPanel();
        panel.setLayout(null);
        this.setContentPane(panel);

        /** 【已选文件】Label设置 **/
        JLabel fileFieldTitle = new JLabel("已选文件：");
        fileFieldTitle.setBounds(80, 150, 70, 30);
        /** 【已选文件】路径设置 **/
        JTextArea jTextArea =  new JTextArea("");
        jTextArea.setBounds(150,150,430,90);
        jTextArea.setLineWrap(true);
        jTextArea.setEditable(false);
        Border border = BorderFactory.createLineBorder(Color.WHITE);
        jTextArea.setBorder(BorderFactory.createCompoundBorder(border,
                BorderFactory.createEmptyBorder(0, 0, 0, 0)));
        JScrollPane scrollPane = new JScrollPane(jTextArea);
        scrollPane.setBounds(150,150,430,90);
        scrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
        scrollPane.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);

        fileChooser = new AtomicReference<>(new JFileChooser());
        JLabel resultFieldText = new JLabel("未开始");
        resultFieldText.setBounds(80,270,500,30);
        //resultFieldText.setBounds(300,270,60,30);

        /** 【选择文件】按钮设置 **/
        JButton openBtn = new JButton("选择文件");
        openBtn.addActionListener(e -> fileChooser.set(this.showFileOpenDialog(this, jTextArea)));
        openBtn.setBounds(160,100,100,30);
        // 绿色背景
        openBtn.setUI(new BEButtonUI().setNormalColor(BEButtonUI.NormalColor.green));
        openBtn.setFont(new Font("宋体", Font.BOLD,15));
        // 字体颜色
        openBtn.setForeground(Color.white);

        /** 【执行转换】按钮设置 **/
        JButton action = new JButton("执行转换");
        action.setBounds(370,100,100,30);
        action.addActionListener(e -> this.action(fileChooser.get(), resultFieldText));
        // 绿色背景
        action.setUI(new BEButtonUI().setNormalColor(BEButtonUI.NormalColor.green));
        action.setFont(new Font("宋体", Font.BOLD,15));
        // 字体颜色
        action.setForeground(Color.white);

        /** 初始化【进度条】 **/
        progressBar = new JProgressBar();
        progressBar.setBounds(80,300,500,30);
        progressBar.setValue(0);
        progressBar.setStringPainted(true);

        /** 所有控件加入panel **/
        panel.add(openBtn);
        panel.add(action);
        panel.add(fileFieldTitle);
        panel.add(scrollPane);
        panel.add(progressBar);
        panel.add(resultFieldText);

        this.setVisible(true);

    }

    /***
     * &History: 2020-08-13 createBy xiongchenghui
     * @param progressBar
     * @return void
     * &Desc: 进度条模拟
     */
    private void doProgressBar(JProgressBar progressBar) {
        new Thread(() -> {
            for (int i = 0; i <= 100; i++) {
                try {
                    Thread.sleep(100);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }
                progressBar.setValue(i);
            }
        }).start();
    }

    /***
     * &History: 2020-08-13 createBy xiongchenghui
     * @param fileChooser
     * @return void
     * &Desc: 执行转换
     */
    private void action(JFileChooser fileChooser, JLabel resultField) {
        if(progressBar.getValue() > 0 && progressBar.getValue() < 100){
            JOptionPane.showMessageDialog(null,
                    "转换过程中，请等待......","警告！",JOptionPane.WARNING_MESSAGE);
            return;
        }
        if (null == fileChooser || null == fileChooser.getSelectedFile()) {
            JOptionPane.showMessageDialog(null,
                    "请先选择要处理的文件！！！", "警告！",JOptionPane.WARNING_MESSAGE);
            return;
        }
        this.doProgressBar(progressBar);

        String xMindFile = fileChooser.getSelectedFile().getAbsolutePath();
        String excelFilePath = xMindFile.replaceAll(FileExtension.XMIND, FileExtension.XLSX);

        logger.info("执行转换：" + fileChooser.getSelectedFile().getAbsolutePath());
        XMind2Excel.setXMindFile(new File(xMindFile));
        XMind2Excel.setExcelFilePath(excelFilePath);
        resultField.setText("转换中...");
        XMind2Excel.xMind2Excel();
        resultField.setText("完成转换; 用例总数：" + XMind2Excel.testCaseInfo.getTestCaseNo() +
                ",测试步骤数：" + XMind2Excel.testCaseInfo.getTestCaseSteps() +
                ",测试验证点：" + XMind2Excel.testCaseInfo.getTestCaseCheckPointers()
        );

    }

    /***
     * &History: 2020-08-13 createBy xiongchenghui
     * @param parent
     * @param textField
     * @return javax.swing.JFileChooser
     * &Desc: 文件选择弹窗
     */
    private JFileChooser showFileOpenDialog(Component parent, JTextArea textField){
        if(progressBar.getValue() != 0 && progressBar.getValue() != 100){
            JOptionPane.showMessageDialog(null,
                    "转换过程中，请等待......","警告！",JOptionPane.WARNING_MESSAGE);
            return null;
        }
        JFileChooser jFileChooser = new JFileChooser();
        // 只能选择文件
        jFileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
        // 只能选择一个文件
        jFileChooser.setMultiSelectionEnabled(false);
        // 设置默认使用的文件过滤器
        jFileChooser.setFileFilter(new FileNameExtensionFilter(
                "excel(*.xmind, *.xmt, *.xmap)", "xmind", "xmt", "xmap"));
        // 打开文件选择框（线程将被阻塞, 直到选择框被关闭）
        int result = jFileChooser.showOpenDialog(parent);

        if(result == JFileChooser.APPROVE_OPTION){
            // 如果点击了"确定", 则获取选择的文件路径
            File file = jFileChooser.getSelectedFile();
            // 如果允许选择多个文件, 则通过下面方法获取选择的所有文件
            // File[] files = jFileChooser.getSelectedFiles();
            textField.setText("");
            textField.setText(file.getAbsolutePath());
        }

        //进度条归零0
        progressBar.setValue(0);
        return jFileChooser;
    }

}